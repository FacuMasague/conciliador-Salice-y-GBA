from __future__ import annotations

import copy
import json
import os
import re
import socket
import ssl
import time
from dataclasses import dataclass
from datetime import date, timedelta
from typing import Any, Dict, List, Tuple
from urllib.error import HTTPError, URLError
from urllib.parse import urlencode
from urllib.request import Request, urlopen

from ..env_loader import load_project_env
from .errors import ExternalConfigError, ExternalProviderError, ExternalTimeoutError

load_project_env()


@dataclass(frozen=True)
class ReceiptsApiResponse:
    payload: Dict[str, Any]
    request_id: str | None
    warnings: list[str]


def _norm_key(s: object) -> str:
    return "".join(ch for ch in str(s or "").lower() if ch.isalnum())


def _dict_get_any(obj: dict, *keys: str) -> Any:
    norm_map = {_norm_key(k): v for k, v in obj.items()}
    for k in keys:
        nk = _norm_key(k)
        if nk in norm_map:
            return norm_map[nk]
    return None


def _parse_empresa_ids(raw: str) -> list[str]:
    out: list[str] = []
    for p in str(raw or "").split(","):
        s = p.strip()
        if s:
            out.append(s)
    return out


def _resolve_empresa_targets(empresa_filter: str | None) -> list[str]:
    defaults = ["2"]
    extra = _parse_empresa_ids(os.getenv("RECEIPTS_API_EMPRESA_IDS", ""))
    for e in extra:
        if e not in defaults:
            defaults.append(e)

    if not empresa_filter:
        return defaults

    f = str(empresa_filter).strip().upper()
    if f.isdigit():
        return [f]
    if f == "GBA":
        return ["2"]
    if f == "SALICE":
        return ["3"]
    if f == "ALARCON":
        return ["6"]
    return defaults


def _empresa_name_from_id(empresa_id: str) -> str:
    eid = str(empresa_id or "").strip()
    if eid == "2":
        return "GBA"
    if eid == "3":
        return "SALICE"
    if eid == "6":
        return "ALARCON"
    return ""


def _parse_paths_csv(raw: str) -> list[str]:
    out: list[str] = []
    for p in str(raw or "").split(","):
        s = p.strip()
        if not s:
            continue
        if not s.startswith("/"):
            s = "/" + s
        if s not in out:
            out.append(s)
    return out


def _resolve_comprobantes_paths() -> list[str]:
    custom = _parse_paths_csv(os.getenv("RECEIPTS_API_COMPROBANTES_PATHS", ""))
    if custom:
        return custom

    default_path = (os.getenv("RECEIPTS_API_COMPROBANTES_PATH", "/api/Ventas/Comprobantes/GetList") or "").strip()
    out = _parse_paths_csv(default_path)
    fallbacks = [
        "/api/Ventas/Comprobantes/GetList",
        "/api/Ventas/Comprobantes/GetListComprobantes",
        "/api/Ventas/Comprobantes/List",
        "/api/Maestros/Comprobantes/GetList",
    ]
    for f in fallbacks:
        if f not in out:
            out.append(f)
    return out


def _resolve_cobros_path() -> str:
    p = (
        os.getenv("RECEIPTS_API_COBROS_PATH", "/api/Ventas/Comprobantes/Cobros/GetList")
        or "/api/Ventas/Comprobantes/Cobros/GetList"
    ).strip()
    if not p.startswith("/"):
        p = "/" + p
    return p


def _prefer_cobros_only(targets: list[str], empresa_filter: str | None) -> bool:
    raw = os.getenv("RECEIPTS_API_FORCE_COBROS_ONLY")
    if raw is not None and raw.strip():
        return _env_bool("RECEIPTS_API_FORCE_COBROS_ONLY", False)

    norm_targets = [str(t or "").strip() for t in targets if str(t or "").strip()]
    empresa = str(empresa_filter or "").strip().upper()
    return norm_targets == ["2"] or empresa == "GBA"


def _getitem_max_keys(targets: list[str], empresa_filter: str | None) -> int:
    raw = os.getenv("RECEIPTS_API_GETITEM_MAX_KEYS")
    if raw is not None and raw.strip():
        try:
            return int(raw.strip())
        except Exception:
            raise ExternalConfigError("RECEIPTS_API_GETITEM_MAX_KEYS inválido")

    if _prefer_cobros_only(targets, empresa_filter):
        # En GBA trabajamos por ventanas diarias y GetItem ya se trocea internamente
        # en chunks chicos. Por default conviene permitir el detalle completo de la
        # ventana para evitar una segunda pasada global sobre todos los recibos.
        return 2000
    return 300


def _page_size_for_targets(targets: list[str], empresa_filter: str | None) -> int:
    raw = os.getenv("RECEIPTS_API_PAGE_SIZE")
    if raw is not None and raw.strip():
        try:
            v = int(raw.strip())
        except Exception:
            raise ExternalConfigError("RECEIPTS_API_PAGE_SIZE inválido")
        if v <= 0:
            raise ExternalConfigError("RECEIPTS_API_PAGE_SIZE inválido")
        return v
    # GBA/Cobros/GetList no responde estable con páginas grandes.
    # En pruebas reales, 100 funciona de manera consistente donde 200/300/2000 se cuelgan.
    return 100 if _prefer_cobros_only(targets, empresa_filter) else 500


def _window_days_for_targets(targets: list[str], empresa_filter: str | None) -> int:
    raw = os.getenv("RECEIPTS_API_WINDOW_DAYS")
    if raw is not None and raw.strip():
        try:
            v = int(raw.strip())
        except Exception:
            raise ExternalConfigError("RECEIPTS_API_WINDOW_DAYS inválido")
        if v < 0:
            raise ExternalConfigError("RECEIPTS_API_WINDOW_DAYS inválido")
        return v
    # Para GBA conviene trocear por día; evita que una página pesada de Cobros/GetList
    # bloquee toda la ventana.
    return 1 if _prefer_cobros_only(targets, empresa_filter) else 0


def _split_date_windows(start: date, end: date, window_days: int) -> list[tuple[date, date]]:
    if window_days <= 0 or start >= end:
        return [(start, end)]
    windows: list[tuple[date, date]] = []
    cur = start
    step = max(int(window_days), 1)
    while cur <= end:
        win_end = min(cur + timedelta(days=step - 1), end)
        windows.append((cur, win_end))
        cur = win_end + timedelta(days=1)
    return windows


def _looks_operational_comprobante(row: dict) -> bool:
    has_cliente = bool(str(_dict_get_any(row, "clienteID", "cliente_id", "clienteId", "ClienteID") or "").strip())
    has_importe = bool(str(_dict_get_any(row, "importeTotal", "importe_total", "ImporteTotal") or "").strip())
    has_fecha = bool(
        str(
            _dict_get_any(
                row,
                "fechaDeEmision",
                "fecha_de_emision",
                "FechaDeEmision",
                "fechaDePrimerVencimiento",
                "fecha_de_primer_vencimiento",
                "FechaDePrimerVencimiento",
            )
            or ""
        ).strip()
    )
    return has_cliente and has_importe and has_fecha


def _normalize_id_text(value: Any) -> str:
    s = str(value or "").strip()
    if not s:
        return ""
    s_num = s.replace(",", ".")
    # Soporta IDs numéricos devueltos como float textual (ej: \"6.0\")
    if re.fullmatch(r"[-+]?[0-9]+(?:\.[0-9]+)?", s_num):
        try:
            return str(int(float(s_num)))
        except Exception:
            pass
    digits = "".join(ch for ch in s if ch.isdigit())
    if digits:
        try:
            return str(int(digits))
        except Exception:
            return digits.lstrip("0") or "0"
    return s


def _row_empresa_id(row: dict, fallback_target: str) -> tuple[str, str]:
    candidates = [
        ("sucursal_id", _dict_get_any(row, "sucursal_id", "sucursalID", "sucursalId", "SucursalID")),
        ("empresaID", _dict_get_any(row, "empresaID", "empresa_id", "empresaId", "EmpresaID")),
    ]
    for source, value in candidates:
        s = _normalize_id_text(value)
        if s:
            return s, source
    return str(fallback_target or "").strip(), "target"


def _row_comprobante_key_for_getitem(row: dict) -> dict[str, Any] | None:
    comp_id = _dict_get_any(row, "ComprobanteID", "comprobanteID", "comprobante_id")
    emp_id = _dict_get_any(
        row,
        "_empresa_id_api_original",
        "EmpresaID",
        "empresaID",
        "empresa_id",
        "sucursalID",
        "sucursal_id",
    )
    serie = _dict_get_any(row, "Serie", "serie")
    pv_id = _dict_get_any(row, "PuntoDeVentaID", "puntoDeVentaID", "punto_de_venta_id")
    numero = _dict_get_any(row, "Numero", "numero")
    if comp_id is None or emp_id is None or pv_id is None or numero is None:
        return None
    comp_n = _normalize_id_text(comp_id)
    emp_n = _normalize_id_text(emp_id)
    pv_n = _normalize_id_text(pv_id)
    num_n = _normalize_id_text(numero)
    if not comp_n or not emp_n or not pv_n or not num_n:
        return None
    try:
        return {
            "ComprobanteID": int(comp_n),
            "EmpresaID": int(emp_n),
            "Serie": str(serie or ""),
            "PuntoDeVentaID": int(pv_n),
            "Numero": int(num_n),
        }
    except Exception:
        return None


def _row_comprobante_match_key(row: dict) -> tuple[str, str, str, str, str]:
    return (
        _normalize_id_text(_dict_get_any(row, "ComprobanteID", "comprobanteID", "comprobante_id")),
        _normalize_id_text(
            _dict_get_any(
                row,
                "_empresa_id_api_original",
                "EmpresaID",
                "empresaID",
                "empresa_id",
                "sucursalID",
                "sucursal_id",
            )
        ),
        str(_dict_get_any(row, "Serie", "serie") or ""),
        _normalize_id_text(_dict_get_any(row, "PuntoDeVentaID", "puntoDeVentaID", "punto_de_venta_id")),
        _normalize_id_text(_dict_get_any(row, "Numero", "numero")),
    )


def _normalize_getitem_key_tuple(k: dict[str, Any]) -> tuple[str, str, str, str, str]:
    return (
        _normalize_id_text(k.get("ComprobanteID")),
        _normalize_id_text(k.get("EmpresaID")),
        str(k.get("Serie") or ""),
        _normalize_id_text(k.get("PuntoDeVentaID")),
        _normalize_id_text(k.get("Numero")),
    )


def _relaxed_key_from_exact_tuple(k: tuple[str, str, str, str, str]) -> tuple[str, str, str, str]:
    return (k[0], k[2], k[3], k[4])


def _row_comprobante_relaxed_key(row: dict) -> tuple[str, str, str, str]:
    k = _row_comprobante_match_key(row)
    return _relaxed_key_from_exact_tuple(k)


def _normalize_recibo_pm(value: Any) -> str:
    s = str(value or "").strip()
    if not s:
        return ""
    digits = "".join(ch for ch in s if ch.isdigit())
    if not digits:
        return ""
    try:
        return str(int(digits))
    except Exception:
        return digits.lstrip("0") or "0"


def _row_has_usable_pm_receipt(row: dict) -> bool:
    numero = _normalize_id_text(_dict_get_any(row, "Numero", "numero", "nro_recibo", "nroRecibo", "NroRecibo"))
    if numero:
        return True
    nro_pm = _normalize_recibo_pm(_dict_get_any(row, "nro_recibo_pm", "nroReciboPm", "NroReciboPm"))
    if nro_pm:
        return True
    codigo = str(
        _dict_get_any(
            row,
            "codigoDeImportacion",
            "codigo_de_importacion",
            "CodigoDeImportacion",
        )
        or ""
    ).strip()
    if not codigo:
        return False
    nro_from_codigo = _normalize_recibo_pm(codigo)
    if not nro_from_codigo:
        return False
    norm = "".join(ch.lower() for ch in codigo if ch.isalnum())
    if "pm" in norm:
        return True
    return codigo.isdigit()


def _row_has_usable_cliente(row: dict) -> bool:
    cli = str(_dict_get_any(row, "clienteID", "cliente_id", "clienteId", "ClienteID") or "").strip()
    if not cli:
        return False
    digits = "".join(ch for ch in cli if ch.isdigit())
    return bool(digits)


def _chunked(items: list[dict], size: int) -> list[list[dict]]:
    if size <= 0:
        size = 100
    return [items[i : i + size] for i in range(0, len(items), size)]


def _get_path(obj: dict, path: tuple[str, ...]) -> Any:
    cur: Any = obj
    for k in path:
        if not isinstance(cur, dict):
            return None
        cur = cur.get(k)
    return cur


def _extract_rows(payload: dict, list_key: str) -> list[dict] | None:
    candidates = [
        (list_key,),
        ("data", list_key),
        ("resultado", list_key),
        ("result", list_key),
        ("datos", list_key),
        ("items",),
        ("data", "items"),
        ("rows",),
        ("data", "rows"),
    ]
    for path in candidates:
        v = _get_path(payload, path)
        if isinstance(v, list):
            return [r for r in v if isinstance(r, dict)]
    return None


def _extract_obj(payload: dict, key: str) -> dict | None:
    candidates = [
        (key,),
        ("data", key),
        ("resultado", key),
        ("result", key),
        ("datos", key),
    ]
    for path in candidates:
        v = _get_path(payload, path)
        if isinstance(v, dict):
            return v
    return None


def _extract_paginacion(payload: dict) -> dict | None:
    candidates = [
        ("paginacion",),
        ("pagination",),
        ("data", "paginacion"),
        ("data", "pagination"),
        ("resultado", "paginacion"),
        ("result", "pagination"),
        ("meta", "paginacion"),
        ("meta", "pagination"),
    ]
    for path in candidates:
        v = _get_path(payload, path)
        if isinstance(v, dict):
            return v
    return None


def _env_bool(name: str, default: bool) -> bool:
    raw = os.getenv(name)
    if raw is None:
        return default
    return raw.strip().lower() in {"1", "true", "yes", "on"}


def _timeout_seconds() -> float:
    raw = (
        os.getenv("RECEIPTS_API_TIMEOUT_SECONDS")
        or os.getenv("HTTP_TIMEOUT_SECONDS")
        or "60"
    ).strip()
    try:
        v = float(raw)
    except Exception:
        raise ExternalConfigError("RECEIPTS_API_TIMEOUT_SECONDS/HTTP_TIMEOUT_SECONDS inválido")
    if v <= 0:
        raise ExternalConfigError("RECEIPTS_API_TIMEOUT_SECONDS/HTTP_TIMEOUT_SECONDS inválido")
    return v


def _page_size_candidates(initial_size: int) -> list[int]:
    raw = (os.getenv("RECEIPTS_API_PAGE_SIZE_FALLBACKS", "") or "").strip()
    fallbacks: list[int] = []
    if raw:
        for part in raw.split(","):
            part = part.strip()
            if not part:
                continue
            try:
                value = int(part)
            except Exception:
                raise ExternalConfigError("RECEIPTS_API_PAGE_SIZE_FALLBACKS inválido")
            if value > 0:
                fallbacks.append(value)
    if not fallbacks:
        fallbacks = [1000, 500, 300, 200, 100, 50]

    out: list[int] = []
    seen: set[int] = set()
    for value in [int(initial_size)] + fallbacks:
        if value <= 0 or value in seen:
            continue
        seen.add(value)
        out.append(value)
    return out


def _fetch_paged_rows_with_fallbacks(
    *,
    base: str,
    path: str,
    method: str,
    headers: dict[str, str],
    list_key: str,
    page_size: int,
    body: dict[str, Any] | None = None,
) -> tuple[list[dict], str | None, bool, list[str]]:
    warnings: list[str] = []
    candidates = _page_size_candidates(int(page_size))
    last_timeout: ExternalTimeoutError | None = None

    for idx, size in enumerate(candidates):
        try:
            rows, rid, ok = _fetch_paged_rows(
                base=base,
                path=path,
                method=method,
                headers=headers,
                list_key=list_key,
                page_size=size,
                body=body,
            )
            if idx > 0:
                warnings.append(f"{path} reintentado con pageSize={size} después de timeout.")
            return rows, rid, ok, warnings
        except ExternalTimeoutError as e:
            last_timeout = e
            warnings.append(f"{path} timeout con pageSize={size}; reintentando con pageSize menor.")
            continue

    if last_timeout is not None:
        raise last_timeout
    return [], None, True, warnings


def _ssl_context() -> ssl.SSLContext:
    verify = _env_bool("API_VERIFY_SSL", True)
    if verify:
        return ssl.create_default_context()
    return ssl._create_unverified_context()


def _allow_insecure_fallback() -> bool:
    return _env_bool("API_SSL_FALLBACK_UNVERIFIED", True)


def _base_url(prefix: str) -> str:
    base = (os.getenv(f"{prefix}_BASE_URL", "") or "").strip().rstrip("/")
    if not base:
        base = "https://m5gba.grupoesi.com.ar"
    return base


def _headers_base(prefix: str) -> dict[str, str]:
    base = _base_url(prefix)
    ua = (
        os.getenv(
            f"{prefix}_USER_AGENT",
            "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) "
            "AppleWebKit/537.36 (KHTML, like Gecko) "
            "Chrome/126.0.0.0 Safari/537.36",
        )
        or ""
    ).strip()
    headers = {
        "Accept": "application/json",
        "Content-Type": "application/json",
        "GrupoESI": (os.getenv(f"{prefix}_GRUPO_ESI_HEADER_VALUE", "true") or "true").strip(),
        "User-Agent": ua,
        "Accept-Language": "es-AR,es;q=0.9,en;q=0.8",
        "Referer": base + "/swagger/index.html",
        "Origin": base,
    }
    empresa_id = (os.getenv(f"{prefix}_EMPRESA_ID", "") or "").strip()
    sucursal_id = (os.getenv(f"{prefix}_SUCURSAL_ID", "") or "").strip()
    if empresa_id:
        headers["empresaID"] = empresa_id
    if sucursal_id:
        headers["sucursalID"] = sucursal_id
    return headers


def _http_json(
    url: str,
    *,
    method: str,
    headers: dict[str, str],
    body: Any | None = None,
) -> tuple[dict[str, Any], str | None]:
    data = None
    if body is not None:
        data = json.dumps(body).encode("utf-8")
    req = Request(url=url, method=method.upper(), headers=headers, data=data)

    try:
        try:
            with urlopen(req, timeout=_timeout_seconds(), context=_ssl_context()) as resp:
                raw = resp.read().decode("utf-8")
                payload = json.loads(raw)
                request_id = resp.headers.get("X-Request-Id") or resp.headers.get("x-request-id")
                if not isinstance(payload, dict):
                    raise ExternalProviderError("receipts", "Respuesta JSON inválida (se esperaba objeto)")
                return payload, request_id
        except URLError as e:
            reason = getattr(e, "reason", None)
            if _allow_insecure_fallback() and isinstance(reason, ssl.SSLCertVerificationError):
                with urlopen(req, timeout=_timeout_seconds(), context=ssl._create_unverified_context()) as resp:
                    raw = resp.read().decode("utf-8")
                    payload = json.loads(raw)
                    request_id = resp.headers.get("X-Request-Id") or resp.headers.get("x-request-id")
                    if not isinstance(payload, dict):
                        raise ExternalProviderError("receipts", "Respuesta JSON inválida (se esperaba objeto)")
                    return payload, request_id
            raise
    except HTTPError as e:
        body_txt = ""
        try:
            body_txt = e.read().decode("utf-8")
        except Exception:
            body_txt = str(e)
        lower_body = body_txt.lower()
        if "cloudflare" in lower_body and ("error 1010" in lower_body or "access denied" in lower_body):
            raise ExternalProviderError(
                "receipts",
                "Cloudflare bloqueó el acceso (Error 1010). "
                "Hay que pedir whitelist de IP/UA del servidor cliente en m5mdp.grupoesi.com.ar.",
                status_code=e.code,
            )
        raise ExternalProviderError("receipts", f"Receipts API HTTP {e.code}: {body_txt}", status_code=e.code)
    except URLError as e:
        if isinstance(e.reason, socket.timeout):
            raise ExternalTimeoutError("receipts", "Timeout en Receipts API")
        raise ExternalProviderError("receipts", f"Receipts API error de red: {e}")
    except socket.timeout:
        raise ExternalTimeoutError("receipts", "Timeout en Receipts API")
    except TimeoutError:
        raise ExternalTimeoutError("receipts", "Timeout en Receipts API")
    except json.JSONDecodeError:
        raise ExternalProviderError("receipts", "Receipts API respondió JSON inválido")


def _login_token(prefix: str, base: str, headers_base: dict[str, str]) -> str:
    auth_mode = (os.getenv(f"{prefix}_AUTH_MODE", "gesi_login") or "gesi_login").strip().lower()
    if auth_mode == "bearer":
        token = (os.getenv(f"{prefix}_TOKEN", "") or "").strip()
        if not token:
            raise ExternalConfigError(f"Falta {prefix}_TOKEN para auth bearer")
        return token
    if auth_mode != "gesi_login":
        raise ExternalConfigError(f"{prefix}_AUTH_MODE inválido: {auth_mode}")

    username = (os.getenv(f"{prefix}_USERNAME", "") or os.getenv("GESI_API_USERNAME", "")).strip()
    password = (os.getenv(f"{prefix}_PASSWORD", "") or os.getenv("GESI_API_PASSWORD", "")).strip()
    if not username or not password:
        raise ExternalConfigError(
            f"Faltan {prefix}_USERNAME/{prefix}_PASSWORD (o GESI_API_USERNAME/GESI_API_PASSWORD)"
        )

    login_path = (os.getenv(f"{prefix}_LOGIN_PATH", "/api/Autenticacion/login") or "/api/Autenticacion/login").strip()
    if not login_path.startswith("/"):
        login_path = "/" + login_path

    payload, _rid = _http_json(
        f"{base}{login_path}",
        method="POST",
        headers=headers_base,
        body={"username": username, "password": password},
    )
    token = str(payload.get("token") or "").strip()
    success = payload.get("success")
    if success is False or not token:
        msg = "Login inválido"
        err = payload.get("error")
        if isinstance(err, dict):
            msg = str(err.get("message") or msg)
        raise ExternalProviderError("receipts", f"No se pudo autenticar en GESI: {msg}")
    return token


def _build_auth_headers_for_empresa(
    *,
    base: str,
    headers_root: dict[str, str],
    empresa_id: str,
    drop_sucursal: bool,
) -> dict[str, str]:
    hdr = dict(headers_root)
    hdr["empresaID"] = str(empresa_id)
    if drop_sucursal:
        hdr.pop("sucursalID", None)
    token = _login_token("RECEIPTS_API", base, hdr)
    hdr["Authorization"] = f"Bearer {token}"
    return hdr


def _status_code_from_exc(exc: Exception) -> int:
    try:
        return int(getattr(exc, "status_code", 0) or 0)
    except Exception:
        return 0


def _is_rate_limited_error(exc: Exception) -> bool:
    if _status_code_from_exc(exc) == 429:
        return True
    txt = str(exc).lower()
    return "cantidad maxima" in txt or "rate" in txt or "too many requests" in txt


def _build_query_for_path(path: str, page: int, size: int) -> dict[str, str]:
    q = {"pageNumber": str(page), "pageSize": str(size)}
    if path.lower() == "/api/maestros/comprobantes/getlist":
        q["comprobanteID"] = "0"
        q["claseDeComprobanteID"] = "0"
    return q


def _fetch_paged_rows(
    *,
    base: str,
    path: str,
    method: str,
    headers: dict[str, str],
    list_key: str,
    page_size: int,
    body: dict[str, Any] | None = None,
) -> tuple[list[dict], str | None, bool]:
    rows: list[dict] = []
    request_id_last: str | None = None
    page = 1

    while True:
        q = _build_query_for_path(path, page, page_size)
        url = f"{base}{path}?{urlencode(q)}"
        payload, rid = _http_json(
            url,
            method=method,
            headers=headers,
            body=(copy.deepcopy(body) if method.upper() == "POST" else None),
        )
        if rid:
            request_id_last = rid

        success = payload.get("success")
        if success is False:
            err = payload.get("error")
            msg = f"GetList devolvió error en {path}"
            if isinstance(err, dict):
                msg = str(err.get("message") or msg)
            raise ExternalProviderError("receipts", msg)

        page_rows = _extract_rows(payload, list_key) or []
        rows.extend(page_rows)

        pag = _extract_paginacion(payload)
        if not isinstance(pag, dict):
            break

        try:
            total_pages = int(
                pag.get("totalPaginas")
                or pag.get("totalpages")
                or pag.get("totalPages")
                or pag.get("pages")
                or 1
            )
        except Exception:
            break

        if page >= total_pages:
            break

        page += 1
        if page > 500:
            break

    return rows, request_id_last, True


def _dedupe_comprobantes(rows: List[dict]) -> tuple[List[dict], int]:
    out: List[dict] = []
    seen: set[tuple[str, str, str, str, str]] = set()
    dups = 0
    for r in rows:
        if not isinstance(r, dict):
            continue
        k = _row_comprobante_match_key(r)
        if not any(k):
            continue
        if k in seen:
            dups += 1
            continue
        seen.add(k)
        out.append(r)
    return out, dups


def _has_useful_valor_ids(value: Any) -> bool:
    if not isinstance(value, list):
        return False
    for it in value:
        if not isinstance(it, dict):
            continue
        vid = _normalize_id_text(_dict_get_any(it, "valorID", "valor_id", "valorId", "ValorID"))
        if vid and vid != "0":
            return True
    return False


def _merge_getitem_into_comprobante(base_row: dict, extra_row: dict) -> None:
    base_norm_map = {_norm_key(k): k for k in base_row.keys()}
    for k, v in extra_row.items():
        nk = _norm_key(k)
        existing_key = base_norm_map.get(nk)
        if existing_key is None:
            base_row[k] = v
            base_norm_map[nk] = k
            continue

        if nk == _norm_key("detalleDeValores"):
            base_has = _has_useful_valor_ids(base_row.get(existing_key))
            extra_has = _has_useful_valor_ids(v)
            if extra_has and not base_has:
                base_row[existing_key] = v
        elif nk in {
            _norm_key("formaDePagoID"),
            _norm_key("DescripcionFormaDePago"),
            _norm_key("formaDePago"),
            _norm_key("medioPago"),
        }:
            if v not in (None, "", [], {}):
                base_row[existing_key] = v
        else:
            cur = base_row.get(existing_key)
            if cur in (None, "", [], {}):
                base_row[existing_key] = v

        if k not in base_row:
            base_row[k] = v


def _fetch_getitem_details(
    *,
    base: str,
    getitem_path: str,
    getitem_query: str,
    headers: dict[str, str],
    keys: list[dict[str, Any]],
    chunk_size: int,
) -> tuple[dict[tuple[str, str, str, str, str], dict], str | None, bool]:
    details: dict[tuple[str, str, str, str, str], dict] = {}
    request_id_last: str | None = None
    rate_limited = False
    retry_429 = int((os.getenv("RECEIPTS_API_GETITEM_429_RETRIES", "0") or "0").strip())
    wait_429 = int((os.getenv("RECEIPTS_API_GETITEM_429_WAIT_SECONDS", "5") or "5").strip())

    for batch in _chunked(keys, chunk_size):
        url = f"{base}{getitem_path}"
        if getitem_query:
            url = f"{url}?{getitem_query}"

        attempt = 0
        while True:
            try:
                payload_item, rid = _http_json(
                    url,
                    method="POST",
                    headers=headers,
                    body=batch,
                )
                if rid:
                    request_id_last = rid
                break
            except ExternalProviderError as e:
                if _is_rate_limited_error(e):
                    if attempt < max(retry_429, 0):
                        attempt += 1
                        time.sleep(max(wait_429, 1))
                        continue
                    rate_limited = True
                    payload_item = {}
                    break
                raise

        if rate_limited and not payload_item:
            break

        if isinstance(payload_item, dict) and payload_item.get("success") is False:
            err = payload_item.get("error")
            msg = ""
            if isinstance(err, dict):
                msg = str(err.get("message") or "").strip()
            else:
                msg = str(payload_item.get("message") or "").strip()
            lower_msg = msg.lower()
            if "cantidad maxima" in lower_msg or "maxima de solicitudes" in lower_msg or "too many requests" in lower_msg:
                rate_limited = True
                break

        rows_item = _extract_rows(payload_item, "comprobantes")
        if rows_item is None and isinstance(payload_item, list):
            rows_item = [x for x in payload_item if isinstance(x, dict)]
        if rows_item is None:
            rows_item = []

        for rr in rows_item:
            mk = _row_comprobante_match_key(rr)
            if any(mk):
                details[mk] = rr

    return details, request_id_last, rate_limited


def _fetch_medios_pago(
    *,
    base: str,
    headers_for_target,
    targets: list[str],
    medios_path: str,
    page_size: int,
) -> tuple[list[dict], str | None, list[str]]:
    warnings: list[str] = []
    request_id_last: str | None = None
    rows: list[dict] = []
    seen: set[tuple[str, str]] = set()

    for target in targets or ["3"]:
        headers = headers_for_target(str(target))
        try:
            page_rows, rid, _ok, retry_warnings = _fetch_paged_rows_with_fallbacks(
                base=base,
                path=medios_path,
                method="GET",
                headers=headers,
                list_key="mediosDePago",
                page_size=page_size,
                body=None,
            )
            warnings.extend(retry_warnings)
            if rid:
                request_id_last = rid
            for r in page_rows:
                emp = _normalize_id_text(_dict_get_any(r, "empresaID", "empresa_id", "empresaId", "EmpresaID"))
                vid = _normalize_id_text(_dict_get_any(r, "valorID", "valor_id", "valorId", "ValorID"))
                k = (emp, vid)
                if k in seen:
                    continue
                seen.add(k)
                rows.append(r)
        except ExternalProviderError as e:
            if _is_rate_limited_error(e):
                warnings.append("No se pudo leer maestro de medios de pago por límite de rate (HTTP 429).")
                break
            warnings.append(f"No se pudo leer maestro de medios de pago: {e}")
        except Exception as e:
            warnings.append(f"No se pudo leer maestro de medios de pago: {e}")

    return rows, request_id_last, warnings


def _get_paged_get(
    *,
    base: str,
    path: str,
    headers: dict[str, str],
    list_key: str,
    page_size: int,
) -> tuple[list[dict], str | None, list[str]]:
    rows, request_id_last, _ok, warnings = _fetch_paged_rows_with_fallbacks(
        base=base,
        path=path,
        method="GET",
        headers=headers,
        list_key=list_key,
        page_size=page_size,
        body=None,
    )
    return rows, request_id_last, warnings


def _fetch_receipts_payload_single_window(
    *,
    days: int,
    empresa_filter: str | None = None,
    start_date: date | None = None,
    end_date: date | None = None,
) -> ReceiptsApiResponse:
    if not _env_bool("API_MODE_ENABLED", True):
        raise ExternalConfigError("API_MODE_ENABLED=false: modo API deshabilitado")

    if int(days) <= 0:
        raise ExternalConfigError("api_receipts_days debe ser > 0")

    base = _base_url("RECEIPTS_API")
    headers_root = _headers_base("RECEIPTS_API")

    comprobantes_paths = _resolve_comprobantes_paths()
    cobros_path = _resolve_cobros_path()

    formas_path = (os.getenv("RECEIPTS_API_FORMAS_PAGO_PATH", "/api/Maestros/FormasDePago/GetList") or "/api/Maestros/FormasDePago/GetList").strip()
    if not formas_path.startswith("/"):
        formas_path = "/" + formas_path

    empresas_path = (os.getenv("RECEIPTS_API_EMPRESAS_PATH", "/api/Maestros/Empresas/GetList") or "/api/Maestros/Empresas/GetList").strip()
    if not empresas_path.startswith("/"):
        empresas_path = "/" + empresas_path

    medios_path = (os.getenv("RECEIPTS_API_MEDIOS_PAGO_PATH", "/api/Maestros/MediosDePago/GetList") or "/api/Maestros/MediosDePago/GetList").strip()
    if not medios_path.startswith("/"):
        medios_path = "/" + medios_path

    getitem_path = (os.getenv("RECEIPTS_API_COBROS_GETITEM_PATH", "/api/Ventas/Comprobantes/Cobros/GetItem") or "/api/Ventas/Comprobantes/Cobros/GetItem").strip()
    if not getitem_path.startswith("/"):
        getitem_path = "/" + getitem_path

    getitem_query = (
        os.getenv(
            "RECEIPTS_API_COBROS_GETITEM_QUERY",
            "incluirDetalleDeMedioDePago=S",
        )
        or "incluirDetalleDeMedioDePago=S"
    ).strip()

    fecha_hasta = end_date or (date.today() - timedelta(days=1))
    fecha_desde = start_date or (fecha_hasta - timedelta(days=max(int(days) - 1, 0)))
    if fecha_desde > fecha_hasta:
        raise ExternalConfigError("Rango API inválido: fecha_desde es posterior a fecha_hasta")

    body_base = {
        "datosOperacion": {
            "FechaDesde": fecha_desde.isoformat(),
            "FechaHasta": fecha_hasta.isoformat(),
        },
        "datosClientes": {},
    }

    targets = _resolve_empresa_targets(empresa_filter)
    prefer_cobros_only = _prefer_cobros_only(targets, empresa_filter)
    getitem_max_keys = _getitem_max_keys(targets, empresa_filter)
    page_size = _page_size_for_targets(targets, empresa_filter)

    warnings: list[str] = []
    request_id_last: str | None = None
    comprobantes: list[dict] = []
    counts_by_target: dict[str, int] = {}
    selected_path: str | None = None
    selected_method: str | None = None

    headers_cache: dict[str, dict[str, str]] = {}

    def _headers_for_target(target_empresa_id: str) -> dict[str, str]:
        key = str(target_empresa_id)
        if key not in headers_cache:
            hdr = _build_auth_headers_for_empresa(
                base=base,
                headers_root=headers_root,
                empresa_id=key,
                drop_sucursal=False,
            )
            headers_cache[key] = dict(hdr)
        return dict(headers_cache[key])

    # 1) Maestro de medios de pago primero (fuente principal para "Medio de pago")
    medios_pago, req_mp, w_mp = _fetch_medios_pago(
        base=base,
        headers_for_target=_headers_for_target,
        targets=targets,
        medios_path=medios_path,
        page_size=page_size,
    )
    if req_mp:
        request_id_last = req_mp
    warnings.extend(w_mp)

    # 2) Intento Comprobantes/GetList
    unsupported_paths: set[tuple[str, str]] = set()
    rate_limited = False
    if prefer_cobros_only:
        warnings.append("API GBA: se omite Comprobantes/GetList y se usa Cobros/GetList directo.")
    else:
        for target in targets:
            added = 0
            target_had_response = False

            for method in ("POST", "GET"):
                for path in comprobantes_paths:
                    if (method, path) in unsupported_paths:
                        continue

                    headers = _headers_for_target(target)
                    body = copy.deepcopy(body_base)
                    body.setdefault("datosOperacion", {})
                    body.setdefault("datosClientes", {})
                    body["datosOperacion"]["EmpresaID"] = int(str(target))
                    body["datosOperacion"]["empresaID"] = int(str(target))
                    body["datosClientes"]["EmpresaID"] = int(str(target))
                    body["empresaID"] = int(str(target))

                    try:
                        rows, rid, _ok, page_warnings = _fetch_paged_rows_with_fallbacks(
                            base=base,
                            path=path,
                            method=method,
                            headers=headers,
                            list_key="comprobantes",
                            page_size=page_size,
                            body=body,
                        )
                        warnings.extend(page_warnings)
                    except ExternalProviderError as e:
                        status = _status_code_from_exc(e)
                        if status in {404, 405}:
                            unsupported_paths.add((method, path))
                            warnings.append(
                                f"Comprobantes/GetList descartado (HTTP {status}, empresaID={target}, method={method.lower()}, path={path})."
                            )
                            continue
                        if _is_rate_limited_error(e):
                            warnings.append("Comprobantes/GetList frenado por límite de rate (HTTP 429).")
                            rate_limited = True
                            break
                        raise

                    target_had_response = True
                    if rid:
                        request_id_last = rid
                    if selected_path is None:
                        selected_path = path
                        selected_method = method.upper()

                    for r in rows:
                        if not isinstance(r, dict):
                            continue
                        rid_original = _normalize_id_text(_dict_get_any(r, "EmpresaID", "empresaID", "empresa_id"))
                        if rid_original:
                            r["_empresa_id_api_original"] = rid_original
                        resolved_id, source_field = _row_empresa_id(r, str(target))
                        r["empresaID"] = resolved_id
                        r["_empresa_id_source"] = source_field
                        if source_field.startswith("sucursal") and not str(_dict_get_any(r, "sucursalID", "sucursal_id") or "").strip():
                            r["sucursalID"] = resolved_id
                        tname = _empresa_name_from_id(resolved_id)
                        if tname:
                            r["_empresa_target_name"] = tname
                        comprobantes.append(r)
                        added += 1

                    # Primer path/método que responde define el target; evitamos sandwich de intentos.
                    break

                if rate_limited or target_had_response:
                    break

            counts_by_target[str(target)] = int(added)
            if not target_had_response:
                warnings.append(f"Comprobantes/GetList empresaID={target}: sin respuesta util en paths configurados.")
            elif added == 0:
                warnings.append(f"Comprobantes/GetList empresaID={target}: sin comprobantes en la ventana solicitada.")

            if rate_limited:
                break

    comprobantes, dups = _dedupe_comprobantes(comprobantes)
    if dups > 0:
        warnings.append(f"Comprobantes duplicados descartados: {dups}")

    operational_rows = [r for r in comprobantes if isinstance(r, dict) and _looks_operational_comprobante(r)]

    # 3) Fallback a Cobros/GetList + Cobros/GetItem si Comprobantes no trajo operativos
    if not operational_rows:
        if not prefer_cobros_only:
            warnings.append("Comprobantes API no devolvió filas operativas (cliente/fecha/importe). Fallback a Cobros/GetList.")
        comprobantes = []
        counts_by_target = {}
        rate_limited = False

        for target in targets:
            added = 0
            target_had_response = False

            for method in ("POST", "GET"):
                headers = _headers_for_target(target)
                body = copy.deepcopy(body_base)
                body.setdefault("datosOperacion", {})
                body.setdefault("datosClientes", {})
                body["datosOperacion"]["EmpresaID"] = int(str(target))
                body["datosOperacion"]["empresaID"] = int(str(target))
                body["datosClientes"]["EmpresaID"] = int(str(target))
                body["empresaID"] = int(str(target))

                try:
                    rows, rid, _ok, page_warnings = _fetch_paged_rows_with_fallbacks(
                        base=base,
                        path=cobros_path,
                        method=method,
                        headers=headers,
                        list_key="comprobantes",
                        page_size=page_size,
                        body=body,
                    )
                    warnings.extend(page_warnings)
                except ExternalProviderError as e:
                    status = _status_code_from_exc(e)
                    if status in {404, 405}:
                        warnings.append(
                            f"Cobros/GetList descartado (HTTP {status}, empresaID={target}, method={method.lower()})."
                        )
                        continue
                    if _is_rate_limited_error(e):
                        warnings.append("Cobros/GetList frenado por límite de rate (HTTP 429).")
                        rate_limited = True
                        break
                    raise

                target_had_response = True
                if rid:
                    request_id_last = rid
                selected_path = cobros_path
                selected_method = method.upper()

                for r in rows:
                    if not isinstance(r, dict):
                        continue
                    rid_original = _normalize_id_text(_dict_get_any(r, "EmpresaID", "empresaID", "empresa_id"))
                    if rid_original:
                        r["_empresa_id_api_original"] = rid_original
                    resolved_id, source_field = _row_empresa_id(r, str(target))
                    r["empresaID"] = resolved_id
                    r["_empresa_id_source"] = source_field
                    tname = _empresa_name_from_id(resolved_id)
                    if tname:
                        r["_empresa_target_name"] = tname
                    comprobantes.append(r)
                    added += 1

                break

            counts_by_target[str(target)] = int(added)
            if not target_had_response:
                warnings.append(f"Cobros/GetList empresaID={target}: sin respuesta util.")
            if rate_limited:
                break

        comprobantes, dups = _dedupe_comprobantes(comprobantes)
        if dups > 0:
            warnings.append(f"Comprobantes duplicados descartados: {dups}")

        if comprobantes:
            warnings.append("Se utilizó Cobros/GetList por fallback para obtener recibos operativos.")

            keys: list[dict[str, Any]] = []
            seen_keys: set[tuple[str, str, str, str, str]] = set()
            for r in comprobantes:
                if not isinstance(r, dict):
                    continue
                # Mantiene foco en los recibos realmente operativos del programa.
                if not _row_has_usable_pm_receipt(r) or not _row_has_usable_cliente(r):
                    continue
                k = _row_comprobante_key_for_getitem(r)
                if k is None:
                    continue
                tk = _normalize_getitem_key_tuple(k)
                if tk in seen_keys:
                    continue
                seen_keys.add(tk)
                keys.append(k)

            if keys and not rate_limited:
                warnings.append(
                    f"Cobros/GetItem keys preparadas: {len(keys)}/{len(comprobantes)} comprobantes."
                )
                detail_by_key: dict[tuple[str, str, str, str, str], dict] = {}
                if getitem_max_keys == 0 or (getitem_max_keys > 0 and len(keys) > getitem_max_keys):
                    warnings.append(
                        f"Cobros/GetItem omitido: {len(keys)} claves exceden el límite configurado ({getitem_max_keys})."
                    )
                else:
                    # En algunos tenants GetItem devuelve de facto ~20 elementos por request
                    # aun enviando hasta 60 claves. Usar 20 mejora cobertura real.
                    getitem_chunk = int((os.getenv("RECEIPTS_API_GETITEM_CHUNK_SIZE", "20") or "20").strip())
                    getitem_chunk = min(max(getitem_chunk, 1), 60)

                    # Ejecuta GetItem por EmpresaID (target) para no perder comprobantes
                    # de otra sucursal/empresa cuando el backend valida encabezados por empresa.
                    keys_by_target: dict[str, list[dict[str, Any]]] = {}
                    for k in keys:
                        target_emp = _normalize_id_text(k.get("EmpresaID"))
                        if not target_emp:
                            target_emp = str(targets[0] if targets else "3")
                        keys_by_target.setdefault(target_emp, []).append(k)

                    for target_emp, keys_target in keys_by_target.items():
                        getitem_headers = _headers_for_target(str(target_emp))
                        details_t, rid_item, rl_item = _fetch_getitem_details(
                            base=base,
                            getitem_path=getitem_path,
                            getitem_query=getitem_query,
                            headers=getitem_headers,
                            keys=keys_target,
                            chunk_size=getitem_chunk,
                        )
                        if rid_item:
                            request_id_last = rid_item
                        if details_t:
                            detail_by_key.update(details_t)
                        if rl_item:
                            warnings.append(
                                f"Cobros/GetItem frenado por límite de rate (HTTP 429, target={target_emp})."
                            )
                            rate_limited = True
                            break

                    all_keys = {_normalize_getitem_key_tuple(k) for k in keys}
                    missing = [k for k in keys if _normalize_getitem_key_tuple(k) not in detail_by_key]
                    ratio = len(detail_by_key) / max(len(all_keys), 1)

                    # Segunda pasada opcional para faltantes con query extendida.
                    if missing and not rate_limited:
                        alt_query = (
                            os.getenv(
                                "RECEIPTS_API_COBROS_GETITEM_QUERY_ALT",
                                "incluirDetalleDeMedioDePago=S&incluirDetalleDeAplicacion=S&incluirDetalleDeImpuestos=S",
                            )
                            or ""
                        ).strip()
                        if alt_query and alt_query != getitem_query:
                            missing_by_target: dict[str, list[dict[str, Any]]] = {}
                            for k in missing:
                                target_emp = _normalize_id_text(k.get("EmpresaID"))
                                if not target_emp:
                                    target_emp = str(targets[0] if targets else "3")
                                missing_by_target.setdefault(target_emp, []).append(k)

                            for target_emp, keys_target in missing_by_target.items():
                                getitem_headers = _headers_for_target(str(target_emp))
                                extra_alt, rid_alt, rl_alt = _fetch_getitem_details(
                                    base=base,
                                    getitem_path=getitem_path,
                                    getitem_query=alt_query,
                                    headers=getitem_headers,
                                    keys=keys_target,
                                    chunk_size=getitem_chunk,
                                )
                                if rid_alt:
                                    request_id_last = rid_alt
                                if extra_alt:
                                    detail_by_key.update(extra_alt)
                                if rl_alt:
                                    warnings.append(
                                        f"Cobros/GetItem segunda pasada frenada por límite de rate (HTTP 429, target={target_emp})."
                                    )
                                    rate_limited = True
                                    break
                            if detail_by_key:
                                missing = [k for k in keys if _normalize_getitem_key_tuple(k) not in detail_by_key]
                                warnings.append(
                                    f"Cobros/GetItem segunda pasada aplicada: cobertura {len(detail_by_key)}/{len(all_keys)}."
                                )
                                warnings.append(
                                    f"Cobros/GetItem retry adaptativo aplicado: cobertura {len(detail_by_key)}/{len(all_keys)}."
                                )

                    # Retry adaptativo acotado (evita explotar rate limit)
                    ratio = len(detail_by_key) / max(len(all_keys), 1)
                    adaptive_limit = int((os.getenv("RECEIPTS_API_GETITEM_ADAPTIVE_MAX_KEYS", "300") or "300").strip())
                    if missing and ratio < 0.95 and not rate_limited:
                        retry_chunk = int((os.getenv("RECEIPTS_API_GETITEM_RETRY_CHUNK_SIZE", "60") or "60").strip())
                        retry_subset = missing[: max(adaptive_limit, 0)] if adaptive_limit > 0 else list(missing)
                        if not retry_subset:
                            retry_subset = list(missing)
                        missing_by_target: dict[str, list[dict[str, Any]]] = {}
                        for k in retry_subset:
                            target_emp = _normalize_id_text(k.get("EmpresaID"))
                            if not target_emp:
                                target_emp = str(targets[0] if targets else "3")
                            missing_by_target.setdefault(target_emp, []).append(k)

                        for target_emp, keys_target in missing_by_target.items():
                            getitem_headers = _headers_for_target(str(target_emp))
                            extra, rid_retry, rl_retry = _fetch_getitem_details(
                                base=base,
                                getitem_path=getitem_path,
                                getitem_query=getitem_query,
                                headers=getitem_headers,
                                keys=keys_target,
                                chunk_size=max(retry_chunk, 1),
                            )
                            if rid_retry:
                                request_id_last = rid_retry
                            if extra:
                                detail_by_key.update(extra)
                            if rl_retry:
                                warnings.append(
                                    f"Cobros/GetItem retry frenado por límite de rate (HTTP 429, target={target_emp})."
                                )
                                rate_limited = True
                                break
                        warnings.append(
                            f"Cobros/GetItem retry adaptativo aplicado: cobertura {len(detail_by_key)}/{len(all_keys)} (reintentados={len(retry_subset)})."
                        )

                    # Pasadas extra sobre faltantes: algunos tenants devuelven cobertura
                    # parcial intermitente aunque el request sea válido.
                    missing = [k for k in keys if _normalize_getitem_key_tuple(k) not in detail_by_key]
                    extra_passes = int((os.getenv("RECEIPTS_API_GETITEM_EXTRA_PASSES", "0") or "0").strip())
                    extra_chunk = int((os.getenv("RECEIPTS_API_GETITEM_EXTRA_CHUNK_SIZE", "10") or "10").strip())
                    extra_chunk = min(max(extra_chunk, 1), 60)
                    extra_wait = float((os.getenv("RECEIPTS_API_GETITEM_EXTRA_WAIT_SECONDS", "0.5") or "0.5").strip())
                    for p in range(max(extra_passes, 0)):
                        if not missing or rate_limited:
                            break
                        before = len(detail_by_key)
                        missing_by_target: dict[str, list[dict[str, Any]]] = {}
                        for k in missing:
                            target_emp = _normalize_id_text(k.get("EmpresaID"))
                            if not target_emp:
                                target_emp = str(targets[0] if targets else "3")
                            missing_by_target.setdefault(target_emp, []).append(k)

                        for target_emp, keys_target in missing_by_target.items():
                            getitem_headers = _headers_for_target(str(target_emp))
                            extra_more, rid_more, rl_more = _fetch_getitem_details(
                                base=base,
                                getitem_path=getitem_path,
                                getitem_query=getitem_query,
                                headers=getitem_headers,
                                keys=keys_target,
                                chunk_size=extra_chunk,
                            )
                            if rid_more:
                                request_id_last = rid_more
                            if extra_more:
                                detail_by_key.update(extra_more)
                            if rl_more:
                                warnings.append(
                                    f"Cobros/GetItem pasada extra frenada por límite de rate (HTTP 429, target={target_emp})."
                                )
                                rate_limited = True
                                break

                        after = len(detail_by_key)
                        warnings.append(
                            f"Cobros/GetItem pasada extra {p + 1}: cobertura {after}/{len(all_keys)}."
                        )
                        if after <= before:
                            break
                        missing = [k for k in keys if _normalize_getitem_key_tuple(k) not in detail_by_key]
                        if extra_wait > 0 and not rate_limited:
                            time.sleep(extra_wait)

                    merged = 0
                    if detail_by_key:
                        detail_by_relaxed: dict[tuple[str, str, str, str], dict] = {}
                        for ek, ev in detail_by_key.items():
                            rk = _relaxed_key_from_exact_tuple(ek)
                            # Prioriza el que trae detalleDeValores útil.
                            if rk not in detail_by_relaxed:
                                detail_by_relaxed[rk] = ev
                                continue
                            cur = detail_by_relaxed[rk]
                            cur_has = _has_useful_valor_ids(cur.get("detalleDeValores"))
                            new_has = _has_useful_valor_ids(ev.get("detalleDeValores"))
                            if new_has and not cur_has:
                                detail_by_relaxed[rk] = ev

                        for r in comprobantes:
                            mk = _row_comprobante_match_key(r)
                            extra = detail_by_key.get(mk)
                            if not isinstance(extra, dict):
                                extra = detail_by_relaxed.get(_row_comprobante_relaxed_key(r))
                            if isinstance(extra, dict):
                                _merge_getitem_into_comprobante(r, extra)
                                merged += 1
                    warnings.append(
                        f"Cobros/GetItem detalle medio de pago: {merged}/{len(comprobantes)} comprobantes enriquecidos."
                    )

    # 4) Catálogos opcionales
    formas: List[dict] = []
    empresas: List[dict] = []

    if targets:
        catalog_headers = _headers_for_target(str(targets[0]))
        if not rate_limited:
            try:
                formas, rid_forms, w_forms = _get_paged_get(
                    base=base,
                    path=formas_path,
                    headers=catalog_headers,
                    list_key="formasDePago",
                    page_size=page_size,
                )
                if rid_forms:
                    request_id_last = rid_forms
                warnings.extend(w_forms)
            except ExternalProviderError as e:
                if _is_rate_limited_error(e):
                    warnings.append("FormasDePago/GetList omitido por límite de rate (HTTP 429).")
                else:
                    raise

        if not rate_limited:
            try:
                empresas, rid_emp, w_emp = _get_paged_get(
                    base=base,
                    path=empresas_path,
                    headers=catalog_headers,
                    list_key="empresas",
                    page_size=page_size,
                )
                if rid_emp:
                    request_id_last = rid_emp
                warnings.extend(w_emp)
            except ExternalProviderError as e:
                if _is_rate_limited_error(e):
                    warnings.append("No se pudo leer maestro de empresas por límite de rate (HTTP 429).")
                else:
                    warnings.append(f"No se pudo leer maestro de empresas: {e}")
            except Exception as e:
                warnings.append(f"No se pudo leer maestro de empresas: {e}")

    return ReceiptsApiResponse(
        payload={
            "comprobantes": comprobantes,
            "formasDePago": formas,
            "empresas": empresas,
            "mediosDePago": medios_pago,
            "empresa_filter": empresa_filter or "",
            "empresa_targets_used": targets,
            "comprobantes_count_by_target": counts_by_target,
            "fecha_desde": fecha_desde.isoformat(),
            "fecha_hasta": fecha_hasta.isoformat(),
            "api_comprobantes_path_used": selected_path,
            "api_comprobantes_method_used": selected_method,
        },
        request_id=request_id_last,
        warnings=warnings,
    )


def fetch_receipts_payload(
    *,
    days: int,
    empresa_filter: str | None = None,
    start_date: date | None = None,
    end_date: date | None = None,
) -> ReceiptsApiResponse:
    if not _env_bool("API_MODE_ENABLED", True):
        raise ExternalConfigError("API_MODE_ENABLED=false: modo API deshabilitado")

    if int(days) <= 0:
        raise ExternalConfigError("api_receipts_days debe ser > 0")

    fecha_hasta = end_date or (date.today() - timedelta(days=1))
    fecha_desde = start_date or (fecha_hasta - timedelta(days=max(int(days) - 1, 0)))
    if fecha_desde > fecha_hasta:
        raise ExternalConfigError("Rango API inválido: fecha_desde es posterior a fecha_hasta")

    targets = _resolve_empresa_targets(empresa_filter)
    window_days = _window_days_for_targets(targets, empresa_filter)
    windows = _split_date_windows(fecha_desde, fecha_hasta, window_days)
    if len(windows) == 1:
        return _fetch_receipts_payload_single_window(
            days=days,
            empresa_filter=empresa_filter,
            start_date=fecha_desde,
            end_date=fecha_hasta,
        )

    all_comprobantes: list[dict] = []
    formas: list[dict] = []
    empresas: list[dict] = []
    medios_pago: list[dict] = []
    warnings: list[str] = [f"Receipts API troceada en {len(windows)} ventanas de {window_days} día(s)."]
    request_id_last: str | None = None
    counts_by_target_total: dict[str, int] = {}
    selected_path: str | None = None
    selected_method: str | None = None
    successful_windows = 0
    last_timeout: ExternalTimeoutError | None = None

    for win_start, win_end in windows:
        try:
            resp = _fetch_receipts_payload_single_window(
                days=((win_end - win_start).days + 1),
                empresa_filter=empresa_filter,
                start_date=win_start,
                end_date=win_end,
            )
        except ExternalTimeoutError as e:
            last_timeout = e
            warnings.append(
                f"Timeout en receipts para ventana {win_start.isoformat()}..{win_end.isoformat()}; se conserva lo ya descargado y se sigue."
            )
            continue
        successful_windows += 1
        request_id_last = resp.request_id or request_id_last
        warnings.append(f"Ventana receipts {win_start.isoformat()}..{win_end.isoformat()} OK.")
        warnings.extend(resp.warnings)
        payload = resp.payload or {}
        all_comprobantes.extend([r for r in (payload.get("comprobantes") or []) if isinstance(r, dict)])
        if not formas:
            formas = [r for r in (payload.get("formasDePago") or []) if isinstance(r, dict)]
        if not empresas:
            empresas = [r for r in (payload.get("empresas") or []) if isinstance(r, dict)]
        if not medios_pago:
            medios_pago = [r for r in (payload.get("mediosDePago") or []) if isinstance(r, dict)]
        for k, v in (payload.get("comprobantes_count_by_target") or {}).items():
            ks = str(k or "").strip()
            counts_by_target_total[ks] = int(counts_by_target_total.get(ks) or 0) + int(v or 0)
        if selected_path is None and payload.get("api_comprobantes_path_used"):
            selected_path = payload.get("api_comprobantes_path_used")
        if selected_method is None and payload.get("api_comprobantes_method_used"):
            selected_method = payload.get("api_comprobantes_method_used")

    comprobantes, dups = _dedupe_comprobantes(all_comprobantes)
    if dups > 0:
        warnings.append(f"Comprobantes duplicados descartados entre ventanas: {dups}")
    if successful_windows == 0 and last_timeout is not None:
        raise last_timeout

    return ReceiptsApiResponse(
        payload={
            "comprobantes": comprobantes,
            "formasDePago": formas,
            "empresas": empresas,
            "mediosDePago": medios_pago,
            "empresa_filter": empresa_filter or "",
            "empresa_targets_used": targets,
            "comprobantes_count_by_target": counts_by_target_total,
            "fecha_desde": fecha_desde.isoformat(),
            "fecha_hasta": fecha_hasta.isoformat(),
            "api_comprobantes_path_used": selected_path,
            "api_comprobantes_method_used": selected_method,
        },
        request_id=request_id_last,
        warnings=warnings,
    )
