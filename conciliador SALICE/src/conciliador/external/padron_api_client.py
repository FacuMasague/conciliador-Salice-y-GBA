from __future__ import annotations

import json
import os
import socket
import ssl
from dataclasses import dataclass
from typing import Any, Dict, List, Tuple
from urllib.error import HTTPError, URLError
from urllib.parse import urlencode
from urllib.request import Request, urlopen

from .errors import ExternalConfigError, ExternalProviderError, ExternalTimeoutError


@dataclass(frozen=True)
class PadronApiResponse:
    payload: Dict[str, Any]
    request_id: str | None
    warnings: list[str]


def _env_bool(name: str, default: bool) -> bool:
    raw = os.getenv(name)
    if raw is None:
        return default
    return raw.strip().lower() in {"1", "true", "yes", "on"}


def _timeout_seconds() -> float:
    raw = os.getenv("HTTP_TIMEOUT_SECONDS", "20").strip()
    try:
        v = float(raw)
        if v <= 0:
            raise ValueError
        return v
    except Exception:
        raise ExternalConfigError("HTTP_TIMEOUT_SECONDS inválido")


def _ssl_context() -> ssl.SSLContext:
    verify = _env_bool("API_VERIFY_SSL", True)
    if verify:
        return ssl.create_default_context()
    return ssl._create_unverified_context()


def _allow_insecure_fallback() -> bool:
    # Default ON para destrabar entornos locales con cadena SSL incompleta.
    return _env_bool("API_SSL_FALLBACK_UNVERIFIED", True)


def _base_url(prefix: str) -> str:
    base = (os.getenv(f"{prefix}_BASE_URL", "") or "").strip().rstrip("/")
    if not base:
        base = "https://m5mdp.grupoesi.com.ar"
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


def _http_json(url: str, *, method: str, headers: dict[str, str], body: dict[str, Any] | None = None) -> tuple[dict[str, Any], str | None]:
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
                    raise ExternalProviderError("padron", "Respuesta JSON inválida (se esperaba objeto)")
                return payload, request_id
        except URLError as e:
            reason = getattr(e, "reason", None)
            if (
                _allow_insecure_fallback()
                and isinstance(reason, ssl.SSLCertVerificationError)
            ):
                with urlopen(req, timeout=_timeout_seconds(), context=ssl._create_unverified_context()) as resp:
                    raw = resp.read().decode("utf-8")
                    payload = json.loads(raw)
                    request_id = resp.headers.get("X-Request-Id") or resp.headers.get("x-request-id")
                    if not isinstance(payload, dict):
                        raise ExternalProviderError("padron", "Respuesta JSON inválida (se esperaba objeto)")
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
                "padron",
                "Cloudflare bloqueó el acceso (Error 1010). "
                "Hay que pedir whitelist de IP/UA del servidor cliente en m5mdp.grupoesi.com.ar.",
                status_code=e.code,
            )
        raise ExternalProviderError("padron", f"Padrón API HTTP {e.code}: {body_txt}", status_code=e.code)
    except URLError as e:
        if isinstance(e.reason, socket.timeout):
            raise ExternalTimeoutError("padron", "Timeout en Padrón API")
        raise ExternalProviderError("padron", f"Padrón API error de red: {e}")
    except socket.timeout:
        raise ExternalTimeoutError("padron", "Timeout en Padrón API")
    except json.JSONDecodeError:
        raise ExternalProviderError("padron", "Padrón API respondió JSON inválido")


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
        raise ExternalConfigError(f"Faltan {prefix}_USERNAME/{prefix}_PASSWORD (o GESI_API_USERNAME/GESI_API_PASSWORD)")

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
    if (success is False) or not token:
        msg = "Login inválido"
        err = payload.get("error")
        if isinstance(err, dict):
            msg = str(err.get("message") or msg)
        raise ExternalProviderError("padron", f"No se pudo autenticar en GESI: {msg}")
    return token


def _get_clientes_getlist(*, base: str, path: str, headers: dict[str, str], empresa_filter: str | None = None) -> Tuple[List[dict], str | None, list[str]]:
    page = 1
    size = int((os.getenv("PADRON_API_PAGE_SIZE", "500") or "500").strip())
    all_rows: List[dict] = []
    request_id_last: str | None = None
    warnings: list[str] = []

    extra_query: dict[str, str] = {}
    if empresa_filter:
        # No existe filtro textual de empresa en docs; puede resolverse por header empresaID.
        pass

    while True:
        q = {"pageNumber": str(page), "pageSize": str(size), **extra_query}
        url = f"{base}{path}?{urlencode(q)}"
        payload, request_id = _http_json(url, method="GET", headers=headers)
        request_id_last = request_id or request_id_last

        success = payload.get("success")
        if success is False:
            err = payload.get("error")
            msg = "Clientes/GetList devolvió error"
            if isinstance(err, dict):
                msg = str(err.get("message") or msg)
            raise ExternalProviderError("padron", msg)

        rows = payload.get("clientes")
        if not isinstance(rows, list):
            raise ExternalProviderError("padron", "Clientes/GetList no devolvió 'clientes' lista")
        all_rows.extend([r for r in rows if isinstance(r, dict)])

        pag = payload.get("paginacion")
        if not isinstance(pag, dict):
            break
        try:
            tp = int(pag.get("totalPaginas"))
        except Exception:
            break
        if page >= tp:
            break
        page += 1
        if page > 1000:
            warnings.append("Corte de paginación de seguridad en Clientes/GetList")
            break

    return all_rows, request_id_last, warnings


def fetch_padron_payload(*, empresa_filter: str | None = None) -> PadronApiResponse:
    if not _env_bool("API_MODE_ENABLED", True):
        raise ExternalConfigError("API_MODE_ENABLED=false: modo API deshabilitado")

    base = _base_url("PADRON_API")
    headers = _headers_base("PADRON_API")
    token = _login_token("PADRON_API", base, headers)
    headers = {**headers, "Authorization": f"Bearer {token}"}

    clientes_path = (os.getenv("PADRON_API_CLIENTES_PATH", "/api/Maestros/Clientes/GetList") or "/api/Maestros/Clientes/GetList").strip()
    if not clientes_path.startswith("/"):
        clientes_path = "/" + clientes_path

    rows, rid, warnings = _get_clientes_getlist(
        base=base,
        path=clientes_path,
        headers=headers,
        empresa_filter=empresa_filter,
    )

    return PadronApiResponse(payload={"clientes": rows}, request_id=rid, warnings=warnings)
