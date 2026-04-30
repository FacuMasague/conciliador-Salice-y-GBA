from __future__ import annotations

import datetime as dt
import os
import re
import unicodedata
from typing import Any, Dict, List, Optional, Tuple

from .errors import ExternalSchemaError
from .padron_api_client import fetch_padron_payload
from .receipts_api_client import (
    fetch_receipts_payload,
    _base_url as _receipts_base_url,
    _headers_base as _receipts_headers_base,
    _build_auth_headers_for_empresa as _receipts_build_auth_headers_for_empresa,
    _fetch_getitem_details as _receipts_fetch_getitem_details,
    _fetch_medios_pago as _receipts_fetch_medios_pago,
    _normalize_getitem_key_tuple as _receipts_normalize_getitem_key_tuple,
    _resolve_empresa_targets as _receipts_resolve_empresa_targets,
    _page_size_for_targets as _receipts_page_size_for_targets,
)
from .types import ExternalPadronEntry, ExternalPayment, ExternalReceipt
from ..pdf_parser import ReceiptPayment


def _normalize_cliente(value: object) -> Optional[str]:
    if value is None:
        return None
    s = str(value).strip()
    if not s:
        return None
    digits = "".join(ch for ch in s if ch.isdigit())
    if not digits:
        return None
    try:
        return str(int(digits))
    except Exception:
        return digits.lstrip("0") or "0"


def _normalize_cuit(value: object) -> Optional[str]:
    if value is None:
        return None
    s = str(value).strip()
    if not s:
        return None
    digits = "".join(ch for ch in s if ch.isdigit())
    if len(digits) == 11:
        return digits
    return None


def _parse_date_yyyy_mm_dd(value: object) -> str:
    s = str(value or "").strip()
    if not s:
        raise ExternalSchemaError("fecha_pago faltante")
    if re.match(r"^\d{4}-\d{2}-\d{2}$", s):
        return s
    # ISO datetime: 2026-02-23T00:00:00(.sss)(Z|+00:00)
    if re.match(r"^\d{4}-\d{2}-\d{2}T", s):
        try:
            s2 = s.replace("Z", "+00:00")
            return dt.datetime.fromisoformat(s2).date().isoformat()
        except Exception:
            pass
    for fmt in ("%d/%m/%Y", "%d-%m-%Y", "%Y/%m/%d"):
        try:
            return dt.datetime.strptime(s, fmt).date().isoformat()
        except Exception:
            pass
    raise ExternalSchemaError(f"fecha_pago inválida: {s}")


def _parse_medio(value: object) -> str:
    s = str(value or "").strip().upper()
    if s in {"TRANSFERENCIA", "MERCADOPAGO"}:
        return s
    if "MERCADO" in s:
        return "MERCADOPAGO"
    if "TRANSFER" in s:
        return "TRANSFERENCIA"
    raise ExternalSchemaError(f"medio_pago inválido: {value}")


def _parse_empresa_map(raw: str) -> Dict[str, str]:
    out: Dict[str, str] = {}
    for part in str(raw or "").split(","):
        p = part.strip()
        if not p or ":" not in p:
            continue
        k, v = p.split(":", 1)
        key = str(k).strip()
        val = str(v).strip().upper()
        if key and val:
            out[key] = val
    return out


def _normalize_text(value: object) -> str:
    s = str(value or "").strip().lower()
    s = "".join(ch for ch in unicodedata.normalize("NFKD", s) if not unicodedata.combining(ch))
    return s


def _clean_text(value: object) -> str:
    return str(value or "").strip()


def _norm_key(s: object) -> str:
    return "".join(ch for ch in str(s or "").lower() if ch.isalnum())


def _get_any(obj: Dict[str, Any], *keys: str) -> Any:
    # Permite leer variantes camel/snake/Pascal con un mismo alias lógico.
    norm_map = {_norm_key(k): v for k, v in obj.items()}
    for k in keys:
        nk = _norm_key(k)
        if nk in norm_map:
            return norm_map[nk]
    return None


def _norm_numeric_id(value: Any) -> str:
    s = str(value or "").strip()
    if not s:
        return ""
    s_num = s.replace(",", ".")
    # Soporta IDs numéricos serializados como float textual (ej: \"2.0\")
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


def _first_nonempty_text(values: List[object]) -> str:
    for v in values:
        s = _clean_text(v)
        if s:
            return s
    return ""


def _vendor_label_from_comprobante(row: Dict[str, Any]) -> str:
    def _walk_vendor_values(value: Any, *, max_depth: int = 4) -> tuple[list[str], list[str]]:
        names: list[str] = []
        ids: list[str] = []
        if max_depth < 0:
            return names, ids
        if isinstance(value, dict):
            for k, v in value.items():
                nk = _norm_key(k)
                if "vendedor" in nk:
                    txt = _clean_text(v)
                    if any(tag in nk for tag in ("nombre", "apellido", "descripcion", "completo")):
                        if txt:
                            names.append(txt)
                    elif nk in {
                        "vendedor",
                        "vendedorid",
                        "idvendedor",
                        "vendedordelclienteid",
                    }:
                        vid = _norm_numeric_id(v)
                        if vid:
                            ids.append(vid)
                        elif txt:
                            names.append(txt)
                if isinstance(v, (dict, list)):
                    child_names, child_ids = _walk_vendor_values(v, max_depth=max_depth - 1)
                    names.extend(child_names)
                    ids.extend(child_ids)
        elif isinstance(value, list):
            for item in value:
                child_names, child_ids = _walk_vendor_values(item, max_depth=max_depth - 1)
                names.extend(child_names)
                ids.extend(child_ids)
        return names, ids

    datos_cliente = row.get("datosClientes")
    if not isinstance(datos_cliente, dict):
        datos_cliente = {}
    vendedor_nombre = _first_nonempty_text(
        [
            _get_any(
                row,
                "vendedor",
                "nombreVendedor",
                "nombre_vendedor",
                "vendedorNombre",
                "vendedor_nombre",
                "apellidoYNombreVendedor",
                "apellido_y_nombre_vendedor",
                "vendedorApellidoYNombre",
                "vendedor_apellido_y_nombre",
                "nombreCompletoVendedor",
                "nombre_completo_vendedor",
                "descripcionVendedor",
                "descripcion_vendedor",
            )
            or _get_any(
                datos_cliente,
                "vendedor",
                "nombreVendedor",
                "nombre_vendedor",
                "vendedorNombre",
                "vendedor_nombre",
                "apellidoYNombreVendedor",
                "apellido_y_nombre_vendedor",
                "vendedorApellidoYNombre",
                "vendedor_apellido_y_nombre",
                "nombreCompletoVendedor",
                "nombre_completo_vendedor",
                "descripcionVendedor",
                "descripcion_vendedor",
            )
        ]
    )
    vendedor_id = _norm_numeric_id(
        _get_any(row, "vendedorID", "vendedor_id", "VendedorID", "idVendedor", "id_vendedor")
        or _get_any(
            datos_cliente,
            "VendedorDelClienteID",
            "vendedorDelClienteID",
            "vendedor_del_cliente_id",
            "vendedorID",
            "vendedor_id",
            "VendedorID",
        )
    )
    if vendedor_nombre and vendedor_id:
        return f"{vendedor_id} - {vendedor_nombre}"
    if vendedor_nombre:
        return vendedor_nombre
    if vendedor_id:
        return vendedor_id
    nested_names, nested_ids = _walk_vendor_values(row)
    nested_name = _first_nonempty_text(nested_names)
    nested_id = _first_nonempty_text(nested_ids)
    if nested_name and nested_id:
        return f"{nested_id} - {nested_name}"
    if nested_name:
        return nested_name
    if nested_id:
        return nested_id
    return ""


def _forma_pago_id_from_row(row: Dict[str, Any]) -> str:
    return str(
        _get_any(
            row,
            "formaDePagoID",
            "forma_de_pago_id",
            "formaPagoID",
            "forma_pago_id",
            "FormaDePagoID",
            "id",
        )
        or ""
    ).strip()


def _forma_pago_desc_from_row(row: Dict[str, Any]) -> str:
    return " | ".join(
        [
            str(_get_any(row, "descripcion", "Descripcion", "description") or "").strip(),
            str(_get_any(row, "denominacion", "Denominacion", "denomination") or "").strip(),
            str(_get_any(row, "nombre", "Nombre", "name") or "").strip(),
            str(_get_any(row, "codigo", "Codigo", "code") or "").strip(),
            str(_get_any(row, "formaDePago", "forma_de_pago", "FormaDePago") or "").strip(),
        ]
    ).strip(" |")


def _parse_medio_map(raw: str) -> Dict[str, str]:
    out: Dict[str, str] = {}
    for part in str(raw or "").split(","):
        p = part.strip()
        if not p or ":" not in p:
            continue
        k, v = p.split(":", 1)
        key = str(k).strip()
        val = str(v).strip().upper()
        if val in {"TRANSFERENCIA", "MERCADOPAGO"} and key:
            out[key] = val
    return out


def _normalize_recibo_pm(value: object) -> Optional[str]:
    if value is None:
        return None
    s = str(value).strip()
    if not s:
        return None
    digits = "".join(ch for ch in s if ch.isdigit())
    if digits:
        try:
            return str(int(digits))
        except Exception:
            return digits.lstrip("0") or "0"
    return None


def _comprobante_empresa_id(c: Dict[str, Any]) -> str:
    # Prioridad de negocio actual: sucursal_id define empresa.
    for v in (
        _get_any(c, "sucursal_id", "sucursalID", "sucursalId", "SucursalID"),
        _get_any(c, "empresaID", "empresa_id", "empresaId", "EmpresaID"),
    ):
        s = _norm_numeric_id(v)
        if s:
            return s
    return ""


def _comprobante_api_key(c: Dict[str, Any]) -> Dict[str, Any] | None:
    comp_id = _norm_numeric_id(_get_any(c, "comprobanteID", "ComprobanteID", "comprobante_id"))
    emp_id = _norm_numeric_id(
        _get_any(c, "_empresa_id_api_original", "EmpresaID", "empresaID", "empresa_id", "EmpresaId")
        or _comprobante_empresa_id(c)
    )
    serie = str(_get_any(c, "serie", "Serie") or "").strip()
    pv_id = _norm_numeric_id(_get_any(c, "puntoDeVentaID", "PuntoDeVentaID", "punto_de_venta_id"))
    numero = _norm_numeric_id(_get_any(c, "numero", "Numero"))
    if not comp_id or not emp_id or not pv_id or not numero:
        return None
    try:
        return {
            "ComprobanteID": int(comp_id),
            "EmpresaID": int(emp_id),
            "Serie": serie,
            "PuntoDeVentaID": int(pv_id),
            "Numero": int(numero),
        }
    except Exception:
        return None


def _empresa_from_text(value: object) -> str:
    t = _normalize_text(value)
    if "gba" in t:
        return "GBA"
    if "alarcon" in t:
        return "ALARCON"
    if "salice" in t:
        return "SALICE"
    return ""


def _is_receipt_forma_pago_supported(desc: object, extra_hint: object = "") -> bool:
    d = _normalize_text(desc)
    h = _normalize_text(extra_hint)
    txt = f"{d} {h}".strip()
    if not txt:
        # v5.0.16: si la API no informa forma de pago, no descartamos el recibo.
        return True

    exclude_markers = (
        "efectivo",
        "cheque",
        "cta cte",
        "cuenta corriente",
        "cdo - cobra vendedor",
        "cdo c/entrega efectivo",
        "cdo c/entrega",
        "30 dias",
        "45 d",
        "60 d",
        "15 d",
        "7 d",
    )
    if any(m in txt for m in exclude_markers):
        return False

    include_markers = (
        "transferencia",
        "transfer",
        "bank transfer",
        "deposito",
        "cbu",
        "interbank",
        "mercado pago",
        "mercadopago",
        "account_money",
        "bank_transfer",
        "qr",
    )
    # Si no hay marker concluyente, lo dejamos pasar igual para no perder recibos
    # cuando la API no informa correctamente la forma de pago.
    _ = any(m in txt for m in include_markers)
    return True


def _classify_bancarizable(desc: object, extra_hint: object = "") -> tuple[bool, str]:
    d = _normalize_text(desc)
    h = _normalize_text(extra_hint)
    txt = f"{d} {h}".strip()
    if not txt:
        return True, "UNKNOWN"

    non_bank_markers = (
        "efectivo",
        "cheque",
        "cta cte",
        "cuenta corriente",
        "cdo - cobra vendedor",
        "cdo c/entrega efectivo",
        "cdo c/entrega",
        "c/ent.cheque",
        "c/entrega de cheque",
        "30 dias",
        "45 d",
        "60 d",
        "15 d",
        "7 d",
    )
    if any(m in txt for m in non_bank_markers):
        return False, "NON_BANKABLE"

    bank_markers = (
        "transferencia",
        "transfer",
        "bank transfer",
        "deposito",
        "cbu",
        "interbank",
        "mercado pago",
        "mercadopago",
        "account_money",
        "bank_transfer",
        "qr",
    )
    if any(m in txt for m in bank_markers):
        return True, "BANKABLE"

    return True, "UNKNOWN"


def _infer_medio_pago(desc: object, hint: object = "") -> str:
    txt = _normalize_text(f"{desc or ''} {hint or ''}")
    mp_markers = (
        "mercado pago",
        "mercadopago",
        "account_money",
        "money in",
        "qr",
        "point",
        "checkout pro",
        "m.pago",
        "mpago",
    )
    tr_markers = (
        "transferencia",
        "transfer",
        "bank transfer",
        "bank_transfer",
        "banelco",
        "cbu",
        "deposito",
        "interbank",
    )
    if any(m in txt for m in mp_markers):
        return "MERCADOPAGO"
    if any(m in txt for m in tr_markers):
        return "TRANSFERENCIA"
    return "NO_INFORMADO"


def _medio_pago_display(
    *,
    forma_desc_catalogo: str,
    forma_de_pago: object,
    forma_pago: object,
    descripcion_forma: object,
    denominacion_forma: object,
    medio_pago: object,
    metodo_pago: object,
    canal_cobro: object,
    tipo_cobro: object,
    origen_cobro: object,
    hint: str,
) -> str:
    # Prioriza el texto específico del comprobante; si no, usa catálogo.
    raw = _first_nonempty_text(
        [
            forma_de_pago,
            forma_pago,
            descripcion_forma,
            denominacion_forma,
            medio_pago,
            metodo_pago,
            canal_cobro,
            tipo_cobro,
            origen_cobro,
            forma_desc_catalogo,
        ]
    )
    if raw:
        return raw
    return "NO_INFORMADO"


def _resolve_empresa(
    empresa_id: object,
    *,
    empresa_filter: str | None = None,
    empresas_by_id: Dict[str, str] | None = None,
    company_hint: str = "",
) -> str:
    # Regla operativa: priorizar siempre sucursal/empresa numérica del comprobante.
    # No forzar por empresa_filter porque rompe el split SALICE/ALARCON cuando
    # llega data mixta en una misma consulta.
    eid = _norm_numeric_id(empresa_id)

    by_hint = _empresa_from_text(company_hint)
    if by_hint:
        return by_hint

    if eid == "3":
        return "SALICE"
    if eid == "6":
        return "ALARCON"

    mapping = _parse_empresa_map(os.getenv("RECEIPTS_API_EMPRESA_MAP", ""))
    if eid in mapping:
        return mapping[eid]

    if empresas_by_id and eid in empresas_by_id:
        by_master = _empresa_from_text(empresas_by_id[eid])
        if by_master:
            return by_master

    direct = _empresa_from_text(eid)
    if direct:
        return direct

    if eid == "2":
        return "GBA"
    if empresa_filter:
        # Solo fallback final si no hubo forma confiable de resolver empresa.
        return str(empresa_filter).strip().upper()
    return ""


def _collect_text_fragments(value: Any) -> List[str]:
    out: List[str] = []
    if isinstance(value, dict):
        for k, v in value.items():
            out.append(str(k))
            out.extend(_collect_text_fragments(v))
    elif isinstance(value, list):
        for it in value:
            out.extend(_collect_text_fragments(it))
    elif isinstance(value, (str, int, float, bool)):
        s = str(value).strip()
        if s:
            out.append(s)
    return out


def _collect_keyed_fragments(value: Any, parent_key: str = "") -> List[tuple[str, str]]:
    out: List[tuple[str, str]] = []
    if isinstance(value, dict):
        for k, v in value.items():
            key = str(k or "")
            if isinstance(v, (str, int, float, bool)):
                sv = str(v).strip()
                if sv:
                    out.append((key, sv))
            out.extend(_collect_keyed_fragments(v, key))
    elif isinstance(value, list):
        for it in value:
            out.extend(_collect_keyed_fragments(it, parent_key))
    return out


def _medio_detail_text(value: Any) -> str:
    pairs = _collect_keyed_fragments(value)
    tokens: list[str] = []
    for k, v in pairs:
        nk = _normalize_text(k)
        # Priorizamos campos realmente vinculados a medios de pago, no texto genérico.
        if any(tag in nk for tag in ("medio", "pago", "forma", "banco", "caja", "valor", "tarjeta", "transfer")):
            tokens.append(v)
    # dedup estable
    out: list[str] = []
    seen: set[str] = set()
    for t in tokens:
        tt = str(t).strip()
        if tt and tt not in seen:
            seen.add(tt)
            out.append(tt)
    return " | ".join(out)


def _extract_forma_pago_from_detail(c: Dict[str, Any]) -> Tuple[str, List[str]]:
    detail_candidates: List[Any] = []
    for k in (
        "detalleDeMedioDePago",
        "detalleDeMediosDePago",
        "detalleDeValores",
        "detalleMedioDePago",
        "mediosDePago",
        "medios_pago",
        "valores",
    ):
        v = _get_any(c, k)
        if isinstance(v, list) and v:
            detail_candidates.append(v)
        elif isinstance(v, dict):
            detail_candidates.append([v])

    best_id = ""
    descs: List[str] = []

    def _add_desc(x: object) -> None:
        s = str(x or "").strip()
        if not s:
            return
        if s not in descs:
            descs.append(s)

    for lst in detail_candidates:
        for it in lst:
            if not isinstance(it, dict):
                continue
            fid = str(
                _get_any(
                    it,
                    "formaDePagoID",
                    "formaPagoID",
                    "forma_de_pago_id",
                    "medioDePagoID",
                    "medio_pago_id",
                )
                or ""
            ).strip()
            if fid and fid not in {"0", "0.0", "0,0"} and not best_id:
                best_id = fid
            _add_desc(_get_any(it, "formaDePago", "forma_pago", "medioPago", "metodoPago"))
            _add_desc(_get_any(it, "descripcion", "descripcionFormaDePago", "denominacion", "nombre"))
            fp_obj = _get_any(it, "formaDePagoObj", "formaDePagoDetalle", "formaDePagoData")
            if isinstance(fp_obj, dict):
                _add_desc(_get_any(fp_obj, "descripcion", "denominacion", "nombre", "formaDePago"))
                fid2 = str(_get_any(fp_obj, "formaDePagoID", "id") or "").strip()
                if fid2 and fid2 not in {"0", "0.0", "0,0"} and not best_id:
                    best_id = fid2

    return best_id, descs


def _medio_from_forma_pago(
    forma_pago_id: object,
    formas_by_id: Dict[str, str],
    *,
    desc_candidates: List[str] | None = None,
    hint_text: str = "",
) -> tuple[str, str | None]:
    key = str(forma_pago_id or "").strip()
    map_by_id = _parse_medio_map(os.getenv("RECEIPTS_API_FORMA_PAGO_MAP", ""))
    if key and key in map_by_id:
        return map_by_id[key], None

    desc_parts: List[str] = []
    base_desc = str(formas_by_id.get(key, "") or "").strip()
    if base_desc:
        desc_parts.append(base_desc)
    if desc_candidates:
        for dsc in desc_candidates:
            ds = str(dsc or "").strip()
            if ds:
                desc_parts.append(ds)
    desc = " | ".join(desc_parts)
    d = _normalize_text(desc)

    mp_markers = (
        "mercado pago",
        "mercadopago",
        "m.pago",
        "mpago",
        "m pago",
        "money in",
        "checkout pro",
        "point",
        "qr",
    )
    trx_markers = (
        "transferencia",
        "transfer",
        "banco",
        "banelco",
        "cbu",
        "deposito",
    )

    has_mp = any(m in d for m in mp_markers) or bool(re.search(r"\bmp\b", d))
    has_trx = any(m in d for m in trx_markers)
    if has_mp and has_trx:
        return "", f"formaDePagoID={key or '?'} desc ambigua ('{desc}')"
    if has_mp:
        return "MERCADOPAGO", None
    if has_trx:
        return "TRANSFERENCIA", None

    # Fallback con pistas del propio comprobante cuando la forma de pago no alcanza.
    h = _normalize_text(hint_text)
    hint_has_mp = any(m in h for m in mp_markers) or bool(re.search(r"\bmp\b", h))
    hint_has_trx = any(m in h for m in trx_markers)
    if hint_has_mp and hint_has_trx:
        return "", "medio ambiguo por texto de comprobante"
    if hint_has_mp:
        return "MERCADOPAGO", "medio inferido por texto del comprobante"
    if hint_has_trx:
        return "TRANSFERENCIA", "medio inferido por texto del comprobante"

    if not key:
        return "", "formaDePagoID faltante"
    if not d:
        return "", f"formaDePagoID={key} sin descripción en formasDePago/comprobante"
    return "", f"formaDePagoID={key} desc='{desc}' no clasificada"


def _medio_pago_desc_from_row(row: Dict[str, Any]) -> str:
    return _first_nonempty_text(
        [
            _get_any(row, "descripcion", "Descripcion", "description"),
            _get_any(row, "denominacion", "Denominacion", "denomination"),
            _get_any(row, "nombre", "Nombre", "name"),
        ]
    )


def _medio_pago_tipo_from_row(row: Dict[str, Any]) -> str:
    return str(_get_any(row, "tipo", "Tipo") or "").strip().upper()


def _build_medios_lookup(rows: Any) -> tuple[Dict[tuple[str, str], Dict[str, Any]], Dict[str, Dict[str, Any]]]:
    by_emp_valor: Dict[tuple[str, str], Dict[str, Any]] = {}
    by_valor: Dict[str, Dict[str, Any]] = {}
    if not isinstance(rows, list):
        return by_emp_valor, by_valor
    for r in rows:
        if not isinstance(r, dict):
            continue
        emp = _norm_numeric_id(_get_any(r, "empresaID", "empresa_id", "empresaId", "EmpresaID"))
        valor = _norm_numeric_id(_get_any(r, "valorID", "valor_id", "valorId", "ValorID"))
        if not valor:
            continue
        if emp:
            by_emp_valor.setdefault((emp, valor), r)
        by_valor.setdefault(valor, r)
    return by_emp_valor, by_valor


def _detalle_valor_ids(c: Dict[str, Any]) -> List[str]:
    out: List[str] = []
    detalle = _get_any(c, "detalleDeValores", "detalle_de_valores", "DetalleDeValores")
    if not isinstance(detalle, list):
        return out
    for it in detalle:
        if not isinstance(it, dict):
            continue
        raw = _get_any(it, "valorID", "valor_id", "valorId", "ValorID")
        sval = _norm_numeric_id(raw)
        if not sval:
            continue
        if sval not in out:
            out.append(sval)
    return out


def _medio_from_valor_ids(
    *,
    empresa_id: str,
    valor_ids: List[str],
    medios_by_emp_valor: Dict[tuple[str, str], Dict[str, Any]],
    medios_by_valor: Dict[str, Dict[str, Any]],
) -> tuple[str, bool | None]:
    if not valor_ids:
        return "", None
    rows: List[Dict[str, Any]] = []
    for vid in valor_ids:
        r = None
        if empresa_id:
            r = medios_by_emp_valor.get((str(empresa_id), str(vid)))
        if r is None:
            r = medios_by_valor.get(str(vid))
        if isinstance(r, dict):
            rows.append(r)
    if not rows:
        return "", None

    descs: List[str] = []
    for r in rows:
        d = _medio_pago_desc_from_row(r)
        if d and d not in descs:
            descs.append(d)

    merged_desc = " + ".join(descs).strip()
    txt = _normalize_text(merged_desc)

    # Clasificación por mediosDePago.descripcion (fuente oficial ESI).
    if "mercado pago" in txt or "mercadopago" in txt:
        return "Mercado Pago", True
    if any(
        m in txt
        for m in (
            "transf. bancaria",
            "transf bancaria",
            "transferencia bancaria",
            "transferencia",
            "transf.",
            "banelco",
            "cbu",
            "deposito banc",
            "dep. banc",
        )
    ):
        return "Transf. Bancaria", True
    if any(m in txt for m in ("echeq", "e-cheq", "cheque electron")):
        return "eCheq", True

    # Si hubo valorID mapeado pero no coincide con medios bancarizables definidos,
    # se considera no bancarizable.
    return merged_desc, False


def _is_bankable_medio_desc(desc: object) -> bool:
    txt = _normalize_text(desc)
    if not txt:
        return False
    if "mercado pago" in txt or "mercadopago" in txt:
        return True
    if any(
        m in txt
        for m in (
            "transf. bancaria",
            "transf bancaria",
            "transferencia bancaria",
            "transferencia",
            "transf.",
            "banelco",
            "cbu",
            "deposito banc",
            "dep. banc",
        )
    ):
        return True
    if any(m in txt for m in ("echeq", "e-cheq", "cheque electron")):
        return True
    return False


def _to_float_or_none(value: Any) -> float | None:
    if value is None:
        return None
    if isinstance(value, (int, float)):
        return float(value)
    s = str(value).strip()
    if not s:
        return None
    s = s.replace("$", "").replace(" ", "")
    if "." in s and "," in s:
        s = s.replace(".", "").replace(",", ".")
    elif "," in s:
        s = s.replace(",", ".")
    try:
        return float(s)
    except Exception:
        return None


def _bankable_amount_from_detalle(
    *,
    c: Dict[str, Any],
    empresa_id: str,
    medios_by_emp_valor: Dict[tuple[str, str], Dict[str, Any]],
    medios_by_valor: Dict[str, Dict[str, Any]],
) -> float | None:
    detalle = _get_any(c, "detalleDeValores", "detalle_de_valores", "DetalleDeValores")
    if not isinstance(detalle, list) or not detalle:
        return None

    total = 0.0
    found_bankable = False
    for it in detalle:
        if not isinstance(it, dict):
            continue
        vid = _norm_numeric_id(_get_any(it, "valorID", "valor_id", "valorId", "ValorID"))
        if not vid:
            continue
        r = None
        if empresa_id:
            r = medios_by_emp_valor.get((str(empresa_id), str(vid)))
        if r is None:
            r = medios_by_valor.get(str(vid))
        if not isinstance(r, dict):
            continue
        desc = _medio_pago_desc_from_row(r)
        if not _is_bankable_medio_desc(desc):
            continue
        imp = _to_float_or_none(_get_any(it, "importe", "Importe"))
        if imp is None:
            continue
        total += float(imp)
        found_bankable = True

    if not found_bankable:
        return None
    return float(total)


def _parse_float(value: object, field_name: str) -> float:
    if isinstance(value, (int, float)):
        return float(value)
    s = str(value or "").strip()
    if not s:
        raise ExternalSchemaError(f"{field_name} faltante")
    s = s.replace("$", "").replace(" ", "")
    if "." in s and "," in s:
        s = s.replace(".", "").replace(",", ".")
    elif "," in s:
        s = s.replace(",", ".")
    try:
        return float(s)
    except Exception:
        raise ExternalSchemaError(f"{field_name} inválido: {value}")


def _gesi_nro_recibo_from_comprobante(c: Dict[str, Any]) -> Optional[str]:
    # En GBA usamos el número del comprobante/recibo devuelto por la API.
    numero = _normalize_recibo_pm(_get_any(c, "numero", "Numero", "nro_recibo", "nroRecibo", "NroRecibo"))
    if numero:
        return numero

    # Compatibilidad con tenants viejos: si no hay número de comprobante, usar PM.
    nro_pm = _normalize_recibo_pm(_get_any(c, "nro_recibo_pm", "nroReciboPm", "NroReciboPm"))
    if nro_pm:
        return nro_pm

    codigo = str(_get_any(c, "codigoDeImportacion", "codigo_de_importacion", "CodigoDeImportacion") or "").strip()
    if not codigo:
        return None
    t = _normalize_text(codigo)
    nro_from_codigo = _normalize_recibo_pm(codigo)
    if not nro_from_codigo:
        return None

    if "pm" in t:
        return nro_from_codigo
    if codigo.isdigit():
        return nro_from_codigo
    return None


def _expect_obj(value: Any, field: str) -> Dict[str, Any]:
    if not isinstance(value, dict):
        raise ExternalSchemaError(f"{field} debe ser objeto")
    return value


def _extract_list(payload: Dict[str, Any], keys: list[str]) -> List[Dict[str, Any]]:
    for k in keys:
        v = payload.get(k)
        if isinstance(v, list):
            out: list[Dict[str, Any]] = []
            for i, it in enumerate(v):
                out.append(_expect_obj(it, f"{k}[{i}]"))
            return out
    raise ExternalSchemaError(f"No se encontró lista requerida en payload ({', '.join(keys)})")


def _to_external_receipt(obj: Dict[str, Any]) -> ExternalReceipt:
    empresa = str(obj.get("empresa") or "").strip().upper()
    nro_recibo = str(obj.get("nro_recibo") or "").strip()
    nro_cliente = str(obj.get("nro_cliente") or "").strip()
    if not empresa or not nro_recibo or not nro_cliente:
        raise ExternalSchemaError("receipt con empresa/nro_recibo/nro_cliente faltante")
    return ExternalReceipt(
        empresa=empresa,
        nro_recibo=nro_recibo,
        nro_cliente=nro_cliente,
        cliente_nombre=str(obj.get("cliente_nombre") or "").strip(),
        vendedor=str(obj.get("vendedor") or "").strip(),
    )


def _to_external_payment(obj: Dict[str, Any], fallback_receipt: ExternalReceipt | None = None) -> ExternalPayment:
    empresa = str(obj.get("empresa") or (fallback_receipt.empresa if fallback_receipt else "")).strip().upper()
    nro_recibo = str(obj.get("nro_recibo") or (fallback_receipt.nro_recibo if fallback_receipt else "")).strip()
    nro_cliente = str(obj.get("nro_cliente") or (fallback_receipt.nro_cliente if fallback_receipt else "")).strip()
    if not empresa or not nro_recibo or not nro_cliente:
        raise ExternalSchemaError("payment con empresa/nro_recibo/nro_cliente faltante")
    return ExternalPayment(
        empresa=empresa,
        nro_recibo=nro_recibo,
        nro_cliente=nro_cliente,
        cliente_nombre=str(obj.get("cliente_nombre") or (fallback_receipt.cliente_nombre if fallback_receipt else "")).strip(),
        vendedor=str(obj.get("vendedor") or (fallback_receipt.vendedor if fallback_receipt else "")).strip(),
        medio_pago=_parse_medio(obj.get("medio_pago")),
        fecha_pago=_parse_date_yyyy_mm_dd(obj.get("fecha_pago")),
        importe_pago=_parse_float(obj.get("importe_pago"), "importe_pago"),
        detalle_pago=str(obj.get("detalle_pago") or "").strip(),
        api_key=obj.get("api_key") if isinstance(obj.get("api_key"), dict) else None,
    )


def _external_payment_to_internal(p: ExternalPayment) -> ReceiptPayment:
    return ReceiptPayment(
        empresa=p.empresa,
        nro_recibo=p.nro_recibo,
        nro_cliente=p.nro_cliente,
        cliente_nombre=p.cliente_nombre or None,
        vendedor=p.vendedor or None,
        medio_pago=p.medio_pago,
        fecha_pago=p.fecha_pago,
        importe_pago=float(p.importe_pago),
        detalle_pago=p.detalle_pago or None,
        api_key=p.api_key if isinstance(p.api_key, dict) else None,
    )


def fetch_receipts_and_payments(
    days: int,
    empresa_filter: str | None = None,
    start_date: dt.date | None = None,
    end_date: dt.date | None = None,
) -> Tuple[List[ReceiptPayment], Dict[str, Any]]:
    if int(days) <= 0:
        raise ExternalSchemaError("api_receipts_days debe ser > 0")

    try:
        resp = fetch_receipts_payload(
            days=int(days),
            empresa_filter=empresa_filter,
            start_date=start_date,
            end_date=end_date,
        )
    except TypeError as e:
        # Compatibilidad con tests/mocks viejos que todavía no aceptan overrides de rango.
        if "end_date" not in str(e) and "start_date" not in str(e):
            raise
        resp = fetch_receipts_payload(days=int(days), empresa_filter=empresa_filter)
    payload = _expect_obj(resp.payload, "payload")

    payments_flat: list[ExternalPayment] = []
    medio_stats = {"BANKABLE": 0, "NON_BANKABLE": 0, "UNKNOWN": 0}

    # Contract option A: payload.payments[]
    if isinstance(payload.get("payments"), list):
        for i, it in enumerate(payload["payments"]):
            payments_flat.append(_to_external_payment(_expect_obj(it, f"payments[{i}]")))

    # Contract option B: payload.receipts[] with nested payments[]
    elif isinstance(payload.get("receipts"), list):
        for i, rit in enumerate(payload["receipts"]):
            robj = _expect_obj(rit, f"receipts[{i}]")
            receipt = _to_external_receipt(robj)
            nested = robj.get("payments")
            if not isinstance(nested, list) or not nested:
                raise ExternalSchemaError(f"receipts[{i}].payments faltante o vacío")
            for j, pit in enumerate(nested):
                payments_flat.append(_to_external_payment(_expect_obj(pit, f"receipts[{i}].payments[{j}]"), receipt))
    # Contract option C (GESI): payload.comprobantes[] + payload.formasDePago[]
    elif isinstance(payload.get("comprobantes"), list):
        formas_rows = payload.get("formasDePago")
        formas_by_id: Dict[str, str] = {}
        if isinstance(formas_rows, list):
            for fr in formas_rows:
                if not isinstance(fr, dict):
                    continue
                fid = _forma_pago_id_from_row(fr)
                desc = _forma_pago_desc_from_row(fr)
                if fid:
                    formas_by_id[fid] = desc
        medios_by_emp_valor, medios_by_valor = _build_medios_lookup(payload.get("mediosDePago"))

        empresas_rows = payload.get("empresas")
        empresas_by_id: Dict[str, str] = {}
        if isinstance(empresas_rows, list):
            for er in empresas_rows:
                if not isinstance(er, dict):
                    continue
                eid = str(
                    er.get("empresaID")
                    or er.get("id")
                    or er.get("empresaId")
                    or er.get("codigo")
                    or ""
                ).strip()
                ename = " | ".join(
                    [
                        str(er.get("descripcion") or "").strip(),
                        str(er.get("denominacion") or "").strip(),
                        str(er.get("nombre") or "").strip(),
                        str(er.get("razonSocial") or "").strip(),
                    ]
                ).strip(" |")
                if eid:
                    empresas_by_id[eid] = ename

        skipped_without_receipt_number = 0
        for i, it in enumerate(payload["comprobantes"]):
            c = _expect_obj(it, f"comprobantes[{i}]")
            nro_recibo = _gesi_nro_recibo_from_comprobante(c)
            nro_cliente = str(
                _get_any(c, "clienteID", "cliente_id", "clienteId", "ClienteID")
                or ""
            ).strip()
            if not nro_recibo or not nro_cliente:
                skipped_without_receipt_number += 1
                continue

            medio_detail = _medio_detail_text(c)
            detail_forma_id, detail_descs = _extract_forma_pago_from_detail(c)
            top_forma_id = str(
                _get_any(
                    c,
                    "formaDePagoID",
                    "forma_de_pago_id",
                    "forma_pago_id",
                    "formaPagoID",
                    "FormaDePagoID",
                )
                or ""
            ).strip()
            forma_id_for_lookup = detail_forma_id or top_forma_id
            catalog_desc = str(formas_by_id.get(forma_id_for_lookup, "") or "")
            hint = " ".join(
                [
                    str(_get_any(c, "notas", "Notas", "nota", "Nota") or ""),
                    str(_get_any(c, "notas2", "Notas2", "nota2", "Nota2") or ""),
                    str(_get_any(c, "subtipo", "Subtipo", "subTipo", "SubTipo") or ""),
                    str(_get_any(c, "serie", "Serie") or ""),
                    str(_get_any(c, "codigoDeImportacion", "codigo_de_importacion", "CodigoDeImportacion") or ""),
                    medio_detail,
                ]
            )
            forma_desc = " | ".join(
                [
                    catalog_desc,
                    str(_get_any(c, "formaDePago", "formaPago", "forma_de_pago", "FormaDePago") or ""),
                    str(_get_any(c, "medioPago", "medio_pago", "MedioPago") or ""),
                    str(_get_any(c, "metodoPago", "metodo_pago", "MetodoPago") or ""),
                    str(_get_any(c, "canalDeCobro", "canal_de_cobro", "CanalDeCobro") or ""),
                    str(_get_any(c, "tipoDeCobro", "tipo_de_cobro", "TipoDeCobro") or ""),
                    str(_get_any(c, "origenDelCobro", "origen_del_cobro", "OrigenDelCobro") or ""),
                    str(_get_any(c, "descripcionFormaDePago", "descripcion_forma_de_pago", "DescripcionFormaDePago") or ""),
                    str(_get_any(c, "denominacionFormaDePago", "denominacion_forma_de_pago", "DenominacionFormaDePago") or ""),
                    " | ".join(detail_descs),
                ]
            )
            medio_hint = " | ".join(
                [
                    str(_get_any(c, "formaDePago", "formaPago", "forma_de_pago", "FormaDePago") or ""),
                    str(_get_any(c, "medioPago", "medio_pago", "MedioPago") or ""),
                    str(_get_any(c, "metodoPago", "metodo_pago", "MetodoPago") or ""),
                    str(_get_any(c, "canalDeCobro", "canal_de_cobro", "CanalDeCobro") or ""),
                    str(_get_any(c, "tipoDeCobro", "tipo_de_cobro", "TipoDeCobro") or ""),
                    str(_get_any(c, "origenDelCobro", "origen_del_cobro", "OrigenDelCobro") or ""),
                    str(_get_any(c, "descripcionFormaDePago", "descripcion_forma_de_pago", "DescripcionFormaDePago") or ""),
                    str(_get_any(c, "denominacionFormaDePago", "denominacion_forma_de_pago", "DenominacionFormaDePago") or ""),
                    " | ".join(detail_descs),
                    medio_detail,
                ]
            )
            empresa_id_for_medio = _comprobante_empresa_id(c)
            valor_ids = _detalle_valor_ids(c)
            medio_from_valor, bancarizable_from_valor = _medio_from_valor_ids(
                empresa_id=empresa_id_for_medio,
                valor_ids=valor_ids,
                medios_by_emp_valor=medios_by_emp_valor,
                medios_by_valor=medios_by_valor,
            )
            importe_total = _parse_float(
                _get_any(c, "importeTotal", "importe_total", "ImporteTotal"),
                "importeTotal",
            )

            if bancarizable_from_valor is True:
                is_bancarizable, medio_class = True, "BANKABLE"
            elif bancarizable_from_valor is False:
                is_bancarizable, medio_class = False, "NON_BANKABLE"
            else:
                if str(empresa_filter or "").strip().upper() == "GBA":
                    # En GBA, GetList arma el universo inicial y GetItem termina de
                    # decidir el medio real. Si no viene detalle suficiente acá,
                    # dejamos el recibo como UNKNOWN para enriquecerlo luego.
                    is_bancarizable, medio_class = True, "UNKNOWN"
                else:
                    is_bancarizable, medio_class = _classify_bancarizable(forma_desc, medio_hint)
            medio_stats[medio_class] = int(medio_stats.get(medio_class, 0) + 1)
            medio_label = _medio_pago_display(
                forma_desc_catalogo=str(
                    catalog_desc
                    or ""
                ),
                forma_de_pago=_first_nonempty_text(
                    [
                        _get_any(c, "formaDePago", "forma_de_pago", "FormaDePago"),
                        " | ".join(detail_descs),
                    ]
                ),
                forma_pago=_get_any(c, "formaPago", "forma_pago"),
                descripcion_forma=_get_any(c, "descripcionFormaDePago", "descripcion_forma_de_pago", "DescripcionFormaDePago"),
                denominacion_forma=_get_any(c, "denominacionFormaDePago", "denominacion_forma_de_pago", "DenominacionFormaDePago"),
                medio_pago=_get_any(c, "medioPago", "medio_pago", "MedioPago"),
                metodo_pago=_get_any(c, "metodoPago", "metodo_pago", "MetodoPago"),
                canal_cobro=_get_any(c, "canalDeCobro", "canal_de_cobro", "CanalDeCobro"),
                tipo_cobro=_get_any(c, "tipoDeCobro", "tipo_de_cobro", "TipoDeCobro"),
                origen_cobro=_get_any(c, "origenDelCobro", "origen_del_cobro", "OrigenDelCobro"),
                hint=hint,
            )
            # Fuente oficial ESI: detalleDeValores.valorID + Maestros/MediosDePago.
            # Si existe mapeo por valorID, lo priorizamos también como etiqueta visible.
            medio_valor_lbl = str(medio_from_valor or "").strip()
            if medio_valor_lbl:
                medio_label = medio_valor_lbl
            elif _normalize_text(medio_label) in {"", "no_informado"}:
                # Si no hay etiqueta textual pero sí valorID en detalle, exponerlo
                # para evitar ocultar información útil en UI/debug.
                if valor_ids:
                    medio_label = "VALOR_ID_" + "+".join(valor_ids)
                else:
                    medio_label = "SIN_MEDIO_API"

            importe_bankable = _bankable_amount_from_detalle(
                c=c,
                empresa_id=empresa_id_for_medio,
                medios_by_emp_valor=medios_by_emp_valor,
                medios_by_valor=medios_by_valor,
            )

            payments_flat.append(
                ExternalPayment(
                    empresa=(
                        "GBA"
                        if str(empresa_filter or "").strip().upper() == "GBA"
                        else _resolve_empresa(
                            _comprobante_empresa_id(c),
                            empresa_filter=empresa_filter,
                            empresas_by_id=empresas_by_id,
                            company_hint=" | ".join(
                                [
                                    str(c.get("_empresa_target_name") or ""),
                                    str(c.get("empresa") or ""),
                                    str(c.get("nombreEmpresa") or ""),
                                    str(c.get("empresaNombre") or ""),
                                    str(c.get("razonSocialEmpresa") or ""),
                                ]
                            ),
                        )
                    ),
                    nro_recibo=nro_recibo,
                    nro_cliente=nro_cliente,
                    cliente_nombre=str(_get_any(c, "razonSocial", "razon_social", "RazonSocial") or "").strip(),
                    vendedor=_vendor_label_from_comprobante(c),
                    medio_pago=medio_label,
                    fecha_pago=_parse_date_yyyy_mm_dd(
                        _get_any(
                            c,
                            "fechaDeEmision",
                            "fecha_de_emision",
                            "FechaDeEmision",
                            "fechaDePrimerVencimiento",
                            "fecha_de_primer_vencimiento",
                            "FechaDePrimerVencimiento",
                        )
                    ),
                    importe_pago=importe_bankable if importe_bankable is not None else importe_total,
                    detalle_pago=str(_get_any(c, "notas", "nota", "Notas", "Nota") or "").strip(),
                    api_key=_comprobante_api_key(c),
                )
            )
        if skipped_without_receipt_number > 0:
            resp.warnings.append(
                f"Se omitieron {skipped_without_receipt_number} comprobantes sin nro de recibo utilizable."
            )
    else:
        _extract_list(payload, ["payments", "receipts", "comprobantes"])

    if not payments_flat:
        meta = {
            "api_request_id": resp.request_id,
            "external_warnings": list(resp.warnings or []) + [
                "La API de recibos no devolvió pagos válidos después de aplicar validaciones/filtros."
            ],
            "payments_count": 0,
            "payments_by_empresa": {},
            "medio_bancarizable_stats": medio_stats,
            "api_empresa_targets_used": payload.get("empresa_targets_used"),
            "api_comprobantes_count_by_target": payload.get("comprobantes_count_by_target"),
            "api_fecha_desde": payload.get("fecha_desde"),
            "api_fecha_hasta": payload.get("fecha_hasta"),
        }
        return [], meta

    out = [_external_payment_to_internal(p) for p in payments_flat]
    payments_by_empresa: Dict[str, int] = {}
    for p in out:
        key = str(p.empresa or "").strip().upper() or "SIN_EMPRESA"
        payments_by_empresa[key] = int(payments_by_empresa.get(key, 0) + 1)
    meta = {
        "api_request_id": resp.request_id,
        "external_warnings": list(resp.warnings or []),
        "payments_count": len(out),
        "payments_by_empresa": payments_by_empresa,
        "medio_bancarizable_stats": medio_stats,
        "api_empresa_targets_used": payload.get("empresa_targets_used"),
        "api_comprobantes_count_by_target": payload.get("comprobantes_count_by_target"),
        "api_fecha_desde": payload.get("fecha_desde"),
        "api_fecha_hasta": payload.get("fecha_hasta"),
    }
    return out, meta


def fetch_payment_detail_map_for_api_keys(
    api_keys: List[Dict[str, Any]],
    empresa_filter: str | None = None,
) -> Tuple[Dict[tuple[str, str, str, str, str], Dict[str, Any]], List[str]]:
    unique_keys: list[dict[str, Any]] = []
    seen: set[tuple[str, str, str, str, str]] = set()
    for raw in api_keys or []:
        if not isinstance(raw, dict):
            continue
        tk = _receipts_normalize_getitem_key_tuple(raw)
        if not any(tk) or tk in seen:
            continue
        seen.add(tk)
        unique_keys.append(raw)

    if not unique_keys:
        return {}, []

    warnings: list[str] = []
    base = _receipts_base_url("RECEIPTS_API")
    headers_root = _receipts_headers_base("RECEIPTS_API")
    explicit_targets = sorted(
        {
            str(_norm_numeric_id(k.get("EmpresaID")) or "").strip()
            for k in unique_keys
            if str(_norm_numeric_id(k.get("EmpresaID")) or "").strip()
        }
    )
    targets = explicit_targets or _receipts_resolve_empresa_targets(empresa_filter)
    page_size = _receipts_page_size_for_targets(targets, empresa_filter)
    getitem_chunk = 15

    headers_cache: Dict[str, dict[str, str]] = {}

    def _headers_for_target(target_empresa_id: str) -> dict[str, str]:
        key = str(target_empresa_id)
        if key not in headers_cache:
            headers_cache[key] = _receipts_build_auth_headers_for_empresa(
                base=base,
                headers_root=headers_root,
                empresa_id=key,
                drop_sucursal=False,
            )
        return dict(headers_cache[key])

    medios_pago, _rid_mp, w_mp = _receipts_fetch_medios_pago(
        base=base,
        headers_for_target=_headers_for_target,
        targets=targets,
        medios_path="/api/Maestros/MediosDePago/GetList",
        page_size=page_size,
    )
    warnings.extend(w_mp)
    medios_by_emp_valor, medios_by_valor = _build_medios_lookup(medios_pago)

    keys_by_target: Dict[str, list[dict[str, Any]]] = {}
    for k in unique_keys:
        target_emp = _norm_numeric_id(k.get("EmpresaID"))
        if not target_emp:
            target_emp = str(targets[0] if targets else "2")
        keys_by_target.setdefault(target_emp, []).append(k)

    def _fetch_complete_details(keys_subset: list[dict[str, Any]], target_emp: str) -> dict[tuple[str, str, str, str, str], dict]:
        acc: dict[tuple[str, str, str, str, str], dict] = {}
        pending = list(keys_subset)
        for _pass in range(3):
            if not pending:
                break
            details_t, _rid_item, rl_item = _receipts_fetch_getitem_details(
                base=base,
                getitem_path="/api/Ventas/Comprobantes/Cobros/GetItem",
                getitem_query="incluirDetalleDeMedioDePago=S",
                headers=_headers_for_target(target_emp),
                keys=pending,
                chunk_size=getitem_chunk,
            )
            if details_t:
                acc.update(details_t)
            if rl_item:
                warnings.append(f"Cobros/GetItem frenado por límite de rate al reconsultar detalle de pagos (target={target_emp}).")
                break
            pending = [k for k in pending if _receipts_normalize_getitem_key_tuple(k) not in acc]
        return acc

    detail_by_key: Dict[tuple[str, str, str, str, str], dict] = {}
    for target_emp, keys_target in keys_by_target.items():
        detail_by_key.update(_fetch_complete_details(keys_target, target_emp))

    out: Dict[tuple[str, str, str, str, str], Dict[str, Any]] = {}
    for tk, c in detail_by_key.items():
        if not isinstance(c, dict):
            continue
        empresa_id_for_medio = _comprobante_empresa_id(c)
        valor_ids = _detalle_valor_ids(c)
        medio_from_valor, _bankable = _medio_from_valor_ids(
            empresa_id=empresa_id_for_medio,
            valor_ids=valor_ids,
            medios_by_emp_valor=medios_by_emp_valor,
            medios_by_valor=medios_by_valor,
        )
        medio = str(medio_from_valor or "").strip()
        if not medio:
            medio = _medio_pago_display(
                forma_desc_catalogo="",
                forma_de_pago=_get_any(c, "formaDePago", "forma_de_pago", "FormaDePago"),
                forma_pago=_get_any(c, "formaPago", "forma_pago"),
                descripcion_forma=_get_any(c, "descripcionFormaDePago", "descripcion_forma_de_pago", "DescripcionFormaDePago"),
                denominacion_forma=_get_any(c, "denominacionFormaDePago", "denominacion_forma_de_pago", "DenominacionFormaDePago"),
                medio_pago=_get_any(c, "medioPago", "medio_pago", "MedioPago"),
                metodo_pago=_get_any(c, "metodoPago", "metodo_pago", "MetodoPago"),
                canal_cobro=_get_any(c, "canalDeCobro", "canal_de_cobro", "CanalDeCobro"),
                tipo_cobro=_get_any(c, "tipoDeCobro", "tipo_de_cobro", "TipoDeCobro"),
                origen_cobro=_get_any(c, "origenDelCobro", "origen_del_cobro", "OrigenDelCobro"),
                hint=_medio_detail_text(c),
            )
        medio = str(medio or "").strip()
        importe_bankable = _bankable_amount_from_detalle(
            c=c,
            empresa_id=empresa_id_for_medio,
            medios_by_emp_valor=medios_by_emp_valor,
            medios_by_valor=medios_by_valor,
        )
        if medio or importe_bankable is not None:
            out[tk] = {
                "medio_pago": medio,
                "importe_bankable": importe_bankable,
            }

    return out, warnings


def _to_padron_entry(obj: Dict[str, Any]) -> ExternalPadronEntry:
    cli = _normalize_cliente(obj.get("nro_cliente") or obj.get("cliente") or obj.get("cliente_id") or obj.get("clienteID"))
    cuit = _normalize_cuit(obj.get("cuit") or obj.get("numero_documento") or obj.get("numeroDeDocumento"))
    if not cli or not cuit:
        raise ExternalSchemaError("entry de padrón inválido (nro_cliente/cuit)")
    return ExternalPadronEntry(nro_cliente=cli, cuit=cuit)


def fetch_cliente_cuit_map(
    empresa_filter: str | None = None,
    cliente_ids: List[str] | None = None,
) -> Tuple[Dict[str, str], Dict[str, Any]]:
    resp = fetch_padron_payload(empresa_filter=empresa_filter, cliente_ids=cliente_ids)
    payload = _expect_obj(resp.payload, "payload")
    raw_rows = _extract_list(payload, ["entries", "padron", "clientes"])

    out: Dict[str, str] = {}
    skipped_invalid = 0
    for i, r in enumerate(raw_rows):
        try:
            entry = _to_padron_entry(_expect_obj(r, f"entries[{i}]"))
        except ExternalSchemaError:
            skipped_invalid += 1
            continue
        if entry.nro_cliente not in out:
            out[entry.nro_cliente] = entry.cuit

    if not out:
        raise ExternalSchemaError("La API de padrón no devolvió mappings cliente↔CUIT válidos")

    warnings = list(resp.warnings or [])
    if skipped_invalid > 0:
        warnings.append(f"Padrón API: se ignoraron {skipped_invalid} filas inválidas sin nro_cliente/cuit utilizable.")

    meta = {
        "api_request_id": resp.request_id,
        "external_warnings": warnings,
        "padron_size": len(out),
    }
    return out, meta
