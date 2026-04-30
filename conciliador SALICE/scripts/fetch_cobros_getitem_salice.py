#!/usr/bin/env python3
from __future__ import annotations

import argparse
import json
import sys
from datetime import date, timedelta
from pathlib import Path
from typing import Any, Dict, List
from urllib.parse import urlencode

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from src.conciliador.external.receipts_api_client import (  # noqa: E402
    _base_url,
    _build_auth_headers_for_empresa,
    _extract_paginacion,
    _extract_rows,
    _headers_base,
    _http_json,
)
from src.conciliador.external.service import (  # noqa: E402
    _comprobante_empresa_id,
    _gesi_nro_recibo_from_comprobante,
    _resolve_empresa,
)


def _clean_text(v: Any) -> str:
    return str(v or "").strip()


def _norm_numeric_id(v: Any) -> str:
    s = str(v or "").strip()
    if not s:
        return ""
    s_num = s.replace(",", ".")
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


def _key_from_row(row: Dict[str, Any]) -> dict | None:
    def _to_int(v: Any) -> int | None:
        s = str(v or "").strip()
        if not s:
            return None
        try:
            return int(float(s))
        except Exception:
            return None

    comp_id = _to_int(row.get("ComprobanteID") or row.get("comprobanteID") or row.get("comprobante_id"))
    emp_id = _to_int(
        row.get("EmpresaID")
        or row.get("empresaID")
        or row.get("empresa_id")
        or row.get("sucursalID")
        or row.get("sucursal_id")
    )
    pv_id = _to_int(row.get("PuntoDeVentaID") or row.get("puntoDeVentaID") or row.get("punto_de_venta_id"))
    numero = _to_int(row.get("Numero") or row.get("numero"))
    serie = _clean_text(row.get("Serie") or row.get("serie"))
    if comp_id is None or emp_id is None or pv_id is None or numero is None:
        return None
    return {
        "ComprobanteID": comp_id,
        "EmpresaID": emp_id,
        "Serie": serie,
        "PuntoDeVentaID": pv_id,
        "Numero": numero,
    }


def _extract_medio_fields(row: Dict[str, Any]) -> Dict[str, Any]:
    out: Dict[str, Any] = {}
    keys = [
        "formaDePagoID",
        "FormaDePagoID",
        "formaPagoID",
        "formaDePago",
        "formaPago",
        "medioPago",
        "metodoPago",
        "canalDeCobro",
        "tipoDeCobro",
        "origenDelCobro",
        "descripcionFormaDePago",
        "denominacionFormaDePago",
        "detalleDeMedioDePago",
        "detalleDeMediosDePago",
        "detalleDeValores",
    ]
    for k in keys:
        if k in row:
            out[k] = row.get(k)
    return out


def _detalle_resumen(row: Dict[str, Any]) -> List[Dict[str, Any]]:
    dv = row.get("detalleDeValores")
    if not isinstance(dv, list):
        return []
    out: List[Dict[str, Any]] = []
    for it in dv:
        if not isinstance(it, dict):
            continue
        out.append(
            {
                "itemID": it.get("itemID"),
                "cajaBancoID": it.get("cajaBancoID"),
                "valorID": it.get("valorID"),
                "tipoMovimiento": it.get("tipoMovimiento"),
                "importe": it.get("importe"),
                "estado": it.get("estado"),
                "bancoID": it.get("bancoID"),
                "descripcionBanco": it.get("descripcionBanco"),
                "codigoDeAutorizacion": it.get("codigoDeAutorizacion"),
                "emisorOBeneficiario": it.get("emisorOBeneficiario"),
            }
        )
    return out


def _empresa_filter_from_id(empresa_id: int) -> str | None:
    if int(empresa_id) == 3:
        return "SALICE"
    if int(empresa_id) == 6:
        return "ALARCON"
    return None


def _row_nro_cliente(row: Dict[str, Any]) -> str:
    for k in ("clienteID", "cliente_id", "clienteId", "ClienteID"):
        s = _clean_text(row.get(k))
        if s:
            return s
    return ""


def _is_useful(value: Any) -> bool:
    if value is None:
        return False
    if isinstance(value, list):
        return len(value) > 0
    if isinstance(value, dict):
        return len(value) > 0
    s = str(value).strip()
    if not s:
        return False
    if s in {"0", "0.0", "0,0"}:
        return False
    if s.replace("|", "").replace(" ", "") == "":
        return False
    return True


def main() -> int:
    parser = argparse.ArgumentParser(
        description="Audita Cobros/GetList + Cobros/GetItem (incluirDetalleDeMedioDePago=S) para SALICE (m5mdp)."
    )
    parser.add_argument("--days", type=int, default=7, help="Ventana de dias hacia atras (default: 7)")
    parser.add_argument("--empresa-id", type=int, default=3, help="Empresa/Sucursal objetivo (SALICE=3)")
    parser.add_argument("--page-size", type=int, default=200, help="Tamano de pagina para GetList (default: 200)")
    parser.add_argument("--max-keys", type=int, default=300, help="Maximo de comprobantes a enviar a GetItem (default: 300)")
    parser.add_argument("--chunk-size", type=int, default=60, help="Tamano de lote para GetItem (maximo: 60, default: 60)")
    parser.add_argument("--sample-limit", type=int, default=4, help="Cantidad de ejemplos en salida (default: 4)")
    parser.add_argument(
        "--include-keys",
        action="store_true",
        help="Incluir keys enviadas a GetItem (default: off, para salida liviana)",
    )
    parser.add_argument(
        "--out",
        default="cobros_getitem_salice_audit.json",
        help="Ruta del archivo JSON de salida",
    )
    args = parser.parse_args()

    base = _base_url("RECEIPTS_API")
    headers_root = _headers_base("RECEIPTS_API")
    headers = _build_auth_headers_for_empresa(
        base=base,
        headers_root=headers_root,
        empresa_id=str(args.empresa_id),
        drop_sucursal=False,
    )
    headers["sucursalID"] = str(args.empresa_id)

    fecha_hasta = date.today() - timedelta(days=1)
    fecha_desde = fecha_hasta - timedelta(days=max(int(args.days) - 1, 0))
    body = {
        "datosOperacion": {
            "FechaDesde": fecha_desde.isoformat(),
            "FechaHasta": fecha_hasta.isoformat(),
            "EmpresaID": int(args.empresa_id),
            "empresaID": int(args.empresa_id),
        },
        "datosClientes": {
            "EmpresaID": int(args.empresa_id),
        },
        "empresaID": int(args.empresa_id),
        "sucursalID": int(args.empresa_id),
    }

    getlist_path = "/api/Ventas/Comprobantes/Cobros/GetList"
    getitem_path = "/api/Ventas/Comprobantes/Cobros/GetItem"
    medios_path = "/api/Maestros/MediosDePago/GetList"
    # Mismo formato conceptual que Postman collection:
    # "query": [{"key":"...","value":"..."}]
    getitem_query_items = [
        {"key": "incluirDetalleDeMedioDePago", "value": "S"},
    ]
    query_getitem = urlencode([(q["key"], q["value"]) for q in getitem_query_items])

    # 1) GetList
    all_rows: List[dict] = []
    page = 1
    request_ids: List[str] = []
    while True:
        url = f"{base}{getlist_path}?{urlencode({'pageNumber': str(page), 'pageSize': str(args.page_size)})}"
        payload, rid = _http_json(url, method="POST", headers=headers, body=body)
        if rid:
            request_ids.append(rid)
        rows = _extract_rows(payload, "comprobantes") or []
        all_rows.extend([r for r in rows if isinstance(r, dict)])
        pag = _extract_paginacion(payload)
        if not isinstance(pag, dict):
            break
        try:
            total_pages = int(pag.get("totalPaginas") or pag.get("totalPages") or pag.get("pages") or 1)
        except Exception:
            total_pages = 1
        if page >= total_pages:
            break
        page += 1
        if page > 500:
            break

    # 2) Preparar keys para GetItem (todos los comprobantes del período)
    empresa_filter = _empresa_filter_from_id(int(args.empresa_id))

    keys: List[dict] = []
    seen: set[tuple[str, str, str, str, str]] = set()
    filtered_rows = 0
    for r in all_rows:
        # Se listan todos los comprobantes del periodo configurado.
        _ = _resolve_empresa(
            _comprobante_empresa_id(r),
            empresa_filter=empresa_filter,
            empresas_by_id=None,
            company_hint=" | ".join(
                [
                    str(r.get("_empresa_target_name") or ""),
                    str(r.get("empresa") or ""),
                    str(r.get("nombreEmpresa") or ""),
                    str(r.get("empresaNombre") or ""),
                    str(r.get("razonSocialEmpresa") or ""),
                ]
            ),
        )
        _ = _clean_text(_gesi_nro_recibo_from_comprobante(r) or "")
        _ = _row_nro_cliente(r)
        filtered_rows += 1

        k = _key_from_row(r)
        if not k:
            continue
        key_tuple = (
            str(k["ComprobanteID"]),
            str(k["EmpresaID"]),
            str(k["Serie"]),
            str(k["PuntoDeVentaID"]),
            str(k["Numero"]),
        )
        if key_tuple in seen:
            continue
        seen.add(key_tuple)
        keys.append(k)
        if len(keys) >= int(args.max_keys):
            break

    # 3) GetItem por lotes (máximo 60 por request)
    detailed_rows: List[dict] = []
    getitem_errors: List[dict] = []
    batch_size = min(max(int(args.chunk_size), 1), 60)
    for idx in range(0, len(keys), batch_size):
        batch = keys[idx : idx + batch_size]
        url = f"{base}{getitem_path}?{query_getitem}"
        try:
            payload_item, rid_item = _http_json(url, method="POST", headers=headers, body=batch)
        except Exception as e:
            getitem_errors.append({"batch_start": idx, "batch_size": len(batch), "error": str(e)})
            continue
        if rid_item:
            request_ids.append(rid_item)
        rows_item = _extract_rows(payload_item, "comprobantes")
        if rows_item is None and isinstance(payload_item, list):
            rows_item = [x for x in payload_item if isinstance(x, dict)]
        if rows_item:
            detailed_rows.extend(rows_item)

    # 4) Analisis de campos de medio
    rows_with_any_useful = 0
    medio_field_counts: Dict[str, int] = {}
    comprobantes: List[dict] = []
    for r in detailed_rows:
        mf = _extract_medio_fields(r)
        useful_keys = [k for k, v in mf.items() if _is_useful(v)]
        if useful_keys:
            rows_with_any_useful += 1
            for k in useful_keys:
                medio_field_counts[k] = int(medio_field_counts.get(k, 0) + 1)
        detalle_raw = r.get("detalleDeValores")
        comprobantes.append(
            {
                "ComprobanteID": r.get("ComprobanteID") or r.get("comprobanteID"),
                "EmpresaID": r.get("EmpresaID") or r.get("empresaID"),
                "sucursalID": r.get("sucursalID") or r.get("sucursal_id"),
                "Numero": r.get("Numero") or r.get("numero"),
                "medio_fields": mf,
                "detalleDeValores_count": len(detalle_raw) if isinstance(detalle_raw, list) else 0,
                "detalleDeValores_resumen": _detalle_resumen(r),
                "detalleDeValores_raw": detalle_raw if isinstance(detalle_raw, list) else [],
            }
        )

    # 5) Maestro de medios de pago (para mapear valorID -> descripcion/tipo)
    medios_rows: List[dict] = []
    page_mp = 1
    while True:
        url = f"{base}{medios_path}?{urlencode({'pageNumber': str(page_mp), 'pageSize': str(args.page_size)})}"
        payload_mp, rid_mp = _http_json(url, method="GET", headers=headers, body=None)
        if rid_mp:
            request_ids.append(rid_mp)
        rows_mp = _extract_rows(payload_mp, "mediosDePago") or []
        medios_rows.extend([r for r in rows_mp if isinstance(r, dict)])
        pag_mp = _extract_paginacion(payload_mp)
        if not isinstance(pag_mp, dict):
            break
        try:
            total_pages_mp = int(pag_mp.get("totalPaginas") or pag_mp.get("totalPages") or pag_mp.get("pages") or 1)
        except Exception:
            total_pages_mp = 1
        if page_mp >= total_pages_mp:
            break
        page_mp += 1
        if page_mp > 500:
            break

    valor_ids_used: set[str] = set()
    for s in comprobantes:
        for it in (s.get("detalleDeValores_raw") or []):
            if not isinstance(it, dict):
                continue
            vid = str(it.get("valorID") or "").strip()
            if vid:
                valor_ids_used.add(vid)

    medios_relevantes: List[dict] = []
    medios_by_valor: Dict[str, List[dict]] = {}
    for r in medios_rows:
        vid = _norm_numeric_id(r.get("valorID") or r.get("valorId") or r.get("valor_id"))
        if vid and vid in valor_ids_used:
            medios_relevantes.append(r)
        if vid:
            medios_by_valor.setdefault(vid, []).append(r)

    for s in comprobantes:
        vids: List[str] = []
        for it in (s.get("detalleDeValores_raw") or []):
            if not isinstance(it, dict):
                continue
            vid = _norm_numeric_id(it.get("valorID"))
            if vid and vid not in vids:
                vids.append(vid)
        matched: List[dict] = []
        seen_pairs: set[tuple[str, str]] = set()
        for vid in vids:
            for mr in medios_by_valor.get(vid, []):
                pair = (
                    _norm_numeric_id(mr.get("empresaID") or mr.get("empresaId") or mr.get("empresa_id")),
                    _norm_numeric_id(mr.get("valorID") or mr.get("valorId") or mr.get("valor_id")),
                )
                if pair in seen_pairs:
                    continue
                seen_pairs.add(pair)
                matched.append(mr)
        s["valorIDs"] = vids
        s["mediosDePago_match"] = matched
    samples = comprobantes[: max(1, int(args.sample_limit))]

    out = {
        "meta": {
            "base_url": base,
            "empresa_id_target": int(args.empresa_id),
            "fecha_desde": fecha_desde.isoformat(),
            "fecha_hasta": fecha_hasta.isoformat(),
            "getlist_path": getlist_path,
            "getitem_path": f"{getitem_path}?{query_getitem}",
            "getitem_query_items": getitem_query_items,
            "medios_path": medios_path,
            "getlist_rows_total": len(all_rows),
            "rows_in_period": filtered_rows,
            "getitem_keys_sent": len(keys),
            "getitem_mode": "batched",
            "getitem_batch_size": batch_size,
            "getitem_rows_total": len(detailed_rows),
            "getitem_errors_total": len(getitem_errors),
            "rows_with_any_useful_medio": rows_with_any_useful,
            "rows_without_useful_medio": int(len(detailed_rows) - rows_with_any_useful),
            "medio_field_useful_counts": medio_field_counts,
            "mediosDePago_total": len(medios_rows),
            "mediosDePago_relevantes_count": len(medios_relevantes),
            "sample_limit": int(args.sample_limit),
            "request_ids": request_ids,
        },
        "samples": samples,
        "comprobantes": comprobantes,
        "mediosDePago_relevantes": medios_relevantes,
        "getitem_errors": getitem_errors,
    }
    if bool(args.include_keys):
        out["keys_sent"] = keys

    out_path = Path(args.out).resolve()
    out_path.write_text(json.dumps(out, ensure_ascii=False, indent=2), encoding="utf-8")
    print(str(out_path))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
