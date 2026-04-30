from __future__ import annotations

from typing import Dict, List, Tuple, Optional
from dataclasses import replace

import datetime as dt
import os
import time
from pathlib import Path

import pandas as pd

from .excel_loader import load_bank_txns
from .utils import _normalize_text, _normalize_cliente, _normalize_cuit, _normalize_recibo
from .pdf_parser import (
    ReceiptPayment,
    parse_receipts_and_payments_from_text,
    pdf_date_range,
    detect_pdf_warnings_from_text,
    report_period_range_from_text,
    extract_pdf_text,
)
from .matcher_hungarian import match_hungarian
from .memdebug import is_mem_debug_enabled, mem_debug_recorder
from .external.service import fetch_cliente_cuit_map as fetch_cliente_cuit_map_api
from .external.service import (
    fetch_receipts_and_payments as fetch_receipts_and_payments_api,
    fetch_payment_detail_map_for_api_keys as fetch_payment_detail_map_for_api_keys_api,
)

API_RECEIPTS_MAX_DAYS = 15


def _medio_is_unknown(value: object) -> bool:
    s = str(value or "").strip().lower()
    return s in {"", "no_informado", "sin_medio_api", "no informado"}


def _medio_from_origen(origen: object) -> str:
    o = str(origen or "").strip().upper()
    if o == "MERCADOPAGO":
        return "Mercado Pago"
    if o in {"BBVA", "GALICIA"}:
        return "Transf. Bancaria"
    return ""


def _medio_applies_to_program(value: object) -> bool:
    txt = _normalize_text(value)
    if not txt:
        return False
    if any(tag in txt for tag in ("efectivo", "redondeo", "bono", "tarjeta", "cheque propio")):
        return False
    if "mercado pago" in txt or "mercadopago" in txt:
        return True
    if any(
        tag in txt
        for tag in (
            "transf",
            "transferencia",
            "banelco",
            "cbu",
            "deposito",
            "depósito",
            "boleta de deposito",
            "boleta de depósito",
            "echeq",
            "e-cheq",
            "cheque electron",
            "bco: cheque",
            "bco cheque",
            "cheques de 3ros",
            "cheque de 3ros",
            "cheque de terceros",
            "cheques terceros",
        )
    ):
        return True
    return False


def _api_detail_key_tuple(api_key: dict) -> tuple[str, str, str, str, str]:
    return (
        str(api_key.get("ComprobanteID") or ""),
        str(api_key.get("EmpresaID") or ""),
        str(api_key.get("Serie") or ""),
        str(api_key.get("PuntoDeVentaID") or ""),
        str(api_key.get("Numero") or ""),
    )


def _select_api_detail_candidate_keys(payments: List[object], txns: List[object], *, max_keys: int = 220, amount_tol: float = 50.0) -> list[dict]:
    selected: list[dict] = []
    seen: set[tuple[str, str, str, str, str]] = set()
    payments_indexed = list(enumerate(payments or []))
    for txn in txns or []:
        if not getattr(txn, "parse_ok", False) or getattr(txn, "was_preconciled", False):
            continue
        txn_candidates: list[tuple[float, int, object]] = []
        for idx, p in payments_indexed:
            api_key = getattr(p, "api_key", None)
            if not isinstance(api_key, dict):
                continue
            pd_raw = getattr(p, "fecha_pago", None)
            try:
                pd = dt.date.fromisoformat(str(pd_raw))
            except Exception:
                continue
            signed = (pd - txn.fecha).days
            if signed < -2 or signed > 10:
                continue
            diff = abs(float(getattr(txn, "importe", 0.0)) - float(getattr(p, "importe_pago", 0.0)))
            if diff > float(amount_tol):
                continue
            txn_candidates.append((diff, abs(signed), idx, p))
        txn_candidates.sort(key=lambda x: (x[0], x[1], x[2]))
        for _diff, _days, _idx, p in txn_candidates[:3]:
            api_key = getattr(p, "api_key", None)
            if not isinstance(api_key, dict):
                continue
            tk = _api_detail_key_tuple(api_key)
            if not any(tk) or tk in seen:
                continue
            seen.add(tk)
            selected.append(api_key)
            if len(selected) >= int(max_keys):
                return selected
    return selected


def _collect_detail_keys(payments: List[object]) -> list[dict]:
    """Recolecta las api_keys de pagos que todavía no tienen medio verificable."""
    all_keys: list[dict] = []
    seen: set[tuple[str, str, str, str, str]] = set()
    for p in payments or []:
        if not _medio_is_unknown(getattr(p, "medio_pago", "")):
            continue
        api_key = getattr(p, "api_key", None)
        if not isinstance(api_key, dict):
            continue
        tk = _api_detail_key_tuple(api_key)
        if not any(tk) or tk in seen:
            continue
        seen.add(tk)
        all_keys.append(api_key)
    return all_keys


def _count_payments_by_unknown_medio(payments: List[object]) -> int:
    """Cuenta pagos cuyo medio sigue siendo desconocido después del enriquecimiento."""
    return sum(1 for p in (payments or []) if _medio_is_unknown(getattr(p, "medio_pago", "")))


def _apply_detail_to_payment(
    p: object,
    detail: dict,
) -> tuple[object | None, bool, str | None]:
    """Aplica el detalle API a un pago individual.

    Devuelve (pago_enriquecido_o_None, fue_actualizado, medio_key_para_stats).
    Retorna None si el pago debe filtrarse.
    """
    medio = str(detail.get("medio_pago") or "").strip()
    importe_bankable = detail.get("importe_bankable")
    new_medio = (
        medio
        if medio and _normalize_text(medio) not in {"", "no_informado", "sin_medio_api"}
        else getattr(p, "medio_pago", "")
    )
    new_importe = (
        float(importe_bankable)
        if isinstance(importe_bankable, (int, float)) and float(importe_bankable) > 0
        else float(getattr(p, "importe_pago"))
    )
    medio_key = str(new_medio or "").strip() or "SIN_MEDIO_API"
    updated = new_medio != getattr(p, "medio_pago") or new_importe != float(getattr(p, "importe_pago"))
    if updated:
        p = replace(p, medio_pago=new_medio, importe_pago=new_importe)
    return p, updated, medio_key


def _enrich_api_payments_before_match(
    payments: List[object],
    txns: List[object],
    empresa_filter: str | None,
) -> tuple[List[object], list[str], Dict[str, object]]:
    """Enriquece los recibos API con GetItem y conserva solo medios bancarizables."""
    stats: Dict[str, object] = {
        "detail_keys_total": 0,
        "detail_keys_resolved": 0,
        "post_detail_medio_counts": {},
        "filtered_non_program_by_medio": {},
        "filtered_unknown_after_detail": 0,
        "kept_after_detail": 0,
    }

    all_keys = _collect_detail_keys(payments)
    stats["detail_keys_total"] = len(all_keys)

    warnings: list[str] = []
    if not all_keys:
        unresolved_unknown = _count_payments_by_unknown_medio(payments)
        filtered_unknown = unresolved_unknown
        if filtered_unknown > 0:
            warnings.append(
                f"Se filtraron {filtered_unknown} recibos sin medio de pago API verificable."
            )
        stats["filtered_unknown_after_detail"] = filtered_unknown
        stats["kept_after_detail"] = len(payments) - filtered_unknown
        stats["unknown_kept_after_detail"] = 0
        return [p for p in (payments or []) if not _medio_is_unknown(getattr(p, "medio_pago", "")) and _medio_applies_to_program(getattr(p, "medio_pago", ""))], warnings, stats

    detail_map, warnings = fetch_payment_detail_map_for_api_keys_api(all_keys, empresa_filter)
    stats["detail_keys_resolved"] = len(detail_map)

    if not detail_map:
        unresolved_unknown = _count_payments_by_unknown_medio(payments)
        warnings = list(warnings)
        filtered_unknown = unresolved_unknown
        if filtered_unknown > 0:
            warnings.append(
                f"Se filtraron {filtered_unknown} recibos sin medio de pago API verificable."
            )
        stats["filtered_unknown_after_detail"] = filtered_unknown
        stats["kept_after_detail"] = len(payments) - filtered_unknown
        stats["unknown_kept_after_detail"] = 0
        return [p for p in (payments or []) if not _medio_is_unknown(getattr(p, "medio_pago", "")) and _medio_applies_to_program(getattr(p, "medio_pago", ""))], warnings, stats

    # 4. Enriquecer / filtrar con el detalle disponible
    enriched: list[object] = []
    updated = 0
    filtered_unknown = 0
    filtered_non_program = 0
    filtered_non_program_by_medio: Dict[str, int] = {}
    post_detail_medio_counts: Dict[str, int] = {}

    for p in payments or []:
        api_key = getattr(p, "api_key", None)

        if not isinstance(api_key, dict):
            if _medio_is_unknown(getattr(p, "medio_pago", "")):
                filtered_unknown += 1
                continue
            if not _medio_applies_to_program(getattr(p, "medio_pago", "")):
                medio_key = str(getattr(p, "medio_pago", "") or "").strip() or "SIN_MEDIO_API"
                filtered_non_program += 1
                filtered_non_program_by_medio[medio_key] = int(filtered_non_program_by_medio.get(medio_key, 0) + 1)
                continue
            enriched.append(p)
            continue

        detail = detail_map.get(_api_detail_key_tuple(api_key))
        if not isinstance(detail, dict):
            if _medio_is_unknown(getattr(p, "medio_pago", "")):
                filtered_unknown += 1
                continue
            if not _medio_applies_to_program(getattr(p, "medio_pago", "")):
                medio_key = str(getattr(p, "medio_pago", "") or "").strip() or "SIN_MEDIO_API"
                filtered_non_program += 1
                filtered_non_program_by_medio[medio_key] = int(filtered_non_program_by_medio.get(medio_key, 0) + 1)
                continue
            enriched.append(p)
            continue

        result_p, was_updated, medio_key = _apply_detail_to_payment(p, detail)
        post_detail_medio_counts[medio_key] = int(post_detail_medio_counts.get(medio_key, 0) + 1)
        if result_p is None:
            continue
        if _medio_is_unknown(medio_key):
            filtered_unknown += 1
            continue
        if not _medio_applies_to_program(medio_key):
            filtered_non_program += 1
            filtered_non_program_by_medio[medio_key] = int(filtered_non_program_by_medio.get(medio_key, 0) + 1)
            continue
        if was_updated:
            updated += 1
        enriched.append(result_p)

    warnings = list(warnings)

    if updated > 0:
        warnings.append(f"Detalle API aplicado a {updated} recibos para recalcular medio/importe bancarizable.")
    if filtered_non_program > 0:
        warnings.append(f"Se filtraron {filtered_non_program} recibos con medios de pago que no terminan como ingreso bancario o de Mercado Pago.")
    if filtered_unknown > 0:
        warnings.append(f"Se filtraron {filtered_unknown} recibos sin medio de pago API verificable.")

    stats["post_detail_medio_counts"] = post_detail_medio_counts
    stats["filtered_non_program_by_medio"] = filtered_non_program_by_medio
    stats["non_program_kept_by_medio"] = {}
    stats["filtered_unknown_after_detail"] = filtered_unknown
    stats["unknown_kept_after_detail"] = 0
    stats["kept_after_detail"] = len(enriched)
    return enriched, warnings, stats


def _enrich_api_result_medios(result: Dict[str, List[dict]], payments: List[object], empresa_filter: str | None) -> list[str]:
    medio_by_recibo: Dict[str, str] = {}
    for p in payments or []:
        nro = _normalize_recibo(getattr(p, "nro_recibo", None))
        medio = str(getattr(p, "medio_pago", "") or "").strip()
        if nro and medio and not _medio_is_unknown(medio):
            medio_by_recibo[nro] = medio

    touched = 0
    for bucket in ("validados", "dudosos", "no_encontrados"):
        for row in result.get(bucket) or []:
            if not _medio_is_unknown(row.get("Medio de pago")):
                continue
            nro = _normalize_recibo(row.get("Nro recibo"))
            medio = medio_by_recibo.get(str(nro or ""))
            if medio:
                row["Medio de pago"] = medio
                touched += 1
    warnings: list[str] = []
    if touched > 0:
        warnings.append(f"Medio de pago API completado en {touched} filas del resultado.")
    return warnings


def _payment_identity_key(p: object) -> tuple:
    api_key = getattr(p, "api_key", None)
    if isinstance(api_key, dict):
        tk = _api_detail_key_tuple(api_key)
        if any(tk):
            return ("api",) + tk
    return (
        "basic",
        str(_normalize_recibo(getattr(p, "nro_recibo", None)) or ""),
        str(_normalize_cliente(getattr(p, "nro_cliente", None)) or ""),
        str(getattr(p, "fecha_pago", "") or ""),
        round(float(getattr(p, "importe_pago", 0.0) or 0.0), 2),
        _normalize_text(getattr(p, "medio_pago", "") or ""),
    )


def _append_unique_payments(base: List[object], extra: List[object]) -> List[object]:
    seen = {_payment_identity_key(p) for p in (base or [])}
    out = list(base or [])
    for p in extra or []:
        k = _payment_identity_key(p)
        if k in seen:
            continue
        seen.add(k)
        out.append(p)
    return out


def _fetch_api_payments_for_date_range(
    *,
    empresa_filter: str | None,
    start_date: dt.date,
    end_date: dt.date,
) -> tuple[List[ReceiptPayment], Dict[str, object]]:
    days = max((end_date - start_date).days + 1, 1)
    try:
        return fetch_receipts_and_payments_api(
            days,
            empresa_filter,
            start_date=start_date,
            end_date=end_date,
        )
    except TypeError as e:
        if "start_date" not in str(e) and "end_date" not in str(e):
            raise
        return fetch_receipts_and_payments_api(days, empresa_filter, end_date=end_date)


def _collect_displaced_preconciled_receipts(result: Dict[str, List[dict]], txns_by_id: Dict[str, object]) -> tuple[set[str], list[object]]:
    displaced_receipts: set[str] = set()
    displaced_txns: list[object] = []
    for bucket in ("validados", "dudosos"):
        for row in (result.get(bucket) or []):
            if bucket == "dudosos" and str(row.get("Tipo fila") or "").strip().upper() != "PRINCIPAL":
                continue
            txn_id = str(row.get("__txn_id") or "").strip()
            if not txn_id:
                continue
            txn = txns_by_id.get(txn_id)
            if txn is None or not getattr(txn, "was_preconciled", False):
                continue
            prev_recibo = _normalize_recibo(getattr(txn, "preconciled_recibo", None))
            curr_recibo = _normalize_recibo(row.get("Nro recibo"))
            if not prev_recibo or not curr_recibo or prev_recibo == curr_recibo:
                continue
            displaced_receipts.add(prev_recibo)
            displaced_txns.append(txn)
    return displaced_receipts, displaced_txns


def _synthesize_preconciled_payment_from_txn(
    txn: object,
) -> tuple[ReceiptPayment | None, str | None]:
    """Sintetiza un ReceiptPayment desde los campos preconciliados de una transacción.

    Devuelve (pago, warning_o_None).
    - Si no hay número de recibo: devuelve (None, None).
    - Si la fecha del recibo preconciliado es desconocida, usa la fecha del ingreso
      como aproximación y devuelve un warning explícito para trazabilidad.
    """
    nro_recibo = _normalize_recibo(getattr(txn, "preconciled_recibo", None))
    if not nro_recibo:
        return None, None
    medio = str(getattr(txn, "preconciled_medio_pago", "") or "").strip()
    if _medio_is_unknown(medio):
        medio = _medio_from_origen(getattr(txn, "origen", ""))
    fecha_pago = _coerce_iso_date_str(getattr(txn, "preconciled_fecha_recibo", None))
    warning: str | None = None
    if not fecha_pago:
        # Usamos la fecha del ingreso como aproximación y lo dejamos explícito
        fecha_ingreso: dt.date | None = getattr(txn, "fecha", None)
        if fecha_ingreso is None:
            return None, (
                f"Recibo {nro_recibo}: no se pudo reconstruir el recibo preconciliado "
                "porque no tiene fecha de recibo ni fecha de ingreso."
            )
        fecha_pago = fecha_ingreso.isoformat()
        warning = (
            f"Recibo {nro_recibo}: fecha de recibo desconocida en el record; "
            f"se usó la fecha del ingreso ({fecha_pago}) como aproximación para el matching."
        )
    importe_pago_raw = getattr(txn, "preconciled_importe_recibo", None)
    try:
        importe_pago = float(importe_pago_raw) if importe_pago_raw is not None else float(getattr(txn, "importe", 0.0) or 0.0)
    except Exception:
        importe_pago = float(getattr(txn, "importe", 0.0) or 0.0)
    payment = ReceiptPayment(
        empresa="GBA",
        nro_recibo=nro_recibo,
        nro_cliente=str(_normalize_cliente(getattr(txn, "preconciled_nro_cliente", None)) or ""),
        cliente_nombre=str(getattr(txn, "preconciled_cliente_nombre", "") or "").strip() or None,
        medio_pago=medio or "SIN_MEDIO_API",
        fecha_pago=fecha_pago,
        importe_pago=importe_pago,
        detalle_pago="RECUPERADO_DEL_RECORD_PRECONCILIADO",
    )
    return payment, warning


def _coerce_iso_date_str(value: object) -> str | None:
    if value is None:
        return None
    if isinstance(value, dt.date):
        return value.isoformat()
    s = str(value).strip()
    if not s:
        return None
    try:
        return dt.date.fromisoformat(s[:10]).isoformat()
    except Exception:
        pass
    try:
        parsed = pd.to_datetime(s, errors="coerce", dayfirst=True)
        if pd.notna(parsed):
            return parsed.date().isoformat()
    except Exception:
        pass
    return None


def _find_padron_column(df: pd.DataFrame, candidates: list[str]) -> Optional[str]:
    wanted = {_normalize_text(c).replace(" ", "") for c in candidates}
    for c in df.columns:
        norm = _normalize_text(c).replace(" ", "")
        if norm in wanted:
            return str(c)
    return None


def _discover_padron_path(excel_path: str) -> Optional[str]:
    roots: list[Path] = []
    try:
        roots.append(Path(excel_path).resolve().parent)
    except Exception:
        pass
    roots.append(Path.cwd())
    roots.append(Path(__file__).resolve().parents[2])

    seen: set[str] = set()
    for root in roots:
        rp = str(root)
        if rp in seen or not root.exists():
            continue
        seen.add(rp)
        for p in root.glob("*.xlsx"):
            n = _normalize_text(p.name)
            if "padron" in n and "basico" in n and "mdp" in n:
                return str(p)
    return None


def _load_cliente_cuit_map(padron_path: str) -> Dict[str, str]:
    if not padron_path or not os.path.exists(padron_path):
        return {}

    df = pd.read_excel(padron_path, dtype=str)
    cliente_col = _find_padron_column(df, ["ClienteID", "Cliente", "Nro cliente", "Nro_cliente"])
    cuit_col = _find_padron_column(df, ["NumeroDeDocumento", "Número de documento", "CUIT", "Número Documento"])
    if not cliente_col or not cuit_col:
        return {}

    out: Dict[str, str] = {}
    for _, row in df.iterrows():
        cli = _normalize_cliente(row.get(cliente_col))
        cuit = _normalize_cuit(row.get(cuit_col))
        if cli and cuit and cli not in out:
            out[cli] = cuit
    return out


def _summarize_excel_errors(txns, limit: int = 20) -> Tuple[int, List[dict]]:
    bad = [t for t in txns if not t.parse_ok]
    sample = [
        {
            "origen": t.origen,
            "row_index": t.row_index,
            "parse_error": t.parse_error,
        }
        for t in bad[:limit]
    ]
    return len(bad), sample


def compare_excel_pdf(
    excel_path: str,
    pdf_path: str,
    *,
    margin_days: int = 5,
    tolerance_days_suspect: int = 7,
    max_options: int = 4,
    # V3.5: multiplicador de días depende del signo (recibo vs banco)
    day_weight_bank_before: float = 20.0,
    day_weight_bank_after: float = 35.0,
    valid_max_peso: float = 260.0,
    dudoso_max_peso: float = 3500.0,
    mp_mismatch_penalty: float = 35.0,
    preconciled_penalty: float = 150.0,
    penalty_salice_to_galicia: float = 45.0,
    penalty_alarcon_to_bbva: float = 45.0,
    cliente_cuit_mismatch_penalty: float = 0.0,
    alternatives_cost_delta: float = 35.0,
    padron_cliente_cuit_path: str | None = None,
    show_peso: bool = False,
    show_cuit: bool = False,
    mem_debug: Optional[bool] = None,
    stage2_candidate_top_k: int = 120,
    receipts_source: str = "pdf",
    api_receipts_days: int = 4,
    api_empresa_filter: str | None = None,
    # V3.0: IA eliminada (se mantiene el parámetro solo por compatibilidad de API).
    enable_ai: bool = False,
) -> Dict[str, List[dict]]:
    # Wrapper para compatibilidad: un solo PDF.
    return compare_excel_pdfs(
        excel_path,
        [(pdf_path, None)],
        margin_days=margin_days,
        tolerance_days_suspect=tolerance_days_suspect,
        day_weight_bank_before=day_weight_bank_before,
        day_weight_bank_after=day_weight_bank_after,
        valid_max_peso=valid_max_peso,
        dudoso_max_peso=dudoso_max_peso,
        mp_mismatch_penalty=mp_mismatch_penalty,
        preconciled_penalty=preconciled_penalty,
        penalty_salice_to_galicia=penalty_salice_to_galicia,
        penalty_alarcon_to_bbva=penalty_alarcon_to_bbva,
        cliente_cuit_mismatch_penalty=cliente_cuit_mismatch_penalty,
        alternatives_cost_delta=alternatives_cost_delta,
        padron_cliente_cuit_path=padron_cliente_cuit_path,
        show_peso=show_peso,
        show_cuit=show_cuit,
        mem_debug=mem_debug,
        stage2_candidate_top_k=stage2_candidate_top_k,
        receipts_source=receipts_source,
        api_receipts_days=api_receipts_days,
        api_empresa_filter=api_empresa_filter,
        max_options=max_options,
        enable_ai=enable_ai,
    )


def compare_excel_pdfs(
    excel_path: str | List[str],
    pdfs: List[Tuple[str, Optional[str]]],
    *,
    margin_days: int = 5,
    tolerance_days_suspect: int = 7,
    max_options: int = 4,
    # V3.5: multiplicador de días depende del signo (recibo vs banco)
    day_weight_bank_before: float = 20.0,
    day_weight_bank_after: float = 35.0,
    valid_max_peso: float = 260.0,
    dudoso_max_peso: float = 3500.0,
    mp_mismatch_penalty: float = 35.0,
    preconciled_penalty: float = 150.0,
    # V3.1: penalización por banco cruzado (empresa ↔ banco)
    penalty_salice_to_galicia: float = 45.0,
    penalty_alarcon_to_bbva: float = 45.0,
    cliente_cuit_mismatch_penalty: float = 0.0,
    alternatives_cost_delta: float = 35.0,
    padron_cliente_cuit_path: str | None = None,
    show_peso: bool = False,
    show_cuit: bool = False,
    mem_debug: Optional[bool] = None,
    stage2_candidate_top_k: int = 120,
    receipts_source: str = "pdf",
    api_receipts_days: int = 4,
    api_empresa_filter: str | None = None,
    api_start_date_override: str | None = None,
    api_end_date_override: str | None = None,
    force_validations: Optional[List[dict]] = None,
    drop_dudosos: Optional[List[dict]] = None,
    excel_record_map: Optional[Dict[str, str]] = None,
    # V3.0: IA eliminada (se mantienen parámetros solo por compatibilidad).
    enable_ai: bool = False,
) -> Dict[str, List[dict]]:
    """Concilia un Excel contra recibos de PDF (legacy) o API (v5).

    receipts_source:
      - "pdf": usa `pdfs` como entrada (modo legacy).
      - "api": ignora PDFs y trae recibos por API.
    """
    t0 = time.perf_counter()
    mem_dbg = is_mem_debug_enabled(mem_debug)
    mem_stages, mem_mark = mem_debug_recorder(mem_dbg)
    mem_mark("start")

    source = str(receipts_source or "pdf").strip().lower()
    if source not in {"pdf", "api"}:
        raise ValueError("receipts_source inválido. Usar 'pdf' o 'api'.")
    effective_api_empresa_filter = (
        (str(api_empresa_filter or "").strip() or "GBA")
        if source == "api"
        else api_empresa_filter
    )

    include_empresa = False
    enable_banco_sin_recibo = True
    receipts_all = []
    payments_all = []
    pdf_warnings: List[str] = []
    external_warnings: List[str] = []
    api_request_ids: List[str] = []
    api_payments_by_empresa: Dict[str, int] = {}
    api_medio_bancarizable_stats: Dict[str, int] = {}
    api_post_detail_stats: Dict[str, object] = {}
    api_targets_used: List[str] = []
    api_count_by_target: Dict[str, int] = {}
    api_fecha_desde: str | None = None
    api_fecha_hasta: str | None = None
    rmins: List[str] = []
    rmaxs: List[str] = []

    effective_api_days = int(api_receipts_days)
    if source == "api":
        if effective_api_days > API_RECEIPTS_MAX_DAYS:
            effective_api_days = API_RECEIPTS_MAX_DAYS
            external_warnings.append(
                f"api_receipts_days limitado a {API_RECEIPTS_MAX_DAYS} días."
            )
        if effective_api_days < 1:
            effective_api_days = 1

    # Excel
    excel_paths = [excel_path] if isinstance(excel_path, str) else [str(p) for p in (excel_path or [])]
    if not excel_paths:
        raise ValueError("No se informaron Excels de movimientos.")
    txns = []
    for p in excel_paths:
        record_key = (excel_record_map or {}).get(str(p))
        if record_key is None:
            txns.extend(load_bank_txns(p))
        else:
            txns.extend(load_bank_txns(p, record_key=record_key))
    mem_mark("excel_loaded", {"txns_total": len(txns), "txns_parse_ok": sum(1 for t in txns if t.parse_ok)})

    bank_txn_dates = [t.fecha for t in txns if t.parse_ok]
    api_start_date: dt.date | None = None
    api_end_date: dt.date | None = None
    if source == "api":
        if api_end_date_override:
            try:
                api_end_date = dt.date.fromisoformat(str(api_end_date_override))
            except Exception as e:
                raise ValueError(f"api_end_date_override inválida: {api_end_date_override}") from e
        elif bank_txn_dates:
            api_end_date = max(bank_txn_dates)
        if api_start_date_override:
            try:
                start_override_date = dt.date.fromisoformat(str(api_start_date_override))
                api_start_date = start_override_date
            except Exception as e:
                raise ValueError(f"api_start_date_override inválida: {api_start_date_override}") from e
        elif api_end_date is not None:
            api_start_date = api_end_date - dt.timedelta(days=max(int(effective_api_days), 0))
        elif bank_txn_dates:
            newest_bank_date = max(bank_txn_dates)
            api_start_date = newest_bank_date - dt.timedelta(days=max(int(effective_api_days), 0))
            api_end_date = newest_bank_date
        else:
            api_end_date = dt.date.today() - dt.timedelta(days=1)
            api_start_date = api_end_date - dt.timedelta(days=max(int(effective_api_days), 0))
            external_warnings.append(
                "No se encontraron fechas válidas en los extractos bancarios; la API usó un rango relativo a ayer."
            )

    if source == "pdf":
        if not pdfs:
            raise ValueError("Se requiere al menos 1 PDF")
        if len(pdfs) > 2:
            raise ValueError("Máximo 2 PDFs")
        include_empresa = (len(pdfs) == 2)
        enable_banco_sin_recibo = (len(pdfs) == 2)

        for pdf_path, empresa_override in pdfs:
            # Leemos cada PDF solo una vez para reducir picos de memoria y CPU.
            text = extract_pdf_text(pdf_path)
            receipts, payments = parse_receipts_and_payments_from_text(text, empresa_override=empresa_override)
            receipts_all.extend(receipts)
            payments_all.extend(payments)
            w = detect_pdf_warnings_from_text(text[:150_000])
            if w:
                pdf_warnings.extend(w)
            rmin, rmax = report_period_range_from_text(text[:60_000])
            if rmin:
                rmins.append(rmin)
            if rmax:
                rmaxs.append(rmax)
        mem_mark("pdf_parsed", {"payments_count": len(payments_all), "receipts_count": len(receipts_all)})
    else:
        try:
            api_payments, api_meta = fetch_receipts_and_payments_api(
                effective_api_days,
                effective_api_empresa_filter,
                start_date=api_start_date,
                end_date=api_end_date,
            )
        except TypeError as e:
            # Compatibilidad con tests/mocks viejos que todavía no aceptan
            # overrides explícitos de rango.
            if "start_date" not in str(e) and "end_date" not in str(e):
                raise
            api_payments, api_meta = fetch_receipts_and_payments_api(
                effective_api_days,
                effective_api_empresa_filter,
                end_date=api_end_date,
            )
        payments_all.extend(api_payments)
        include_empresa = False
        enable_banco_sin_recibo = True
        rid = str(api_meta.get("api_request_id") or "").strip()
        if rid:
            api_request_ids.append(rid)
        external_warnings.extend([str(x) for x in (api_meta.get("external_warnings") or []) if str(x).strip()])
        if isinstance(api_meta.get("payments_by_empresa"), dict):
            api_payments_by_empresa = {str(k): int(v) for k, v in api_meta.get("payments_by_empresa").items()}
        if isinstance(api_meta.get("medio_bancarizable_stats"), dict):
            api_medio_bancarizable_stats = {str(k): int(v) for k, v in api_meta.get("medio_bancarizable_stats").items()}
        if isinstance(api_meta.get("api_empresa_targets_used"), list):
            api_targets_used = [str(x) for x in api_meta.get("api_empresa_targets_used")]
        if isinstance(api_meta.get("api_comprobantes_count_by_target"), dict):
            api_count_by_target = {str(k): int(v) for k, v in api_meta.get("api_comprobantes_count_by_target").items()}
        api_fecha_desde = str(api_meta.get("api_fecha_desde") or "") or None
        api_fecha_hasta = str(api_meta.get("api_fecha_hasta") or "") or None
        mem_mark("api_receipts_parsed", {"payments_count": len(payments_all), "window_days": effective_api_days})

    payments_filtered_old = 0
    payments_filtered_after_bank_max = 0
    payments_filtered_before_bank_min = 0
    payments_filtered_preconciled = 0
    # Caché de todos los pagos API antes de filtrar preconciliados (usado más abajo para
    # recuperar recibos desplazados sin necesidad de una segunda llamada a la API).
    _api_payments_full_cache: List[object] = []
    if source == "api":
        upper = api_end_date or (dt.date.today() - dt.timedelta(days=1))
        cutoff = api_start_date or (upper - dt.timedelta(days=max(effective_api_days - 1, 0)))
        keep_recent = []
        old_outside = 0
        for p in payments_all:
            try:
                pd = dt.date.fromisoformat(str(p.fecha_pago))
            except Exception:
                # Si viene una fecha inválida, ya fue validada aguas arriba.
                pd = dt.date.min
            if pd > upper:
                payments_filtered_after_bank_max += 1
                continue
            if pd < cutoff:
                old_outside += 1
                payments_filtered_before_bank_min += 1
                continue
            keep_recent.append(p)
        payments_filtered_old = int(old_outside)
        payments_all = keep_recent
        if payments_filtered_old > 0:
            external_warnings.append(
                f"Se filtraron {payments_filtered_old} recibos anteriores al rango derivado de los extractos ({cutoff.isoformat()} a {upper.isoformat()})."
            )
        if payments_filtered_after_bank_max > 0:
            external_warnings.append(
                f"Se filtraron {payments_filtered_after_bank_max} recibos posteriores a la última fecha bancaria detectada ({upper.isoformat()})."
            )

        conciliated_receipts = {
            nr
            for t in txns
            if t.parse_ok and t.was_preconciled
            for nr in [_normalize_recibo(t.preconciled_recibo)]
            if nr
        }
        if conciliated_receipts:
            # Guardamos el estado ANTES de filtrar preconciliados: si un recibo
            # preconciliado es desplazado más adelante, lo encontramos en este caché
            # sin tener que hacer otra llamada a la API.
            _api_payments_full_cache = list(payments_all)
            keep_not_conciliated = []
            for p in payments_all:
                nro = _normalize_recibo(p.nro_recibo)
                if nro and nro in conciliated_receipts:
                    continue
                keep_not_conciliated.append(p)
            payments_filtered_preconciled = max(0, len(payments_all) - len(keep_not_conciliated))
            payments_all = keep_not_conciliated
            if payments_filtered_preconciled > 0:
                external_warnings.append(
                    f"Se filtraron {payments_filtered_preconciled} recibos ya conciliados en el record consolidado."
                )
    if source == "api" and payments_all:
        payments_all, medio_filter_warnings, api_post_detail_stats = _enrich_api_payments_before_match(
            payments_all,
            txns,
            effective_api_empresa_filter,
        )
        if medio_filter_warnings:
            external_warnings.extend([str(x) for x in medio_filter_warnings if str(x).strip()])

    padron_warnings: List[str] = []
    cliente_to_cuit_map: Dict[str, str] = {}
    resolved_padron_path: str | None = None
    if source == "api":
        padron_target_clientes = sorted(
            {
                str(p.nro_cliente).strip()
                for p in payments_all
                if str(getattr(p, "nro_cliente", "") or "").strip()
            }
        )
        padron_map, padron_meta = fetch_cliente_cuit_map_api(
            effective_api_empresa_filter,
            cliente_ids=padron_target_clientes,
        )
        cliente_to_cuit_map = padron_map
        rid = str(padron_meta.get("api_request_id") or "").strip()
        if rid:
            api_request_ids.append(rid)
        external_warnings.extend([str(x) for x in (padron_meta.get("external_warnings") or []) if str(x).strip()])
    else:
        padron_seed_path = excel_paths[0]
        resolved_padron_path = (padron_cliente_cuit_path or "").strip() or _discover_padron_path(padron_seed_path)
        if resolved_padron_path:
            try:
                cliente_to_cuit_map = _load_cliente_cuit_map(resolved_padron_path)
                if not cliente_to_cuit_map:
                    padron_warnings.append("No se encontraron mappings válidos cliente↔CUIT en el padrón.")
            except Exception as e:
                padron_warnings.append(f"No se pudo leer el padrón cliente↔CUIT: {e}")
        else:
            padron_warnings.append("No se encontró el archivo de padrón básico MDP para validar cliente↔CUIT.")

    effective_mp_mismatch_penalty = 0.0 if source == "api" else float(mp_mismatch_penalty)
    effective_penalty_salice_to_galicia = 0.0 if source == "api" else float(penalty_salice_to_galicia)
    effective_penalty_alarcon_to_bbva = 0.0 if source == "api" else float(penalty_alarcon_to_bbva)
    banco_sin_recibo_grace_days = 10 if source == "api" else 0
    recibo_sin_banco_grace_days = 2 if source == "api" else 0

    if source == "api" and payments_all and any(t.parse_ok and t.was_preconciled for t in txns):
        prelim = match_hungarian(
            txns,
            payments_all,
            margin_days=margin_days,
            tolerance_days_suspect=tolerance_days_suspect,
            day_weight_bank_before=day_weight_bank_before,
            day_weight_bank_after=day_weight_bank_after,
            valid_max_peso=valid_max_peso,
            dudoso_max_peso=dudoso_max_peso,
            mp_mismatch_penalty=effective_mp_mismatch_penalty,
            preconciled_penalty=preconciled_penalty,
            penalty_salice_to_galicia=effective_penalty_salice_to_galicia,
            penalty_alarcon_to_bbva=effective_penalty_alarcon_to_bbva,
            cliente_cuit_mismatch_penalty=cliente_cuit_mismatch_penalty,
            cliente_to_cuit_map=cliente_to_cuit_map,
            alternatives_cost_delta=alternatives_cost_delta,
            max_alternatives=max(0, int(max_options) - 1),
            report_date_min=(api_start_date.isoformat() if api_start_date else None),
            report_date_max=(api_end_date.isoformat() if api_end_date else None),
            include_empresa=include_empresa,
            enable_banco_sin_recibo=enable_banco_sin_recibo,
            banco_sin_recibo_grace_days=banco_sin_recibo_grace_days,
            recibo_sin_banco_grace_days=recibo_sin_banco_grace_days,
            stage2_candidate_top_k=stage2_candidate_top_k,
            exclude_preconciled_txns=False,
            mem_debug=False,
            validated_allow_all_receipts=False,
            non_bankable_receipt_cost_multiplier=1.0,
            suspects_and_no_bankable_only=False,
            no_encontrados_bankable_only=False,
        )
        txns_by_id = {str(t.txn_id): t for t in txns if t.parse_ok}
        displaced_receipts, displaced_txns = _collect_displaced_preconciled_receipts(prelim, txns_by_id)
        current_receipts = {
            nr
            for nr in (_normalize_recibo(getattr(p, "nro_recibo", None)) for p in payments_all)
            if nr
        }
        displaced_receipts = {nr for nr in displaced_receipts if nr and nr not in current_receipts}
        if displaced_receipts and displaced_txns:
            before_count = len(payments_all)

            # 1. Primero intentamos recuperar del caché del payload API original
            #    (gratis, sin llamada extra a la API).
            recovered_from_cache = [
                p
                for p in _api_payments_full_cache
                if (_normalize_recibo(getattr(p, "nro_recibo", None)) or "") in displaced_receipts
            ]
            payments_all = _append_unique_payments(payments_all, recovered_from_cache)
            recovered_receipts = {
                nr
                for nr in (_normalize_recibo(getattr(p, "nro_recibo", None)) for p in payments_all)
                if nr and nr in displaced_receipts
            }
            missing_after_cache = set(displaced_receipts) - recovered_receipts

            # 2. Solo consultamos la API por los recibos que no estaban en el caché
            #    (fecha fuera de la ventana original o nunca descargados).
            recovered_from_api: list[object] = []
            if missing_after_cache:
                fetch_start = min(t.fecha for t in displaced_txns) - dt.timedelta(days=2)
                fetch_end = max(t.fecha for t in displaced_txns) + dt.timedelta(days=10)
                extra_payments, extra_meta = _fetch_api_payments_for_date_range(
                    empresa_filter=effective_api_empresa_filter,
                    start_date=fetch_start,
                    end_date=fetch_end,
                )
                external_warnings.extend([str(x) for x in (extra_meta.get("external_warnings") or []) if str(x).strip()])
                recovered_from_api = [
                    p
                    for p in (extra_payments or [])
                    if (_normalize_recibo(getattr(p, "nro_recibo", None)) or "") in missing_after_cache
                ]
                payments_all = _append_unique_payments(payments_all, recovered_from_api)
                recovered_receipts = {
                    nr
                    for nr in (_normalize_recibo(getattr(p, "nro_recibo", None)) for p in payments_all)
                    if nr and nr in displaced_receipts
                }

            # 3. Si todavía faltan recibos, los sintetizamos desde los campos del record.
            synthetic_payments: list[ReceiptPayment] = []
            missing_receipts = set(displaced_receipts) - recovered_receipts
            if missing_receipts:
                seen_synth: set[str] = set()
                for txn in displaced_txns:
                    prev_rec = _normalize_recibo(getattr(txn, "preconciled_recibo", None))
                    if not prev_rec or prev_rec not in missing_receipts or prev_rec in seen_synth:
                        continue
                    synth, synth_warning = _synthesize_preconciled_payment_from_txn(txn)
                    if synth is None:
                        continue
                    if synth_warning:
                        external_warnings.append(synth_warning)
                    synthetic_payments.append(synth)
                    seen_synth.add(prev_rec)
                payments_all = _append_unique_payments(payments_all, synthetic_payments)
                recovered_receipts = {
                    nr
                    for nr in (_normalize_recibo(getattr(p, "nro_recibo", None)) for p in payments_all)
                    if nr and nr in displaced_receipts
                }
                missing_receipts = set(displaced_receipts) - recovered_receipts

            external_warnings.append(
                f"Se reabrieron {len(displaced_receipts)} recibos ya conciliados del record porque un ingreso preconciliado tuvo un candidato nuevo mejor."
            )
            added_from_api = len(recovered_from_api)
            added_from_cache = len(recovered_from_cache)
            if added_from_cache > 0:
                external_warnings.append(
                    f"Se recuperaron {added_from_cache} recibos preconciliados desplazados desde la respuesta API original (sin llamada extra)."
                )
            if added_from_api > 0:
                external_warnings.append(
                    f"Se recuperaron {added_from_api} pagos API fuera del rango manual porque estaban conciliados en ingresos del record reabiertos."
                )
            if synthetic_payments:
                external_warnings.append(
                    f"Se reconstruyeron {len(synthetic_payments)} recibos preconciliados desde el record actualizado porque la API no los devolvió en la ventana ampliada."
                )
            if missing_receipts:
                external_warnings.append(
                    f"No se pudieron recuperar {len(missing_receipts)} recibos preconciliados desplazados."
                )

    dmin, dmax = pdf_date_range(payments_all)
    if source == "api":
        rmin_global = (api_start_date.isoformat() if api_start_date else dmin)
        rmax_global = (api_end_date.isoformat() if api_end_date else dmax)
    else:
        rmin_global = min(rmins) if rmins else dmin
        rmax_global = max(rmaxs) if rmaxs else dmax

    # Matching
    # Matching (nota: el rango efectivo de Excel se define por el PDF y el margen,
    # pero el matching "dudoso" usa una ventana más amplia: tolerance_days_suspect)
    res = match_hungarian(
        txns,
        payments_all,
        margin_days=margin_days,
        tolerance_days_suspect=tolerance_days_suspect,
        day_weight_bank_before=day_weight_bank_before,
        day_weight_bank_after=day_weight_bank_after,
        valid_max_peso=valid_max_peso,
        dudoso_max_peso=dudoso_max_peso,
        mp_mismatch_penalty=effective_mp_mismatch_penalty,
        preconciled_penalty=preconciled_penalty,
        penalty_salice_to_galicia=effective_penalty_salice_to_galicia,
        penalty_alarcon_to_bbva=effective_penalty_alarcon_to_bbva,
        cliente_cuit_mismatch_penalty=cliente_cuit_mismatch_penalty,
        cliente_to_cuit_map=cliente_to_cuit_map,
        alternatives_cost_delta=alternatives_cost_delta,
        max_alternatives=max(0, int(max_options) - 1),
        report_date_min=rmin_global,
        report_date_max=rmax_global,
        include_empresa=include_empresa,
        enable_banco_sin_recibo=enable_banco_sin_recibo,
        banco_sin_recibo_grace_days=banco_sin_recibo_grace_days,
        recibo_sin_banco_grace_days=recibo_sin_banco_grace_days,
        stage2_candidate_top_k=stage2_candidate_top_k,
        exclude_preconciled_txns=False,
        mem_debug=mem_dbg,
        validated_allow_all_receipts=False,
        non_bankable_receipt_cost_multiplier=1.0,
        suspects_and_no_bankable_only=False,
        no_encontrados_bankable_only=False,
    )
    if isinstance(res.get("no_encontrados"), list):
        res["no_encontrados"] = [
            row
            for row in (res.get("no_encontrados") or [])
            if str(row.get("Tipo no encontrado", "")).upper() == "RECIBO_SIN_BANCO"
        ]
    mem_mark("matched", {
        "validados": len(res.get("validados") or []),
        "dudosos": len(res.get("dudosos") or []),
        "no_encontrados": len(res.get("no_encontrados") or []),
    })
    if source == "api":
        medio_warnings = _enrich_api_result_medios(res, payments_all, effective_api_empresa_filter)
        if medio_warnings:
            external_warnings.extend([str(x) for x in medio_warnings if str(x).strip()])

    # -----------------
    # V3.0: IA eliminada. Dejamos trazabilidad en meta para la UI.
    # -----------------
    res.setdefault("meta", {})
    res["meta"]["ai_enabled"] = False

    # -----------------
    # V3.4: columna Peso opcional (por defecto oculta).
    # -----------------
    if not show_peso:
        for k in ("validados", "dudosos", "no_encontrados"):
            for r in (res.get(k) or []):
                r.pop("Peso", None)
    res.setdefault("meta", {})["show_peso"] = bool(show_peso)

    if not show_cuit:
        for k in ("validados", "dudosos", "no_encontrados"):
            for r in (res.get(k) or []):
                r.pop("CUIT recibo", None)
                r.pop("CUIT ingreso", None)
    res.setdefault("meta", {})["show_cuit"] = bool(show_cuit)

    # -----------------
    # V3.7: Forzar validación manual desde la UI (promover Dudosos a Validados).
    # Se identifica por (Fila Excel, Nro recibo, Medio de pago).
    # -----------------
    if force_validations:
        forced_keys = set()
        forced_exact_keys = set()
        for it in force_validations:
            try:
                case_id = str(it.get("case_id", "") or "")
                fila_excel = str(it.get("fila_excel", "") or "")
                ranking = str(it.get("ranking", "") or "")
                if case_id and fila_excel:
                    forced_exact_keys.add((case_id, fila_excel, ranking))
                forced_keys.add((
                    fila_excel,
                    str(it.get("nro_recibo", "") or ""),
                    str(it.get("medio_pago", "") or ""),
                ))
            except Exception:
                continue

        def _k(row: dict) -> tuple[str, str, str]:
            return (
                str(row.get("Fila Excel", "")),
                str(row.get("Nro recibo", "")),
                str(row.get("Medio de pago", "")),
            )

        def _k_exact(row: dict) -> tuple[str, str, str]:
            return (
                str(row.get("__case_id", "") or ""),
                str(row.get("Fila Excel", "") or ""),
                str(row.get("Ranking", "") or ""),
            )

        moved: List[dict] = []
        kept: List[dict] = []
        for row in (res.get("dudosos") or []):
            if _k_exact(row) in forced_exact_keys or _k(row) in forced_keys:
                moved.append(row)
            else:
                kept.append(row)
        if moved:
            res["dudosos"] = kept
            res.setdefault("validados", [])
            res["validados"].extend(moved)

    # -----------------
    # V4.4.1: borrar casos dudosos desde UI.
    # Se identifica por (case_id, fila_excel, ranking) o fallback (fila_excel, nro_recibo, medio_pago).
    # -----------------
    if drop_dudosos:
        drop_keys = set()
        drop_exact_keys = set()
        for it in drop_dudosos:
            try:
                case_id = str(it.get("case_id", "") or "")
                fila_excel = str(it.get("fila_excel", "") or "")
                ranking = str(it.get("ranking", "") or "")
                if case_id and fila_excel:
                    drop_exact_keys.add((case_id, fila_excel, ranking))
                drop_keys.add((
                    fila_excel,
                    str(it.get("nro_recibo", "") or ""),
                    str(it.get("medio_pago", "") or ""),
                ))
            except Exception:
                continue

        def _drop_k(row: dict) -> tuple[str, str, str]:
            return (
                str(row.get("Fila Excel", "")),
                str(row.get("Nro recibo", "")),
                str(row.get("Medio de pago", "")),
            )

        def _drop_k_exact(row: dict) -> tuple[str, str, str]:
            return (
                str(row.get("__case_id", "") or ""),
                str(row.get("Fila Excel", "") or ""),
                str(row.get("Ranking", "") or ""),
            )

        kept_dudosos: List[dict] = []
        for row in (res.get("dudosos") or []):
            if _drop_k_exact(row) in drop_exact_keys or _drop_k(row) in drop_keys:
                continue
            kept_dudosos.append(row)
        res["dudosos"] = kept_dudosos


    # Attach run metadata (lightweight)
    excel_errors_count, excel_errors_sample = _summarize_excel_errors(txns)
    meta = dict(res.get("meta", {}) or {})
    meta.update({
        "pdf_date_min": dmin,
        "pdf_date_max": dmax,
        "pdf_report_date_min": rmin_global,
        "pdf_report_date_max": rmax_global,
        "margin_days": margin_days,
        "tolerance_days_suspect": tolerance_days_suspect,
        "day_weight_bank_before": day_weight_bank_before,
        "day_weight_bank_after": day_weight_bank_after,
        "valid_max_peso": valid_max_peso,
        "dudoso_max_peso": dudoso_max_peso,
        "mp_mismatch_penalty": effective_mp_mismatch_penalty,
        "mp_mismatch_penalty_effective": effective_mp_mismatch_penalty,
        "preconciled_penalty": preconciled_penalty,
        "penalty_salice_to_galicia": effective_penalty_salice_to_galicia,
        "penalty_alarcon_to_bbva": effective_penalty_alarcon_to_bbva,
        "cliente_cuit_mismatch_penalty": cliente_cuit_mismatch_penalty,
        "banco_sin_recibo_grace_days": banco_sin_recibo_grace_days,
        "recibo_sin_banco_grace_days": recibo_sin_banco_grace_days,
        "alternatives_cost_delta": alternatives_cost_delta,
        "max_bank_before_receipt_days": 10,
        "max_bank_after_receipt_days": 2,
        "excel_path": excel_path,
        "pdf_paths": [p for p, _ in pdfs],
        "receipts_source_used": source,
        "api_receipts_window_days": (effective_api_days if source == "api" else None),
        "api_request_id": (", ".join(dict.fromkeys([x for x in api_request_ids if x])) if api_request_ids else None),
        "external_warnings": external_warnings,
        "payments_filtered_old_window": int(payments_filtered_old),
        "payments_filtered_current_day": 0,
        "payments_filtered_before_bank_min_date": int(payments_filtered_before_bank_min),
        "payments_filtered_after_bank_max_date": int(payments_filtered_after_bank_max),
        "payments_filtered_preconciled": int(payments_filtered_preconciled),
        "api_start_date_from_bank": (api_start_date.isoformat() if api_start_date else None),
        "api_end_date_from_bank": (api_end_date.isoformat() if api_end_date else None),
        "api_start_date_override": (str(api_start_date_override) if api_start_date_override else None),
        "api_end_date_override": (str(api_end_date_override) if api_end_date_override else None),
        "api_payments_by_empresa": api_payments_by_empresa,
        "api_medio_bancarizable_stats": api_medio_bancarizable_stats,
        "api_post_detail_stats": api_post_detail_stats,
        "api_empresa_targets_used": api_targets_used,
        "api_comprobantes_count_by_target": api_count_by_target,
        "api_fecha_desde": api_fecha_desde,
        "api_fecha_hasta": api_fecha_hasta,
        "receipts_count": len(receipts_all),
        "payments_count": len(payments_all),
        "txns_count": sum(1 for t in txns if t.parse_ok),
        "txns_preconciled_count": sum(1 for t in txns if t.parse_ok and t.was_preconciled),
        "excel_parse_errors_count": excel_errors_count,
        "excel_parse_errors_sample": excel_errors_sample,
        "pdf_warnings": pdf_warnings,
        "padron_warnings": padron_warnings,
        "padron_cliente_cuit_path": resolved_padron_path,
        "padron_cliente_cuit_size": len(cliente_to_cuit_map),
        "txns_with_cuit_count": sum(1 for t in txns if t.parse_ok and t.cuit),
        "modo": "COMPLETO" if include_empresa else "SIMPLE",
        "mem_debug_enabled": bool(mem_dbg),
        "stage2_candidate_top_k": int(stage2_candidate_top_k),
        "exclude_preconciled_txns": False,
    })
    if mem_dbg:
        meta["mem_stages"] = mem_stages
    meta["elapsed_seconds"] = round(time.perf_counter() - t0, 3)
    res["meta"] = meta
    return res
