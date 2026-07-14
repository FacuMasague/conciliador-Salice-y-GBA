from __future__ import annotations

from typing import Dict, List, Tuple, Optional

import datetime as dt
import os
import time
import unicodedata
from dataclasses import replace
from pathlib import Path

import pandas as pd

from .excel_loader import load_bank_txns
from .pdf_parser import (
    parse_receipts_and_payments_from_text,
    pdf_date_range,
    detect_pdf_warnings_from_text,
    report_period_range_from_text,
    extract_pdf_text,
)
from .matcher_hungarian import match_hungarian
from .memdebug import is_mem_debug_enabled, mem_debug_recorder
from .external.service import fetch_cliente_cuit_map as fetch_cliente_cuit_map_api
from .external.service import fetch_receipts_and_payments as fetch_receipts_and_payments_api
from .collector_catalog import load_internal_collector_receipts

API_RECEIPTS_MAX_DAYS = 15


def _normalize_text(value: object) -> str:
    s = unicodedata.normalize("NFKD", str(value or ""))
    s = "".join(ch for ch in s if not unicodedata.combining(ch))
    return s.strip().lower()


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


def _normalize_recibo(value: object) -> Optional[str]:
    if value is None:
        return None
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        try:
            return str(int(float(value)))
        except Exception:
            pass
    s = str(value).strip()
    if not s:
        return None
    s_num = s.replace("$", "").replace(" ", "")
    if s_num:
        parsed_float: float | None = None
        try:
            if "." in s_num and "," in s_num:
                # Asumimos formato contable local: 68.734,00 -> 68734.00
                parsed_float = float(s_num.replace(".", "").replace(",", "."))
            elif "," in s_num:
                if s_num.count(",") == 1 and len(s_num.rsplit(",", 1)[-1]) <= 2:
                    parsed_float = float(s_num.replace(".", "").replace(",", "."))
                elif all(part.isdigit() for part in s_num.split(",")):
                    parsed_float = float(s_num.replace(",", ""))
            elif "." in s_num:
                if s_num.count(".") == 1 and len(s_num.rsplit(".", 1)[-1]) <= 2:
                    parsed_float = float(s_num)
                elif all(part.isdigit() for part in s_num.split(".")):
                    parsed_float = float(s_num.replace(".", ""))
            elif s_num.isdigit():
                parsed_float = float(s_num)
        except Exception:
            parsed_float = None
        if parsed_float is not None:
            try:
                return str(int(parsed_float))
            except Exception:
                pass
    digits = "".join(ch for ch in s if ch.isdigit())
    if digits:
        try:
            return str(int(digits))
        except Exception:
            return digits.lstrip("0") or "0"
    return s.upper()


def enrich_api_payments_with_collectors(
    payments: List,
    collector_receipts: List,
) -> tuple[List, Dict[str, int], List[str]]:
    """Completa el cobrador de pagos API usando una relación exacta por recibo.

    La fuente de cobradores no reemplaza importes, fechas ni medios de pago de
    GESI. Sólo aporta el usuario/cobrador asociado al número de recibo.
    """
    exact: Dict[tuple[str, str], str] = {}
    by_number: Dict[str, set[str]] = {}
    conflicts = 0
    pdf_receipt_numbers: set[tuple[str, str]] = set()

    for receipt in collector_receipts:
        number = _normalize_recibo(getattr(receipt, "nro_recibo", None))
        collector = str(getattr(receipt, "vendedor", "") or "").strip()
        company = str(getattr(receipt, "empresa", "") or "").strip().upper()
        if not number or not collector:
            continue
        key = (company, number)
        previous = exact.get(key)
        if previous and previous != collector:
            conflicts += 1
            continue
        exact[key] = collector
        pdf_receipt_numbers.add(key)
        by_number.setdefault(number, set()).add(collector)

    enriched: List = []
    matched_pdf_keys: set[tuple[str, str]] = set()
    enriched_count = 0
    for payment in payments:
        number = _normalize_recibo(getattr(payment, "nro_recibo", None))
        company = str(getattr(payment, "empresa", "") or "").strip().upper()
        collector = exact.get((company, number or ""))
        if collector:
            matched_pdf_keys.add((company, number or ""))
        elif number and len(by_number.get(number, set())) == 1:
            collector = next(iter(by_number[number]))
            for key in pdf_receipt_numbers:
                if key[1] == number:
                    matched_pdf_keys.add(key)
                    break

        if collector:
            enriched.append(replace(payment, vendedor=collector))
            enriched_count += 1
        else:
            enriched.append(payment)

    with_collector = sum(
        bool(str(getattr(payment, "vendedor", "") or "").strip())
        for payment in enriched
    )
    stats = {
        "pdf_collectors_count": len(exact),
        "api_payments_enriched_count": enriched_count,
        "api_collectors_count": with_collector,
        "api_collectors_missing_count": len(enriched) - with_collector,
        "pdf_collectors_not_in_api_count": len(pdf_receipt_numbers - matched_pdf_keys),
        "pdf_collector_conflicts_count": conflicts,
    }
    warnings: List[str] = []
    if exact:
        warnings.append(
            f"Cobradores: {enriched_count}/{len(enriched)} pagos API fueron enriquecidos "
            f"desde {len(exact)} asignaciones exactas de la fuente interna."
        )
        missing = stats["api_collectors_missing_count"]
        if missing:
            warnings.append(
                f"Cobradores: {missing} pagos API quedaron sin cobrador porque su recibo "
                "no aparece identificado en la fuente interna."
            )
        if stats["pdf_collectors_not_in_api_count"]:
            warnings.append(
                "Cobradores: "
                f"{stats['pdf_collectors_not_in_api_count']} asignaciones internas no fueron "
                "devueltos por GESI en el rango consultado."
            )
        if conflicts:
            warnings.append(
                f"Cobradores: se ignoraron {conflicts} asignaciones contradictorias."
            )
    else:
        warnings.append(
            "Cobradores: la fuente interna no contiene asignaciones para este rango. "
            "GESI no informa el cobrador del recibo, por lo que "
            "los casos sin dato directo quedan vacíos."
        )
    return enriched, stats, warnings


# Nombre conservado para integraciones y pruebas de la versión 5.2.2.
enrich_api_payments_with_pdf_collectors = enrich_api_payments_with_collectors


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
    day_weight_bank_before: float = 40.0,
    day_weight_bank_after: float = 50.0,
    valid_max_peso: float = 150.0,
    dudoso_max_peso: float = 3500.0,
    mp_mismatch_penalty: float = 35.0,
    preconciled_penalty: float = 150.0,
    alternatives_cost_delta: float = 50.0,
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
    excel_path: str,
    pdfs: List[Tuple[str, Optional[str]]],
    *,
    margin_days: int = 5,
    tolerance_days_suspect: int = 7,
    max_options: int = 4,
    # V3.5: multiplicador de días depende del signo (recibo vs banco)
    day_weight_bank_before: float = 40.0,
    day_weight_bank_after: float = 50.0,
    valid_max_peso: float = 150.0,
    dudoso_max_peso: float = 3500.0,
    mp_mismatch_penalty: float = 35.0,
    preconciled_penalty: float = 150.0,
    alternatives_cost_delta: float = 50.0,
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
    # V3.0: IA eliminada (se mantienen parámetros solo por compatibilidad).
    enable_ai: bool = False,
) -> Dict[str, List[dict]]:
    """Concilia un Excel contra recibos de PDF (legacy) o API (v5).

    receipts_source:
      - "pdf": usa `pdfs` como entrada (modo legacy).
      - "api": trae recibos por API y completa el cobrador automáticamente
        desde la fuente interna administrada por el sistema.
    """
    t0 = time.perf_counter()
    mem_dbg = is_mem_debug_enabled(mem_debug)
    mem_stages, mem_mark = mem_debug_recorder(mem_dbg)
    mem_mark("start")

    source = str(receipts_source or "pdf").strip().lower()
    if source not in {"pdf", "api"}:
        raise ValueError("receipts_source inválido. Usar 'pdf' o 'api'.")

    enable_banco_sin_recibo = True
    receipts_all = []
    payments_all = []
    pdf_warnings: List[str] = []
    external_warnings: List[str] = []
    api_request_ids: List[str] = []
    api_payments_by_empresa: Dict[str, int] = {}
    api_medio_bancarizable_stats: Dict[str, int] = {}
    api_targets_used: List[str] = []
    api_count_by_target: Dict[str, int] = {}
    api_fleteros_count = 0
    api_fleteros_missing_count = 0
    api_collector_stats: Dict[str, int] = {
        "pdf_collectors_count": 0,
        "api_payments_enriched_count": 0,
        "api_collectors_count": 0,
        "api_collectors_missing_count": 0,
        "pdf_collectors_not_in_api_count": 0,
        "pdf_collector_conflicts_count": 0,
    }
    internal_collector_meta: Dict[str, object] = {
        "internal_collector_catalog_count": 0,
        "internal_collector_catalog_files": 0,
        "internal_collector_catalog_loaded": False,
    }
    api_fecha_desde: str | None = None
    api_fecha_hasta: str | None = None
    api_cliente_cuit_map: Dict[str, str] = {}
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
    txns = load_bank_txns(excel_path)
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
                api_empresa_filter,
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
                api_empresa_filter,
                end_date=api_end_date,
            )
        payments_all.extend(api_payments)
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
        api_fleteros_count = int(api_meta.get("fleteros_count") or 0)
        api_fleteros_missing_count = int(api_meta.get("fleteros_missing_count") or 0)
        api_fecha_desde = str(api_meta.get("api_fecha_desde") or "") or None
        api_fecha_hasta = str(api_meta.get("api_fecha_hasta") or "") or None
        if isinstance(api_meta.get("cliente_cuit_map"), dict):
            api_cliente_cuit_map = {
                str(k): str(v)
                for k, v in api_meta["cliente_cuit_map"].items()
                if str(k).strip() and str(v).strip()
            }

        # GESI no devuelve el usuario que generó los recibos importados desde
        # Pedidos Móviles. La aplicación administra internamente la relación
        # exacta nro_recibo -> [ID - nombre]; el operador no sube controles.
        collector_receipts, internal_collector_meta, internal_collector_warnings = (
            load_internal_collector_receipts()
        )
        external_warnings.extend(internal_collector_warnings)

        # Compatibilidad de backend: una integración antigua todavía puede
        # adjuntar reportes, aunque la interfaz 5.2.3 ya no los solicita.
        for pdf_path, empresa_override in pdfs:
            text = extract_pdf_text(pdf_path)
            pdf_receipts, _pdf_payments_ignored = parse_receipts_and_payments_from_text(
                text,
                empresa_override=empresa_override,
            )
            receipts_all.extend(pdf_receipts)
            collector_receipts.extend(pdf_receipts)
            pdf_warnings.extend(detect_pdf_warnings_from_text(text[:150_000]))
            rmin, rmax = report_period_range_from_text(text[:60_000])
            if rmin:
                rmins.append(rmin)
            if rmax:
                rmaxs.append(rmax)

        payments_all, api_collector_stats, collector_warnings = (
            enrich_api_payments_with_collectors(payments_all, collector_receipts)
        )
        external_warnings.extend(collector_warnings)
        api_fleteros_count = int(api_collector_stats["api_collectors_count"])
        api_fleteros_missing_count = int(
            api_collector_stats["api_collectors_missing_count"]
        )
        mem_mark("api_receipts_parsed", {"payments_count": len(payments_all), "window_days": effective_api_days})

    payments_filtered_old = 0
    payments_filtered_after_bank_max = 0
    payments_filtered_before_bank_min = 0
    payments_filtered_preconciled = 0
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
    dmin, dmax = pdf_date_range(payments_all)
    rmin_global = min(rmins) if rmins else dmin
    if source == "api":
        rmax_global = (api_end_date.isoformat() if api_end_date else dmax)
    else:
        rmax_global = max(rmaxs) if rmaxs else dmax

    padron_warnings: List[str] = []
    cliente_to_cuit_map: Dict[str, str] = {}
    resolved_padron_path: str | None = None
    if source == "api":
        if api_cliente_cuit_map:
            cliente_to_cuit_map = api_cliente_cuit_map
        else:
            padron_map, padron_meta = fetch_cliente_cuit_map_api(api_empresa_filter)
            cliente_to_cuit_map = padron_map
            rid = str(padron_meta.get("api_request_id") or "").strip()
            if rid:
                api_request_ids.append(rid)
            external_warnings.extend([str(x) for x in (padron_meta.get("external_warnings") or []) if str(x).strip()])
    else:
        resolved_padron_path = (padron_cliente_cuit_path or "").strip() or _discover_padron_path(excel_path)
        if resolved_padron_path:
            try:
                cliente_to_cuit_map = _load_cliente_cuit_map(resolved_padron_path)
                if not cliente_to_cuit_map:
                    padron_warnings.append("No se encontraron mappings válidos cliente↔CUIT en el padrón.")
            except Exception as e:
                padron_warnings.append(f"No se pudo leer el padrón cliente↔CUIT: {e}")
        else:
            padron_warnings.append("No se encontró el archivo de padrón básico MDP para validar cliente↔CUIT.")

    # Matching
    # Matching (nota: el rango efectivo de Excel se define por el PDF y el margen,
    # pero el matching "dudoso" usa una ventana más amplia: tolerance_days_suspect)
    effective_mp_mismatch_penalty = 0.0 if source == "api" else float(mp_mismatch_penalty)
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
        cliente_to_cuit_map=cliente_to_cuit_map,
        alternatives_cost_delta=alternatives_cost_delta,
        max_alternatives=max(0, int(max_options) - 1),
        report_date_min=rmin_global,
        report_date_max=rmax_global,
        enable_banco_sin_recibo=enable_banco_sin_recibo,
        stage2_candidate_top_k=stage2_candidate_top_k,
        exclude_preconciled_txns=(source == "api"),
        mem_debug=mem_dbg,
    )
    mem_mark("matched", {
        "validados": len(res.get("validados") or []),
        "dudosos": len(res.get("dudosos") or []),
        "no_encontrados": len(res.get("no_encontrados") or []),
    })

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
        "mp_mismatch_penalty": mp_mismatch_penalty,
        "mp_mismatch_penalty_effective": effective_mp_mismatch_penalty,
        "preconciled_penalty": preconciled_penalty,
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
        "api_empresa_targets_used": api_targets_used,
        "api_comprobantes_count_by_target": api_count_by_target,
        "api_fleteros_count": api_fleteros_count,
        "api_fleteros_missing_count": api_fleteros_missing_count,
        "api_cobradores_count": api_collector_stats["api_collectors_count"],
        "api_cobradores_missing_count": api_collector_stats["api_collectors_missing_count"],
        "api_payments_enriched_from_pdf_count": api_collector_stats["api_payments_enriched_count"],
        "api_payments_enriched_from_internal_count": api_collector_stats["api_payments_enriched_count"],
        "pdf_cobradores_count": api_collector_stats["pdf_collectors_count"],
        "pdf_cobradores_not_in_api_count": api_collector_stats["pdf_collectors_not_in_api_count"],
        "pdf_cobradores_conflicts_count": api_collector_stats["pdf_collector_conflicts_count"],
        **internal_collector_meta,
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
        "modo": "UNICO",
        "mem_debug_enabled": bool(mem_dbg),
        "stage2_candidate_top_k": int(stage2_candidate_top_k),
        "exclude_preconciled_txns": bool(source == "api"),
    })
    if mem_dbg:
        meta["mem_stages"] = mem_stages
    meta["elapsed_seconds"] = round(time.perf_counter() - t0, 3)
    res["meta"] = meta
    return res
