from __future__ import annotations

import datetime as dt
from typing import Dict, List, Tuple

from .excel_loader import BankTxn
from .pdf_parser import ReceiptPayment


def _days_between(d1: dt.date, d2: dt.date) -> int:
    return abs((d1 - d2).days)


def _medio_from_origen(origen: str) -> str:
    """Map bank origin to receipt medio_pago."""
    return "MERCADOPAGO" if origen == "MERCADOPAGO" else "TRANSFERENCIA"


def _amount_tolerance_suspect(amount: float) -> float:
    # Default rule: max($50, 0.5% del importe)
    return max(50.0, 0.005 * abs(amount))


def _is_numeric_str(s: str) -> bool:
    return s.isdigit()


def _recibo_sort_key(nro_recibo: str) -> Tuple[int, str]:
    # Sort primarily by numeric value when possible.
    if _is_numeric_str(nro_recibo):
        return (0, f"{int(nro_recibo):020d}")
    return (1, nro_recibo)


def match(
    txns: List[BankTxn],
    payments: List[ReceiptPayment],
    *,
    margin_days: int = 5,
    tolerance_days_valid: int = 1,
    tolerance_days_suspect: int = 7,
    tolerance_amount_valid_abs: float = 1.0,
    max_alternatives: int = 3,
    report_date_min: str | None = None,
    report_date_max: str | None = None,
    include_empresa: bool = False,
    enable_banco_sin_recibo: bool = True,
) -> Dict[str, List[dict]]:
    """Conciliación bi-direccional con salida enfocada en el recibo.

    - La salida principal (validados/dudosos) está en términos de *pagos del recibo*
      (Transferencia / Mercado Pago) y su mejor candidato en el banco.
    - Para evitar "ruido", los candidatos alternativos NO crean recibos extra:
      se emiten como filas adicionales *justo debajo* del candidato principal,
      sin repetir datos del recibo (solo datos del candidato y diferencias).
    - "Raros" (outliers) siempre van a dudosos (is raro = True).

    Devuelve:
      - validados: filas (principal + alternativos si aplica, aunque en validado suele no haber)
      - dudosos: filas (principal + alternativos)
      - no_encontrados: filas (BANCO_SIN_RECIBO y RECIBO_SIN_BANCO)
    """

    # Precompute date range from payments
    pay_dates = [dt.date.fromisoformat(p.fecha_pago) for p in payments]
    if pay_dates:
        min_d, max_d = min(pay_dates), max(pay_dates)
    else:
        min_d = max_d = None

    # IMPORTANT: el rango efectivo para tomar movimientos del Excel debe cubrir
    # tanto el margen como la ventana de "dudoso".
    effective_margin = max(margin_days, tolerance_days_suspect)

    def in_range(d: dt.date) -> bool:
        if not min_d or not max_d:
            return True
        return (min_d - dt.timedelta(days=effective_margin)) <= d <= (max_d + dt.timedelta(days=effective_margin))

    txns_match_scope = [t for t in txns if t.parse_ok and in_range(t.fecha)]

    # Para NO_ENCONTRADOS (BANCO_SIN_RECIBO) listamos **solo** movimientos dentro del rango exacto
    # informado en el encabezado del PDF (Desde/Hasta). Si no se puede detectar, caemos al
    # rango de fechas efectivas de pagos.
    report_min = dt.date.fromisoformat(report_date_min) if report_date_min else min_d
    report_max = dt.date.fromisoformat(report_date_max) if report_date_max else max_d

    def in_report_range_strict(d: dt.date) -> bool:
        if not report_min or not report_max:
            return False
        return report_min <= d <= report_max

    txns_report_scope = [t for t in txns if t.parse_ok and in_report_range_strict(t.fecha)]

    # Index txns by medio to reduce comparisons
    txns_by_medio: Dict[str, List[BankTxn]] = {"TRANSFERENCIA": [], "MERCADOPAGO": []}
    for t in txns_match_scope:
        txns_by_medio[_medio_from_origen(t.origen)].append(t)

    # Helper to build a candidate dict
    def build_candidate(t: BankTxn, p: ReceiptPayment) -> dict:
        pd = dt.date.fromisoformat(p.fecha_pago)
        dd = _days_between(t.fecha, pd)
        di = abs(float(t.importe) - float(p.importe_pago))
        return {
            "Origen": t.origen,
            "Fecha movimiento": t.fecha.isoformat(),
            "Importe movimiento": round(float(t.importe), 2),
            "Detalle movimiento": t.texto_ref,
            "Fila Excel": t.row_index,
            "Dif días": int(dd),
            "Dif importe": round(float(di), 2),
                    }

    # Produce rows for validados/dudosos (receipt-centric)
    validated_rows: List[dict] = []
    suspect_rows: List[dict] = []

    # Track used bank txns.
    # - used_valid_txn_ids: txns consumidos por VALIDADO (se bloquean para dudosos/no encontrados)
    # - used_principal_txn_ids: txns usados como candidato principal (validado o dudoso)
    used_valid_txn_ids: set[str] = set()
    used_principal_txn_ids: set[str] = set()

    # Helper: compute candidate list for a payment, optionally excluding already-used VALIDADO txns.
    def candidates_for_payment(p: ReceiptPayment, *, exclude_validated: bool = True) -> List[Tuple[float, int, BankTxn]]:
        pd = dt.date.fromisoformat(p.fecha_pago)
        out: List[Tuple[float, int, BankTxn]] = []
        for t in txns_by_medio.get(p.medio_pago, []):
            if exclude_validated and (t.txn_id in used_valid_txn_ids):
                continue
            dd = _days_between(t.fecha, pd)
            if dd > tolerance_days_suspect:
                continue
            di = abs(float(t.importe) - float(p.importe_pago))
            out.append((float(di), int(dd), t))
        out.sort(key=lambda x: (x[0], x[1]))
        return out

    # ------------------------------------------------------------------
    # FASE 1: Resolver VALIDADO globalmente (bloqueo efectivo)
    # ------------------------------------------------------------------
    # Problema que corrige: si se procesa por orden de recibo, un ingreso bancario
    # puede aparecer como dudoso en un recibo anterior y luego quedar VALIDADO
    # en un recibo posterior. Para evitarlo, primero detectamos todos los matches
    # VALIDADO fuertes y bloqueamos sus txns antes de construir DUDOSOS.

    pre_valid_best: List[Tuple[float, int, ReceiptPayment, BankTxn]] = []
    for p in [pp for pp in payments if pp.medio_pago in {"TRANSFERENCIA", "MERCADOPAGO"}]:
        cands = candidates_for_payment(p, exclude_validated=False)
        if not cands:
            continue
        di, dd, t = cands[0]
        # Solo consideramos como "validado fuerte" si cumple umbrales de validación.
        if dd <= tolerance_days_valid and di <= tolerance_amount_valid_abs:
            pre_valid_best.append((di, dd, p, t))

    # Asignación codiciosa: prioriza menor diff_importe y luego menor diff_dias.
    # Esto es suficiente para el caso operativo (y evita re-uso de un ingreso validado).
    pre_valid_best.sort(key=lambda x: (x[0], x[1]))

    validated_assignment: Dict[Tuple[str, str, str], BankTxn] = {}
    # key = (nro_recibo, nro_cliente, fecha_pago_iso, medio_pago)
    for _di, _dd, p, t in pre_valid_best:
        if t.txn_id in used_valid_txn_ids:
            continue
        key = (str(p.nro_recibo), str(p.nro_cliente), str(p.fecha_pago), str(p.medio_pago))
        validated_assignment[key] = t
        used_valid_txn_ids.add(t.txn_id)

    # Only consider relevant receipt payments
    relevant_payments = [p for p in payments if p.medio_pago in {"TRANSFERENCIA", "MERCADOPAGO"}]
    relevant_payments.sort(key=lambda p: _recibo_sort_key(str(p.nro_recibo)))

    # Para aclarar "recibo con múltiples transferencias/MP" sin confundir con alternativos,
    # precomputamos cuántos pagos relevantes hay por recibo y medio.
    counts_by_recibo_medio: Dict[tuple[str, str], int] = {}
    for p0 in relevant_payments:
        key = (str(p0.nro_recibo), p0.medio_pago)
        counts_by_recibo_medio[key] = counts_by_recibo_medio.get(key, 0) + 1

    # Índice (1..n) del pago dentro del recibo por medio
    idx_by_recibo_medio: Dict[tuple[str, str], int] = {}
    matched_payment_keys: set[Tuple[str, str, str, str]] = set()
    for p in relevant_payments:
        pay_key = (str(p.nro_recibo), str(p.nro_cliente), str(p.fecha_pago), str(p.medio_pago))

        # Base receipt columns (for principal rows)
        key_rm = (str(p.nro_recibo), p.medio_pago)
        idx_by_recibo_medio[key_rm] = idx_by_recibo_medio.get(key_rm, 0) + 1
        total_same = counts_by_recibo_medio.get(key_rm, 1)
        if total_same > 1:
            item_label = f"{idx_by_recibo_medio[key_rm]}/{total_same}"
        else:
            item_label = ""

        receipt_base = {
            **({"Empresa": str(p.empresa)} if include_empresa else {}),
            "Nro recibo": str(p.nro_recibo),
            "Nro cliente": str(p.nro_cliente),
            "Medio de pago": p.medio_pago,
            "Fecha recibo": p.fecha_pago,
            "Importe recibo": round(float(p.importe_pago), 2),
            "Vendedor": str(getattr(p, "vendedor", "") or ""),
            "Ítem en recibo": item_label,
        }

        # --------------------------------------------------------------
        # Si este pago ya quedó VALIDADO en la fase 1, lo emitimos aquí
        # y evitamos que quede "sin candidato" por el bloqueo.
        # --------------------------------------------------------------
        if pay_key in validated_assignment:
            best_t = validated_assignment[pay_key]
            best_dd = _days_between(best_t.fecha, dt.date.fromisoformat(p.fecha_pago))
            best_di = abs(float(best_t.importe) - float(p.importe_pago))

            principal = {
                "Tipo fila": "PRINCIPAL",
                "Ranking": 1,
                **receipt_base,
                **build_candidate(best_t, p),
            }
            validated_rows.append(principal)
            matched_payment_keys.add(pay_key)
            used_principal_txn_ids.add(best_t.txn_id)
            continue

        # Find candidates within "dudoso" window (± tolerance_days_suspect)
        candidates = candidates_for_payment(p, exclude_validated=True)

        if not candidates:
            # Receipt payment without bank
            # Goes to NO ENCONTRADOS (RECIBO_SIN_BANCO)
            # We'll emit it later in the no_encontrados list.
            continue

        # Determine best candidate
        best_di, best_dd, best_t = candidates[0]

        # Decide status for principal
        is_valid = (best_dd <= tolerance_days_valid) and (best_di <= tolerance_amount_valid_abs)

        # Determine "raro" and motivo (determinístico)
        motivo_parts: List[str] = []
        is_raro = False

        # Rare checks
        if float(best_t.importe) < 0:
            is_raro = True
            motivo_parts.append("raro: importe negativo")

        # Difference thresholds
        pct = (best_di / abs(float(p.importe_pago))) if float(p.importe_pago) != 0 else 0.0
        if best_di > 2000.0 or pct > 0.10:
            is_raro = True
            motivo_parts.append("raro: diferencia grande de importe")
        if best_dd > 3:
            # dentro de 7 días sigue siendo dudoso, pero ya es "raro" en la práctica
            is_raro = True
            motivo_parts.append("raro: fecha muy corrida")

        # Non-rare reasons
        if not is_raro:
            if best_dd > tolerance_days_valid:
                motivo_parts.append("fecha corrida")
            if best_di > tolerance_amount_valid_abs:
                motivo_parts.append("importe similar")

        # Multiple candidate ambiguity
        include_alternatives = False
        close_alts: List[Tuple[float, int, BankTxn]] = []
        if len(candidates) > 1:
            # Si el recibo tiene múltiples pagos (por ejemplo, 2 transferencias),
            # NO mostramos alternativos porque no hay "distintos casos posibles":
            # son varios ingresos para el mismo recibo.
            if total_same == 1:
                # Ambigüedad real si hay otros candidatos "cerca" del mejor.
                for di2, dd2, t2 in candidates[1:]:
                    if len(close_alts) >= max_alternatives:
                        break
                    if abs(di2 - best_di) <= 1.0 and abs(dd2 - best_dd) <= 1:
                        close_alts.append((di2, dd2, t2))

                if close_alts:
                    motivo_parts.append("múltiples candidatos")
                    include_alternatives = True

        motivo = "; ".join(dict.fromkeys(motivo_parts)) if motivo_parts else ("ok" if is_valid else "dudoso")

        # If not valid or rare, it is DUDOSO
        status = "VALIDADO" if (is_valid and not is_raro) else "DUDOSO"

        # Build principal row. En VALIDADOS no incluimos columnas "Motivo" ni "Es raro".
        principal = {
            "Tipo fila": "PRINCIPAL",
            "Ranking": 1,
            **receipt_base,
            **build_candidate(best_t, p),
        }
        if status != "VALIDADO":
            principal["Motivo"] = motivo
            principal["Es raro"] = bool(is_raro)

        # For principal matches, mark txn used.
        # Solo los VALIDADO se bloquean globalmente.
        used_principal_txn_ids.add(best_t.txn_id)
        if status == "VALIDADO":
            used_valid_txn_ids.add(best_t.txn_id)

        # Alternativos rows (solo si hay ambigüedad real): no repetir datos del recibo
        alternativos_rows: List[dict] = []
        if include_alternatives:
            for i, (_di, _dd, t_alt) in enumerate(close_alts, start=2):
                cand = build_candidate(t_alt, p)
                # Short reason for appearance
                why = []
                if cand["Dif importe"] == principal["Dif importe"]:
                    why.append("mismo importe")
                if cand["Dif días"] == principal["Dif días"]:
                    why.append("misma fecha")
                if not why:
                    why.append("candidato alternativo cercano")
                alternativos_rows.append({
                    "Tipo fila": "ALTERNATIVO",
                    "Ranking": i,
                    # Blank receipt columns
                    **({"Empresa": ""} if include_empresa else {}),
                    "Nro recibo": "",
                    "Nro cliente": "",
                    "Medio de pago": "",
                    "Fecha recibo": "",
                    "Importe recibo": "",
                    "Vendedor": "",
                    "Ítem en recibo": "",
                    **cand,
                    "Motivo": "Alternativo: " + ", ".join(why),
                    "Es raro": False,
                })

        # Append to correct list with alternativos just below
        if status == "VALIDADO":
            validated_rows.append(principal)
            # (en validado, alternativos normalmente no se muestran; pero si querés, se puede habilitar)
        else:
            suspect_rows.append(principal)
            suspect_rows.extend(alternativos_rows)

        matched_payment_keys.add(pay_key)

    # NO ENCONTRADOS
    no_encontrados: List[dict] = []

    # A) RECIBO_SIN_BANCO (por cada pago del recibo sin match)
    # We need to re-evaluate those payments we skipped earlier
    for p in relevant_payments:
        pay_key = (str(p.nro_recibo), str(p.nro_cliente), str(p.fecha_pago), str(p.medio_pago))
        if pay_key in matched_payment_keys:
            continue
        candidates = candidates_for_payment(p, exclude_validated=True)
        if not candidates:
            no_encontrados.append({
                "Tipo no encontrado": "RECIBO_SIN_BANCO",
                **({"Empresa": str(p.empresa)} if include_empresa else {}),
                "Nro recibo": str(p.nro_recibo),
                "Nro cliente": str(p.nro_cliente),
                "Medio de pago": p.medio_pago,
                "Fecha recibo": p.fecha_pago,
                "Importe recibo": round(float(p.importe_pago), 2),
                "Vendedor": str(getattr(p, "vendedor", "") or ""),
                "Ítem en recibo": "",
                "Origen": "NO APLICA",
                "Fecha movimiento": "NO APLICA",
                "Importe movimiento": "NO APLICA",
                "Detalle movimiento": "NO APLICA",
                "Fila Excel": "NO APLICA",
            })

    # B) BANCO_SIN_RECIBO (movimientos bancarios no usados como principal)
    if enable_banco_sin_recibo:
        for t in txns_report_scope:
            # Un ingreso usado como principal (validado o dudoso) no debería aparecer como
            # "banco sin recibo" porque ya fue propuesto/asignado a un recibo.
            if t.txn_id in used_principal_txn_ids:
                continue
            medio = _medio_from_origen(t.origen)
            no_encontrados.append({
                "Tipo no encontrado": "BANCO_SIN_RECIBO",
                **({"Empresa": ""} if include_empresa else {}),
                "Nro recibo": "NO APLICA",
                "Nro cliente": "NO APLICA",
                "Medio de pago": medio,
                "Fecha recibo": "NO APLICA",
                "Importe recibo": "NO APLICA",
                "Vendedor": "",
                "Ítem en recibo": "",
                "Origen": t.origen,
                "Fecha movimiento": t.fecha.isoformat(),
                "Importe movimiento": round(float(t.importe), 2),
                "Detalle movimiento": t.texto_ref,
                "Fila Excel": t.row_index,
            })

    # Sort no_encontrados by receipt when present; otherwise by date
    def ne_sort_key(r: dict) -> Tuple:
        tr = r.get("Tipo no encontrado")
        nro = r.get("Nro recibo")
        if tr == "RECIBO_SIN_BANCO":
            return (0, _recibo_sort_key(str(nro)))
        # bank only
        return (1, str(r.get("Fecha movimiento", "")), str(r.get("Origen", "")))

    no_encontrados.sort(key=ne_sort_key)

    return {
        "validados": validated_rows,
        "dudosos": suspect_rows,
        "no_encontrados": no_encontrados,
    }
