from __future__ import annotations

import datetime as dt
from collections import defaultdict, deque
from typing import Dict, List, Tuple, Optional

import numpy as np
from scipy.optimize import linear_sum_assignment

from .excel_loader import BankTxn
from .pdf_parser import ReceiptPayment
from .memdebug import mem_debug_recorder


def _days_between(d1: dt.date, d2: dt.date) -> int:
    return abs((d1 - d2).days)


def _signed_days(receipt_date: dt.date, bank_date: dt.date) -> int:
    """Días (con signo) entre recibo y banco.

    > 0  => el recibo es posterior al ingreso bancario.
    = 0  => mismo día.
    < 0  => el recibo es anterior al ingreso (posible "delay" del banco).
    """
    return (receipt_date - bank_date).days


def _amount_tolerance_suspect(amount: float) -> float:
    # Default rule: max($50, 0.5% del importe)
    return max(50.0, 0.005 * abs(amount))


def _is_numeric_str(s: str) -> bool:
    return s.isdigit()


def _recibo_sort_key(nro_recibo: str) -> Tuple[int, str]:
    if _is_numeric_str(nro_recibo):
        return (0, f"{int(nro_recibo):020d}")
    return (1, nro_recibo)


def _mp_mismatch_penalty(p_medio: str, origen: str, penalty: float) -> float:
    """Penaliza cuando el medio del recibo no coincide con el "origen" del Excel.

    Regla de negocio: un cobro por Mercado Pago puede impactar en el banco,
    y una transferencia puede aparecer como ingreso en Mercado Pago. No lo prohibimos;
    solo lo hacemos menos preferible.

    - Recibo MERCADOPAGO + origen != MERCADOPAGO -> penaliza
    - Recibo TRANSFERENCIA + origen == MERCADOPAGO -> penaliza
    """
    if p_medio == "MERCADOPAGO" and origen != "MERCADOPAGO":
        return float(penalty)
    if p_medio != "MERCADOPAGO" and origen == "MERCADOPAGO":
        return float(penalty)
    return 0.0



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


def _cost(
    t: BankTxn,
    p: ReceiptPayment,
    *,
    day_weight_bank_before: float,
    day_weight_bank_after: float,
    mp_mismatch_penalty: float,
    preconciled_penalty: float,
) -> float:
    pd = dt.date.fromisoformat(p.fecha_pago)
    dd_signed = _signed_days(pd, t.fecha)
    dd = abs(dd_signed)
    # V3.5: multiplicador de días según signo.
    #  >= 0: ingreso bancario anterior (o mismo día) que el recibo
    #   < 0: ingreso bancario posterior al recibo (delay del banco)
    dw = float(day_weight_bank_before) if dd_signed >= 0 else float(day_weight_bank_after)
    di = abs(float(t.importe) - float(p.importe_pago))
    return (
        float(dw) * float(dd)
        + float(di)
        + _mp_mismatch_penalty(p.medio_pago, t.origen, mp_mismatch_penalty)
        + (float(preconciled_penalty) if t.was_preconciled else 0.0)
    )


def _build_candidate(
    t: BankTxn,
    p: ReceiptPayment,
    *,
    day_weight_bank_before: float,
    day_weight_bank_after: float,
    mp_mismatch_penalty: float,
    preconciled_penalty: float,
) -> dict:
    pd = dt.date.fromisoformat(p.fecha_pago)
    dd_signed = _signed_days(pd, t.fecha)
    dd = abs(dd_signed)
    di = abs(float(t.importe) - float(p.importe_pago))
    cost = _cost(
        t,
        p,
        day_weight_bank_before=day_weight_bank_before,
        day_weight_bank_after=day_weight_bank_after,
        mp_mismatch_penalty=mp_mismatch_penalty,
        preconciled_penalty=preconciled_penalty,
    )
    return {
        "Origen": t.origen,
        "Fecha movimiento": t.fecha.isoformat(),
        "Importe movimiento": round(float(t.importe), 2),
        "Detalle movimiento": t.texto_ref,
        "CUIT ingreso": str(t.cuit or ""),
        "Fila Excel": t.row_index,
        # Nota: dejamos el signo para poder auditar rápidamente el sentido.
        # Negativo => el banco impactó después del recibo (delay del banco).
        "Dif días": int(dd_signed),
        "Dif importe": round(float(di), 2),
        "Peso": round(float(cost), 2),
    }


def _build_motivo(
    p_medio: str,
    origen: str,
    dd: int,
    di: float,
    *,
    prefix: str,
    was_preconciled: bool = False,
    preconciled_recibo: str | None = None,
    medium_cross_enabled: bool = True,
) -> str:
    """Motivo corto y operativo (sin IA)."""
    why: List[str] = []
    if int(dd) != 0:
        why.append(f"fecha corrida {int(dd)} días")
    if float(di) > 0:
        why.append(f"importe distinto ${float(di):.2f}")
    # Medio cruzado (no es prohibido, pero explica por qué no validó directo)
    if medium_cross_enabled:
        if p_medio == "MERCADOPAGO" and origen != "MERCADOPAGO":
            why.append("medio cruzado")
        if p_medio != "MERCADOPAGO" and origen == "MERCADOPAGO":
            why.append("medio cruzado")
    if preconciled_recibo:
        why.append(f"ingreso ya conciliado con recibo {preconciled_recibo}")
    elif was_preconciled:
        why.append("ingreso ya conciliado previamente")
    if not why:
        why.append("costo alto")
    return prefix + ": " + ", ".join(why)


def _multi_info(relevant_payments: List[ReceiptPayment]) -> Tuple[Dict[tuple[str, str], int], Dict[tuple[str, str], int]]:
    counts_by_recibo_medio: Dict[tuple[str, str], int] = {}
    for p in relevant_payments:
        k = (str(p.nro_recibo), p.medio_pago)
        counts_by_recibo_medio[k] = counts_by_recibo_medio.get(k, 0) + 1
    idx_by_recibo_medio: Dict[tuple[str, str], int] = {}
    return counts_by_recibo_medio, idx_by_recibo_medio


def match_hungarian(
    txns: List[BankTxn],
    payments: List[ReceiptPayment],
    *,
    margin_days: int = 5,
    tolerance_days_suspect: int = 7,
    # V3.5: se separa el multiplicador de días según el signo (recibo vs banco).
    day_weight_bank_before: float = 40.0,
    day_weight_bank_after: float = 50.0,
    valid_max_peso: float = 150.0,
    dudoso_max_peso: float = 3500.0,
    mp_mismatch_penalty: float = 35.0,
    preconciled_penalty: float = 150.0,
    cuit_mismatch_penalty: float = 75.0,
    cliente_to_cuit_map: Dict[str, str] | None = None,
    max_alternatives: int = 3,
    report_date_min: str | None = None,
    report_date_max: str | None = None,
    enable_banco_sin_recibo: bool = True,
    alternatives_cost_delta: float = 50.0,
    stage2_candidate_top_k: int = 120,
    exclude_preconciled_txns: bool = False,
    mem_debug: bool = False,
) -> Dict[str, List[dict]]:
    """Prototipo de conciliación usando Hungarian algorithm.

    Cambios clave para tu caso:
    - Etapa 1: fija VALIDADOS por peso (≤ umbral) con desempate por menor peso.
    - Etapa 2: Hungarian para el resto.
    - Los ingresos usados en VALIDADOS no pueden aparecer en DUDOSOS (principal ni alternativos).

    El "Peso" es el costo literal: (días*day_weight) + dif_importe + penalizaciones de reglas de negocio.
    En V3.5, day_weight depende del signo de "Dif días":
      - Dif días >= 0 => day_weight_bank_before
      - Dif días  < 0 => day_weight_bank_after
    Menor es mejor.
    """

    mem_stages, mem_mark = mem_debug_recorder(bool(mem_debug))
    mem_mark("matcher_start")

    # -----------------
    # REGLA V3.0 (producto):
    # Para cada recibo, solo consideramos movimientos bancarios dentro de:
    #   banco ∈ [recibo - 10 días, recibo + 2 días]
    # Esto evita matches imposibles y acota el universo de candidatos.
    # -----------------
    MAX_BANK_BEFORE_RECEIPT_DAYS = 10
    MAX_BANK_AFTER_RECEIPT_DAYS = 5        # ventana general: banco puede llegar hasta 5 días después del recibo
    MAX_BANK_AFTER_RECEIPT_DAYS_VALID = 2  # para VALIDADOS: máximo 2 días después del recibo

    def _is_valid_time_window(receipt_date: dt.date, bank_date: dt.date) -> bool:
        sd = _signed_days(receipt_date, bank_date)
        return (-MAX_BANK_AFTER_RECEIPT_DAYS) <= sd <= MAX_BANK_BEFORE_RECEIPT_DAYS

    def _is_valid_for_validado(receipt_date: dt.date, bank_date: dt.date) -> bool:
        """Ventana más estricta para VALIDADOS: banco no puede llegar más de 2 días después del recibo."""
        sd = _signed_days(receipt_date, bank_date)
        return sd >= -MAX_BANK_AFTER_RECEIPT_DAYS_VALID

    # Relevant receipt payments
    relevant_payments = list(payments)
    if not relevant_payments:
        return {"validados": [], "dudosos": [], "no_encontrados": []}

    # Stable order by recibo
    relevant_payments.sort(key=lambda p: _recibo_sort_key(str(p.nro_recibo)))

    # Date range from payments (fallback); match scope uses max(margin, tolerance_suspect)
    pay_dates = [dt.date.fromisoformat(p.fecha_pago) for p in relevant_payments]
    min_d, max_d = min(pay_dates), max(pay_dates)
    effective_margin = max(int(margin_days), int(tolerance_days_suspect))

    def in_match_scope(d: dt.date) -> bool:
        return (min_d - dt.timedelta(days=effective_margin)) <= d <= (max_d + dt.timedelta(days=effective_margin))

    txns_match_scope = [t for t in txns if t.parse_ok and in_match_scope(t.fecha)]
    if exclude_preconciled_txns:
        txns_match_scope = [t for t in txns_match_scope if not t.was_preconciled]

    # V5.2: exclusividad CUIT.
    # Si un recibo tiene un CUIT esperado (del padrón) que aparece en algún movimiento bancario,
    # ese recibo SOLO puede conciliarse con movimientos que compartan ese mismo CUIT.
    txn_cuit_set: set[str] = {
        c for t in txns_match_scope if (c := _normalize_cuit(t.cuit))
    }

    def _cuit_mismatch_extra(t: BankTxn, p: ReceiptPayment) -> float:
        """Devuelve una penalidad adicional cuando el CUIT del movimiento bancario no coincide
        con el CUIT esperado del cliente (según el padrón).

        Lógica:
        - Txns sin CUIT: penalidad 0 (no podemos restringir).
        - CUIT del txn == CUIT esperado: penalidad 0 (match perfecto).
        - CUIT del txn != CUIT esperado Y el CUIT esperado está en scope: penalidad moderada.
          Esto permite que un cliente que paga desde otra cuenta igual matchee, pero hace
          que el algoritmo prefiera el txn con CUIT correcto cuando ambos están disponibles.
        """
        if not cliente_to_cuit_map:
            return 0.0
        txn_cuit = _normalize_cuit(t.cuit)
        if not txn_cuit:
            return 0.0  # txn sin CUIT: sin penalidad
        cli = _normalize_cliente(p.nro_cliente)
        if not cli:
            return 0.0
        expected = _normalize_cuit(cliente_to_cuit_map.get(cli))
        if not expected:
            return 0.0
        if expected not in txn_cuit_set:
            return 0.0  # ningún txn tiene el CUIT esperado → sin restricción
        if txn_cuit == expected:
            return 0.0  # match perfecto
        return float(cuit_mismatch_penalty)

    # Strict report range (for BANCO_SIN_RECIBO)
    report_min = dt.date.fromisoformat(report_date_min) if report_date_min else min_d
    report_max = dt.date.fromisoformat(report_date_max) if report_date_max else max_d

    def in_report_range(d: dt.date) -> bool:
        return report_min <= d <= report_max

    txns_report_scope = [t for t in txns if t.parse_ok and in_report_range(t.fecha)]
    if exclude_preconciled_txns:
        txns_report_scope = [t for t in txns_report_scope if not t.was_preconciled]

    def _cuit_recibo_for_payment(p: ReceiptPayment) -> str:
        if not cliente_to_cuit_map:
            return ""
        cli = _normalize_cliente(p.nro_cliente)
        if not cli:
            return ""
        return str(_normalize_cuit(cliente_to_cuit_map.get(cli)) or "")

    def _cost_for(t: BankTxn, p: ReceiptPayment) -> float:
        return _cost(
            t,
            p,
            day_weight_bank_before=day_weight_bank_before,
            day_weight_bank_after=day_weight_bank_after,
            mp_mismatch_penalty=mp_mismatch_penalty,
            preconciled_penalty=preconciled_penalty,
        ) + _cuit_mismatch_extra(t, p)

    def _candidate_for(t: BankTxn, p: ReceiptPayment) -> dict:
        result = _build_candidate(
            t,
            p,
            day_weight_bank_before=day_weight_bank_before,
            day_weight_bank_after=day_weight_bank_after,
            mp_mismatch_penalty=mp_mismatch_penalty,
            preconciled_penalty=preconciled_penalty,
        )
        extra = _cuit_mismatch_extra(t, p)
        if extra:
            result["Peso"] = round(result.get("Peso", 0.0) + extra, 2)
        return result

    # Multi-transfer notes per recibo/medio
    counts_by_recibo_medio, idx_by_recibo_medio = _multi_info(relevant_payments)

    def receipt_base_for(p: ReceiptPayment) -> dict:
        key_rm = (str(p.nro_recibo), p.medio_pago)
        idx_by_recibo_medio[key_rm] = idx_by_recibo_medio.get(key_rm, 0) + 1
        total_same = counts_by_recibo_medio.get(key_rm, 1)
        if total_same > 1:
            item_label = f"{idx_by_recibo_medio[key_rm]}/{total_same}"
        else:
            item_label = ""

        return {
            "Nro recibo": str(p.nro_recibo),
            "Nro cliente": str(p.nro_cliente),
            "Cliente": str(getattr(p, "cliente_nombre", "") or ""),
            "Vendedor": str(getattr(p, "vendedor", "") or ""),
            "CUIT recibo": _cuit_recibo_for_payment(p),
            "Medio de pago": p.medio_pago,
            "Fecha recibo": p.fecha_pago,
            "Importe recibo": round(float(p.importe_pago), 2),
            "Ítem en recibo": item_label,
        }

    # -----------------
    # ETAPA 1: VALIDADOS por peso (≤ umbral) con desempate por menor peso
    # -----------------
    edges_valid: List[Tuple[float, int, int]] = []  # (peso, i_payment, j_txn)
    for i, p in enumerate(relevant_payments):
        pd = dt.date.fromisoformat(p.fecha_pago)
        for j, t in enumerate(txns_match_scope):
            # Ventana temporal de negocio (V3.0): recibo vs banco
            if not _is_valid_time_window(pd, t.fecha):
                continue
            # Para VALIDADOS, el banco no puede llegar más de 2 días después del recibo
            if not _is_valid_for_validado(pd, t.fecha):
                continue
            c = _cost_for(t, p)
            # Si el ingreso ya estaba conciliado y hay penalización activa,
            # evitamos validarlo automáticamente (debe quedar para revisión manual).
            if t.was_preconciled and float(preconciled_penalty) > 0:
                continue
            if c <= valid_max_peso:
                edges_valid.append((c, i, j))

    edges_valid.sort(key=lambda x: x[0])

    used_payments: set[int] = set()
    used_txns: set[int] = set()
    used_txn_ids_validated: set[str] = set()

    validated_rows: List[dict] = []

    for c, i, j in edges_valid:
        if i in used_payments or j in used_txns:
            continue
        p = relevant_payments[i]
        t = txns_match_scope[j]
        row = {
            "Tipo fila": "PRINCIPAL",
            "Ranking": 1,
            **receipt_base_for(p),
            **_candidate_for(t, p),
        }
        validated_rows.append(row)
        used_payments.add(i)
        used_txns.add(j)
        used_txn_ids_validated.add(t.txn_id)

    mem_mark("stage1_validated", {"validated_edges": len(edges_valid), "validated_count": len(validated_rows)})

    # Remaining payments/txns for Hungarian
    rem_payments: List[ReceiptPayment] = [p for k, p in enumerate(relevant_payments) if k not in used_payments]
    rem_txns: List[BankTxn] = [t for k, t in enumerate(txns_match_scope) if k not in used_txns]

    # -----------------
    # ETAPA 2: Hungarian para el resto
    # -----------------
    suspect_rows: List[dict] = []
    assignment_by_payment: Dict[int, int | None] = {}

    n = len(rem_payments)
    m = len(rem_txns)
    CUTOFF_MULTIPLIER = 50.0
    UNASSIGNED_COST = 1e8

    if n > 0 and m > 0:
        candidates_by_payment: List[List[Tuple[float, int]]] = [[] for _ in range(n)]
        payments_by_txn: Dict[int, List[int]] = defaultdict(list)

        for i, p in enumerate(rem_payments):
            pd = dt.date.fromisoformat(p.fecha_pago)
            edges_i: List[Tuple[float, int]] = []
            max_di = _amount_tolerance_suspect(float(p.importe_pago)) * CUTOFF_MULTIPLIER
            for j, t in enumerate(rem_txns):
                if not _is_valid_time_window(pd, t.fecha):
                    continue
                di = abs(float(t.importe) - float(p.importe_pago))
                if di > max_di:
                    continue
                c = _cost_for(t, p)
                edges_i.append((float(c), j))

            if edges_i:
                edges_i.sort(key=lambda x: x[0])
                if int(stage2_candidate_top_k) > 0:
                    edges_i = edges_i[: int(stage2_candidate_top_k)]
                candidates_by_payment[i] = edges_i
                for _c, j in edges_i:
                    payments_by_txn[j].append(i)

        mem_mark(
            "stage2_candidates",
            {
                "rem_payments": n,
                "rem_txns": m,
                "edges_total": sum(len(x) for x in candidates_by_payment),
                "top_k": int(stage2_candidate_top_k),
            },
        )

        for i in range(n):
            assignment_by_payment[i] = None

        visited_p: set[int] = set()
        visited_t: set[int] = set()
        components: List[Tuple[List[int], List[int]]] = []

        for p0 in range(n):
            if p0 in visited_p or not candidates_by_payment[p0]:
                continue
            q = deque([("p", p0)])
            comp_p: set[int] = set()
            comp_t: set[int] = set()
            while q:
                kind, idx = q.popleft()
                if kind == "p":
                    if idx in visited_p:
                        continue
                    visited_p.add(idx)
                    comp_p.add(idx)
                    for _cost_val, tj in candidates_by_payment[idx]:
                        if tj not in visited_t:
                            q.append(("t", tj))
                else:
                    if idx in visited_t:
                        continue
                    visited_t.add(idx)
                    comp_t.add(idx)
                    for pi in payments_by_txn.get(idx, []):
                        if pi not in visited_p:
                            q.append(("p", pi))
            if comp_p:
                components.append((sorted(comp_p), sorted(comp_t)))

        total_cells = 0
        for comp_payments, comp_txns in components:
            npay = len(comp_payments)
            ntxn = len(comp_txns)
            if npay == 0 or ntxn == 0:
                continue
            size = max(npay, ntxn)
            total_cells += size * size
            cost_matrix = np.full((size, size), UNASSIGNED_COST, dtype=np.float32)

            pay_local_idx = {pidx: i_local for i_local, pidx in enumerate(comp_payments)}
            txn_local_idx = {tidx: j_local for j_local, tidx in enumerate(comp_txns)}
            for pidx in comp_payments:
                i_local = pay_local_idx[pidx]
                for c, tidx in candidates_by_payment[pidx]:
                    j_local = txn_local_idx.get(tidx)
                    if j_local is not None:
                        cost_matrix[i_local, j_local] = float(c)

            row_ind, col_ind = linear_sum_assignment(cost_matrix)
            for r, c in zip(row_ind.tolist(), col_ind.tolist()):
                if r >= npay:
                    continue
                pidx = comp_payments[r]
                if c >= ntxn or float(cost_matrix[r, c]) >= UNASSIGNED_COST:
                    assignment_by_payment[pidx] = None
                else:
                    assignment_by_payment[pidx] = comp_txns[c]

        mem_mark(
            "stage2_hungarian",
            {"components": len(components), "matrix_cells_total": total_cells},
        )
    else:
        for i in range(n):
            assignment_by_payment[i] = None

    # Process Hungarian assignments into DUDOSOS (and drop overly-high matches to NO_ENCONTRADO)
    for i, p in enumerate(rem_payments):
        j = assignment_by_payment.get(i)
        if j is None:
            continue
        t = rem_txns[j]
        case_id = f"{p.nro_recibo}|{p.nro_cliente}|{p.medio_pago}|{p.fecha_pago}|{i}"

        dd_signed = _signed_days(dt.date.fromisoformat(p.fecha_pago), t.fecha)
        dd = abs(dd_signed)
        di = abs(float(t.importe) - float(p.importe_pago))
        cost = _cost_for(t, p)

        if cost > dudoso_max_peso:
            assignment_by_payment[i] = None
            continue

        principal_row = {
            "Tipo fila": "PRINCIPAL",
            "Ranking": 1,
            "__case_id": case_id,
            **receipt_base_for(p),
            **_candidate_for(t, p),
        }

        principal_row["Motivo"] = _build_motivo(
            p.medio_pago,
            t.origen,
            dd,
            di,
            prefix="dudoso",
            was_preconciled=t.was_preconciled,
            preconciled_recibo=t.preconciled_recibo,
            medium_cross_enabled=(float(mp_mismatch_penalty) > 0),
        )
        suspect_rows.append(principal_row)

        # Alternatives: only if similar cost AND never use validated bank txns
        candidates: List[Tuple[float, BankTxn]] = []
        pd = dt.date.fromisoformat(p.fecha_pago)
        for t2 in rem_txns:
            if t2.txn_id == t.txn_id:
                continue
            if t2.txn_id in used_txn_ids_validated:
                continue
            if not _is_valid_time_window(pd, t2.fecha):
                continue
            di2 = abs(float(t2.importe) - float(p.importe_pago))
            if di2 > _amount_tolerance_suspect(float(p.importe_pago)) * CUTOFF_MULTIPLIER:
                continue
            c2 = _cost_for(t2, p)
            candidates.append((c2, t2))

        candidates.sort(key=lambda x: x[0])
        alt_rank = 2
        for c2, t2 in candidates:
            if alt_rank > (1 + max_alternatives):
                break
            if c2 > (cost + alternatives_cost_delta):
                break
            alt_row = {
                "Tipo fila": "ALTERNATIVO",
                "Ranking": alt_rank,
                "__case_id": case_id,
                "Nro recibo": "",
                "Nro cliente": "",
                "Cliente": "",
                "CUIT recibo": "",
                "Medio de pago": "",
                "Fecha recibo": "",
                "Importe recibo": "",
                "Vendedor": "",
                "Ítem en recibo": "",
                **_candidate_for(t2, p),
                "Motivo": _build_motivo(
                    p.medio_pago,
                    t2.origen,
                    abs(_signed_days(dt.date.fromisoformat(p.fecha_pago), t2.fecha)),
                    abs(float(t2.importe) - float(p.importe_pago)),
                    prefix="alternativo",
                    was_preconciled=t2.was_preconciled,
                    preconciled_recibo=t2.preconciled_recibo,
                    medium_cross_enabled=(float(mp_mismatch_penalty) > 0),
                ),
                            }
            suspect_rows.append(alt_row)
            alt_rank += 1

    # -----------------
    # NO ENCONTRADOS
    # -----------------
    no_rows: List[dict] = []

    # RECIBO_SIN_BANCO: remaining payments without an assigned txn
    for i, p in enumerate(rem_payments):
        if assignment_by_payment.get(i) is not None:
            continue
        best_cost = None
        pd = dt.date.fromisoformat(p.fecha_pago)
        for t in rem_txns:
            if t.txn_id in used_txn_ids_validated:
                continue
            if not _is_valid_time_window(pd, t.fecha):
                continue
            di = abs(float(t.importe) - float(p.importe_pago))
            if di > _amount_tolerance_suspect(float(p.importe_pago)) * CUTOFF_MULTIPLIER:
                continue
            c = _cost_for(t, p)
            if best_cost is None or c < best_cost:
                best_cost = c
        row = {
            "Tipo no encontrado": "RECIBO_SIN_BANCO",
            "Nro recibo": str(p.nro_recibo),
            "Nro cliente": str(p.nro_cliente),
            "Cliente": str(getattr(p, "cliente_nombre", "") or ""),
            "Vendedor": str(getattr(p, "vendedor", "") or ""),
            "CUIT recibo": _cuit_recibo_for_payment(p),
            "Medio de pago": p.medio_pago,
            "Fecha recibo": p.fecha_pago,
            "Importe recibo": round(float(p.importe_pago), 2),
            "Ítem en recibo": "",
            "CUIT ingreso": "",
            "Peso": (round(float(best_cost), 2) if best_cost is not None else ""),
        }
        no_rows.append(row)

    # BANCO_SIN_RECIBO (global) only if enabled
    if enable_banco_sin_recibo:
        used_txn_ids = set(used_txn_ids_validated)
        # also count Hungarian accepted matches
        for i, j in assignment_by_payment.items():
            if j is None:
                continue
            used_txn_ids.add(rem_txns[j].txn_id)

        for t in txns_report_scope:
            if t.txn_id in used_txn_ids:
                continue
            no_rows.append(
                {
                    "Tipo no encontrado": "BANCO_SIN_RECIBO",
                    "Cliente": "",
                    "Origen": t.origen,
                    "Fecha movimiento": t.fecha.isoformat(),
                    "Importe movimiento": round(float(t.importe), 2),
                    "Detalle movimiento": t.texto_ref,
                    "Vendedor": "",
                    "Ítem en recibo": "",
                    "CUIT recibo": "",
                    "CUIT ingreso": str(t.cuit or ""),
                    "Fila Excel": t.row_index,
                }
            )

    # Orden final
    # - Validados: por Nro recibo
    # - Dudosos: por menor Peso (revisar primero lo más cercano)
    validated_rows.sort(key=lambda r: _recibo_sort_key(str(r.get("Nro recibo", ""))))

    def _peso_key(row: dict) -> float:
        v = row.get("Peso", None)
        if isinstance(v, (int, float)):
            return float(v)
        try:
            return float(str(v).replace(",", "."))
        except Exception:
            return float("inf")

    suspect_rows.sort(key=_peso_key)

    result = {
        "validados": validated_rows,
        "dudosos": suspect_rows,
        "no_encontrados": no_rows,
    }
    if mem_debug:
        result["meta"] = {"matcher_mem_stages": mem_stages}
    return result
