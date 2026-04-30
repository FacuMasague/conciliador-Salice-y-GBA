from __future__ import annotations

import datetime as dt
from collections import defaultdict, deque
from typing import Dict, List, Tuple, Optional

import numpy as np
from scipy.optimize import linear_sum_assignment

from .excel_loader import BankTxn
from .pdf_parser import ReceiptPayment
from .memdebug import mem_debug_recorder
from .utils import _normalize_cliente, _normalize_cuit, _normalize_recibo


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


def _amount_difference_penalty(diff_amount: float) -> float:
    """Penalización convexa por diferencia de importe.

    Objetivo operativo:
    - 0..1 peso: penalización baja, pero no nula.
    - 1..10 pesos: sube de forma gradual.
    - >10 pesos: crece cada vez más rápido.

    Esto evita que diferencias medianas de importe queden "baratas"
    frente a una simple diferencia de días.
    """
    di = max(0.0, float(diff_amount))
    if di <= 1.0:
        return 0.35 * di
    if di <= 10.0:
        x = di - 1.0
        return 0.35 + (0.65 * x) + (0.09 * x * x)
    y = di - 10.0
    base = 0.35 + (0.65 * 9.0) + (0.09 * 9.0 * 9.0)
    return base + (2.2 * y) + (0.11 * y * y)


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


def _empresa_bank_cross_penalty(
    empresa: str,
    origen: str,
    *,
    penalty_salice_to_galicia: float,
    penalty_alarcon_to_bbva: float,
) -> float:
    """Penaliza asignaciones "banco cruzado" por empresa/banco.

    Regla de negocio (V3.1):
    - Recibos de ALARCON tienen prioridad en ingresos GALICIA.
    - Recibos de SALICE tienen prioridad en ingresos BBVA.

    No se prohíbe el cruce (puede pasar), pero se lo vuelve menos preferible.
    """
    e = (empresa or "").strip().upper()
    o = (origen or "").strip().upper()
    if e == "SALICE" and o == "GALICIA":
        return float(penalty_salice_to_galicia)
    if e == "ALARCON" and o == "BBVA":
        return float(penalty_alarcon_to_bbva)
    return 0.0


def _payment_is_bankable_for_stage2(medio_pago: object) -> bool:
    txt = str(medio_pago or "").strip().lower()
    if txt in {"", "no_informado", "sin_medio_api", "no informado"}:
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


def _payment_is_bankable(medio_pago: object) -> bool:
    return _payment_is_bankable_for_stage2(medio_pago)


def _is_same_preconciled_receipt(t: BankTxn, p: ReceiptPayment) -> bool:
    if not t.was_preconciled:
        return False
    prev_rec = _normalize_recibo(t.preconciled_recibo)
    curr_rec = _normalize_recibo(p.nro_recibo)
    return bool(prev_rec and curr_rec and prev_rec == curr_rec)


def _has_cliente_cuit_mismatch(
    nro_cliente: object,
    cuit_movimiento: object,
    cliente_to_cuit_map: Dict[str, str] | None,
) -> bool:
    if not cliente_to_cuit_map:
        return False
    cli = _normalize_cliente(nro_cliente)
    if not cli:
        return False
    expected = _normalize_cuit(cliente_to_cuit_map.get(cli))
    current = _normalize_cuit(cuit_movimiento)
    # V4.3.1:
    # también penalizamos cuando falta uno de los dos CUIT (ingreso o recibo).
    return expected != current


def _cliente_cuit_matches(
    nro_cliente: object,
    cuit_movimiento: object,
    cliente_to_cuit_map: Dict[str, str] | None,
) -> bool:
    if not cliente_to_cuit_map:
        return False
    cli = _normalize_cliente(nro_cliente)
    current = _normalize_cuit(cuit_movimiento)
    if not cli or not current:
        return False
    expected = _normalize_cuit(cliente_to_cuit_map.get(cli))
    return bool(expected and expected == current)


def _cost(
    t: BankTxn,
    p: ReceiptPayment,
    *,
    day_weight_bank_before: float,
    day_weight_bank_after: float,
    mp_mismatch_penalty: float,
    penalty_salice_to_galicia: float,
    penalty_alarcon_to_bbva: float,
    preconciled_penalty: float,
    cliente_to_cuit_map: Dict[str, str] | None = None,
    cliente_cuit_mismatch_penalty: float = 0.0,
    non_bankable_receipt_cost_multiplier: float = 1.0,
) -> float:
    pd = dt.date.fromisoformat(p.fecha_pago)
    dd_signed = _signed_days(pd, t.fecha)
    dd = abs(dd_signed)
    # V3.5: multiplicador de días según signo.
    #  >= 0: ingreso bancario anterior (o mismo día) que el recibo
    #   < 0: ingreso bancario posterior al recibo (delay del banco)
    dw = float(day_weight_bank_before) if dd_signed >= 0 else float(day_weight_bank_after)
    di = abs(float(t.importe) - float(p.importe_pago))
    total = (
        float(dw) * float(dd)
        + _amount_difference_penalty(di)
        + _mp_mismatch_penalty(p.medio_pago, t.origen, mp_mismatch_penalty)
        + _empresa_bank_cross_penalty(
            p.empresa,
            t.origen,
            penalty_salice_to_galicia=penalty_salice_to_galicia,
            penalty_alarcon_to_bbva=penalty_alarcon_to_bbva,
        )
        + (float(preconciled_penalty) if (t.was_preconciled and not _is_same_preconciled_receipt(t, p)) else 0.0)
    )
    if not _payment_is_bankable_for_stage2(getattr(p, "medio_pago", "")):
        total *= float(non_bankable_receipt_cost_multiplier)
    return total


def _build_candidate(
    t: BankTxn,
    p: ReceiptPayment,
    *,
    day_weight_bank_before: float,
    day_weight_bank_after: float,
    mp_mismatch_penalty: float,
    penalty_salice_to_galicia: float,
    penalty_alarcon_to_bbva: float,
    preconciled_penalty: float,
    cliente_to_cuit_map: Dict[str, str] | None = None,
    cliente_cuit_mismatch_penalty: float = 0.0,
    non_bankable_receipt_cost_multiplier: float = 1.0,
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
        penalty_salice_to_galicia=penalty_salice_to_galicia,
        penalty_alarcon_to_bbva=penalty_alarcon_to_bbva,
        preconciled_penalty=preconciled_penalty,
        cliente_to_cuit_map=cliente_to_cuit_map,
        cliente_cuit_mismatch_penalty=cliente_cuit_mismatch_penalty,
        non_bankable_receipt_cost_multiplier=non_bankable_receipt_cost_multiplier,
    )
    return {
        "Origen": t.origen,
        "Fecha movimiento": t.fecha.isoformat(),
        "Importe movimiento": round(float(t.importe), 2),
        "Detalle movimiento": t.texto_ref,
        "CUIT ingreso": str(t.cuit or ""),
        "Fila Excel": t.row_index,
        "__txn_id": t.txn_id,
        "__sheet_name": str(t.sheet_name or ""),
        "__record_key": str(t.record_key or ""),
        # Nota: dejamos el signo para poder auditar rápidamente el sentido.
        # Negativo => el banco impactó después del recibo (delay del banco).
        "Dif días": int(dd_signed),
        "Dif importe": round(float(di), 2),
        "Peso": round(float(cost), 2),
    }


def _build_motivo(
    empresa: str,
    p_medio: str,
    origen: str,
    dd: int,
    di: float,
    *,
    prefix: str,
    was_preconciled: bool = False,
    preconciled_recibo: str | None = None,
    cliente_cuit_mismatch: bool = False,
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

    # Banco cruzado por empresa (V3.1)
    e = (empresa or "").strip().upper()
    o = (origen or "").strip().upper()
    if e == "SALICE" and o == "GALICIA":
        why.append("banco cruzado")
    if e == "ALARCON" and o == "BBVA":
        why.append("banco cruzado")
    if cliente_cuit_mismatch:
        why.append("cliente/cuit no coincide")
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
    day_weight_bank_before: float = 20.0,
    day_weight_bank_after: float = 35.0,
    valid_max_peso: float = 260.0,
    dudoso_max_peso: float = 3500.0,
    mp_mismatch_penalty: float = 35.0,
    preconciled_penalty: float = 150.0,
    penalty_salice_to_galicia: float = 45.0,
    penalty_alarcon_to_bbva: float = 45.0,
    cliente_cuit_mismatch_penalty: float = 0.0,
    cliente_to_cuit_map: Dict[str, str] | None = None,
    max_alternatives: int = 3,
    report_date_min: str | None = None,
    report_date_max: str | None = None,
    current_date_override: str | None = None,
    include_empresa: bool = False,
    enable_banco_sin_recibo: bool = True,
    banco_sin_recibo_grace_days: int = 0,
    recibo_sin_banco_grace_days: int = 0,
    alternatives_cost_delta: float = 35.0,
    stage2_candidate_top_k: int = 120,
    stage2_matrix_cells_limit: int = 250000,
    exclude_preconciled_txns: bool = False,
    mem_debug: bool = False,
    validated_allow_all_receipts: bool = False,
    non_bankable_receipt_cost_multiplier: float = 1.0,
    suspects_and_no_bankable_only: bool = False,
    no_encontrados_bankable_only: bool = False,
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
    MAX_BANK_AFTER_RECEIPT_DAYS = 2

    def _is_valid_time_window(receipt_date: dt.date, bank_date: dt.date) -> bool:
        sd = _signed_days(receipt_date, bank_date)
        return (-MAX_BANK_AFTER_RECEIPT_DAYS) <= sd <= MAX_BANK_BEFORE_RECEIPT_DAYS

    def _index_txns_by_date(txns_list: List[BankTxn]) -> Dict[dt.date, List[Tuple[int, BankTxn]]]:
        out: Dict[dt.date, List[Tuple[int, BankTxn]]] = defaultdict(list)
        for idx, txn in enumerate(txns_list):
            out[txn.fecha].append((idx, txn))
        return out

    def _candidate_txns_for_payment(
        p: ReceiptPayment,
        txns_by_date: Dict[dt.date, List[Tuple[int, BankTxn]]],
    ) -> List[Tuple[int, BankTxn]]:
        pd = dt.date.fromisoformat(p.fecha_pago)
        out: List[Tuple[int, BankTxn]] = []
        for delta in range(-MAX_BANK_BEFORE_RECEIPT_DAYS, MAX_BANK_AFTER_RECEIPT_DAYS + 1):
            out.extend(txns_by_date.get(pd + dt.timedelta(days=delta), []))
        return out

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
    CUTOFF_MULTIPLIER = 50.0

    def in_match_scope(d: dt.date) -> bool:
        return (min_d - dt.timedelta(days=effective_margin)) <= d <= (max_d + dt.timedelta(days=effective_margin))

    txns_match_scope = [t for t in txns if t.parse_ok and in_match_scope(t.fecha)]
    if exclude_preconciled_txns:
        txns_match_scope = [t for t in txns_match_scope if not t.was_preconciled]
    txns_match_scope_by_date = _index_txns_by_date(txns_match_scope)

    # Strict report range (for BANCO_SIN_RECIBO)
    report_min = dt.date.fromisoformat(report_date_min) if report_date_min else min_d
    report_max = dt.date.fromisoformat(report_date_max) if report_date_max else max_d
    if current_date_override:
        current_date = dt.date.fromisoformat(str(current_date_override))
    else:
        current_date = dt.date.today()

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

    def _is_plausible_candidate(t: BankTxn, p: ReceiptPayment) -> bool:
        pd = dt.date.fromisoformat(p.fecha_pago)
        if not _is_valid_time_window(pd, t.fecha):
            return False
        di = abs(float(t.importe) - float(p.importe_pago))
        max_di = _amount_tolerance_suspect(float(p.importe_pago)) * CUTOFF_MULTIPLIER
        return di <= max_di

    txn_requires_matching_cuit: Dict[str, bool] = {t.txn_id: False for t in txns_match_scope}
    plausible_match_candidates_by_payment: List[List[Tuple[int, BankTxn]]] = []
    for p in relevant_payments:
        candidates_p: List[Tuple[int, BankTxn]] = []
        for j, t in _candidate_txns_for_payment(p, txns_match_scope_by_date):
            if not _is_plausible_candidate(t, p):
                continue
            candidates_p.append((j, t))
            if _cliente_cuit_matches(p.nro_cliente, t.cuit, cliente_to_cuit_map):
                txn_requires_matching_cuit[t.txn_id] = True
        plausible_match_candidates_by_payment.append(candidates_p)

    def _pair_allowed(t: BankTxn, p: ReceiptPayment) -> bool:
        if not txn_requires_matching_cuit.get(t.txn_id, False):
            return True
        return _cliente_cuit_matches(p.nro_cliente, t.cuit, cliente_to_cuit_map)

    def _cost_for(t: BankTxn, p: ReceiptPayment) -> float:
        return _cost(
            t,
            p,
            day_weight_bank_before=day_weight_bank_before,
            day_weight_bank_after=day_weight_bank_after,
            mp_mismatch_penalty=mp_mismatch_penalty,
            penalty_salice_to_galicia=penalty_salice_to_galicia,
            penalty_alarcon_to_bbva=penalty_alarcon_to_bbva,
            preconciled_penalty=preconciled_penalty,
            cliente_to_cuit_map=cliente_to_cuit_map,
            cliente_cuit_mismatch_penalty=cliente_cuit_mismatch_penalty,
            non_bankable_receipt_cost_multiplier=non_bankable_receipt_cost_multiplier,
        )

    def _candidate_for(t: BankTxn, p: ReceiptPayment) -> dict:
        return _build_candidate(
            t,
            p,
            day_weight_bank_before=day_weight_bank_before,
            day_weight_bank_after=day_weight_bank_after,
            mp_mismatch_penalty=mp_mismatch_penalty,
            penalty_salice_to_galicia=penalty_salice_to_galicia,
            penalty_alarcon_to_bbva=penalty_alarcon_to_bbva,
            preconciled_penalty=preconciled_penalty,
            cliente_to_cuit_map=cliente_to_cuit_map,
            cliente_cuit_mismatch_penalty=cliente_cuit_mismatch_penalty,
            non_bankable_receipt_cost_multiplier=non_bankable_receipt_cost_multiplier,
        )

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
            **({"Empresa": str(p.empresa)} if include_empresa else {}),
            "Nro recibo": str(p.nro_recibo),
            "Nro cliente": str(p.nro_cliente),
            "Cliente": str(getattr(p, "cliente_nombre", "") or ""),
            "CUIT recibo": _cuit_recibo_for_payment(p),
            "Medio de pago": p.medio_pago,
            "Fecha recibo": p.fecha_pago,
            "Importe recibo": round(float(p.importe_pago), 2),
            "Divisor": item_label,
            "__payment_lookup_key": str(id(p)),
        }

    # -----------------
    # ETAPA 1: VALIDADOS por peso (≤ umbral) con matching óptimo
    # -----------------
    valid_candidates_by_payment: List[List[Tuple[float, int]]] = [[] for _ in range(len(relevant_payments))]
    payments_by_valid_txn: Dict[int, List[int]] = defaultdict(list)
    for i, p in enumerate(relevant_payments):
        if (not validated_allow_all_receipts) and (not _payment_is_bankable(getattr(p, "medio_pago", ""))):
            continue
        edges_i: List[Tuple[float, int]] = []
        for j, t in plausible_match_candidates_by_payment[i]:
            if not _pair_allowed(t, p):
                continue
            c = _cost_for(t, p)
            if c <= valid_max_peso:
                edges_i.append((float(c), j))
        edges_i.sort(key=lambda x: x[0])
        valid_candidates_by_payment[i] = edges_i
        for _c, j in edges_i:
            payments_by_valid_txn[j].append(i)

    used_payments: set[int] = set()
    used_txns: set[int] = set()
    used_txn_ids_validated: set[str] = set()
    validated_rows: List[dict] = []

    visited_p_valid: set[int] = set()
    visited_t_valid: set[int] = set()
    valid_components: List[Tuple[List[int], List[int]]] = []
    for p0 in range(len(relevant_payments)):
        if p0 in visited_p_valid or not valid_candidates_by_payment[p0]:
            continue
        q = deque([("p", p0)])
        comp_p: set[int] = set()
        comp_t: set[int] = set()
        while q:
            kind, idx = q.popleft()
            if kind == "p":
                if idx in visited_p_valid:
                    continue
                visited_p_valid.add(idx)
                comp_p.add(idx)
                for _cost_val, tj in valid_candidates_by_payment[idx]:
                    if tj not in visited_t_valid:
                        q.append(("t", tj))
            else:
                if idx in visited_t_valid:
                    continue
                visited_t_valid.add(idx)
                comp_t.add(idx)
                for pi in payments_by_valid_txn.get(idx, []):
                    if pi not in visited_p_valid:
                        q.append(("p", pi))
        if comp_p:
            valid_components.append((sorted(comp_p), sorted(comp_t)))

    dummy_cost = float(valid_max_peso) + 1.0
    for comp_payments, comp_txns in valid_components:
        npay = len(comp_payments)
        ntxn = len(comp_txns)
        if npay == 0:
            continue
        cost_matrix = np.full((npay, ntxn + npay), dummy_cost, dtype=np.float32)
        pay_local_idx = {pidx: i_local for i_local, pidx in enumerate(comp_payments)}
        txn_local_idx = {tidx: j_local for j_local, tidx in enumerate(comp_txns)}
        for pidx in comp_payments:
            i_local = pay_local_idx[pidx]
            for c, tidx in valid_candidates_by_payment[pidx]:
                j_local = txn_local_idx.get(tidx)
                if j_local is not None:
                    cost_matrix[i_local, j_local] = float(c)

        row_ind, col_ind = linear_sum_assignment(cost_matrix)
        for r_local, c_local in zip(row_ind.tolist(), col_ind.tolist()):
            if r_local >= npay or c_local >= ntxn:
                continue
            chosen_cost = float(cost_matrix[r_local, c_local])
            if chosen_cost > float(valid_max_peso):
                continue
            i = comp_payments[r_local]
            j = comp_txns[c_local]
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

    mem_mark(
        "stage1_validated",
        {
            "validated_components": len(valid_components),
            "validated_edges": sum(len(x) for x in valid_candidates_by_payment),
            "validated_count": len(validated_rows),
        },
    )

    # Remaining payments/txns for Hungarian
    rem_payments: List[ReceiptPayment] = [
        p
        for k, p in enumerate(relevant_payments)
        if k not in used_payments
        and (
            not suspects_and_no_bankable_only
            or _payment_is_bankable_for_stage2(getattr(p, "medio_pago", ""))
        )
    ]
    rem_txns: List[BankTxn] = [t for k, t in enumerate(txns_match_scope) if k not in used_txns]

    # -----------------
    # ETAPA 2: Hungarian para el resto
    # -----------------
    suspect_rows: List[dict] = []
    assignment_by_payment: Dict[int, int | None] = {}
    candidates_by_payment: List[List[Tuple[float, int]]] = []

    n = len(rem_payments)
    m = len(rem_txns)
    UNASSIGNED_COST = 1e8

    if n > 0 and m > 0:
        candidates_by_payment = [[] for _ in range(n)]
        payments_by_txn: Dict[int, List[int]] = defaultdict(list)
        rem_txns_by_date = _index_txns_by_date(rem_txns)

        for i, p in enumerate(rem_payments):
            edges_i: List[Tuple[float, int]] = []
            for j, t in _candidate_txns_for_payment(p, rem_txns_by_date):
                if not _pair_allowed(t, p):
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
            if int(stage2_matrix_cells_limit) > 0 and (size * size) > int(stage2_matrix_cells_limit):
                # Fallback pragmático: evita que un componente enorme deje la corrida colgada.
                # Asigna greedy por menor costo respetando unicidad de pago/txn.
                greedy_edges: list[tuple[float, int, int]] = []
                for pidx in comp_payments:
                    for c, tidx in candidates_by_payment[pidx]:
                        if tidx in comp_txns:
                            greedy_edges.append((float(c), pidx, tidx))
                greedy_edges.sort(key=lambda x: x[0])
                used_local_p: set[int] = set()
                used_local_t: set[int] = set()
                for c, pidx, tidx in greedy_edges:
                    if pidx in used_local_p or tidx in used_local_t:
                        continue
                    assignment_by_payment[pidx] = tidx
                    used_local_p.add(pidx)
                    used_local_t.add(tidx)
                continue
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
            p.empresa,
            p.medio_pago,
            t.origen,
            dd,
            di,
            prefix="dudoso",
            was_preconciled=t.was_preconciled,
            preconciled_recibo=(None if _is_same_preconciled_receipt(t, p) else t.preconciled_recibo),
            cliente_cuit_mismatch=_has_cliente_cuit_mismatch(p.nro_cliente, t.cuit, cliente_to_cuit_map),
            medium_cross_enabled=(float(mp_mismatch_penalty) > 0),
        )
        suspect_rows.append(principal_row)

        # Alternatives: only if similar cost AND never use validated bank txns
        candidates: List[Tuple[float, BankTxn]] = []
        for c2, j2 in candidates_by_payment[i]:
            t2 = rem_txns[j2]
            if t2.txn_id == t.txn_id:
                continue
            if t2.txn_id in used_txn_ids_validated:
                continue
            candidates.append((c2, t2))

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
                **({"Empresa": ""} if include_empresa else {}),
                "Nro recibo": "",
                "Nro cliente": "",
                "Cliente": "",
                "CUIT recibo": "",
                "Medio de pago": "",
                "Fecha recibo": "",
                "Importe recibo": "",
                "Divisor": "",
                **_candidate_for(t2, p),
                "Motivo": _build_motivo(
                    p.empresa,
                    p.medio_pago,
                    t2.origen,
                    abs(_signed_days(dt.date.fromisoformat(p.fecha_pago), t2.fecha)),
                    abs(float(t2.importe) - float(p.importe_pago)),
                    prefix="alternativo",
                    was_preconciled=t2.was_preconciled,
                    preconciled_recibo=(None if _is_same_preconciled_receipt(t2, p) else t2.preconciled_recibo),
                    cliente_cuit_mismatch=_has_cliente_cuit_mismatch(p.nro_cliente, t2.cuit, cliente_to_cuit_map),
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
    recibo_sin_banco_cutoff = min(
        report_max,
        current_date - dt.timedelta(days=max(int(recibo_sin_banco_grace_days), 0)),
    )
    for i, p in enumerate(rem_payments):
        if assignment_by_payment.get(i) is not None:
            continue
        if no_encontrados_bankable_only and (not _payment_is_bankable(getattr(p, "medio_pago", ""))):
            continue
        try:
            pdate = dt.date.fromisoformat(str(p.fecha_pago))
        except Exception:
            pdate = report_min
        if pdate > recibo_sin_banco_cutoff:
            continue
        best_cost = candidates_by_payment[i][0][0] if i < len(candidates_by_payment) and candidates_by_payment[i] else None
        row = {
            "Tipo no encontrado": "RECIBO_SIN_BANCO",
            **({"Empresa": str(p.empresa)} if include_empresa else {}),
            "Nro recibo": str(p.nro_recibo),
            "Nro cliente": str(p.nro_cliente),
            "Cliente": str(getattr(p, "cliente_nombre", "") or ""),
            "CUIT recibo": _cuit_recibo_for_payment(p),
            "Medio de pago": p.medio_pago,
            "Fecha recibo": p.fecha_pago,
            "Importe recibo": round(float(p.importe_pago), 2),
            "Divisor": "",
            "CUIT ingreso": "",
            "Peso": (round(float(best_cost), 2) if best_cost is not None else ""),
            "__payment_lookup_key": str(id(p)),
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

        banco_sin_recibo_cutoff = min(
            report_max,
            current_date - dt.timedelta(days=max(int(banco_sin_recibo_grace_days), 0)),
        )
        for t in txns_report_scope:
            if t.txn_id in used_txn_ids:
                continue
            if t.fecha > banco_sin_recibo_cutoff:
                continue
            no_rows.append(
                {
                    "Tipo no encontrado": "BANCO_SIN_RECIBO",
                    "Cliente": "",
                    "Origen": t.origen,
                    "Fecha movimiento": t.fecha.isoformat(),
                    "Importe movimiento": round(float(t.importe), 2),
                    "Detalle movimiento": t.texto_ref,
                    "Divisor": "",
                    "CUIT recibo": "",
                    "CUIT ingreso": str(t.cuit or ""),
                    "Fila Excel": t.row_index,
                    "__sheet_name": str(t.sheet_name or ""),
                    "__record_key": str(t.record_key or ""),
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
