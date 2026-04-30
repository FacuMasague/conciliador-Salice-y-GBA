from __future__ import annotations

import datetime as dt

import openpyxl
import pandas as pd

from src.conciliador.excel_loader import BankTxn
from src.conciliador.matcher_hungarian import match_hungarian
from src.conciliador.pdf_parser import ReceiptPayment
from src.conciliador.pipeline import _load_cliente_cuit_map
from src.conciliador.excel_loader import load_bank_txns


def _sample_txn(*, cuit: str) -> BankTxn:
    return BankTxn(
        txn_id="T1",
        origen="BBVA",
        fecha=dt.date(2026, 2, 10),
        hora=None,
        importe=1000.0,
        texto_ref=f"TRANSFERENCIA {cuit}",
        row_index=10,
        parse_ok=True,
        parse_error=None,
        was_preconciled=False,
        preconciled_recibo=None,
        cuit=cuit,
    )


def _sample_payment() -> ReceiptPayment:
    return ReceiptPayment(
        empresa="SALICE",
        nro_recibo="100",
        nro_cliente="1234",
        cliente_nombre="Cliente Test",
        medio_pago="TRANSFERENCIA",
        fecha_pago="2026-02-10",
        importe_pago=1000.0,
        vendedor="203 - Edgardo Larrea",
    )


def test_matcher_cuit_prefers_correct_cuit_over_wrong_cuit():
    """Si el CUIT esperado del recibo aparece en algún txn, ese txn se prefiere.
    El txn con CUIT distinto NO se bloquea pero tiene mayor costo (penalidad moderada).
    Con ambos disponibles, el algoritmo debe elegir el de CUIT correcto.
    """
    txn_correct = BankTxn(
        txn_id="T_correct",
        origen="BBVA",
        fecha=dt.date(2026, 2, 10),
        hora=None,
        importe=1000.0,
        texto_ref="TRANSFERENCIA 20301020304",
        row_index=10,
        parse_ok=True,
        parse_error=None,
        was_preconciled=False,
        preconciled_recibo=None,
        cuit="20301020304",
    )
    txn_wrong = BankTxn(
        txn_id="T_wrong",
        origen="BBVA",
        fecha=dt.date(2026, 2, 10),
        hora=None,
        importe=1000.0,
        texto_ref="TRANSFERENCIA 20999999999",
        row_index=11,
        parse_ok=True,
        parse_error=None,
        was_preconciled=False,
        preconciled_recibo=None,
        cuit="20999999999",
    )
    # Con ambos disponibles, el algoritmo debe elegir el de CUIT correcto (costo 0 vs 75).
    res = match_hungarian(
        [txn_wrong, txn_correct],
        [_sample_payment()],
        valid_max_peso=200,
        dudoso_max_peso=500,
        cliente_to_cuit_map={"1234": "20301020304"},
        enable_banco_sin_recibo=False,
    )
    assert len(res["validados"]) == 1
    assert res["validados"][0]["CUIT ingreso"] == "20301020304"


def test_matcher_cuit_wrong_cuit_still_matches_when_correct_unavailable():
    """Cuando el txn con CUIT correcto no está disponible, el txn con CUIT distinto
    puede matchear (con penalidad). Antes era hard-bloqueado; ahora se permite."""
    txn_wrong = BankTxn(
        txn_id="T_wrong",
        origen="BBVA",
        fecha=dt.date(2026, 2, 10),
        hora=None,
        importe=1000.0,
        texto_ref="TRANSFERENCIA 20999999999",
        row_index=11,
        parse_ok=True,
        parse_error=None,
        was_preconciled=False,
        preconciled_recibo=None,
        cuit="20999999999",
    )
    txn_correct_for_other_receipt = BankTxn(
        txn_id="T_correct",
        origen="BBVA",
        fecha=dt.date(2026, 2, 10),
        hora=None,
        importe=1000.0,
        texto_ref="TRANSFERENCIA 20301020304",
        row_index=10,
        parse_ok=True,
        parse_error=None,
        was_preconciled=False,
        preconciled_recibo=None,
        cuit="20301020304",
    )
    # txn_correct ya está "ocupado" por otro recibo. Solo txn_wrong disponible.
    # Con penalidad moderada (75) y valid_max_peso=200, debe seguir matcheando.
    from src.conciliador.pdf_parser import ReceiptPayment
    other_payment = ReceiptPayment(
        empresa="SALICE", nro_recibo="999", nro_cliente="9999",
        cliente_nombre="Otro", medio_pago="TRANSFERENCIA",
        fecha_pago="2026-02-10", importe_pago=1000.0, vendedor="",
    )
    res = match_hungarian(
        [txn_wrong, txn_correct_for_other_receipt],
        [_sample_payment(), other_payment],
        valid_max_peso=200,
        dudoso_max_peso=500,
        cuit_mismatch_penalty=75.0,
        cliente_to_cuit_map={"1234": "20301020304", "9999": "20000000000"},
        enable_banco_sin_recibo=False,
    )
    # Ambos recibos deben matchear: 1234→T_correct (costo 0), 9999→T_wrong (costo 75)
    assert len(res["validados"]) == 2


def test_matcher_cuit_exclusivity_matches_correct_cuit():
    """Cuando CUIT coincide, el match se produce con costo 0."""
    res = match_hungarian(
        [_sample_txn(cuit="20301020304")],
        [_sample_payment()],
        valid_max_peso=200,
        dudoso_max_peso=500,
        cliente_to_cuit_map={"1234": "20301020304"},
        enable_banco_sin_recibo=False,
    )
    assert len(res["validados"]) == 1
    assert res["validados"][0]["Peso"] == 0.0


def test_matcher_cuit_exclusivity_no_restriction_when_cuit_not_in_txns():
    """Si el CUIT esperado del recibo no aparece en ningún txn, no hay restricción."""
    # El txn tiene CUIT 20301020304 pero el cliente espera 20111222333 que no está en ningún txn.
    res = match_hungarian(
        [_sample_txn(cuit="20301020304")],
        [_sample_payment()],
        valid_max_peso=200,
        dudoso_max_peso=500,
        cliente_to_cuit_map={"1234": "20111222333"},
        enable_banco_sin_recibo=False,
    )
    # Sin restricción → match permitido
    assert len(res["validados"]) == 1
    assert res["validados"][0]["Peso"] == 0.0


def test_matcher_cuit_exclusivity_no_restriction_when_txn_cuit_empty():
    """Txn sin CUIT se permite siempre, aunque otro txn tenga el CUIT esperado en el scope."""
    # Hay un txn con el CUIT correcto (popula txn_cuit_set) y uno sin CUIT.
    # El recibo debería poder matchear el txn sin CUIT también.
    txn_with_cuit = BankTxn(
        txn_id="T_cuit",
        origen="BBVA",
        fecha=dt.date(2026, 2, 10),
        hora=None,
        importe=999.0,  # importe distinto para que no se lleve el match
        texto_ref="TRANSFERENCIA 20301020304",
        row_index=9,
        parse_ok=True,
        parse_error=None,
        was_preconciled=False,
        preconciled_recibo=None,
        cuit="20301020304",
    )
    txn_no_cuit = BankTxn(
        txn_id="T_no_cuit",
        origen="BBVA",  # mismo banco para evitar mp_mismatch_penalty
        fecha=dt.date(2026, 2, 10),
        hora=None,
        importe=1000.0,  # importe exacto → match ideal (costo=0 vs costo=1 del otro)
        texto_ref="TRANSFERENCIA SIN CUIT",
        row_index=10,
        parse_ok=True,
        parse_error=None,
        was_preconciled=False,
        preconciled_recibo=None,
        cuit=None,
    )
    # txn_with_cuit popula txn_cuit_set con "20301020304".
    # El recibo espera "20301020304". txn_no_cuit no tiene CUIT → debe ser permitido.
    res = match_hungarian(
        [txn_with_cuit, txn_no_cuit],
        [_sample_payment()],
        valid_max_peso=200,
        dudoso_max_peso=500,
        cliente_to_cuit_map={"1234": "20301020304"},
        enable_banco_sin_recibo=False,
    )
    assert len(res["validados"]) == 1
    # El match ideal es con txn_no_cuit (importe exacto = 0 dif), no con txn_with_cuit (dif $1)
    assert res["validados"][0]["CUIT ingreso"] == ""


def test_no_encontrados_include_cuit_fields():
    res = match_hungarian(
        [],
        [_sample_payment()],
        valid_max_peso=200,
        dudoso_max_peso=500,
        cliente_to_cuit_map={"1234": "20301020304"},
        enable_banco_sin_recibo=False,
    )
    assert len(res["no_encontrados"]) == 1
    row = res["no_encontrados"][0]
    assert row["Vendedor"] == "203 - Edgardo Larrea"
    assert "Aclaración recibo" not in row
    assert row["CUIT recibo"] == "20301020304"
    assert row["CUIT ingreso"] == ""


def test_load_cliente_cuit_map_reads_expected_columns(tmp_path):
    padron_path = tmp_path / "Padron Basico MDP.xlsx"
    df = pd.DataFrame(
        {
            "ClienteID": ["001234", "5678", "9999"],
            "RazonSocial": ["A", "B", "C"],
            "NumeroDeDocumento": ["20-11112222-3", " - ", "30-12345678-9"],
        }
    )
    df.to_excel(padron_path, index=False)

    m = _load_cliente_cuit_map(str(padron_path))
    assert m["1234"] == "20111122223"
    assert m["9999"] == "30123456789"
    assert "5678" not in m


def test_excel_loader_mp_prefers_cuit_column_over_operacion_relacionada(tmp_path):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "MercadoPago "
    ws.append(
        [
            "Fecha de Pago",
            "Tipo de Operación",
            "Número de Movimiento",
            "Operación Relacionada",
            "Unnamed: 4",
            "Recibio?",
            "Cliente",
            "Recibo",
            "NÚMERO DE IDENTIFICACIÓN DEL PAGADOR",
        ]
    )
    ws.append(
        [
            "20/02/2026 00:00:00",
            "Cobro",
            "",
            "145515211432",  # ID operación (no es CUIT)
            1000.0,
            "",
            "",
            "",
            "20-95448273-2",  # CUIT correcto
        ]
    )
    p = tmp_path / "mp_runtime.xlsx"
    wb.save(p)

    txns = load_bank_txns(str(p))
    mp = [t for t in txns if t.origen == "MERCADOPAGO"]
    assert len(mp) == 1
    assert mp[0].cuit == "20954482732"


def test_excel_loader_mp_does_not_infer_cuit_from_operacion_id_when_column_missing(tmp_path):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "MercadoPago "
    ws.append(
        [
            "Fecha de Pago",
            "Tipo de Operación",
            "Número de Movimiento",
            "Operación Relacionada",
            "Unnamed: 4",
            "Recibio?",
            "Cliente",
            "Recibo",
        ]
    )
    ws.append(["20/02/2026 00:00:00", "Cobro", "", "145515211432", 1000.0, "", "", ""])
    p = tmp_path / "mp_runtime_no_cuit_col.xlsx"
    wb.save(p)

    txns = load_bank_txns(str(p))
    mp = [t for t in txns if t.origen == "MERCADOPAGO"]
    assert len(mp) == 1
    assert mp[0].cuit is None
