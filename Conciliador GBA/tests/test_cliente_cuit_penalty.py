from __future__ import annotations

import datetime as dt

import openpyxl
import pandas as pd
import pytest

from src.conciliador.excel_loader import BankTxn
from src.conciliador.matcher_hungarian import _amount_difference_penalty, match_hungarian
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


def test_matcher_restricts_txn_to_same_cuit_when_plausible_match_exists():
    res = match_hungarian(
        [_sample_txn(cuit="20301020304")],
        [
            _sample_payment(),
            ReceiptPayment(
                empresa="SALICE",
                nro_recibo="101",
                nro_cliente="9999",
                cliente_nombre="Cliente Mismo Cuit",
                medio_pago="TRANSFERENCIA",
                fecha_pago="2026-02-10",
                importe_pago=1020.0,
                vendedor="203 - Edgardo Larrea",
            ),
        ],
        valid_max_peso=200,
        dudoso_max_peso=500,
        cliente_to_cuit_map={"1234": "20111222333", "9999": "20301020304"},
        cliente_cuit_mismatch_penalty=50.0,
        enable_banco_sin_recibo=False,
    )
    assert len(res["validados"]) == 1
    assert res["validados"][0]["Nro cliente"] == "9999"
    assert res["validados"][0]["Nro recibo"] == "101"
    assert res["validados"][0]["Peso"] == pytest.approx(round(_amount_difference_penalty(20.0), 2))
    assert "Divisor" not in res["validados"][0]
    assert "Aclaración recibo" not in res["validados"][0]
    assert res["validados"][0]["CUIT recibo"] == "20301020304"
    assert res["validados"][0]["CUIT ingreso"] == "20301020304"
    assert len(res["no_encontrados"]) == 1
    assert res["no_encontrados"][0]["Nro recibo"] == "100"


def test_matcher_does_not_restrict_when_no_plausible_same_cuit_exists():
    res = match_hungarian(
        [_sample_txn(cuit="20301020304")],
        [
            _sample_payment(),
            ReceiptPayment(
                empresa="SALICE",
                nro_recibo="101",
                nro_cliente="9999",
                cliente_nombre="Cliente Mismo Cuit",
                medio_pago="TRANSFERENCIA",
                fecha_pago="2026-02-10",
                importe_pago=10000.0,
                vendedor="203 - Edgardo Larrea",
            ),
        ],
        valid_max_peso=200,
        dudoso_max_peso=10000,
        cliente_to_cuit_map={"1234": "20111222333", "9999": "20301020304"},
        cliente_cuit_mismatch_penalty=50.0,
        enable_banco_sin_recibo=False,
    )
    assert len(res["validados"]) == 1
    assert res["validados"][0]["Peso"] == 0.0
    assert res["validados"][0]["Nro recibo"] == "100"
    assert res["validados"][0]["CUIT recibo"] == "20111222333"


def test_matcher_keeps_zero_penalty_when_ingreso_cuit_missing():
    res = match_hungarian(
        [_sample_txn(cuit="")],
        [_sample_payment()],
        valid_max_peso=200,
        dudoso_max_peso=500,
        cliente_to_cuit_map={"1234": "20301020304"},
        cliente_cuit_mismatch_penalty=50.0,
        enable_banco_sin_recibo=False,
    )
    assert len(res["validados"]) == 1
    assert res["validados"][0]["Peso"] == 0.0
    assert res["validados"][0]["CUIT recibo"] == "20301020304"
    assert res["validados"][0]["CUIT ingreso"] == ""


def test_stage1_validated_uses_optimal_matching_not_greedy():
    txns = [
        BankTxn(
            txn_id="T1",
            origen="BBVA",
            fecha=dt.date(2026, 2, 10),
            hora=None,
            importe=1000.0,
            texto_ref="TRANSFERENCIA 1",
            row_index=10,
            parse_ok=True,
            parse_error=None,
            was_preconciled=False,
            preconciled_recibo=None,
            cuit=None,
        ),
        BankTxn(
            txn_id="T2",
            origen="BBVA",
            fecha=dt.date(2026, 2, 10),
            hora=None,
            importe=1020.0,
            texto_ref="TRANSFERENCIA 2",
            row_index=11,
            parse_ok=True,
            parse_error=None,
            was_preconciled=False,
            preconciled_recibo=None,
            cuit=None,
        ),
    ]
    payments = [
        ReceiptPayment(
            empresa="GBA",
            nro_recibo="100",
            nro_cliente="1",
            cliente_nombre="A",
            medio_pago="Transf. Bancaria",
            fecha_pago="2026-02-10",
            importe_pago=1000.0,
        ),
        ReceiptPayment(
            empresa="GBA",
            nro_recibo="101",
            nro_cliente="2",
            cliente_nombre="B",
            medio_pago="Transf. Bancaria",
            fecha_pago="2026-02-10",
            importe_pago=980.0,
        ),
    ]

    res = match_hungarian(
        txns,
        payments,
        valid_max_peso=100,
        dudoso_max_peso=500,
        enable_banco_sin_recibo=False,
    )

    assert len(res["validados"]) == 2
    assert {row["Nro recibo"] for row in res["validados"]} == {"100", "101"}


def test_matcher_keeps_zero_penalty_when_recibo_cuit_missing():
    res = match_hungarian(
        [_sample_txn(cuit="20301020304")],
        [_sample_payment()],
        valid_max_peso=200,
        dudoso_max_peso=500,
        cliente_to_cuit_map={"9999": "20301020304"},
        cliente_cuit_mismatch_penalty=50.0,
        enable_banco_sin_recibo=False,
    )
    assert len(res["validados"]) == 1
    assert res["validados"][0]["Peso"] == 0.0
    assert res["validados"][0]["CUIT recibo"] == ""
    assert res["validados"][0]["CUIT ingreso"] == "20301020304"


def test_no_encontrados_include_cuit_fields():
    res = match_hungarian(
        [],
        [_sample_payment()],
        valid_max_peso=200,
        dudoso_max_peso=500,
        cliente_to_cuit_map={"1234": "20301020304"},
        cliente_cuit_mismatch_penalty=50.0,
        enable_banco_sin_recibo=False,
    )
    assert len(res["no_encontrados"]) == 1
    row = res["no_encontrados"][0]
    assert "Divisor" not in row
    assert "Aclaración recibo" not in row
    assert row["CUIT recibo"] == "20301020304"
    assert row["CUIT ingreso"] == ""


def test_banco_sin_recibo_grace_days_defers_recent_bank_movements():
    txn = BankTxn(
        txn_id="T1",
        origen="BBVA",
        fecha=dt.date(2026, 2, 20),
        hora=None,
        importe=1000.0,
        texto_ref="TRANSFERENCIA",
        row_index=10,
        parse_ok=True,
        parse_error=None,
        was_preconciled=False,
        preconciled_recibo=None,
        cuit=None,
    )
    payment = ReceiptPayment(
        empresa="SALICE",
        nro_recibo="100",
        nro_cliente="1234",
        cliente_nombre="Cliente Test",
        medio_pago="TRANSFERENCIA",
        fecha_pago="2026-02-20",
        importe_pago=50000.0,
        vendedor="203 - Edgardo Larrea",
    )

    res_without_grace = match_hungarian(
        [txn],
        [payment],
        valid_max_peso=170,
        dudoso_max_peso=500,
        enable_banco_sin_recibo=True,
        banco_sin_recibo_grace_days=0,
        current_date_override="2026-02-20",
    )
    res_with_grace = match_hungarian(
        [txn],
        [payment],
        valid_max_peso=170,
        dudoso_max_peso=500,
        enable_banco_sin_recibo=True,
        banco_sin_recibo_grace_days=10,
        current_date_override="2026-02-20",
    )

    assert any(r["Tipo no encontrado"] == "BANCO_SIN_RECIBO" for r in res_without_grace["no_encontrados"])
    assert not any(r["Tipo no encontrado"] == "BANCO_SIN_RECIBO" for r in res_with_grace["no_encontrados"])


def test_banco_sin_recibo_grace_does_not_hide_old_closed_ranges_forever():
    txn = BankTxn(
        txn_id="T1",
        origen="BBVA",
        fecha=dt.date(2026, 4, 1),
        hora=None,
        importe=1000.0,
        texto_ref="TRANSFERENCIA",
        row_index=10,
        parse_ok=True,
        parse_error=None,
        was_preconciled=False,
        preconciled_recibo=None,
        cuit=None,
    )
    payment = ReceiptPayment(
        empresa="GBA",
        nro_recibo="100",
        nro_cliente="1234",
        cliente_nombre="Cliente Test",
        medio_pago="Mercado Pago",
        fecha_pago="2026-04-01",
        importe_pago=50000.0,
        vendedor="203 - Edgardo Larrea",
    )

    res = match_hungarian(
        [txn],
        [payment],
        report_date_min="2026-03-30",
        report_date_max="2026-04-01",
        valid_max_peso=170,
        dudoso_max_peso=500,
        enable_banco_sin_recibo=True,
        banco_sin_recibo_grace_days=10,
        current_date_override="2026-04-15",
    )

    assert any(r["Tipo no encontrado"] == "BANCO_SIN_RECIBO" for r in res["no_encontrados"])


def test_amount_difference_penalty_grows_by_ranges():
    p05 = _amount_difference_penalty(0.5)
    p1 = _amount_difference_penalty(1.0)
    p5 = _amount_difference_penalty(5.0)
    p10 = _amount_difference_penalty(10.0)
    p30 = _amount_difference_penalty(30.0)

    assert p05 > 0.0
    assert p05 < 1.0
    assert p1 < p5 < p10 < p30
    assert p10 > 10.0
    assert p30 > 40.0


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
