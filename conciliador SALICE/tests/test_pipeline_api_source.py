from __future__ import annotations

import datetime as dt

from src.conciliador.excel_loader import BankTxn
from src.conciliador.pdf_parser import ReceiptPayment
from src.conciliador.pipeline import compare_excel_pdfs


def test_compare_excel_pdfs_api_source_uses_external_fetch(monkeypatch):
    called = {"receipts": False, "padron": False}
    expected_end_date = dt.date(2026, 2, 18)

    def _fake_fetch_receipts(days, empresa_filter, end_date=None):
        called["receipts"] = True
        assert days == 15
        assert end_date == expected_end_date
        return (
            [
                ReceiptPayment(
                    empresa="SALICE",
                    nro_recibo="68734",
                    nro_cliente="33119",
                    cliente_nombre="X",
                    medio_pago="TRANSFERENCIA",
                    fecha_pago="2026-02-18",
                    importe_pago=100.0,
                )
            ],
            {"api_request_id": "rid-rec", "external_warnings": []},
        )

    def _fake_fetch_padron(empresa_filter):
        called["padron"] = True
        return ({"33119": "20123456789"}, {"api_request_id": "rid-pad", "external_warnings": []})

    def _fake_load_bank_txns(_excel_path):
        return [
            BankTxn(
                txn_id="t1",
                origen="BBVA",
                fecha=dt.date(2026, 2, 18),
                hora=None,
                importe=100.0,
                texto_ref="TRANSFERENCIA",
                row_index=10,
                parse_ok=True,
                parse_error=None,
            )
        ]

    def _fake_match_hungarian(txns, payments, **kwargs):
        assert len(txns) == 1
        assert len(payments) == 1
        assert kwargs.get("exclude_preconciled_txns") is True
        return {"validados": [], "dudosos": [], "no_encontrados": []}

    monkeypatch.setattr("src.conciliador.pipeline.fetch_receipts_and_payments_api", _fake_fetch_receipts)
    monkeypatch.setattr("src.conciliador.pipeline.fetch_cliente_cuit_map_api", _fake_fetch_padron)
    monkeypatch.setattr("src.conciliador.pipeline.load_bank_txns", _fake_load_bank_txns)
    monkeypatch.setattr("src.conciliador.pipeline.match_hungarian", _fake_match_hungarian)

    out = compare_excel_pdfs(
        "fake.xlsx",
        [],
        receipts_source="api",
        api_receipts_days=60,
    )

    assert called["receipts"]
    assert called["padron"]
    assert out["meta"]["receipts_source_used"] == "api"
    assert out["meta"]["api_receipts_window_days"] == 15
    assert out["meta"]["api_request_id"] == "rid-rec, rid-pad"


def test_compare_excel_pdfs_api_source_filters_old_and_preconciled(monkeypatch):
    bank_max_date = dt.date(2026, 3, 3)
    old_date = (bank_max_date - dt.timedelta(days=120)).isoformat()
    recent_date = (bank_max_date - dt.timedelta(days=5)).isoformat()
    after_bank_date = (bank_max_date + dt.timedelta(days=1)).isoformat()

    def _fake_fetch_receipts(days, empresa_filter, end_date=None):
        assert end_date == bank_max_date
        return (
            [
                ReceiptPayment(
                    empresa="SALICE",
                    nro_recibo="100",
                    nro_cliente="33119",
                    cliente_nombre="X",
                    medio_pago="TRANSFERENCIA",
                    fecha_pago=old_date,
                    importe_pago=100.0,
                ),
                ReceiptPayment(
                    empresa="SALICE",
                    nro_recibo="200",
                    nro_cliente="33119",
                    cliente_nombre="X",
                    medio_pago="TRANSFERENCIA",
                    fecha_pago=recent_date,
                    importe_pago=100.0,
                ),
                ReceiptPayment(
                    empresa="SALICE",
                    nro_recibo="300",
                    nro_cliente="33119",
                    cliente_nombre="X",
                    medio_pago="TRANSFERENCIA",
                    fecha_pago=recent_date,
                    importe_pago=100.0,
                ),
                ReceiptPayment(
                    empresa="SALICE",
                    nro_recibo="400",
                    nro_cliente="33119",
                    cliente_nombre="X",
                    medio_pago="TRANSFERENCIA",
                    fecha_pago=after_bank_date,
                    importe_pago=100.0,
                ),
            ],
            {"api_request_id": "rid-rec", "external_warnings": []},
        )

    def _fake_fetch_padron(empresa_filter):
        return ({}, {"api_request_id": "rid-pad", "external_warnings": []})

    def _fake_load_bank_txns(_excel_path):
        return [
            BankTxn(
                txn_id="t1",
                origen="BBVA",
                fecha=bank_max_date,
                hora=None,
                importe=10.0,
                texto_ref="X",
                row_index=10,
                parse_ok=True,
                parse_error=None,
                was_preconciled=True,
                preconciled_recibo="200",
            )
        ]

    observed = {"payments_len": -1}

    def _fake_match_hungarian(txns, payments, **kwargs):
        observed["payments_len"] = len(payments)
        assert kwargs.get("exclude_preconciled_txns") is True
        return {"validados": [], "dudosos": [], "no_encontrados": []}

    monkeypatch.setattr("src.conciliador.pipeline.fetch_receipts_and_payments_api", _fake_fetch_receipts)
    monkeypatch.setattr("src.conciliador.pipeline.fetch_cliente_cuit_map_api", _fake_fetch_padron)
    monkeypatch.setattr("src.conciliador.pipeline.load_bank_txns", _fake_load_bank_txns)
    monkeypatch.setattr("src.conciliador.pipeline.match_hungarian", _fake_match_hungarian)

    out = compare_excel_pdfs(
        "fake.xlsx",
        [],
        receipts_source="api",
        api_receipts_days=60,
    )

    assert observed["payments_len"] == 1
    assert out["meta"]["payments_filtered_old_window"] == 1
    assert out["meta"]["payments_filtered_current_day"] == 0
    assert out["meta"]["payments_filtered_after_bank_max_date"] == 1
    assert out["meta"]["payments_filtered_preconciled"] == 1


def test_compare_excel_pdfs_api_source_prefers_raw_extract_max_date_override(monkeypatch):
    record_max_date = dt.date(2026, 3, 4)
    raw_extract_max_date = dt.date(2026, 3, 3)
    observed = {"end_date": None}

    def _fake_fetch_receipts(days, empresa_filter, end_date=None):
        observed["end_date"] = end_date
        return ([], {"api_request_id": "rid-rec", "external_warnings": []})

    def _fake_fetch_padron(empresa_filter):
        return ({}, {"api_request_id": "rid-pad", "external_warnings": []})

    def _fake_load_bank_txns(_excel_path):
        return [
            BankTxn(
                txn_id="t1",
                origen="BBVA",
                fecha=record_max_date,
                hora=None,
                importe=100.0,
                texto_ref="TRANSFERENCIA",
                row_index=10,
                parse_ok=True,
                parse_error=None,
            )
        ]

    def _fake_match_hungarian(txns, payments, **kwargs):
        return {"validados": [], "dudosos": [], "no_encontrados": []}

    monkeypatch.setattr("src.conciliador.pipeline.fetch_receipts_and_payments_api", _fake_fetch_receipts)
    monkeypatch.setattr("src.conciliador.pipeline.fetch_cliente_cuit_map_api", _fake_fetch_padron)
    monkeypatch.setattr("src.conciliador.pipeline.load_bank_txns", _fake_load_bank_txns)
    monkeypatch.setattr("src.conciliador.pipeline.match_hungarian", _fake_match_hungarian)

    out = compare_excel_pdfs(
        "fake.xlsx",
        [],
        receipts_source="api",
        api_receipts_days=15,
        api_end_date_override=raw_extract_max_date.isoformat(),
    )

    assert observed["end_date"] == raw_extract_max_date
    assert out["meta"]["api_end_date_from_bank"] == raw_extract_max_date.isoformat()
    assert out["meta"]["api_end_date_override"] == raw_extract_max_date.isoformat()


def test_compare_excel_pdfs_api_source_uses_end_date_override_as_window_anchor(monkeypatch):
    raw_extract_max_date = dt.date(2026, 3, 6)
    observed = {"start_date": None, "end_date": None}

    def _fake_fetch_receipts(days, empresa_filter, start_date=None, end_date=None):
        observed["start_date"] = start_date
        observed["end_date"] = end_date
        return ([], {"api_request_id": "rid-rec", "external_warnings": []})

    def _fake_fetch_padron(empresa_filter):
        return ({}, {"api_request_id": "rid-pad", "external_warnings": []})

    def _fake_load_bank_txns(_excel_path):
        return [
            BankTxn(
                txn_id="t1",
                origen="BBVA",
                fecha=dt.date(2026, 3, 20),
                hora=None,
                importe=100.0,
                texto_ref="TRANSFERENCIA",
                row_index=10,
                parse_ok=True,
                parse_error=None,
            )
        ]

    def _fake_match_hungarian(txns, payments, **kwargs):
        return {"validados": [], "dudosos": [], "no_encontrados": []}

    monkeypatch.setattr("src.conciliador.pipeline.fetch_receipts_and_payments_api", _fake_fetch_receipts)
    monkeypatch.setattr("src.conciliador.pipeline.fetch_cliente_cuit_map_api", _fake_fetch_padron)
    monkeypatch.setattr("src.conciliador.pipeline.load_bank_txns", _fake_load_bank_txns)
    monkeypatch.setattr("src.conciliador.pipeline.match_hungarian", _fake_match_hungarian)

    out = compare_excel_pdfs(
        "fake.xlsx",
        [],
        receipts_source="api",
        api_receipts_days=4,
        api_end_date_override=raw_extract_max_date.isoformat(),
    )

    assert observed["start_date"] == dt.date(2026, 3, 2)
    assert observed["end_date"] == raw_extract_max_date
    assert out["meta"]["api_start_date_from_bank"] == "2026-03-02"
    assert out["meta"]["api_end_date_from_bank"] == raw_extract_max_date.isoformat()


def test_compare_excel_pdfs_api_source_filters_preconciled_accounting_receipt_numbers(monkeypatch):
    bank_max_date = dt.date(2026, 3, 3)

    def _fake_fetch_receipts(days, empresa_filter, end_date=None):
        return (
            [
                ReceiptPayment(
                    empresa="SALICE",
                    nro_recibo="68734",
                    nro_cliente="33119",
                    cliente_nombre="X",
                    medio_pago="TRANSFERENCIA",
                    fecha_pago="2026-03-02",
                    importe_pago=100.0,
                ),
                ReceiptPayment(
                    empresa="SALICE",
                    nro_recibo="200",
                    nro_cliente="33119",
                    cliente_nombre="X",
                    medio_pago="TRANSFERENCIA",
                    fecha_pago="2026-03-02",
                    importe_pago=120.0,
                ),
            ],
            {"api_request_id": "rid-rec", "external_warnings": []},
        )

    def _fake_fetch_padron(empresa_filter):
        return ({}, {"api_request_id": "rid-pad", "external_warnings": []})

    def _fake_load_bank_txns(_excel_path):
        return [
            BankTxn(
                txn_id="t1",
                origen="BBVA",
                fecha=bank_max_date,
                hora=None,
                importe=10.0,
                texto_ref="X",
                row_index=10,
                parse_ok=True,
                parse_error=None,
                was_preconciled=True,
                preconciled_recibo="68.734,00",
            ),
            BankTxn(
                txn_id="t2",
                origen="BBVA",
                fecha=bank_max_date,
                hora=None,
                importe=11.0,
                texto_ref="Y",
                row_index=11,
                parse_ok=True,
                parse_error=None,
                was_preconciled=True,
                preconciled_recibo="200.0",
            ),
        ]

    observed = {"payments": None}

    def _fake_match_hungarian(txns, payments, **kwargs):
        observed["payments"] = [p.nro_recibo for p in payments]
        return {"validados": [], "dudosos": [], "no_encontrados": []}

    monkeypatch.setattr("src.conciliador.pipeline.fetch_receipts_and_payments_api", _fake_fetch_receipts)
    monkeypatch.setattr("src.conciliador.pipeline.fetch_cliente_cuit_map_api", _fake_fetch_padron)
    monkeypatch.setattr("src.conciliador.pipeline.load_bank_txns", _fake_load_bank_txns)
    monkeypatch.setattr("src.conciliador.pipeline.match_hungarian", _fake_match_hungarian)

    out = compare_excel_pdfs(
        "fake.xlsx",
        [],
        receipts_source="api",
        api_receipts_days=15,
    )

    assert observed["payments"] == []
    assert out["meta"]["payments_filtered_preconciled"] == 2
