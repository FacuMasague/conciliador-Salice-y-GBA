from __future__ import annotations

import datetime as dt

from src.conciliador.excel_loader import BankTxn
from src.conciliador.pdf_parser import ReceiptPayment
from src.conciliador.pipeline import compare_excel_pdfs, _medio_applies_to_program


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

    def _fake_fetch_padron(empresa_filter, cliente_ids=None):
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
        assert kwargs.get("exclude_preconciled_txns") is False
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

    def _fake_fetch_padron(empresa_filter, cliente_ids=None):
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
        assert kwargs.get("exclude_preconciled_txns") is False
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


def test_compare_excel_pdfs_api_source_reopens_displaced_preconciled_receipt(monkeypatch):
    observed_calls = []

    def _fake_fetch_receipts(days, empresa_filter, start_date=None, end_date=None):
        if start_date == dt.date(2026, 3, 16) and end_date == dt.date(2026, 3, 16):
            return (
                [
                    ReceiptPayment(
                        empresa="GBA",
                        nro_recibo="300",
                        nro_cliente="30",
                        cliente_nombre="Nuevo",
                        medio_pago="Mercado Pago",
                        fecha_pago="2026-03-16",
                        importe_pago=100.0,
                    )
                ],
                {"api_request_id": "rid-rec", "external_warnings": []},
            )
        return (
            [
                ReceiptPayment(
                    empresa="GBA",
                    nro_recibo="200",
                    nro_cliente="20",
                    cliente_nombre="Viejo",
                    medio_pago="Mercado Pago",
                    fecha_pago="2026-03-15",
                    importe_pago=100.0,
                ),
                ReceiptPayment(
                    empresa="GBA",
                    nro_recibo="999",
                    nro_cliente="99",
                    cliente_nombre="No deberia entrar",
                    medio_pago="Mercado Pago",
                    fecha_pago="2026-03-15",
                    importe_pago=999.0,
                ),
            ],
            {"api_request_id": "rid-rec-extra", "external_warnings": []},
        )

    def _fake_fetch_padron(empresa_filter, cliente_ids=None):
        return ({}, {"api_request_id": "rid-pad", "external_warnings": []})

    def _fake_load_bank_txns(_excel_path):
        return [
            BankTxn(
                txn_id="t1",
                origen="BBVA",
                fecha=dt.date(2026, 3, 16),
                hora=None,
                importe=100.0,
                texto_ref="X",
                row_index=10,
                parse_ok=True,
                parse_error=None,
                was_preconciled=True,
                preconciled_recibo="200",
                preconciled_nro_cliente="20",
                preconciled_cliente_nombre="Viejo",
                preconciled_fecha_recibo="2026-03-15",
                preconciled_medio_pago="Mercado Pago",
                preconciled_importe_recibo=100.0,
            )
        ]

    def _fake_match_hungarian(txns, payments, **kwargs):
        observed_calls.append([str(p.nro_recibo) for p in payments])
        assert kwargs.get("exclude_preconciled_txns") is False
        if len(observed_calls) == 1:
            return {
                "validados": [
                    {
                        "Tipo fila": "PRINCIPAL",
                        "Nro recibo": "300",
                        "__txn_id": "t1",
                    }
                ],
                "dudosos": [],
                "no_encontrados": [],
            }
        return {"validados": [], "dudosos": [], "no_encontrados": []}

    monkeypatch.setattr("src.conciliador.pipeline.fetch_receipts_and_payments_api", _fake_fetch_receipts)
    monkeypatch.setattr("src.conciliador.pipeline.fetch_cliente_cuit_map_api", _fake_fetch_padron)
    monkeypatch.setattr("src.conciliador.pipeline.load_bank_txns", _fake_load_bank_txns)
    monkeypatch.setattr("src.conciliador.pipeline.match_hungarian", _fake_match_hungarian)

    out = compare_excel_pdfs(
        "fake.xlsx",
        [],
        receipts_source="api",
        api_start_date_override="2026-03-16",
        api_end_date_override="2026-03-16",
    )

    assert observed_calls == [["300"], ["300", "200"]]
    assert any("Se reabrieron 1 recibos ya conciliados del record" in w for w in out["meta"]["external_warnings"])


def test_medio_applies_to_program_gba_allowlist():
    assert _medio_applies_to_program("Transf. Bancaria") is True
    assert _medio_applies_to_program("Mercado Pago") is True
    assert _medio_applies_to_program("Boleta de Deposito") is True
    assert _medio_applies_to_program("Cheques de 3ros") is True
    assert _medio_applies_to_program("Cheque Electronico") is True
    assert _medio_applies_to_program("Efectivo") is False
    assert _medio_applies_to_program("Bonos") is False
    assert _medio_applies_to_program("Tarjeta Credito") is False
    assert _medio_applies_to_program("Cheque propio") is False
    assert _medio_applies_to_program("Redondeo") is False


def test_compare_excel_pdfs_api_source_filters_non_program_media_with_detail(monkeypatch):
    observed = {"payments_len": -1, "medios": []}

    def _fake_fetch_receipts(days, empresa_filter, start_date=None, end_date=None):
        return (
            [
                ReceiptPayment(
                    empresa="GBA",
                    nro_recibo="100",
                    nro_cliente="1",
                    cliente_nombre="A",
                    medio_pago="SIN_MEDIO_API",
                    fecha_pago="2026-03-16",
                    importe_pago=100.0,
                    api_key={"ComprobanteID": 1, "EmpresaID": 2, "Serie": "", "PuntoDeVentaID": 1, "Numero": 100},
                ),
                ReceiptPayment(
                    empresa="GBA",
                    nro_recibo="200",
                    nro_cliente="2",
                    cliente_nombre="B",
                    medio_pago="SIN_MEDIO_API",
                    fecha_pago="2026-03-16",
                    importe_pago=200.0,
                    api_key={"ComprobanteID": 2, "EmpresaID": 2, "Serie": "", "PuntoDeVentaID": 1, "Numero": 200},
                ),
                ReceiptPayment(
                    empresa="GBA",
                    nro_recibo="300",
                    nro_cliente="3",
                    cliente_nombre="C",
                    medio_pago="SIN_MEDIO_API",
                    fecha_pago="2026-03-16",
                    importe_pago=300.0,
                    api_key={"ComprobanteID": 3, "EmpresaID": 2, "Serie": "", "PuntoDeVentaID": 1, "Numero": 300},
                ),
            ],
            {"api_request_id": "rid-rec", "external_warnings": []},
        )

    def _fake_fetch_padron(empresa_filter, cliente_ids=None):
        return ({}, {"api_request_id": "rid-pad", "external_warnings": []})

    def _fake_fetch_payment_detail_map(api_keys, empresa_filter):
        assert len(api_keys) == 3
        return (
            {
                ("1", "2", "", "1", "100"): {"medio_pago": "Efectivo", "importe_bankable": 100.0},
                ("2", "2", "", "1", "200"): {"medio_pago": "Mercado Pago", "importe_bankable": 200.0},
                ("3", "2", "", "1", "300"): {"medio_pago": "Transf. Bancaria", "importe_bankable": 300.0},
            },
            [],
        )

    def _fake_load_bank_txns(_excel_path):
        return [
            BankTxn(
                txn_id="t1",
                origen="BBVA",
                fecha=dt.date(2026, 3, 16),
                hora=None,
                importe=200.0,
                texto_ref="X",
                row_index=10,
                parse_ok=True,
                parse_error=None,
            )
        ]

    def _fake_match_hungarian(txns, payments, **kwargs):
        observed["payments_len"] = len(payments)
        observed["medios"] = [str(p.medio_pago) for p in payments]
        return {"validados": [], "dudosos": [], "no_encontrados": []}

    monkeypatch.setattr("src.conciliador.pipeline.fetch_receipts_and_payments_api", _fake_fetch_receipts)
    monkeypatch.setattr("src.conciliador.pipeline.fetch_cliente_cuit_map_api", _fake_fetch_padron)
    monkeypatch.setattr("src.conciliador.pipeline.fetch_payment_detail_map_for_api_keys_api", _fake_fetch_payment_detail_map)
    monkeypatch.setattr("src.conciliador.pipeline.load_bank_txns", _fake_load_bank_txns)
    monkeypatch.setattr("src.conciliador.pipeline.match_hungarian", _fake_match_hungarian)

    out = compare_excel_pdfs(
        "fake.xlsx",
        [],
        receipts_source="api",
        api_start_date_override="2026-03-16",
        api_end_date_override="2026-03-16",
    )

    assert observed["payments_len"] == 2
    assert observed["medios"] == ["Mercado Pago", "Transf. Bancaria"]
    assert any("Se filtraron 1 recibos con medios de pago" in w for w in out["meta"]["external_warnings"])


def test_compare_excel_pdfs_api_source_filters_unknown_media_after_detail(monkeypatch):
    observed = {"payments_len": -1, "medios": []}

    def _fake_fetch_receipts(days, empresa_filter, start_date=None, end_date=None):
        return (
            [
                ReceiptPayment(
                    empresa="GBA",
                    nro_recibo="100",
                    nro_cliente="1",
                    cliente_nombre="A",
                    medio_pago="SIN_MEDIO_API",
                    fecha_pago="2026-03-16",
                    importe_pago=100.0,
                    api_key={"ComprobanteID": 1, "EmpresaID": 2, "Serie": "", "PuntoDeVentaID": 1, "Numero": 100},
                ),
                ReceiptPayment(
                    empresa="GBA",
                    nro_recibo="200",
                    nro_cliente="2",
                    cliente_nombre="B",
                    medio_pago="SIN_MEDIO_API",
                    fecha_pago="2026-03-16",
                    importe_pago=200.0,
                    api_key={"ComprobanteID": 2, "EmpresaID": 2, "Serie": "", "PuntoDeVentaID": 1, "Numero": 200},
                ),
            ],
            {"api_request_id": "rid-rec", "external_warnings": []},
        )

    def _fake_fetch_padron(empresa_filter, cliente_ids=None):
        return ({}, {"api_request_id": "rid-pad", "external_warnings": []})

    def _fake_fetch_payment_detail_map(api_keys, empresa_filter):
        return (
            {
                ("1", "2", "", "1", "100"): {"medio_pago": "SIN_MEDIO_API", "importe_bankable": 100.0},
                ("2", "2", "", "1", "200"): {"medio_pago": "Mercado Pago", "importe_bankable": 200.0},
            },
            [],
        )

    def _fake_load_bank_txns(_excel_path):
        return [
            BankTxn(
                txn_id="t1",
                origen="MERCADOPAGO",
                fecha=dt.date(2026, 3, 16),
                hora=None,
                importe=200.0,
                texto_ref="X",
                row_index=10,
                parse_ok=True,
                parse_error=None,
            )
        ]

    def _fake_match_hungarian(txns, payments, **kwargs):
        observed["payments_len"] = len(payments)
        observed["medios"] = [str(p.medio_pago) for p in payments]
        return {"validados": [], "dudosos": [], "no_encontrados": []}

    monkeypatch.setattr("src.conciliador.pipeline.fetch_receipts_and_payments_api", _fake_fetch_receipts)
    monkeypatch.setattr("src.conciliador.pipeline.fetch_cliente_cuit_map_api", _fake_fetch_padron)
    monkeypatch.setattr("src.conciliador.pipeline.fetch_payment_detail_map_for_api_keys_api", _fake_fetch_payment_detail_map)
    monkeypatch.setattr("src.conciliador.pipeline.load_bank_txns", _fake_load_bank_txns)
    monkeypatch.setattr("src.conciliador.pipeline.match_hungarian", _fake_match_hungarian)

    out = compare_excel_pdfs(
        "fake.xlsx",
        [],
        receipts_source="api",
        api_start_date_override="2026-03-16",
        api_end_date_override="2026-03-16",
    )

    assert observed["payments_len"] == 1
    assert observed["medios"] == ["Mercado Pago"]
    assert any("sin medio de pago API verificable" in w for w in out["meta"]["external_warnings"])


def test_compare_excel_pdfs_api_source_filters_unknown_media_when_detail_missing(monkeypatch):
    observed = {"payments_len": -1}

    def _fake_fetch_receipts(days, empresa_filter, start_date=None, end_date=None):
        return (
            [
                ReceiptPayment(
                    empresa="GBA",
                    nro_recibo="100",
                    nro_cliente="1",
                    cliente_nombre="A",
                    medio_pago="SIN_MEDIO_API",
                    fecha_pago="2026-03-16",
                    importe_pago=100.0,
                    api_key={"ComprobanteID": 1, "EmpresaID": 2, "Serie": "", "PuntoDeVentaID": 1, "Numero": 100},
                ),
                ReceiptPayment(
                    empresa="GBA",
                    nro_recibo="200",
                    nro_cliente="2",
                    cliente_nombre="B",
                    medio_pago="Mercado Pago",
                    fecha_pago="2026-03-16",
                    importe_pago=200.0,
                    api_key={"ComprobanteID": 2, "EmpresaID": 2, "Serie": "", "PuntoDeVentaID": 1, "Numero": 200},
                ),
            ],
            {"api_request_id": "rid-rec", "external_warnings": []},
        )

    def _fake_fetch_padron(empresa_filter, cliente_ids=None):
        return ({}, {"api_request_id": "rid-pad", "external_warnings": []})

    def _fake_fetch_payment_detail_map(api_keys, empresa_filter):
        return ({}, [])

    def _fake_load_bank_txns(_excel_path):
        return [
            BankTxn(
                txn_id="t1",
                origen="MERCADOPAGO",
                fecha=dt.date(2026, 3, 16),
                hora=None,
                importe=200.0,
                texto_ref="X",
                row_index=10,
                parse_ok=True,
                parse_error=None,
            )
        ]

    def _fake_match_hungarian(txns, payments, **kwargs):
        observed["payments_len"] = len(payments)
        return {"validados": [], "dudosos": [], "no_encontrados": []}

    monkeypatch.setattr("src.conciliador.pipeline.fetch_receipts_and_payments_api", _fake_fetch_receipts)
    monkeypatch.setattr("src.conciliador.pipeline.fetch_cliente_cuit_map_api", _fake_fetch_padron)
    monkeypatch.setattr("src.conciliador.pipeline.fetch_payment_detail_map_for_api_keys_api", _fake_fetch_payment_detail_map)
    monkeypatch.setattr("src.conciliador.pipeline.load_bank_txns", _fake_load_bank_txns)
    monkeypatch.setattr("src.conciliador.pipeline.match_hungarian", _fake_match_hungarian)

    out = compare_excel_pdfs(
        "fake.xlsx",
        [],
        receipts_source="api",
        api_start_date_override="2026-03-16",
        api_end_date_override="2026-03-16",
    )

    assert observed["payments_len"] == 1
    assert any("sin medio de pago API verificable" in w for w in out["meta"]["external_warnings"])

def test_compare_excel_pdfs_api_source_requests_detail_for_all_amounts(monkeypatch):
    observed = {"payments_len": -1, "detail_keys": None}

    def _fake_fetch_receipts(days, empresa_filter, start_date=None, end_date=None):
        return (
            [
                ReceiptPayment(
                    empresa="GBA",
                    nro_recibo="100",
                    nro_cliente="1",
                    cliente_nombre="A",
                    medio_pago="SIN_MEDIO_API",
                    fecha_pago="2026-03-16",
                    importe_pago=100.0,
                    api_key={"ComprobanteID": 1, "EmpresaID": 2, "Serie": "", "PuntoDeVentaID": 1, "Numero": 100},
                ),
                ReceiptPayment(
                    empresa="GBA",
                    nro_recibo="101",
                    nro_cliente="2",
                    cliente_nombre="B",
                    medio_pago="SIN_MEDIO_API",
                    fecha_pago="2026-03-16",
                    importe_pago=101.5,
                    api_key={"ComprobanteID": 2, "EmpresaID": 2, "Serie": "", "PuntoDeVentaID": 1, "Numero": 101},
                ),
            ],
            {"api_request_id": "rid-rec", "external_warnings": []},
        )

    def _fake_fetch_padron(empresa_filter, cliente_ids=None):
        return ({}, {"api_request_id": "rid-pad", "external_warnings": []})

    def _fake_fetch_payment_detail_map(api_keys, empresa_filter):
        observed["detail_keys"] = list(api_keys)
        return ({("1", "2", "", "1", "100"): {"medio_pago": "Mercado Pago", "importe_bankable": 100.0}}, [])

    def _fake_load_bank_txns(_excel_path):
        return [BankTxn(txn_id="t1", origen="MERCADOPAGO", fecha=dt.date(2026, 3, 16), hora=None, importe=101.5, texto_ref="X", row_index=10, parse_ok=True, parse_error=None)]

    def _fake_match_hungarian(txns, payments, **kwargs):
        observed["payments_len"] = len(payments)
        return {"validados": [], "dudosos": [], "no_encontrados": []}

    monkeypatch.setattr("src.conciliador.pipeline.fetch_receipts_and_payments_api", _fake_fetch_receipts)
    monkeypatch.setattr("src.conciliador.pipeline.fetch_cliente_cuit_map_api", _fake_fetch_padron)
    monkeypatch.setattr("src.conciliador.pipeline.fetch_payment_detail_map_for_api_keys_api", _fake_fetch_payment_detail_map)
    monkeypatch.setattr("src.conciliador.pipeline.load_bank_txns", _fake_load_bank_txns)
    monkeypatch.setattr("src.conciliador.pipeline.match_hungarian", _fake_match_hungarian)

    out = compare_excel_pdfs(
        "fake.xlsx",
        [],
        receipts_source="api",
        api_start_date_override="2026-03-16",
        api_end_date_override="2026-03-16",
    )

    assert observed["payments_len"] == 1
    assert observed["detail_keys"] == [
        {"ComprobanteID": 1, "EmpresaID": 2, "Serie": "", "PuntoDeVentaID": 1, "Numero": 100},
        {"ComprobanteID": 2, "EmpresaID": 2, "Serie": "", "PuntoDeVentaID": 1, "Numero": 101},
    ]
    assert out["meta"]["api_post_detail_stats"]["detail_keys_total"] == 2


def test_compare_excel_pdfs_api_source_filters_unknown_without_detail_for_all_amounts(monkeypatch):
    observed = {"payments_len": -1, "recibos": []}

    def _fake_fetch_receipts(days, empresa_filter, start_date=None, end_date=None):
        return (
            [
                ReceiptPayment(
                    empresa="GBA",
                    nro_recibo="100",
                    nro_cliente="1",
                    cliente_nombre="A",
                    medio_pago="SIN_MEDIO_API",
                    fecha_pago="2026-03-16",
                    importe_pago=100.0,
                    api_key={"ComprobanteID": 1, "EmpresaID": 2, "Serie": "", "PuntoDeVentaID": 1, "Numero": 100},
                ),
                ReceiptPayment(
                    empresa="GBA",
                    nro_recibo="101",
                    nro_cliente="2",
                    cliente_nombre="B",
                    medio_pago="SIN_MEDIO_API",
                    fecha_pago="2026-03-16",
                    importe_pago=101.5,
                    api_key={"ComprobanteID": 2, "EmpresaID": 2, "Serie": "", "PuntoDeVentaID": 1, "Numero": 101},
                ),
            ],
            {"api_request_id": "rid-rec", "external_warnings": []},
        )

    def _fake_fetch_padron(empresa_filter, cliente_ids=None):
        return ({}, {"api_request_id": "rid-pad", "external_warnings": []})

    def _fake_fetch_payment_detail_map(api_keys, empresa_filter):
        return ({}, [])

    def _fake_load_bank_txns(_excel_path):
        return [BankTxn(txn_id="t1", origen="BBVA", fecha=dt.date(2026, 3, 16), hora=None, importe=101.5, texto_ref="X", row_index=10, parse_ok=True, parse_error=None)]

    def _fake_match_hungarian(txns, payments, **kwargs):
        observed["payments_len"] = len(payments)
        observed["recibos"] = [p.nro_recibo for p in payments]
        return {"validados": [], "dudosos": [], "no_encontrados": []}

    monkeypatch.setattr("src.conciliador.pipeline.fetch_receipts_and_payments_api", _fake_fetch_receipts)
    monkeypatch.setattr("src.conciliador.pipeline.fetch_cliente_cuit_map_api", _fake_fetch_padron)
    monkeypatch.setattr("src.conciliador.pipeline.fetch_payment_detail_map_for_api_keys_api", _fake_fetch_payment_detail_map)
    monkeypatch.setattr("src.conciliador.pipeline.load_bank_txns", _fake_load_bank_txns)
    monkeypatch.setattr("src.conciliador.pipeline.match_hungarian", _fake_match_hungarian)

    out = compare_excel_pdfs(
        "fake.xlsx",
        [],
        receipts_source="api",
        api_start_date_override="2026-03-16",
        api_end_date_override="2026-03-16",
    )

    assert observed["payments_len"] == 0
    assert observed["recibos"] == []
    assert any("sin medio de pago API verificable" in w for w in out["meta"]["external_warnings"])


def test_compare_excel_pdfs_api_source_result_uses_api_medium_not_origen(monkeypatch):
    def _fake_fetch_receipts(days, empresa_filter, start_date=None, end_date=None):
        return (
            [
                ReceiptPayment(
                    empresa="GBA",
                    nro_recibo="200",
                    nro_cliente="2",
                    cliente_nombre="B",
                    medio_pago="SIN_MEDIO_API",
                    fecha_pago="2026-03-16",
                    importe_pago=200.0,
                    api_key={"ComprobanteID": 2, "EmpresaID": 2, "Serie": "", "PuntoDeVentaID": 1, "Numero": 200},
                ),
            ],
            {"api_request_id": "rid-rec", "external_warnings": []},
        )

    def _fake_fetch_padron(empresa_filter, cliente_ids=None):
        return ({}, {"api_request_id": "rid-pad", "external_warnings": []})

    def _fake_fetch_payment_detail_map(api_keys, empresa_filter):
        return (
            {
                ("2", "2", "", "1", "200"): {"medio_pago": "Mercado Pago", "importe_bankable": 200.0},
            },
            [],
        )

    def _fake_load_bank_txns(_excel_path):
        return [BankTxn(txn_id="t1", origen="BBVA", fecha=dt.date(2026, 3, 16), hora=None, importe=200.0, texto_ref="X", row_index=10, parse_ok=True, parse_error=None)]

    def _fake_match_hungarian(txns, payments, **kwargs):
        return {
            "validados": [
                {
                    "Nro recibo": "200",
                    "Medio de pago": "SIN_MEDIO_API",
                    "Origen": "BBVA",
                }
            ],
            "dudosos": [],
            "no_encontrados": [],
        }

    monkeypatch.setattr("src.conciliador.pipeline.fetch_receipts_and_payments_api", _fake_fetch_receipts)
    monkeypatch.setattr("src.conciliador.pipeline.fetch_cliente_cuit_map_api", _fake_fetch_padron)
    monkeypatch.setattr("src.conciliador.pipeline.fetch_payment_detail_map_for_api_keys_api", _fake_fetch_payment_detail_map)
    monkeypatch.setattr("src.conciliador.pipeline.load_bank_txns", _fake_load_bank_txns)
    monkeypatch.setattr("src.conciliador.pipeline.match_hungarian", _fake_match_hungarian)

    out = compare_excel_pdfs(
        "fake.xlsx",
        [],
        receipts_source="api",
        api_start_date_override="2026-03-16",
        api_end_date_override="2026-03-16",
    )

    assert out["validados"][0]["Medio de pago"] == "Mercado Pago"
    assert any("Medio de pago API completado" in w for w in out["meta"]["external_warnings"])


def test_matcher_hides_recent_recibo_sin_banco_with_grace():
    txns = [
        BankTxn(
            txn_id="t1",
            origen="BBVA",
            fecha=dt.date(2026, 3, 20),
            hora=None,
            importe=100.0,
            texto_ref="X",
            row_index=10,
            parse_ok=True,
            parse_error=None,
        )
    ]
    payments = [
        ReceiptPayment(
            empresa="GBA",
            nro_recibo="1",
            nro_cliente="1",
            cliente_nombre="A",
            medio_pago="Mercado Pago",
            fecha_pago="2026-03-21",
            importe_pago=999.0,
        )
    ]

    from src.conciliador.matcher_hungarian import match_hungarian

    out = match_hungarian(
        txns,
        payments,
        report_date_min="2026-03-16",
        report_date_max="2026-03-22",
        recibo_sin_banco_grace_days=2,
        banco_sin_recibo_grace_days=10,
        current_date_override="2026-03-22",
    )

    assert out["no_encontrados"] == []


def test_compare_excel_pdfs_drop_dudoso_moves_receipt_and_bank_to_no_encontrados(monkeypatch):
    def _fake_fetch_receipts(days, empresa_filter, start_date=None, end_date=None):
        return ([], {"api_request_id": "rid-rec", "external_warnings": []})

    def _fake_fetch_padron(empresa_filter, cliente_ids=None):
        return ({}, {"api_request_id": "rid-pad", "external_warnings": []})

    def _fake_load_bank_txns(_excel_path, **kwargs):
        return [
            BankTxn(
                txn_id="t1",
                origen="BBVA",
                fecha=dt.date(2026, 4, 6),
                hora=None,
                importe=100.0,
                texto_ref="TRANSFERENCIA",
                row_index=10,
                parse_ok=True,
                parse_error=None,
            )
        ]

    def _fake_match_hungarian(txns, payments, **kwargs):
        return {
            "validados": [],
            "dudosos": [
                {
                    "Tipo fila": "PRINCIPAL",
                    "Ranking": 1,
                    "__case_id": "case-1",
                    "Nro recibo": "100",
                    "Nro cliente": "10",
                    "Cliente": "Cliente",
                    "Medio de pago": "Transferencia",
                    "Fecha recibo": "2026-04-06",
                    "Importe recibo": 100.0,
                    "Origen": "BBVA",
                    "Fecha movimiento": "2026-04-06",
                    "Importe movimiento": 100.0,
                    "Detalle movimiento": "TRANSFERENCIA",
                    "Fila Excel": 10,
                    "Peso": 25,
                }
            ],
            "no_encontrados": [],
        }

    monkeypatch.setattr("src.conciliador.pipeline.fetch_receipts_and_payments_api", _fake_fetch_receipts)
    monkeypatch.setattr("src.conciliador.pipeline.fetch_cliente_cuit_map_api", _fake_fetch_padron)
    monkeypatch.setattr("src.conciliador.pipeline.load_bank_txns", _fake_load_bank_txns)
    monkeypatch.setattr("src.conciliador.pipeline.match_hungarian", _fake_match_hungarian)

    out = compare_excel_pdfs(
        "fake.xlsx",
        [],
        receipts_source="api",
        api_start_date_override="2026-04-06",
        api_end_date_override="2026-04-06",
        drop_dudosos=[{"case_id": "case-1", "fila_excel": 10, "ranking": 1}],
    )

    assert out["dudosos"] == []
    assert len(out["dudosos_borrados"]) == 1
    tipos = [r["Tipo no encontrado"] for r in out["no_encontrados"]]
    assert tipos == ["RECIBO_SIN_BANCO", "BANCO_SIN_RECIBO"]


def test_compare_excel_pdfs_api_source_prefers_raw_extract_max_date_override(monkeypatch):
    record_max_date = dt.date(2026, 3, 4)
    raw_extract_max_date = dt.date(2026, 3, 3)
    observed = {"end_date": None}

    def _fake_fetch_receipts(days, empresa_filter, end_date=None):
        observed["end_date"] = end_date
        return ([], {"api_request_id": "rid-rec", "external_warnings": []})

    def _fake_fetch_padron(empresa_filter, cliente_ids=None):
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

    def _fake_fetch_padron(empresa_filter, cliente_ids=None):
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

    def _fake_fetch_padron(empresa_filter, cliente_ids=None):
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
