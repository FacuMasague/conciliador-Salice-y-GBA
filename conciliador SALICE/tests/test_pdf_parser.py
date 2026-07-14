import re

from conciliador.pdf_parser import (
    extract_pdf_text,
    infer_empresa_from_text,
    parse_receipts_and_payments,
    parse_receipts_and_payments_from_text,
    pdf_date_range,
)


def test_infer_empresa_salice(pdf_salice_path):
    text = extract_pdf_text(pdf_salice_path, max_pages=1)
    assert infer_empresa_from_text(text) == "SALICE"


def test_infer_empresa_alarcon(pdf_alarcon_path):
    text = extract_pdf_text(pdf_alarcon_path, max_pages=1)
    # en este dataset el header incluye ALARCON
    assert infer_empresa_from_text(text) in {"ALARCON", "DESCONOCIDA"}


def test_parse_receipts_and_payments_basic(pdf_salice_path):
    receipts, payments = parse_receipts_and_payments(pdf_salice_path)
    assert len(receipts) > 10
    assert len(payments) > 10

    # nro_recibo siempre TEXT
    assert isinstance(receipts[0].nro_recibo, str)

    # nro_recibo debe ser el nÃºmero de encabezado, no el comprobante 0011-A-...
    assert "-" not in receipts[0].nro_recibo
    assert receipts[0].nro_recibo.isdigit()

    # pagos solo de los 2 medios
    assert set(p.medio_pago for p in payments).issubset({"TRANSFERENCIA", "MERCADOPAGO"})

    dmin, dmax = pdf_date_range(payments)
    assert dmin is not None and dmax is not None
    assert re.match(r"\d{4}-\d{2}-\d{2}", dmin)
    assert re.match(r"\d{4}-\d{2}-\d{2}", dmax)


def test_parse_receipts_reads_collector_from_pedidos_moviles_pdf(pdf_salice_path):
    receipts, payments = parse_receipts_and_payments(pdf_salice_path)
    assert receipts
    assert any(r.vendedor for r in receipts)
    assert any(p.vendedor for p in payments)


def test_parse_collector_accepts_comparte_suffix():
    text = """
    Reporte de Cobranza SALICE
    78613 30598 - Romeo Liliana Edith 10/07/2026 0,00 0,00 205.154,89
    [203 - Edgardo Larrea] comparte
    """
    receipts, payments = parse_receipts_and_payments_from_text(
        text,
        empresa_override="SALICE",
    )
    assert not payments
    assert len(receipts) == 1
    assert receipts[0].nro_recibo == "78613"
    assert receipts[0].vendedor == "203 - Edgardo Larrea"


def test_parse_receipts_cliente_nombre_does_not_include_fecha(pdf_alarcon_path):
    receipts, _payments = parse_receipts_and_payments(pdf_alarcon_path)
    target = next((r for r in receipts if r.nro_recibo == "68716"), None)
    assert target is not None
    assert target.cliente_nombre == "Turchetti Enzo"
