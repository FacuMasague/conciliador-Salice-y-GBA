import datetime as dt

import openpyxl

from conciliador.excel_loader import load_bank_txns


def test_load_bank_txns_has_multiple_sources(excel_path):
    txns = load_bank_txns(excel_path)
    assert len(txns) > 100
    origins = {t.origen for t in txns}
    # deberÃÂ­amos ver al menos BBVA/GALICIA/MP en este archivo
    assert "BBVA" in origins
    assert "GALICIA" in origins
    assert "MERCADOPAGO" in origins


def test_load_bank_txns_parsing_flags(excel_path):
    txns = load_bank_txns(excel_path)
    assert any(t.parse_ok for t in txns)
    # No queremos que falle: los errores deben quedar marcados, no crashear
    if any((not t.parse_ok) for t in txns):
        assert all(t.parse_error for t in txns if not t.parse_ok)
    else:
        # dataset puede venir limpio; igual es válido
        assert True


def test_load_bank_txns_date_types(excel_path):
    txns = load_bank_txns(excel_path)
    ok = [t for t in txns if t.parse_ok]
    assert isinstance(ok[0].fecha, dt.date)
    assert isinstance(ok[0].importe, float)


def test_load_bank_txns_bbva_excludes_impuesto_ley(tmp_path):
    p = tmp_path / "bbva_impuesto.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "SALICE BBVA"
    ws.append(["BBVA Frances"])
    ws.append(["Fecha", "Número Documento", "Oficina", "Débitos", "Importe", "ok", "cliente", "recibo"])
    ws.append(["23-02-2026", "IMPUESTO LEY 20/02/26 00002", "", "", "1,40", "", "", ""])
    ws.append(["23-02-2026", "TRANSF.BANEL 20302962112", "", "", "88200,00", "", "", ""])
    wb.save(str(p))

    txns = load_bank_txns(str(p))
    bbva = [t for t in txns if t.origen == "BBVA" and t.parse_ok]
    textos = [t.texto_ref for t in bbva]

    assert any("TRANSF.BANEL" in t for t in textos)
    assert not any("IMPUESTO LEY" in t.upper() for t in textos)


def test_load_bank_txns_bbva_keeps_transf_credito_banelco(tmp_path):
    p = tmp_path / "bbva_banelco.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "SALICE BBVA"
    ws.append(["BBVA Frances"])
    ws.append(["Fecha", "Número Documento", "Oficina", "Débitos", "Importe", "ok", "cliente", "recibo"])
    ws.append(["23-02-2026", "TRANSF CREDITO BANELCO", "", "", "1.000,00", "", "", ""])
    ws.append(["23-02-2026", "TRANSF.BANEL 20302962112", "", "", "88200,00", "", "", ""])
    wb.save(str(p))

    txns = load_bank_txns(str(p))
    bbva = [t for t in txns if t.origen == "BBVA" and t.parse_ok]
    textos = [t.texto_ref for t in bbva]

    assert any("TRANSF.BANEL" in t for t in textos)
    assert any("TRANSF CREDITO BANELCO" in t.upper() for t in textos)
