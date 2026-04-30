from conciliador.pipeline import compare_excel_pdf


def _assert_has_keys(row, keys):
    missing = [k for k in keys if k not in row]
    assert not missing, f"Faltan keys: {missing} en row={row}"


def test_contract_returns_three_lists(excel_path, pdf_salice_path):
    res = compare_excel_pdf(excel_path, pdf_salice_path, margin_days=5)
    assert "validados" in res
    assert "dudosos" in res
    assert "no_encontrados" in res


def test_contract_validados_and_dudosos_min_keys(excel_path, pdf_salice_path):
    res = compare_excel_pdf(excel_path, pdf_salice_path, margin_days=5)
    base_keys = {
        "Nro recibo",
        "Nro cliente",
        "Medio de pago",
        "Fecha recibo",
        "Importe recibo",
        "Origen",
        "Fecha movimiento",
        "Importe movimiento",
        "Detalle movimiento",
        "Fila Excel",
    }
    if res["validados"]:
        _assert_has_keys(res["validados"][0], base_keys)
        assert isinstance(res["validados"][0]["Nro recibo"], str)
    if res["dudosos"]:
        _assert_has_keys(res["dudosos"][0], base_keys | {"Motivo", "Dif días", "Dif importe"})


def test_contract_no_encontrados_types(excel_path, pdf_salice_path):
    res = compare_excel_pdf(excel_path, pdf_salice_path, margin_days=5)
    if not res["no_encontrados"]:
        # permitido: algunos meses pueden conciliar al 100%
        return
    row = res["no_encontrados"][0]
    _assert_has_keys(row, {"Tipo no encontrado"})
    assert row["Tipo no encontrado"] in {"BANCO_SIN_RECIBO", "RECIBO_SIN_BANCO"}


def test_must_have_some_validated_or_suspect(excel_path, pdf_salice_path):
    res = compare_excel_pdf(excel_path, pdf_salice_path, margin_days=5)
    assert len(res["validados"]) + len(res["dudosos"]) > 0
