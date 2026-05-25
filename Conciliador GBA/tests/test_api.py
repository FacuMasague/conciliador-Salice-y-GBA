from fastapi.testclient import TestClient

from app import app


def test_compare_endpoint_returns_contract(excel_path, pdf_salice_path):
    client = TestClient(app)

    with open(excel_path, "rb") as f_excel, open(pdf_salice_path, "rb") as f_pdf:
        files = {
            "excel": ("Movimientos bancarios 2026.xlsx", f_excel, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"),
            "pdf_salice": ("reporte salice.pdf", f_pdf, "application/pdf"),
        }
        r = client.post("/compare?margin_days=5", files=files)

    assert r.status_code == 200, r.text
    data = r.json()

    # Listas base
    for k in ("validados", "dudosos", "no_encontrados", "meta"):
        assert k in data

    # Chequeo minimo de filas (si hay)
    # La salida usa nombres con espacios (pensado para empleados).
    if data["validados"]:
        row = data["validados"][0]
        for key in (
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
        ):
            assert key in row

    # No encontrados deben tener tipo
    if data["no_encontrados"]:
        row = data["no_encontrados"][0]
        assert row["Tipo no encontrado"] in ("BANCO_SIN_RECIBO", "RECIBO_SIN_BANCO")
