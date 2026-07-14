from __future__ import annotations

from io import BytesIO

import openpyxl
from fastapi.testclient import TestClient

import app as app_module


def _minimal_xlsx_bytes() -> bytes:
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Sheet1"
    ws.append(["a"])
    bio = BytesIO()
    wb.save(bio)
    return bio.getvalue()


def test_compare_api_mode_accepts_no_pdfs(monkeypatch):
    client = TestClient(app_module.app)

    def _fake_compare(*args, **kwargs):
        assert kwargs.get("receipts_source") == "api"
        return {"validados": [], "dudosos": [], "no_encontrados": [], "meta": {"ok": True}}

    monkeypatch.setattr(app_module, "compare_excel_pdfs", _fake_compare)

    xbytes = _minimal_xlsx_bytes()
    files = {
        "excel": ("legacy.xlsx", xbytes, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"),
    }

    r = client.post("/compare?receipts_source=api&api_receipts_days=15", files=files)
    assert r.status_code == 200, r.text
    body = r.json()
    assert body["meta"]["app_version"] == app_module.APP_VERSION


def test_web_does_not_ask_user_for_collector_control():
    client = TestClient(app_module.app)

    response = client.get("/")

    assert response.status_code == 200
    assert 'id="pdf_salice"' not in response.text
    assert 'id="pdf_alarcon"' not in response.text
    assert "Control de cobradores" not in response.text
    assert "Subí el PDF de Pedidos Móviles" not in response.text


def test_compare_api_mode_passes_optional_collector_pdf(monkeypatch):
    client = TestClient(app_module.app)

    def _fake_compare(_excel_path, pdfs, **kwargs):
        assert kwargs.get("receipts_source") == "api"
        assert len(pdfs) == 1
        pdf_path, company = pdfs[0]
        assert company == "SALICE"
        with open(pdf_path, "rb") as fh:
            assert fh.read() == b"collector-pdf"
        return {"validados": [], "dudosos": [], "no_encontrados": [], "meta": {}}

    monkeypatch.setattr(app_module, "compare_excel_pdfs", _fake_compare)
    files = {
        "excel": (
            "legacy.xlsx",
            _minimal_xlsx_bytes(),
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        ),
        "pdf_salice": ("cobradores.pdf", b"collector-pdf", "application/pdf"),
    }

    response = client.post("/compare?receipts_source=api", files=files)

    assert response.status_code == 200, response.text


def test_compare_pdf_mode_requires_pdf(monkeypatch):
    client = TestClient(app_module.app)

    xbytes = _minimal_xlsx_bytes()
    files = {
        "excel": ("legacy.xlsx", xbytes, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"),
    }

    r = client.post("/compare?receipts_source=pdf", files=files)
    assert r.status_code == 400
    assert "al menos 1 PDF" in r.text


def test_compare_api_mode_passes_manual_receipts_date_range(monkeypatch):
    client = TestClient(app_module.app)
    observed = {"start": None, "end": None}

    def _fake_compare(*args, **kwargs):
        observed["start"] = kwargs.get("api_start_date_override")
        observed["end"] = kwargs.get("api_end_date_override")
        return {"validados": [], "dudosos": [], "no_encontrados": [], "meta": {"ok": True}}

    monkeypatch.setattr(app_module, "compare_excel_pdfs", _fake_compare)

    xbytes = _minimal_xlsx_bytes()
    files = {
        "excel": ("legacy.xlsx", xbytes, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"),
    }

    r = client.post(
        "/compare?receipts_source=api&api_start_date=2026-03-01&api_end_date=2026-03-07",
        files=files,
    )
    assert r.status_code == 200, r.text
    assert observed["start"] == "2026-03-01"
    assert observed["end"] == "2026-03-07"
