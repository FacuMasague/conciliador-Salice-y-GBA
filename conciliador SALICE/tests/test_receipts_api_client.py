from __future__ import annotations

import datetime as dt
from urllib.parse import parse_qs, urlparse

from src.conciliador.external.receipts_api_client import fetch_receipts_payload


def test_fetch_receipts_payload_queries_dual_company_and_until_yesterday(monkeypatch):
    called_empresas: list[str] = []

    def _fake_build_auth_headers_for_empresa(*, base, headers_root, empresa_id, drop_sucursal):
        called_empresas.append(str(empresa_id))
        return {"Authorization": "Bearer x", "empresaID": str(empresa_id)}

    def _fake_http_json(url, *, method, headers, body=None):
        if "Comprobantes/GetList" in url:
            eid = str(headers.get("empresaID") or "")
            q = parse_qs(urlparse(url).query)
            page = int((q.get("pageNumber") or ["1"])[0])
            assert page == 1
            return (
                {
                    "success": True,
                    "comprobantes": [
                        {
                            "empresaID": 99,
                            "sucursal_id": int(eid),
                            "numero": 1000 + int(eid),
                            "codigoDeImportacion": f"PMCBR_{1000 + int(eid)}",
                            "clienteID": 1,
                            "fechaDeEmision": "2026-02-18",
                            "importeTotal": 10.0,
                            "formaDePagoID": 5,
                        }
                    ],
                    "paginacion": {"totalPaginas": 1},
                },
                f"rid-{eid}",
            )
        if "FormasDePago/GetList" in url:
            return (
                {"formasDePago": [{"formaDePagoID": 5, "descripcion": "Transferencia"}], "paginacion": {"totalPaginas": 1}},
                "rid-formas",
            )
        if "Empresas/GetList" in url:
            return (
                {
                    "empresas": [
                        {"empresaID": 3, "descripcion": "SALICE"},
                        {"empresaID": 6, "descripcion": "ALARCON"},
                    ],
                    "paginacion": {"totalPaginas": 1},
                },
                "rid-emp",
            )
        raise AssertionError(f"url inesperada: {url}")

    monkeypatch.setattr("src.conciliador.external.receipts_api_client._build_auth_headers_for_empresa", _fake_build_auth_headers_for_empresa)
    monkeypatch.setattr("src.conciliador.external.receipts_api_client._http_json", _fake_http_json)
    monkeypatch.setenv("API_MODE_ENABLED", "true")

    resp = fetch_receipts_payload(days=15, empresa_filter=None)
    payload = resp.payload

    assert called_empresas[:2] == ["3", "6"]
    assert payload.get("empresa_targets_used") == ["3", "6"]
    assert payload.get("comprobantes_count_by_target") == {"3": 1, "6": 1}
    assert len(payload.get("comprobantes") or []) == 2
    assert payload.get("fecha_desde") == (dt.date.today() - dt.timedelta(days=15)).isoformat()
    assert payload.get("fecha_hasta") == (dt.date.today() - dt.timedelta(days=1)).isoformat()
    assert payload.get("api_comprobantes_path_used") == "/api/Ventas/Comprobantes/GetList"
    assert payload.get("api_comprobantes_method_used") == "POST"
    assert [str(r.get("empresaID")) for r in (payload.get("comprobantes") or [])] == ["3", "6"]


def test_fetch_receipts_payload_respects_explicit_end_date(monkeypatch):
    observed_ranges: list[tuple[str, str]] = []

    def _fake_build_auth_headers_for_empresa(*, base, headers_root, empresa_id, drop_sucursal):
        return {"Authorization": "Bearer x", "empresaID": str(empresa_id)}

    def _fake_http_json(url, *, method, headers, body=None):
        if "Ventas/Comprobantes/GetList" in url or "Cobros/GetList" in url:
            datos = (body or {}).get("datosOperacion") or {}
            observed_ranges.append((str(datos.get("FechaDesde") or ""), str(datos.get("FechaHasta") or "")))
            return ({"success": True, "comprobantes": [], "paginacion": {"totalPaginas": 1}}, "rid-1")
        if "FormasDePago/GetList" in url:
            return ({"formasDePago": [], "paginacion": {"totalPaginas": 1}}, "rid-formas")
        if "Empresas/GetList" in url:
            return ({"empresas": [], "paginacion": {"totalPaginas": 1}}, "rid-emp")
        raise AssertionError(f"url inesperada: {url}")

    monkeypatch.setattr("src.conciliador.external.receipts_api_client._build_auth_headers_for_empresa", _fake_build_auth_headers_for_empresa)
    monkeypatch.setattr("src.conciliador.external.receipts_api_client._http_json", _fake_http_json)
    monkeypatch.setenv("API_MODE_ENABLED", "true")

    resp = fetch_receipts_payload(days=15, empresa_filter="SALICE", end_date=dt.date(2026, 3, 3))

    assert resp.payload.get("fecha_desde") == "2026-02-17"
    assert resp.payload.get("fecha_hasta") == "2026-03-03"
    assert observed_ranges
    assert all(r == ("2026-02-17", "2026-03-03") for r in observed_ranges)


def test_fetch_receipts_payload_fallback_cobros_enriches_with_getitem(monkeypatch):
    def _fake_build_auth_headers_for_empresa(*, base, headers_root, empresa_id, drop_sucursal):
        return {"Authorization": "Bearer x", "empresaID": str(empresa_id), "sucursalID": str(empresa_id)}

    def _fake_http_json(url, *, method, headers, body=None):
        if "Ventas/Comprobantes/GetList" in url:
            from src.conciliador.external.errors import ExternalProviderError
            raise ExternalProviderError("receipts", f"Receipts API HTTP 404: {url}", status_code=404)
        if "FormasDePago/GetList" in url:
            # Empty to force GetItem-enriched fields to be used.
            return ({"formasDePago": [], "paginacion": {"totalPaginas": 1}}, "rid-formas")
        if "Empresas/GetList" in url:
            return ({"empresas": [], "paginacion": {"totalPaginas": 1}}, "rid-emp")
        if "Cobros/GetList" in url:
            return (
                {
                    "success": True,
                    "comprobantes": [
                        {
                            "ComprobanteID": 124,
                            "EmpresaID": 3,
                            "Serie": "X",
                            "PuntoDeVentaID": 5,
                            "Numero": 78898,
                            "formaDePagoID": 0,
                            "clienteID": 10,
                            "fechaDeEmision": "2026-02-20",
                            "importeTotal": 100.0,
                            "codigoDeImportacion": "PMCBR_78898",
                        }
                    ],
                    "paginacion": {"totalPaginas": 1},
                },
                "rid-cobros-list",
            )
        if "Cobros/GetItem" in url:
            assert method.upper() == "POST"
            assert isinstance(body, list) and len(body) == 1
            return (
                {
                    "success": True,
                    "comprobantes": [
                        {
                            "ComprobanteID": 124,
                            "EmpresaID": 3,
                            "Serie": "X",
                            "PuntoDeVentaID": 5,
                            "Numero": 78898,
                            "FormaDePagoID": 9,
                            "DescripcionFormaDePago": "Transferencia Bancaria",
                        }
                    ],
                },
                "rid-cobros-item",
            )
        # Force 404-like fallback in paths that are not Cobros.
        from src.conciliador.external.errors import ExternalProviderError
        raise ExternalProviderError("receipts", f"Receipts API HTTP 404: {url}", status_code=404)

    monkeypatch.setattr("src.conciliador.external.receipts_api_client._build_auth_headers_for_empresa", _fake_build_auth_headers_for_empresa)
    monkeypatch.setattr("src.conciliador.external.receipts_api_client._http_json", _fake_http_json)
    monkeypatch.setenv("API_MODE_ENABLED", "true")

    resp = fetch_receipts_payload(days=15, empresa_filter="SALICE")
    payload = resp.payload
    rows = payload.get("comprobantes") or []
    assert len(rows) == 1
    assert payload.get("api_comprobantes_path_used") == "/api/Ventas/Comprobantes/Cobros/GetList"
    assert str(rows[0].get("FormaDePagoID") or "") == "9"
    assert str(rows[0].get("formaDePagoID") or "") in {"9", "9.0"}


def test_fetch_receipts_payload_fallback_cobros_getitem_adaptive_retry(monkeypatch):
    def _fake_build_auth_headers_for_empresa(*, base, headers_root, empresa_id, drop_sucursal):
        return {"Authorization": "Bearer x", "empresaID": str(empresa_id), "sucursalID": str(empresa_id)}

    def _fake_http_json(url, *, method, headers, body=None):
        if "Ventas/Comprobantes/GetList" in url:
            from src.conciliador.external.errors import ExternalProviderError
            raise ExternalProviderError("receipts", f"Receipts API HTTP 404: {url}", status_code=404)
        if "FormasDePago/GetList" in url:
            return ({"formasDePago": [], "paginacion": {"totalPaginas": 1}}, "rid-formas")
        if "Empresas/GetList" in url:
            return ({"empresas": [], "paginacion": {"totalPaginas": 1}}, "rid-emp")
        if "Cobros/GetList" in url:
            return (
                {
                    "success": True,
                    "comprobantes": [
                        {
                            "ComprobanteID": 124,
                            "EmpresaID": 3,
                            "Serie": "X",
                            "PuntoDeVentaID": 5,
                            "Numero": 78898,
                            "formaDePagoID": 0,
                            "clienteID": 10,
                            "fechaDeEmision": "2026-02-20",
                            "importeTotal": 100.0,
                            "codigoDeImportacion": "PMCBR_78898",
                        },
                        {
                            "ComprobanteID": 124,
                            "EmpresaID": 3,
                            "Serie": "X",
                            "PuntoDeVentaID": 5,
                            "Numero": 78899,
                            "formaDePagoID": 0,
                            "clienteID": 11,
                            "fechaDeEmision": "2026-02-20",
                            "importeTotal": 200.0,
                            "codigoDeImportacion": "PMCBR_78899",
                        },
                    ],
                    "paginacion": {"totalPaginas": 1},
                },
                "rid-cobros-list",
            )
        if "Cobros/GetItem" in url:
            # Simula truncado por lote grande: con body > 1 devuelve solo una fila.
            assert method.upper() == "POST"
            assert isinstance(body, list)
            if len(body) > 1:
                b = body[0]
                return (
                    {
                        "success": True,
                        "comprobantes": [
                            {
                                "ComprobanteID": b["ComprobanteID"],
                                "EmpresaID": b["EmpresaID"],
                                "Serie": b["Serie"],
                                "PuntoDeVentaID": b["PuntoDeVentaID"],
                                "Numero": b["Numero"],
                                "FormaDePagoID": 9,
                            }
                        ],
                    },
                    "rid-cobros-item-bulk",
                )
            b = body[0]
            return (
                {
                    "success": True,
                    "comprobantes": [
                        {
                            "ComprobanteID": b["ComprobanteID"],
                            "EmpresaID": b["EmpresaID"],
                            "Serie": b["Serie"],
                            "PuntoDeVentaID": b["PuntoDeVentaID"],
                            "Numero": b["Numero"],
                            "FormaDePagoID": 9,
                        }
                    ],
                },
                "rid-cobros-item-single",
            )
        from src.conciliador.external.errors import ExternalProviderError
        raise ExternalProviderError("receipts", f"Receipts API HTTP 404: {url}", status_code=404)

    monkeypatch.setattr("src.conciliador.external.receipts_api_client._build_auth_headers_for_empresa", _fake_build_auth_headers_for_empresa)
    monkeypatch.setattr("src.conciliador.external.receipts_api_client._http_json", _fake_http_json)
    monkeypatch.setenv("API_MODE_ENABLED", "true")

    resp = fetch_receipts_payload(days=15, empresa_filter="SALICE")
    rows = resp.payload.get("comprobantes") or []
    assert len(rows) == 2
    # Debe enriquecer ambos tras retry con lotes más chicos.
    assert all(str(r.get("formaDePagoID") or "") in {"9", "9.0"} for r in rows)
    assert any("retry adaptativo aplicado" in w for w in (resp.warnings or []))
