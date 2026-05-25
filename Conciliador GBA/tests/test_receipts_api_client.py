from __future__ import annotations

import datetime as dt

from src.conciliador.external.errors import ExternalTimeoutError
from src.conciliador.external.receipts_api_client import (
    _get_paged_get,
    _http_json,
    _page_size_for_targets,
    fetch_receipts_payload,
)


def _sample_cobro(numero: int) -> dict:
    return {
        "empresaID": 2,
        "comprobanteID": numero,
        "serie": "X",
        "puntoDeVentaID": 1,
        "numero": numero,
        "clienteID": 1000 + numero,
        "fechaDeEmision": "2026-03-16",
        "importeTotal": 100.0,
        "formaDePagoID": 0,
        "detalleDeValores": [],
    }


def test_fetch_receipts_payload_prefers_cobros_for_gba(monkeypatch):
    calls: list[tuple[str, str]] = []
    getitem_calls: list[list[dict]] = []

    monkeypatch.setenv("API_MODE_ENABLED", "true")
    monkeypatch.delenv("RECEIPTS_API_FORCE_COBROS_ONLY", raising=False)
    monkeypatch.delenv("RECEIPTS_API_GETITEM_MAX_KEYS", raising=False)
    monkeypatch.setattr(
        "src.conciliador.external.receipts_api_client._build_auth_headers_for_empresa",
        lambda **kwargs: {},
    )
    monkeypatch.setattr(
        "src.conciliador.external.receipts_api_client._fetch_medios_pago",
        lambda **kwargs: ([], None, []),
    )
    monkeypatch.setattr(
        "src.conciliador.external.receipts_api_client._get_paged_get",
        lambda **kwargs: ([], None, []),
    )

    def _fake_fetch_paged_rows(*, path, method, **kwargs):
        calls.append((path, method))
        if path == "/api/Ventas/Comprobantes/Cobros/GetList":
            return ([_sample_cobro(1)], None, True)
        return ([], None, True)

    monkeypatch.setattr(
        "src.conciliador.external.receipts_api_client._fetch_paged_rows",
        _fake_fetch_paged_rows,
    )
    monkeypatch.setattr(
        "src.conciliador.external.receipts_api_client._fetch_getitem_details",
        lambda **kwargs: (getitem_calls.append(list(kwargs["keys"])) or {}, None, False),
    )

    resp = fetch_receipts_payload(
        days=7,
        empresa_filter="GBA",
        start_date=dt.date(2026, 3, 16),
        end_date=dt.date(2026, 3, 22),
    )

    assert len(resp.payload["comprobantes"]) == 1
    assert any("omite Comprobantes/GetList" in w for w in resp.warnings)
    assert ("/api/Ventas/Comprobantes/Cobros/GetList", "POST") in calls
    assert len(getitem_calls) >= 1
    assert all(len(chunk) == 1 for chunk in getitem_calls)
    assert not any(
        path in {
            "/api/Ventas/Comprobantes/GetList",
            "/api/Ventas/Comprobantes/GetListComprobantes",
            "/api/Ventas/Comprobantes/List",
            "/api/Maestros/Comprobantes/GetList",
        }
        for path, _method in calls
    )


def test_fetch_receipts_payload_skips_getitem_when_limit_is_zero(monkeypatch):
    monkeypatch.setenv("API_MODE_ENABLED", "true")
    monkeypatch.setenv("RECEIPTS_API_FORCE_COBROS_ONLY", "true")
    monkeypatch.setenv("RECEIPTS_API_GETITEM_MAX_KEYS", "0")
    monkeypatch.setattr(
        "src.conciliador.external.receipts_api_client._build_auth_headers_for_empresa",
        lambda **kwargs: {},
    )
    monkeypatch.setattr(
        "src.conciliador.external.receipts_api_client._fetch_medios_pago",
        lambda **kwargs: ([], None, []),
    )
    monkeypatch.setattr(
        "src.conciliador.external.receipts_api_client._get_paged_get",
        lambda **kwargs: ([], None, []),
    )

    def _fake_fetch_paged_rows(*, path, **kwargs):
        if path == "/api/Ventas/Comprobantes/Cobros/GetList":
            return ([_sample_cobro(1), _sample_cobro(2)], None, True)
        return ([], None, True)

    monkeypatch.setattr(
        "src.conciliador.external.receipts_api_client._fetch_paged_rows",
        _fake_fetch_paged_rows,
    )
    monkeypatch.setattr(
        "src.conciliador.external.receipts_api_client._fetch_getitem_details",
        lambda **kwargs: (_ for _ in ()).throw(AssertionError("GetItem debería haberse omitido")),
    )

    resp = fetch_receipts_payload(
        days=7,
        empresa_filter="GBA",
        start_date=dt.date(2026, 3, 16),
        end_date=dt.date(2026, 3, 22),
    )

    assert len(resp.payload["comprobantes"]) == 2
    assert any("Cobros/GetItem omitido" in w for w in resp.warnings)


def test_getitem_max_keys_for_gba_default_is_enabled(monkeypatch):
    monkeypatch.delenv("RECEIPTS_API_GETITEM_MAX_KEYS", raising=False)
    from src.conciliador.external.receipts_api_client import _getitem_max_keys

    assert _getitem_max_keys(["2"], "GBA") == 2000


def test_fetch_receipts_payload_retries_with_smaller_page_size_on_timeout(monkeypatch):
    calls: list[int] = []

    monkeypatch.setenv("API_MODE_ENABLED", "true")
    monkeypatch.setenv("RECEIPTS_API_FORCE_COBROS_ONLY", "true")
    monkeypatch.setenv("RECEIPTS_API_PAGE_SIZE", "2000")
    monkeypatch.setenv("RECEIPTS_API_PAGE_SIZE_FALLBACKS", "300,200")
    monkeypatch.setattr(
        "src.conciliador.external.receipts_api_client._build_auth_headers_for_empresa",
        lambda **kwargs: {},
    )
    monkeypatch.setattr(
        "src.conciliador.external.receipts_api_client._fetch_medios_pago",
        lambda **kwargs: ([], None, []),
    )
    monkeypatch.setattr(
        "src.conciliador.external.receipts_api_client._get_paged_get",
        lambda **kwargs: ([], None, []),
    )

    def _fake_fetch_paged_rows(*, path, page_size, **kwargs):
        calls.append(int(page_size))
        if path == "/api/Ventas/Comprobantes/Cobros/GetList" and int(page_size) > 300:
            raise ExternalTimeoutError("receipts", "Timeout en Receipts API")
        if path == "/api/Ventas/Comprobantes/Cobros/GetList":
            return ([_sample_cobro(1), _sample_cobro(2)], None, True)
        return ([], None, True)

    monkeypatch.setattr(
        "src.conciliador.external.receipts_api_client._fetch_paged_rows",
        _fake_fetch_paged_rows,
    )
    monkeypatch.setattr(
        "src.conciliador.external.receipts_api_client._fetch_getitem_details",
        lambda **kwargs: ({}, None, False),
    )

    resp = fetch_receipts_payload(
        days=7,
        empresa_filter="GBA",
        start_date=dt.date(2026, 3, 16),
        end_date=dt.date(2026, 3, 22),
    )

    assert len(resp.payload["comprobantes"]) == 2
    assert calls[:2] == [2000, 300]
    assert any("timeout con pageSize=2000" in w for w in resp.warnings)


def test_get_paged_get_retries_with_smaller_page_size_on_timeout(monkeypatch):
    calls: list[int] = []

    monkeypatch.setenv("RECEIPTS_API_PAGE_SIZE_FALLBACKS", "300,200")

    def _fake_fetch_paged_rows(*, path, page_size, **kwargs):
        calls.append(int(page_size))
        if path == "/api/Maestros/Empresas/GetList" and int(page_size) > 300:
            raise ExternalTimeoutError("receipts", "Timeout en Receipts API")
        return ([{"empresaID": 2, "descripcion": "GBA"}], None, True)

    monkeypatch.setattr(
        "src.conciliador.external.receipts_api_client._fetch_paged_rows",
        _fake_fetch_paged_rows,
    )

    rows, _rid, warnings = _get_paged_get(
        base="https://m5gba.grupoesi.com.ar",
        path="/api/Maestros/Empresas/GetList",
        headers={},
        list_key="empresas",
        page_size=2000,
    )

    assert rows == [{"empresaID": 2, "descripcion": "GBA"}]
    assert calls[:2] == [2000, 300]
    assert any("timeout con pageSize=2000" in w for w in warnings)


def test_page_size_for_gba_defaults_to_100(monkeypatch):
    monkeypatch.delenv("RECEIPTS_API_PAGE_SIZE", raising=False)
    assert _page_size_for_targets(["2"], "GBA") == 100


def test_http_json_maps_builtin_timeout_to_external_timeout(monkeypatch):
    monkeypatch.setattr(
        "src.conciliador.external.receipts_api_client.urlopen",
        lambda *args, **kwargs: (_ for _ in ()).throw(TimeoutError("The read operation timed out")),
    )
    monkeypatch.setattr(
        "src.conciliador.external.receipts_api_client._ssl_context",
        lambda: None,
    )

    try:
        _http_json("https://m5gba.grupoesi.com.ar/api/test", method="GET", headers={})
    except ExternalTimeoutError:
        pass
    else:
        raise AssertionError("Se esperaba ExternalTimeoutError")


def test_fetch_receipts_payload_splits_gba_windows_and_accumulates(monkeypatch):
    calls: list[tuple[str, str]] = []

    def _fake_single(*, days, empresa_filter=None, start_date=None, end_date=None):
        calls.append((start_date.isoformat(), end_date.isoformat()))
        numero = int(start_date.strftime("%d"))
        return type("Resp", (), {
            "payload": {
                "comprobantes": [_sample_cobro(numero)],
                "formasDePago": [{"formaDePagoID": 1}],
                "empresas": [{"empresaID": 2}],
                "mediosDePago": [{"valorID": 9}],
                "comprobantes_count_by_target": {"2": 1},
                "api_comprobantes_path_used": "/api/Ventas/Comprobantes/Cobros/GetList",
                "api_comprobantes_method_used": "POST",
            },
            "request_id": f"rid-{numero}",
            "warnings": [],
        })()

    monkeypatch.setenv("API_MODE_ENABLED", "true")
    monkeypatch.delenv("RECEIPTS_API_WINDOW_DAYS", raising=False)
    monkeypatch.setattr(
        "src.conciliador.external.receipts_api_client._fetch_receipts_payload_single_window",
        _fake_single,
    )

    resp = fetch_receipts_payload(
        days=3,
        empresa_filter="GBA",
        start_date=dt.date(2026, 4, 6),
        end_date=dt.date(2026, 4, 8),
    )

    assert calls == [("2026-04-06", "2026-04-06"), ("2026-04-07", "2026-04-07"), ("2026-04-08", "2026-04-08")]
    assert len(resp.payload["comprobantes"]) == 3
    assert resp.payload["comprobantes_count_by_target"] == {"2": 3}
    assert any("troceada en 3 ventanas" in w for w in resp.warnings)


def test_fetch_receipts_payload_keeps_partial_windows_on_timeout(monkeypatch):
    calls: list[str] = []

    def _fake_single(*, days, empresa_filter=None, start_date=None, end_date=None):
        day = start_date.isoformat()
        calls.append(day)
        if day == "2026-04-07":
            raise ExternalTimeoutError("receipts", "Timeout en Receipts API")
        numero = int(start_date.strftime("%d"))
        return type("Resp", (), {
            "payload": {
                "comprobantes": [_sample_cobro(numero)],
                "formasDePago": [],
                "empresas": [],
                "mediosDePago": [],
                "comprobantes_count_by_target": {"2": 1},
                "api_comprobantes_path_used": "/api/Ventas/Comprobantes/Cobros/GetList",
                "api_comprobantes_method_used": "POST",
            },
            "request_id": f"rid-{numero}",
            "warnings": [],
        })()

    monkeypatch.setenv("API_MODE_ENABLED", "true")
    monkeypatch.delenv("RECEIPTS_API_WINDOW_DAYS", raising=False)
    monkeypatch.setattr(
        "src.conciliador.external.receipts_api_client._fetch_receipts_payload_single_window",
        _fake_single,
    )

    resp = fetch_receipts_payload(
        days=3,
        empresa_filter="GBA",
        start_date=dt.date(2026, 4, 6),
        end_date=dt.date(2026, 4, 8),
    )

    assert calls == ["2026-04-06", "2026-04-07", "2026-04-08"]
    assert [r["numero"] for r in resp.payload["comprobantes"]] == [6, 8]
    assert resp.payload["comprobantes_count_by_target"] == {"2": 2}
    assert any("se conserva lo ya descargado" in w for w in resp.warnings)
