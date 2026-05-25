from __future__ import annotations

from src.conciliador.external.errors import ExternalTimeoutError
from src.conciliador.external.padron_api_client import (
    _get_clientes_getlist,
    _http_json,
    _page_size,
    fetch_padron_payload,
)


def test_page_size_for_gba_defaults_to_500(monkeypatch):
    monkeypatch.delenv("PADRON_API_PAGE_SIZE", raising=False)
    assert _page_size("GBA") == 500


def test_http_json_maps_builtin_timeout_to_external_timeout(monkeypatch):
    monkeypatch.setattr(
        "src.conciliador.external.padron_api_client.urlopen",
        lambda *args, **kwargs: (_ for _ in ()).throw(TimeoutError("The read operation timed out")),
    )
    monkeypatch.setattr(
        "src.conciliador.external.padron_api_client._ssl_context",
        lambda: None,
    )

    try:
        _http_json("https://m5gba.grupoesi.com.ar/api/test", method="GET", headers={})
    except ExternalTimeoutError:
        pass
    else:
        raise AssertionError("Se esperaba ExternalTimeoutError")


def test_get_clientes_getlist_retries_with_smaller_page_size_on_timeout(monkeypatch):
    calls: list[int] = []

    monkeypatch.delenv("PADRON_API_PAGE_SIZE", raising=False)
    monkeypatch.setenv("PADRON_API_PAGE_SIZE_FALLBACKS", "300,200")

    def _fake_once(*, page_size, **kwargs):
        calls.append(int(page_size))
        if int(page_size) > 300:
            raise ExternalTimeoutError("padron", "Timeout en Padrón API")
        return ([{"clienteID": 1}], "rid-1", [])

    monkeypatch.setattr(
        "src.conciliador.external.padron_api_client._get_clientes_getlist_once",
        _fake_once,
    )

    rows, rid, warnings = _get_clientes_getlist(
        base="https://m5gba.grupoesi.com.ar",
        path="/api/Maestros/Clientes/GetList",
        headers={},
        empresa_filter="GBA",
    )

    assert rows == [{"clienteID": 1}]
    assert rid == "rid-1"
    assert calls[:2] == [500, 300]
    assert any("timeout con pageSize=500" in w for w in warnings)


def test_fetch_padron_payload_uses_targeted_getitem_when_cliente_ids_present(monkeypatch):
    calls = []

    monkeypatch.setenv("API_MODE_ENABLED", "true")
    monkeypatch.setattr(
        "src.conciliador.external.padron_api_client._base_url",
        lambda prefix: "https://m5gba.grupoesi.com.ar",
    )
    monkeypatch.setattr(
        "src.conciliador.external.padron_api_client._headers_base",
        lambda prefix: {},
    )
    monkeypatch.setattr(
        "src.conciliador.external.padron_api_client._login_token",
        lambda *args, **kwargs: "tok",
    )
    monkeypatch.setattr(
        "src.conciliador.external.padron_api_client._get_clientes_targeted",
        lambda **kwargs: (calls.append(kwargs["cliente_ids"]) or ([{"clienteID": 1, "numeroDeDocumento": "20-12345678-9"}], "rid-pad", [])),
    )
    monkeypatch.setattr(
        "src.conciliador.external.padron_api_client._get_clientes_getlist",
        lambda **kwargs: (_ for _ in ()).throw(AssertionError("No debería usar GetList")),
    )

    resp = fetch_padron_payload(empresa_filter="GBA", cliente_ids=["1", "1", "2"])

    assert len(resp.payload["clientes"]) == 1
    assert calls == [["1", "1", "2"]]
    assert any("dirigido a 2 cliente(s)" in w for w in resp.warnings)
