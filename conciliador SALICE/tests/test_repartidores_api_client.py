from __future__ import annotations

from src.conciliador.external.padron_api_client import fetch_repartidores_payload


def test_fetch_repartidores_payload_uses_official_getlist_and_paginates(monkeypatch):
    calls: list[str] = []

    monkeypatch.setenv("API_MODE_ENABLED", "true")
    monkeypatch.setenv("PADRON_API_BASE_URL", "https://gesi.example")
    monkeypatch.setattr(
        "src.conciliador.external.padron_api_client._login_token",
        lambda prefix, base, headers: "token",
    )

    def fake_http(url, *, method, headers, body=None):
        calls.append(url)
        page = 2 if "pageNumber=2" in url else 1
        return (
            {
                "success": True,
                "repartidores": [
                    {"repartidorID": 64 + page, "razonSocial": f"Repartidor {page}"}
                ],
                "paginacion": {"totalPaginas": 2},
            },
            f"request-{page}",
        )

    monkeypatch.setattr(
        "src.conciliador.external.padron_api_client._http_json", fake_http
    )

    response = fetch_repartidores_payload()

    assert [row["repartidorID"] for row in response.payload["repartidores"]] == [65, 66]
    assert response.request_id == "request-2"
    assert calls == [
        "https://gesi.example/api/Maestros/Repartidores/GetList?pageNumber=1&pageSize=500",
        "https://gesi.example/api/Maestros/Repartidores/GetList?pageNumber=2&pageSize=500",
    ]
