from __future__ import annotations

import datetime as dt

from src.conciliador.external.errors import ExternalProviderError
from src.conciliador.external.repartos_api_client import (
    fetch_repartos_detail,
    match_receipts_to_fleteros,
)


def _foja(foja_id, day, repartidor_id, name, invoices, empresa_id=3):
    return {
        "empresaID": empresa_id,
        "fojaID": foja_id,
        "desde": f"{day}T00:00:00",
        "repartidorID": repartidor_id,
        "descripcionRepartidor": name,
        "listaDeComprobantesAsignados": [
            {"clienteID": client, "importeTotal": amount}
            for client, amount in invoices
        ],
    }


def _receipt(number, client, amount, day="2026-07-10", empresa_id=3):
    return {
        "empresaID": empresa_id,
        "numero": number,
        "clienteID": client,
        "importeTotal": amount,
        "fechaDeEmision": f"{day}T00:00:00",
        "codigoDeImportacion": f"PMCBR_{number}",
    }


def test_same_client_and_different_amount_can_have_different_fleteros():
    receipts = [
        _receipt(1, 16610, 174592.79),
        _receipt(2, 16610, 107758.64),
    ]
    fojas = [
        _foja(10, "2026-07-09", 381, "Gonzalez Adrian", [(16610, 174592.79)]),
        _foja(11, "2026-07-09", 56, "Lencina Ramon", [(16610, 107758.64)]),
    ]

    matches, _ = match_receipts_to_fleteros(receipts, fojas, {})

    assert matches[0].label == "381 - Gonzalez Adrian"
    assert matches[0].source == "invoice_exact"
    assert matches[1].label == "56 - Lencina Ramon"


def test_sum_of_multiple_invoices_identifies_fletero():
    receipts = [_receipt(1, 17939, 188781.92)]
    fojas = [
        _foja(
            30,
            "2026-07-07",
            102,
            "Diaz David",
            [(17939, 125261.85), (17939, 63520.07)],
        )
    ]

    matches, _ = match_receipts_to_fleteros(receipts, fojas, {})

    assert matches[0].label == "102 - Diaz David"
    assert matches[0].source == "foja_sum_exact"


def test_partial_payment_uses_unique_recent_route_but_not_contradictory_routes():
    receipts = [_receipt(1, 100, 25), _receipt(2, 200, 25)]
    fojas = [
        _foja(40, "2026-07-09", 10, "Ruta Unica", [(100, 100)]),
        _foja(41, "2026-07-09", 20, "Ruta A", [(200, 100)]),
        _foja(42, "2026-07-08", 21, "Ruta B", [(200, 80)]),
    ]

    matches, _ = match_receipts_to_fleteros(receipts, fojas, {})

    assert matches[0].label == "10 - Ruta Unica"
    assert matches[0].source == "unique_client_route"
    assert 1 not in matches


def test_vendor_zone_subzone_fallback_learns_only_from_strong_same_batch_matches():
    receipts = [
        _receipt(1, 100, 100),
        _receipt(2, 101, 200),
        _receipt(3, 102, 50),
        _receipt(4, 103, 60),
    ]
    fojas = [
        _foja(50, "2026-07-09", 30, "Ruta Aprendida", [(100, 100), (101, 200)]),
        # El cliente 103 sí aparece en una foja: un perfil genérico no puede taparlo.
        _foja(51, "2026-07-09", 31, "Ruta Incierta", [(103, 999)]),
    ]
    features = {
        str(cid): {"vendedor_id": "7", "zona_id": "8", "subzona_id": "9"}
        for cid in (100, 101, 102, 103)
    }

    matches, _ = match_receipts_to_fleteros(receipts, fojas, features)

    assert matches[2].label == "30 - Ruta Aprendida"
    assert matches[2].source == "batch_vendor_zone_subzone"
    assert matches[3].label == "31 - Ruta Incierta"
    assert matches[3].source == "unique_client_route"


def test_stale_wrong_route_does_not_displace_strong_current_assignment():
    receipts = [_receipt(1, 17204, 175)]
    fojas = [
        _foja(60, "2026-07-09", 109, "Ruta Actual", [(17204, 175)]),
        _foja(61, "2026-07-08", 420, "Ruta Historica", [(17204, 999)]),
    ]

    matches, _ = match_receipts_to_fleteros(receipts, fojas, {})

    assert matches[0].label == "109 - Ruta Actual"
    assert matches[0].source == "invoice_exact"


def test_no_trustworthy_route_returns_empty_instead_of_vendor():
    receipts = [_receipt(1, 500, 50)]
    fojas = [
        _foja(70, "2026-07-09", 1, "Ruta A", [(500, 100)]),
        _foja(71, "2026-07-08", 2, "Ruta B", [(500, 100)]),
    ]
    features = {"500": {"vendedor_id": "999", "zona_id": "8", "subzona_id": "9"}}

    matches, warnings = match_receipts_to_fleteros(receipts, fojas, features)

    assert matches == {}
    assert any("1 recibos quedaron sin identificar" in warning for warning in warnings)


def test_same_client_id_isolated_by_empresa():
    receipts = [_receipt(1, 100, 50, empresa_id=3), _receipt(2, 100, 50, empresa_id=6)]
    fojas = [
        _foja(80, "2026-07-09", 3, "Salice", [(100, 50)], empresa_id=3),
        _foja(81, "2026-07-09", 6, "Alarcon", [(100, 50)], empresa_id=6),
    ]

    matches, _ = match_receipts_to_fleteros(receipts, fojas, {})

    assert matches[0].label == "3 - Salice"
    assert matches[1].label == "6 - Alarcon"


def test_fetch_repartos_uses_salice_empresa_and_skips_individual_500(monkeypatch):
    seen_headers = []

    def _headers(**kwargs):
        seen_headers.append(kwargs["empresa_id"])
        return {"empresaID": kwargs["empresa_id"]}

    def _http(url, **kwargs):
        if "GetList" in url:
            return ({
                "fojasReparto": [
                    {"empresaID": 3, "fojaID": 1, "desde": "2026-07-09"},
                    {"empresaID": 3, "fojaID": 2, "desde": "2026-07-09"},
                ],
                "paginacion": {"totalPaginas": 1},
            }, None)
        if "fojaID=1" in url:
            return ({"fojaReparto": _foja(1, "2026-07-09", 10, "Ruta", [(1, 100)])}, None)
        raise ExternalProviderError("receipts", "falló foja", status_code=500)

    monkeypatch.setattr(
        "src.conciliador.external.repartos_api_client._build_auth_headers_for_empresa",
        _headers,
    )
    monkeypatch.setattr("src.conciliador.external.repartos_api_client._http_json", _http)

    details, warnings = fetch_repartos_detail(
        start_date=dt.date(2026, 7, 10),
        end_date=dt.date(2026, 7, 10),
        empresa_filter="SALICE",
    )

    assert seen_headers == ["3"]
    assert len(details) == 1
    assert any("1 fojas no pudieron leerse" in warning for warning in warnings)
