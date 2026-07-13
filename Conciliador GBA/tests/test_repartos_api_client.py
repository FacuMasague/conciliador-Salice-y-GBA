from __future__ import annotations

from src.conciliador.external.repartos_api_client import match_receipts_to_fleteros


def _foja(foja_id, day, repartidor_id, name, invoices):
    return {
        "fojaID": foja_id,
        "desde": f"{day}T00:00:00",
        "repartidorID": repartidor_id,
        "descripcionRepartidor": name,
        "listaDeComprobantesAsignados": [
            {"clienteID": client, "importeTotal": amount}
            for client, amount in invoices
        ],
    }


def _receipt(number, client, amount, day="2026-07-10"):
    return {
        "numero": number,
        "clienteID": client,
        "importeTotal": amount,
        "fechaDeEmision": f"{day}T00:00:00",
    }


def test_same_client_can_be_assigned_to_different_fleteros_by_exact_amount():
    receipts = [
        _receipt(1, 16610, 174592.79),
        _receipt(2, 16610, 107758.64),
    ]
    fojas = [
        _foja(10, "2026-07-09", 381, "Gonzalez Adrian", [(16610, 174592.79)]),
        _foja(11, "2026-07-09", 56, "Lencina Ramon", [(16610, 107758.64)]),
    ]

    matches, _warnings = match_receipts_to_fleteros(receipts, fojas, {})

    assert matches[0].label == "381 - Gonzalez Adrian"
    assert matches[0].source == "invoice_exact"
    assert matches[1].label == "56 - Lencina Ramon"


def test_profile_from_exact_fojas_overrides_stale_route_and_fills_missing_clients():
    receipts = [
        _receipt(1, 17000, 100),
        _receipt(2, 17001, 200),
        _receipt(3, 17204, 175),
        _receipt(4, 17002, 250),
        _receipt(5, 15812, 300),
        _receipt(6, 15824, 400),
        _receipt(7, 18657, 199),
    ]
    for receipt in receipts:
        receipt["codigoDeImportacion"] = f"PMCBR_{receipt['numero']}"
    fojas = [
        _foja(20, "2026-07-09", 109, "Pita Carlos", [(17000, 100), (17001, 200), (17002, 250)]),
        _foja(21, "2026-07-09", 102, "Diaz David", [(15812, 300), (15824, 400)]),
        # Ruta histórica incorrecta para el lote actual: no debe ganar por ser la última.
        _foja(22, "2026-07-08", 420, "Buletti Horacio", [(17204, 999)]),
    ]
    client_features = {
        "17000": {"vendedor_id": "65", "zona_id": "7613", "subzona_id": "4"},
        "17001": {"vendedor_id": "65", "zona_id": "7613", "subzona_id": "4"},
        "17002": {"vendedor_id": "65", "zona_id": "7613", "subzona_id": "4"},
        "17204": {"vendedor_id": "65", "zona_id": "7613", "subzona_id": "4"},
        "15812": {"vendedor_id": "1011", "zona_id": "7613", "subzona_id": "3"},
        "15824": {"vendedor_id": "1011", "zona_id": "7613", "subzona_id": "3"},
        "18657": {"vendedor_id": "1011", "zona_id": "7613", "subzona_id": "3"},
    }

    matches, _warnings = match_receipts_to_fleteros(receipts, fojas, client_features)

    assert matches[2].label == "109 - Pita Carlos"
    assert matches[2].source == "batch_neighbor_profile"
    assert matches[6].label == "102 - Diaz David"
    assert matches[6].source == "batch_vendor_zone_subzone"


def test_sum_of_invoices_can_identify_fletero():
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

    matches, _warnings = match_receipts_to_fleteros(receipts, fojas, {})

    assert matches[0].label == "102 - Diaz David"
    assert matches[0].source == "foja_sum_exact"
