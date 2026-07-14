from __future__ import annotations

import json

from src.conciliador.collector_catalog import load_internal_collector_receipts


def test_bundled_internal_catalog_has_validated_july_collectors(monkeypatch):
    monkeypatch.delenv("CONCILIADOR_COBRADORES_PATH", raising=False)

    receipts, meta, warnings = load_internal_collector_receipts()
    by_number = {receipt.nro_recibo: receipt.vendedor for receipt in receipts}

    assert len(receipts) == 181
    assert by_number["78535"] == "206 - Andres Dominguez"
    assert by_number["78718"] == "206 - Andres Dominguez"
    assert meta["internal_collector_catalog_loaded"] is True
    assert any("181" in warning for warning in warnings)


def test_internal_catalog_can_be_replaced_server_side(monkeypatch, tmp_path):
    catalog = tmp_path / "cobradores.json"
    catalog.write_text(
        json.dumps({"companies": {"SALICE": {"90001": "337 - Agustin Rodriguez"}}}),
        encoding="utf-8",
    )
    monkeypatch.setenv("CONCILIADOR_COBRADORES_PATH", str(catalog))

    receipts, meta, _warnings = load_internal_collector_receipts()

    assert [(row.empresa, row.nro_recibo, row.vendedor) for row in receipts] == [
        ("SALICE", "90001", "337 - Agustin Rodriguez")
    ]
    assert meta["internal_collector_catalog_count"] == 1
