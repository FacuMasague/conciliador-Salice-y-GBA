from __future__ import annotations

import json
import os
from pathlib import Path
from typing import Any

from .pdf_parser import Receipt


DEFAULT_CATALOG_PATH = Path(__file__).resolve().parents[2] / "data" / "cobradores.json"


def _catalog_paths() -> list[Path]:
    configured = str(os.getenv("CONCILIADOR_COBRADORES_PATH", "") or "").strip()
    if not configured:
        return [DEFAULT_CATALOG_PATH]
    return [Path(value.strip()) for value in configured.split(os.pathsep) if value.strip()]


def _normalize_catalog(payload: Any) -> dict[str, dict[str, str]]:
    if not isinstance(payload, dict):
        raise ValueError("el catálogo debe ser un objeto JSON")
    companies = payload.get("companies", payload)
    if not isinstance(companies, dict):
        raise ValueError("falta el objeto 'companies'")

    normalized: dict[str, dict[str, str]] = {}
    for raw_company, raw_receipts in companies.items():
        company = str(raw_company or "").strip().upper()
        if not company or not isinstance(raw_receipts, dict):
            continue
        company_map: dict[str, str] = {}
        for raw_receipt, raw_collector in raw_receipts.items():
            receipt = str(raw_receipt or "").strip()
            collector = str(raw_collector or "").strip()
            if receipt and collector:
                company_map[receipt] = collector
        if company_map:
            normalized[company] = company_map
    return normalized


def load_internal_collector_receipts() -> tuple[list[Receipt], dict[str, Any], list[str]]:
    """Carga la relación interna recibo -> cobrador administrada por el sistema.

    El operador de la web nunca tiene que subir este archivo. En Render se usa
    el catálogo incluido en el despliegue, o una ruta privada indicada mediante
    ``CONCILIADOR_COBRADORES_PATH`` cuando se necesite actualizarlo sin tocar la UI.
    """
    merged: dict[tuple[str, str], str] = {}
    loaded_paths: list[str] = []
    warnings: list[str] = []

    for path in _catalog_paths():
        try:
            payload = json.loads(path.read_text(encoding="utf-8"))
            catalog = _normalize_catalog(payload)
        except FileNotFoundError:
            warnings.append(f"Cobradores: no se encontró el catálogo interno {path.name}.")
            continue
        except Exception as exc:
            warnings.append(f"Cobradores: no se pudo leer el catálogo interno {path.name}: {exc}")
            continue

        loaded_paths.append(str(path))
        for company, receipts in catalog.items():
            for receipt, collector in receipts.items():
                merged[(company, receipt)] = collector

    rows = [
        Receipt(
            empresa=company,
            nro_recibo=receipt,
            nro_cliente="",
            cliente_nombre=None,
            vendedor=collector,
        )
        for (company, receipt), collector in sorted(merged.items())
    ]
    if rows:
        warnings.append(
            f"Cobradores: se cargaron {len(rows)} asignaciones desde la fuente interna."
        )
    elif not warnings:
        warnings.append("Cobradores: la fuente interna no contiene asignaciones.")

    meta = {
        "internal_collector_catalog_count": len(rows),
        "internal_collector_catalog_files": len(loaded_paths),
        "internal_collector_catalog_loaded": bool(rows),
    }
    return rows, meta, warnings
