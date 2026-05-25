from pathlib import Path
import sys
import os

# Asegura que el root del repo esté importable (para `app.py`).
REPO_ROOT = Path(__file__).resolve().parents[1]
if str(REPO_ROOT) not in sys.path:
    sys.path.insert(0, str(REPO_ROOT))

import pytest


def _first_existing(candidates: list[Path]) -> Path | None:
    for p in candidates:
        if p.exists():
            return p
    return None


def _resolve_test_file(env_name: str, fallback_names: list[str]) -> str:
    env_value = os.getenv(env_name, "").strip()
    if env_value:
        p = Path(env_value)
        if p.exists():
            return str(p)
        pytest.skip(f"{env_name} apunta a archivo inexistente: {p}")

    roots = [
        Path("/mnt/data"),
        Path("/Users/facundomasague/Documents/Comparacion Ingresos Salice Alarcon/Datos 20-23"),
        Path("/Users/facundomasague/Documents/Comparacion Ingresos Salice Alarcon/Datos 16-26"),
        REPO_ROOT,
    ]
    candidates: list[Path] = []
    for r in roots:
        for name in fallback_names:
            candidates.append(r / name)

    found = _first_existing(candidates)
    if found:
        return str(found)
    pytest.skip(f"No se encontró archivo de fixture para {env_name} en rutas conocidas.")


@pytest.fixture(scope="session")
def excel_path() -> str:
    return _resolve_test_file(
        "CONCILIADOR_TEST_EXCEL",
        [
            "Movimientos bancarios 2026.xlsx",
            "mov bancarios 2026 ia.xlsx",
        ],
    )


@pytest.fixture(scope="session")
def pdf_salice_path() -> str:
    return _resolve_test_file(
        "CONCILIADOR_TEST_PDF_SALICE",
        [
            "reporte salice.pdf",
            "salice.pdf",
        ],
    )


@pytest.fixture(scope="session")
def pdf_alarcon_path() -> str:
    return _resolve_test_file(
        "CONCILIADOR_TEST_PDF_ALARCON",
        [
            "reporte alarcon.pdf",
            "alarcon.pdf",
        ],
    )
