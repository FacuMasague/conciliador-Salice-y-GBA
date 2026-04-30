"""utils.py — Funciones de normalización compartidas.

Definición única de las funciones de normalización usadas en pipeline.py y
matcher_hungarian.py. Cualquier cambio de lógica se hace aquí y se propaga
automáticamente a ambos módulos.
"""
from __future__ import annotations

import unicodedata
from typing import Optional


def _normalize_text(value: object) -> str:
    """Minúsculas, strip de acentos y normalización unicode."""
    s = unicodedata.normalize("NFKD", str(value or ""))
    s = "".join(ch for ch in s if not unicodedata.combining(ch))
    return s.strip().lower()


def _normalize_cliente(value: object) -> Optional[str]:
    """Forma numérica canónica de un ID de cliente (solo dígitos, sin ceros a la izquierda)."""
    if value is None:
        return None
    s = str(value).strip()
    if not s:
        return None
    digits = "".join(ch for ch in s if ch.isdigit())
    if not digits:
        return None
    try:
        return str(int(digits))
    except Exception:
        return digits.lstrip("0") or "0"


def _normalize_cuit(value: object) -> Optional[str]:
    """CUIT de 11 dígitos como string, o None si el valor no es válido."""
    if value is None:
        return None
    s = str(value).strip()
    if not s:
        return None
    digits = "".join(ch for ch in s if ch.isdigit())
    if len(digits) == 11:
        return digits
    return None


def _normalize_recibo(value: object) -> Optional[str]:
    """Forma canónica de un número de recibo.

    Maneja:
    - int/float directamente
    - Strings numéricos simples ("12345")
    - Formato contable local ("68.734,00" → "68734")
    - Cadenas mixtas (extrae dígitos como fallback)
    """
    if value is None:
        return None
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        try:
            return str(int(float(value)))
        except Exception:
            pass
    s = str(value).strip()
    if not s:
        return None
    s_num = s.replace("$", "").replace(" ", "")
    if s_num:
        parsed_float: float | None = None
        try:
            if "." in s_num and "," in s_num:
                # Formato contable local: 68.734,00 → 68734.00
                parsed_float = float(s_num.replace(".", "").replace(",", "."))
            elif "," in s_num:
                if s_num.count(",") == 1 and len(s_num.rsplit(",", 1)[-1]) <= 2:
                    parsed_float = float(s_num.replace(".", "").replace(",", "."))
                elif all(part.isdigit() for part in s_num.split(",")):
                    parsed_float = float(s_num.replace(",", ""))
            elif "." in s_num:
                if s_num.count(".") == 1 and len(s_num.rsplit(".", 1)[-1]) <= 2:
                    parsed_float = float(s_num)
                elif all(part.isdigit() for part in s_num.split(".")):
                    parsed_float = float(s_num.replace(".", ""))
            elif s_num.isdigit():
                parsed_float = float(s_num)
        except Exception:
            parsed_float = None
        if parsed_float is not None:
            try:
                return str(int(parsed_float))
            except Exception:
                pass
    digits = "".join(ch for ch in s if ch.isdigit())
    if digits:
        try:
            return str(int(digits))
        except Exception:
            return digits.lstrip("0") or "0"
    return s.upper()
