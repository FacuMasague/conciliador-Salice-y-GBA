from __future__ import annotations

from dataclasses import dataclass
from typing import Any


@dataclass(frozen=True)
class ExternalReceipt:
    empresa: str
    nro_recibo: str
    nro_cliente: str
    cliente_nombre: str
    vendedor: str


@dataclass(frozen=True)
class ExternalPayment:
    empresa: str
    nro_recibo: str
    nro_cliente: str
    cliente_nombre: str
    vendedor: str
    medio_pago: str
    fecha_pago: str
    importe_pago: float
    detalle_pago: str = ""
    api_key: dict[str, Any] | None = None


@dataclass(frozen=True)
class ExternalPadronEntry:
    nro_cliente: str
    cuit: str
