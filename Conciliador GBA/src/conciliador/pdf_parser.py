import re
from dataclasses import dataclass
from typing import Any, List, Optional, Tuple

import pdfplumber


@dataclass(frozen=True)
class Receipt:
    empresa: str
    nro_recibo: str  # TEXT
    nro_cliente: str
    cliente_nombre: Optional[str]
    vendedor: Optional[str] = None


@dataclass(frozen=True)
class ReceiptPayment:
    empresa: str
    nro_recibo: str
    nro_cliente: str
    cliente_nombre: Optional[str]
    medio_pago: str  # TRANSFERENCIA | MERCADOPAGO
    fecha_pago: str  # YYYY-MM-DD
    importe_pago: float
    detalle_pago: Optional[str] = None
    vendedor: Optional[str] = None
    api_key: dict[str, Any] | None = None


_RE_HEADER_WITH_DATE = re.compile(r"^(?P<recibo>\d+)\s+(?P<cliente>\d+)\s+-\s+(?P<nombre>.+?)\s+\d{2}/\d{2}/\d{4}\b")
_RE_HEADER = re.compile(r"^(?P<recibo>\d+)\s+(?P<cliente>\d+)\s+-\s+(?P<nombre>.*)$")
_RE_VENDOR = re.compile(r"^\[(?P<codigo>\d+)\s*-\s*(?P<nombre>[^\]]+)\](?:.*)?$")
_RE_PAY_LINE = re.compile(
    r"^(?P<prefix>Transferencia(?:\s+Bancar)?|Mercado\s+Pago).*?\|\s*(?P<fecha>\d{4}-\d{2}-\d{2})\s*\|\s*(?P<importe>-?\d+(?:\.\d+)?)\s*$",
    re.IGNORECASE,
)


def _clean_cliente_nombre(raw: str) -> str:
    """Limpia basura numérica que a veces queda pegada al nombre al extraer PDF.

    Ejemplo a limpiar:
      "Martinez Jorge Sebastian 0,00 0,62 681.460,43 681.461,05"
    """
    s = (raw or "").strip()
    if not s:
        return ""

    s = re.sub(r"\s+\d{2}/\d{2}/\d{4}.*$", "", s).strip()
    num_like = re.compile(r"^[+-]?\d+(?:[.,]\d+)*$")
    tokens = s.split()
    kept: list[str] = []
    for tok in tokens:
        # Si ya empezó una cola numérica (sin letras), cortamos.
        if num_like.match(tok):
            break
        kept.append(tok)

    cleaned = " ".join(kept).strip()
    return cleaned


def infer_empresa_from_text(text: str) -> str:
    t = text.lower()
    # Algunos reportes mantienen el título 'Cobranza SALICE' aunque la sucursal sea ALARCON.
    if "sucursal:" in t and "alarcon" in t:
        return "ALARCON"
    if "cobranza alarcon" in t or "reporte de cobranza alarcon" in t:
        return "ALARCON"
    if "cobranza salice" in t or "salice distribuciones" in t:
        return "SALICE"
    if "alarcon" in t:
        return "ALARCON"
    return "DESCONOCIDA"



def _medio_pago(prefix: str) -> str:
    p = prefix.lower()
    if "mercado" in p:
        return "MERCADOPAGO"
    if "transfer" in p:
        return "TRANSFERENCIA"
    return "OTRO"


def extract_pdf_text(pdf_path: str, max_pages: Optional[int] = None) -> str:
    parts: List[str] = []
    with pdfplumber.open(pdf_path) as pdf:
        pages = pdf.pages if max_pages is None else pdf.pages[:max_pages]
        for page in pages:
            parts.append(page.extract_text() or "")
    return "\n".join(parts)


def parse_receipts_and_payments_from_text(text: str, *, empresa_override: str | None = None) -> Tuple[List[Receipt], List[ReceiptPayment]]:
    empresa = (empresa_override or '').strip().upper() or infer_empresa_from_text(text)
    if empresa not in {"SALICE", "ALARCON"}:
        # Mantener el valor detectado (DESCONOCIDA) si no se puede inferir.
        empresa = infer_empresa_from_text(text)

    receipts: List[Receipt] = []
    payments: List[ReceiptPayment] = []
    current: Optional[Receipt] = None

    for raw_line in text.splitlines():
        line = raw_line.strip()
        if not line:
            continue

        m = _RE_HEADER_WITH_DATE.match(line) or _RE_HEADER.match(line)
        if m:
            nro_recibo = m.group("recibo")
            nro_cliente = m.group("cliente")
            nombre = (m.group("nombre") or "").strip()
            nombre_clean = _clean_cliente_nombre(nombre)
            current = Receipt(
                empresa=empresa,
                nro_recibo=str(nro_recibo),
                nro_cliente=str(nro_cliente),
                cliente_nombre=nombre_clean if nombre_clean else None,
                vendedor=None,
            )
            receipts.append(current)
            continue

        vm = _RE_VENDOR.match(line)
        if vm and current:
            vend = f"{vm.group('codigo').strip()} - {vm.group('nombre').strip()}"
            current = Receipt(
                empresa=current.empresa,
                nro_recibo=current.nro_recibo,
                nro_cliente=current.nro_cliente,
                cliente_nombre=current.cliente_nombre,
                vendedor=vend,
            )
            if receipts:
                receipts[-1] = current
            continue

        pm = _RE_PAY_LINE.match(line)
        if pm and current:
            medio = _medio_pago(pm.group("prefix"))
            if medio not in ("TRANSFERENCIA", "MERCADOPAGO"):
                continue
            payments.append(
                ReceiptPayment(
                    empresa=current.empresa,
                    nro_recibo=current.nro_recibo,
                    nro_cliente=current.nro_cliente,
                    cliente_nombre=current.cliente_nombre,
                    vendedor=current.vendedor,
                    medio_pago=medio,
                    fecha_pago=pm.group("fecha"),
                    importe_pago=float(pm.group("importe")),
                )
            )

    return receipts, payments


def parse_receipts_and_payments(pdf_path: str, *, empresa_override: str | None = None) -> Tuple[List[Receipt], List[ReceiptPayment]]:
    """Parse receipts and payment lines from the standardized report PDFs.

    Captures only TRANSFERENCIA and MERCADOPAGO payment lines.
    """
    text = extract_pdf_text(pdf_path)
    return parse_receipts_and_payments_from_text(text, empresa_override=empresa_override)


def pdf_date_range(payments: List[ReceiptPayment]) -> Tuple[Optional[str], Optional[str]]:
    if not payments:
        return None, None
    fechas = sorted(p.fecha_pago for p in payments)
    return fechas[0], fechas[-1]


def detect_pdf_warnings(pdf_path: str) -> List[str]:
    """Detect simple warning flags in the PDF text.

    The reports sometimes state that they include annulled payments. We don't
    parse annulled blocks perfectly yet, so we surface it as a warning for the UI.
    """
    text = extract_pdf_text(pdf_path, max_pages=2).lower()
    return detect_pdf_warnings_from_text(text)


def detect_pdf_warnings_from_text(text: str) -> List[str]:
    text = (text or "").lower()
    warnings: List[str] = []
    if "incluye pagos anulados" in text:
        warnings.append("El PDF indica que incluye pagos anulados.")
    if "monto pago anulado" in text:
        warnings.append("Se detectó texto de pagos anulados (MONTO PAGO ANULADO...).")
    return warnings


def report_period_range(pdf_path: str) -> Tuple[Optional[str], Optional[str]]:
    """Extrae el rango "Desde/Hasta" informado en el encabezado del reporte.

    Esto NO necesariamente coincide con el rango de fechas efectivas en las líneas de pago.
    Para reportes de "No encontrados" (BANCO_SIN_RECIBO) usamos este rango estricto,
    tal como lo ve el empleado en el PDF.
    """
    text = extract_pdf_text(pdf_path, max_pages=1)
    return report_period_range_from_text(text)


def report_period_range_from_text(text: str) -> Tuple[Optional[str], Optional[str]]:
    # Captura dd/mm/yyyy o dd-mm-yyyy.
    # En estos reportes el encabezado suele venir como "Desde Fecha: 01/02/2026 Hasta: 09/02/2026".
    # Por eso contemplamos el literal "Fecha" entre Desde/Hasta y los dos puntos.
    m_from = re.search(r"\bDesde\s*(?:Fecha)?\s*[:]?\s*(\d{2}[/-]\d{2}[/-]\d{4})\b", text or "", re.IGNORECASE)
    m_to = re.search(r"\bHasta\s*(?:Fecha)?\s*[:]?\s*(\d{2}[/-]\d{2}[/-]\d{4})\b", text or "", re.IGNORECASE)

    def _to_iso(s: str) -> Optional[str]:
        if not s:
            return None
        s = s.replace("-", "/")
        try:
            import datetime as dt

            d = dt.datetime.strptime(s, "%d/%m/%Y").date()
            return d.isoformat()
        except Exception:
            return None

    return (_to_iso(m_from.group(1)) if m_from else None, _to_iso(m_to.group(1)) if m_to else None)
