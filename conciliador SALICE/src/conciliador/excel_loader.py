from __future__ import annotations

import datetime as dt
import re
import unicodedata
from dataclasses import dataclass
from typing import List, Optional

import pandas as pd


@dataclass(frozen=True)
class BankTxn:
    txn_id: str
    origen: str  # BBVA | GALICIA | MERCADOPAGO
    fecha: dt.date
    hora: Optional[dt.time]
    importe: float
    texto_ref: str
    row_index: int
    parse_ok: bool
    parse_error: Optional[str]
    # Marca si el ingreso ya estaba conciliado en el Excel (columna ok == "ok")
    was_preconciled: bool = False
    preconciled_recibo: Optional[str] = None
    cuit: Optional[str] = None


def _parse_date(value) -> Optional[dt.date]:
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return None
    if isinstance(value, dt.datetime):
        return value.date()
    if isinstance(value, dt.date):
        return value
    s = str(value).strip()
    if not s:
        return None
    for fmt in ("%d-%m-%Y", "%d/%m/%Y", "%Y-%m-%d"):
        try:
            return dt.datetime.strptime(s[:10], fmt).date()
        except Exception:
            pass
    return None


def _parse_datetime(value) -> tuple[Optional[dt.date], Optional[dt.time]]:
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return None, None
    if isinstance(value, dt.datetime):
        return value.date(), value.time().replace(microsecond=0)
    s = str(value).strip()
    if not s:
        return None, None
    for fmt in ("%d/%m/%Y %H:%M:%S", "%d-%m-%Y %H:%M:%S", "%Y-%m-%d %H:%M:%S"):
        try:
            d = dt.datetime.strptime(s, fmt)
            return d.date(), d.time()
        except Exception:
            pass
    d = _parse_date(s)
    return d, None


def _parse_amount(value) -> Optional[float]:
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return None
    if isinstance(value, (int, float)):
        return float(value)
    s = str(value).strip()
    if not s:
        return None
    s = s.replace("$", "").replace(" ", "")
    if re.match(r"^-?\d{1,3}(?:\.\d{3})*,\d{1,2}$", s):
        s = s.replace(".", "").replace(",", ".")
    else:
        s = s.replace(",", "")
    try:
        return float(s)
    except Exception:
        return None


def _norm_col(name: object) -> str:
    if name is None:
        return ""
    return str(name).strip().lower()


def _find_col(df: pd.DataFrame, candidates: list[str]) -> Optional[str]:
    wanted = {c.strip().lower() for c in candidates}
    for c in df.columns:
        if _norm_col(c) in wanted:
            return c
    return None


def _is_ok_marker(value: object) -> bool:
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return False
    s = str(value).strip().lower()
    if not s:
        return False
    # "ok", "OK", "ok " o variaciones con texto adicional.
    return re.search(r"\bok\b", s) is not None


def _as_clean_text(value: object) -> Optional[str]:
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return None
    s = str(value).strip()
    return s if s else None


def _norm_text(value: object) -> str:
    s = str(value or "").strip().lower()
    return "".join(c for c in unicodedata.normalize("NFKD", s) if not unicodedata.combining(c))


def _is_excluded_bbva_concept(texto_ref: object) -> bool:
    t = _norm_text(texto_ref)
    # V4.5.4: excluir impuestos automáticos de BBVA (no son ingresos conciliables).
    return "impuesto ley" in t


def _extract_cuit11(value: object) -> Optional[str]:
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return None
    s = str(value).strip()
    if not s:
        return None
    digits = "".join(ch for ch in s if ch.isdigit())
    if len(digits) == 11:
        return digits
    if len(digits) > 11:
        for i in range(0, len(digits) - 10):
            cand = digits[i : i + 11]
            if len(cand) == 11:
                return cand
    return None


def load_bank_txns(excel_path: str) -> List[BankTxn]:
    """Load and normalize bank txns from workbook.

    Además de fecha/importe, detecta si la fila ya estaba conciliada
    (columna ok == "ok") y guarda el recibo previo para penalizar en matcher.
    """
    txns: List[BankTxn] = []

    # Galicia
    for sheet in ["SALICE GALICIA (ALARCON)"]:
        try:
            df = pd.read_excel(excel_path, sheet_name=sheet)
        except Exception:
            continue
        ok_col = _find_col(df, ["ok", "recibio", "recibio?"])
        rec_col = _find_col(df, ["recibo", "nro recibo", "nro_recibo", "nro. recibo"])
        cuit_col = _find_col(df, ["cuit", "leyendas adicionales 2", "numero de documento", "número documento"])
        for idx, row in df.iterrows():
            d = _parse_date(row.get("Fecha"))
            amt = _parse_amount(row.get("Importe"))
            texto = " ".join(str(row.get(c, "") or "") for c in ["Concepto", "Razon social", "CUIT"]).strip()
            cuit = _extract_cuit11(row.get(cuit_col)) if cuit_col else None
            if not cuit:
                cuit = _extract_cuit11(texto)
            was_preconciled = _is_ok_marker(row.get(ok_col)) if ok_col else False
            preconciled_recibo = _as_clean_text(row.get(rec_col)) if rec_col else None
            if d and amt is not None:
                txns.append(
                    BankTxn(
                        f"GALICIA:{int(idx)}",
                        "GALICIA",
                        d,
                        None,
                        amt,
                        texto,
                        int(idx) + 2,
                        True,
                        None,
                        was_preconciled,
                        preconciled_recibo,
                        cuit,
                    )
                )
            else:
                txns.append(
                    BankTxn(
                        f"GALICIA:{int(idx)}",
                        "GALICIA",
                        d or dt.date(1900, 1, 1),
                        None,
                        amt or 0.0,
                        texto,
                        int(idx) + 2,
                        False,
                        "fecha inválida" if not d else "importe inválido",
                        was_preconciled,
                        preconciled_recibo,
                        cuit,
                    )
                )

    # MercadoPago
    for sheet in ["MercadoPago "]:
        try:
            df = pd.read_excel(excel_path, sheet_name=sheet)
        except Exception:
            continue
        ok_col = _find_col(df, ["ok", "recibio", "recibio?"])
        rec_col = _find_col(df, ["recibo", "nro recibo", "nro_recibo", "nro. recibo"])
        cuit_col = _find_col(
            df,
            [
                "NÚMERO DE IDENTIFICACIÓN DEL PAGADOR",
                "NUMERO DE IDENTIFICACION DEL PAGADOR",
                "Número de identificación del pagador",
            ],
        )
        for idx, row in df.iterrows():
            d, t = _parse_datetime(row.get("Fecha de Pago"))
            amt = _parse_amount(row.get("Unnamed: 4"))
            texto = str(row.get("Operación Relacionada", row.get("OperaciÃ³n Relacionada", "")) or "").strip()
            cuit = _extract_cuit11(row.get(cuit_col)) if cuit_col else None
            was_preconciled = _is_ok_marker(row.get(ok_col)) if ok_col else False
            preconciled_recibo = _as_clean_text(row.get(rec_col)) if rec_col else None
            if d and amt is not None:
                txns.append(
                    BankTxn(
                        f"MERCADOPAGO:{int(idx)}",
                        "MERCADOPAGO",
                        d,
                        t,
                        amt,
                        texto,
                        int(idx) + 2,
                        True,
                        None,
                        was_preconciled,
                        preconciled_recibo,
                        cuit,
                    )
                )
            else:
                txns.append(
                    BankTxn(
                        f"MERCADOPAGO:{int(idx)}",
                        "MERCADOPAGO",
                        d or dt.date(1900, 1, 1),
                        t,
                        amt or 0.0,
                        texto,
                        int(idx) + 2,
                        False,
                        "fecha inválida" if not d else "importe inválido",
                        was_preconciled,
                        preconciled_recibo,
                        cuit,
                    )
                )

    # BBVA
    for sheet, origen in [("SALICE BBVA", "BBVA"), (" ALARCON BBVA", "BBVA")]:
        try:
            raw = pd.read_excel(excel_path, sheet_name=sheet, header=None, dtype=str)
        except Exception:
            continue
        header_row = None
        for r in range(min(25, len(raw))):
            row_vals = [str(x) for x in raw.iloc[r].tolist() if x is not None and str(x) != "nan"]
            joined = " ".join(row_vals).lower()
            if ("número documento" in joined or "numero documento" in joined) and "importe" in joined:
                header_row = r
                break
        if header_row is None:
            continue
        df = pd.read_excel(excel_path, sheet_name=sheet, header=header_row, dtype=str)
        date_col = df.columns[0]
        importe_col = None
        for c in df.columns:
            if str(c).strip().lower() == "importe":
                importe_col = c
        if importe_col is None and len(df.columns) >= 5:
            importe_col = df.columns[4]
        ref_col = df.columns[1] if len(df.columns) > 1 else df.columns[0]
        ok_col = _find_col(df, ["ok", "ok?", "recibio", "recibio?"])
        rec_col = _find_col(df, ["recibo", "nro recibo", "nro_recibo", "nro. recibo"])
        for idx, row in df.iterrows():
            d = _parse_date(row.get(date_col))
            amt = _parse_amount(row.get(importe_col))
            texto = str(row.get(ref_col, "") or "").strip()
            if _is_excluded_bbva_concept(texto):
                continue
            cuit = _extract_cuit11(texto)
            excel_row = int(idx) + int(header_row) + 2
            was_preconciled = _is_ok_marker(row.get(ok_col)) if ok_col else False
            preconciled_recibo = _as_clean_text(row.get(rec_col)) if rec_col else None
            if d and amt is not None:
                txns.append(
                    BankTxn(
                        f"{origen}:{int(idx)}",
                        origen,
                        d,
                        None,
                        amt,
                        texto,
                        excel_row,
                        True,
                        None,
                        was_preconciled,
                        preconciled_recibo,
                        cuit,
                    )
                )
            else:
                txns.append(
                    BankTxn(
                        f"{origen}:{int(idx)}",
                        origen,
                        d or dt.date(1900, 1, 1),
                        None,
                        amt or 0.0,
                        texto,
                        excel_row,
                        False,
                        "fecha inválida" if not d else "importe inválido",
                        was_preconciled,
                        preconciled_recibo,
                        cuit,
                    )
                )

    return txns
