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
    was_preconciled: bool = False
    preconciled_recibo: Optional[str] = None
    preconciled_nro_cliente: Optional[str] = None
    preconciled_cliente_nombre: Optional[str] = None
    preconciled_fecha_recibo: Optional[str] = None
    preconciled_medio_pago: Optional[str] = None
    preconciled_importe_recibo: Optional[float] = None
    cuit: Optional[str] = None
    sheet_name: Optional[str] = None
    record_key: Optional[str] = None


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
    for fmt in ("%d-%m-%Y", "%d/%m/%Y", "%Y-%m-%d", "%Y/%m/%d"):
        try:
            return dt.datetime.strptime(s[:10], fmt).date()
        except Exception:
            pass
    try:
        d = pd.to_datetime(s, errors="coerce", dayfirst=True)
        if pd.notna(d):
            return d.date()
    except Exception:
        pass
    return None


def _parse_datetime(value) -> tuple[Optional[dt.date], Optional[dt.time]]:
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return None, None
    if isinstance(value, dt.datetime):
        return value.date(), value.time().replace(microsecond=0)
    if isinstance(value, dt.date):
        return value, None
    s = str(value).strip()
    if not s:
        return None, None
    try:
        iso = dt.datetime.fromisoformat(s.replace("Z", "+00:00"))
        return iso.date(), iso.timetz().replace(microsecond=0, tzinfo=None)
    except Exception:
        pass
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
    return re.search(r"\bok\b", s) is not None or s in {"si", "s", "yes", "true", "1"}


def _as_clean_text(value: object) -> Optional[str]:
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return None
    s = str(value).strip()
    return s if s else None


def _coerce_export_date_str(value: object) -> Optional[str]:
    d = _parse_date(value)
    if d is not None:
        return d.isoformat()
    s = _as_clean_text(value)
    return s if s else None


def _norm_text(value: object) -> str:
    s = str(value or "").strip().lower()
    return "".join(c for c in unicodedata.normalize("NFKD", s) if not unicodedata.combining(c))


def _is_excluded_bbva_concept(texto_ref: object) -> bool:
    t = _norm_text(texto_ref)
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


def _plain_id(value: object) -> str:
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return ""
    if isinstance(value, bool):
        return str(value)
    if isinstance(value, int):
        return str(value)
    if isinstance(value, float):
        if value.is_integer():
            return str(int(value))
        return format(value, "f").rstrip("0").rstrip(".")
    s = str(value).strip()
    if not s:
        return ""
    s_compact = s.replace(" ", "")
    if "e" in s_compact.lower():
        try:
            f = float(s_compact.replace(",", "."))
            if abs(f - round(f)) < 1e-6:
                return str(int(round(f)))
            return format(f, "f").rstrip("0").rstrip(".")
        except Exception:
            return s
    if s_compact.endswith(".0") and s_compact[:-2].isdigit():
        return s_compact[:-2]
    return s


def _join_parts(*values: object) -> str:
    out: list[str] = []
    seen: set[str] = set()
    for value in values:
        s = _as_clean_text(value)
        if not s:
            continue
        if s in seen:
            continue
        seen.add(s)
        out.append(s)
    return " | ".join(out)


def _txn(
    *,
    txn_id: str,
    origen: str,
    sheet_name: str,
    record_key: str | None,
    row_index: int,
    fecha: Optional[dt.date],
    hora: Optional[dt.time],
    importe: Optional[float],
    texto_ref: str,
    was_preconciled: bool,
    preconciled_recibo: Optional[str],
    preconciled_nro_cliente: Optional[str],
    preconciled_cliente_nombre: Optional[str],
    preconciled_fecha_recibo: Optional[str],
    preconciled_medio_pago: Optional[str],
    preconciled_importe_recibo: Optional[float],
    cuit: Optional[str],
    parse_error: Optional[str],
) -> BankTxn:
    return BankTxn(
        txn_id=txn_id,
        origen=origen,
        fecha=fecha or dt.date(1900, 1, 1),
        hora=hora,
        importe=float(importe or 0.0),
        texto_ref=texto_ref,
        row_index=row_index,
        parse_ok=(parse_error is None),
        parse_error=parse_error,
        was_preconciled=was_preconciled,
        preconciled_recibo=preconciled_recibo,
        preconciled_nro_cliente=preconciled_nro_cliente,
        preconciled_cliente_nombre=preconciled_cliente_nombre,
        preconciled_fecha_recibo=preconciled_fecha_recibo,
        preconciled_medio_pago=preconciled_medio_pago,
        preconciled_importe_recibo=preconciled_importe_recibo,
        cuit=cuit,
        sheet_name=sheet_name,
        record_key=record_key,
    )


def _load_galicia_df(df: pd.DataFrame, *, sheet_name: str, record_key: str | None) -> List[BankTxn]:
    txns: List[BankTxn] = []
    fecha_col = _find_col(df, ["Fecha"])
    detalle_col = _find_col(df, ["Concepto", "Descripcion", "Descripción"])
    importe_col = _find_col(df, ["Importe", "Creditos", "Créditos"])
    razon_col = _find_col(df, ["Razon social", "Razón social", "Leyendas Adicionales 1"])
    cuit_col = _find_col(df, ["CUIT", "Leyendas Adicionales 2"])
    ok_col = _find_col(df, ["ok", "recibio", "recibio?", "acreditado?"])
    rec_col = _find_col(df, ["recibo", "nro recibo", "nro_recibo", "nro. recibo"])
    cli_col = _find_col(df, ["cliente", "nro cliente", "nro_cliente", "nro. cliente"])
    cli_nombre_col = _find_col(df, ["cliente nombre", "nombre cliente", "cliente_nombre"])
    fecha_rec_col = _find_col(df, ["fecha recibo", "fecha_recibo"])
    medio_rec_col = _find_col(df, ["medio de pago", "medio_pago"])
    imp_rec_col = _find_col(df, ["importe recibo", "importe_recibo"])
    if not fecha_col or not detalle_col or not importe_col:
        return txns

    for idx, row in df.iterrows():
        amt = _parse_amount(row.get(importe_col))
        if amt is None or amt <= 0:
            continue
        d = _parse_date(row.get(fecha_col))
        texto = _join_parts(row.get(detalle_col), row.get(razon_col), row.get(cuit_col))
        cuit = _extract_cuit11(row.get(cuit_col)) or _extract_cuit11(texto)
        was_preconciled = _is_ok_marker(row.get(ok_col)) if ok_col else False
        preconciled_recibo = _as_clean_text(row.get(rec_col)) if rec_col else None
        preconciled_nro_cliente = _as_clean_text(row.get(cli_col)) if cli_col else None
        preconciled_cliente_nombre = _as_clean_text(row.get(cli_nombre_col)) if cli_nombre_col else None
        preconciled_fecha_recibo = _coerce_export_date_str(row.get(fecha_rec_col)) if fecha_rec_col else None
        preconciled_medio_pago = _as_clean_text(row.get(medio_rec_col)) if medio_rec_col else None
        preconciled_importe_recibo = _parse_amount(row.get(imp_rec_col)) if imp_rec_col else None
        txns.append(
            _txn(
                txn_id=f"GALICIA:{sheet_name}:{int(idx)}",
                origen="GALICIA",
                sheet_name=sheet_name,
                record_key=record_key,
                row_index=int(idx) + 2,
                fecha=d,
                hora=None,
                importe=amt,
                texto_ref=texto,
                was_preconciled=was_preconciled,
                preconciled_recibo=preconciled_recibo,
                preconciled_nro_cliente=preconciled_nro_cliente,
                preconciled_cliente_nombre=preconciled_cliente_nombre,
                preconciled_fecha_recibo=preconciled_fecha_recibo,
                preconciled_medio_pago=preconciled_medio_pago,
                preconciled_importe_recibo=preconciled_importe_recibo,
                cuit=cuit,
                parse_error=None if d else "fecha invalida",
            )
        )
    return txns


def _load_mp_df(df: pd.DataFrame, *, sheet_name: str, record_key: str | None) -> List[BankTxn]:
    txns: List[BankTxn] = []
    fecha_col = _find_col(df, ["Fecha de Pago"])
    if not fecha_col:
        return txns
    tipo_col = _find_col(
        df,
        [
            "Tipo de Operacion",
            "Tipo de Operación",
            "ID DE OPERACION EN MERCADO PAGO",
            "ID DE OPERACIÓN EN MERCADO PAGO",
        ],
    )
    related_col = _find_col(df, ["Operacion Relacionada", "Operación Relacionada"])
    importe_col = _find_col(df, ["Importe", "Unnamed: 4", "VALOR DE LA COMPRA"])
    cuit_col = _find_col(
        df,
        [
            "NÚMERO DE IDENTIFICACIÓN DEL PAGADOR",
            "NUMERO DE IDENTIFICACION DEL PAGADOR",
            "Numero de identificacion del pagador",
        ],
    )
    ok_col = _find_col(df, ["ok", "recibio", "recibio?", "control logistica", "acreditado?"])
    rec_col = _find_col(df, ["recibo", "nro recibo", "nro_recibo", "nro. recibo"])
    cli_col = _find_col(df, ["cliente", "nro cliente", "nro_cliente", "nro. cliente"])
    cli_nombre_col = _find_col(df, ["cliente nombre", "nombre cliente", "cliente_nombre"])
    fecha_rec_col = _find_col(df, ["fecha recibo", "fecha_recibo"])
    medio_rec_col = _find_col(df, ["medio de pago", "medio_pago"])
    imp_rec_col = _find_col(df, ["importe recibo", "importe_recibo"])
    if not importe_col:
        return txns

    for idx, row in df.iterrows():
        amt = _parse_amount(row.get(importe_col))
        if amt is None or amt <= 0:
            continue
        d, t = _parse_datetime(row.get(fecha_col))
        op_id = _plain_id(row.get(tipo_col)) if tipo_col else ""
        related = _as_clean_text(row.get(related_col)) or ""
        texto = _join_parts(op_id, related)
        cuit = _extract_cuit11(row.get(cuit_col)) if cuit_col else None
        was_preconciled = _is_ok_marker(row.get(ok_col)) if ok_col else False
        preconciled_recibo = _as_clean_text(row.get(rec_col)) if rec_col else None
        preconciled_nro_cliente = _as_clean_text(row.get(cli_col)) if cli_col else None
        preconciled_cliente_nombre = _as_clean_text(row.get(cli_nombre_col)) if cli_nombre_col else None
        preconciled_fecha_recibo = _coerce_export_date_str(row.get(fecha_rec_col)) if fecha_rec_col else None
        preconciled_medio_pago = _as_clean_text(row.get(medio_rec_col)) if medio_rec_col else None
        preconciled_importe_recibo = _parse_amount(row.get(imp_rec_col)) if imp_rec_col else None
        txns.append(
            _txn(
                txn_id=f"MERCADOPAGO:{sheet_name}:{int(idx)}",
                origen="MERCADOPAGO",
                sheet_name=sheet_name,
                record_key=record_key,
                row_index=int(idx) + 2,
                fecha=d,
                hora=t,
                importe=amt,
                texto_ref=texto or related or op_id,
                was_preconciled=was_preconciled,
                preconciled_recibo=preconciled_recibo,
                preconciled_nro_cliente=preconciled_nro_cliente,
                preconciled_cliente_nombre=preconciled_cliente_nombre,
                preconciled_fecha_recibo=preconciled_fecha_recibo,
                preconciled_medio_pago=preconciled_medio_pago,
                preconciled_importe_recibo=preconciled_importe_recibo,
                cuit=cuit,
                parse_error=None if d else "fecha invalida",
            )
        )
    return txns


def _load_gba_bbva_df(df: pd.DataFrame, *, sheet_name: str, record_key: str | None) -> List[BankTxn]:
    txns: List[BankTxn] = []
    fecha_col = _find_col(df, ["Fecha"])
    concepto_col = _find_col(df, ["Concepto"])
    credito_col = _find_col(df, ["Credito", "Crédito"])
    numero_doc_col = _find_col(df, ["Numero Documento", "Número Documento"])
    detalle_col = _find_col(df, ["Detalle"])
    ok_col = _find_col(df, ["ok", "recibio", "recibio?", "acreditado?"])
    rec_col = _find_col(df, ["recibo", "nro recibo", "nro_recibo", "nro. recibo"])
    cli_col = _find_col(df, ["cliente", "nro cliente", "nro_cliente", "nro. cliente"])
    cli_nombre_col = _find_col(df, ["cliente nombre", "nombre cliente", "cliente_nombre"])
    fecha_rec_col = _find_col(df, ["fecha recibo", "fecha_recibo"])
    medio_rec_col = _find_col(df, ["medio de pago", "medio_pago"])
    imp_rec_col = _find_col(df, ["importe recibo", "importe_recibo"])
    if not fecha_col or not concepto_col or not credito_col:
        return txns

    for idx, row in df.iterrows():
        amt = _parse_amount(row.get(credito_col))
        if amt is None or amt <= 0:
            continue
        d = _parse_date(row.get(fecha_col))
        concepto = _as_clean_text(row.get(concepto_col)) or ""
        detalle = _as_clean_text(row.get(detalle_col)) or ""
        numero_doc = _as_clean_text(row.get(numero_doc_col)) or ""
        texto = _join_parts(concepto, numero_doc, detalle)
        if _is_excluded_bbva_concept(texto):
            continue
        cuit = _extract_cuit11(row.get(numero_doc_col)) or _extract_cuit11(detalle) or _extract_cuit11(texto)
        was_preconciled = _is_ok_marker(row.get(ok_col)) if ok_col else False
        preconciled_recibo = _as_clean_text(row.get(rec_col)) if rec_col else None
        preconciled_nro_cliente = _as_clean_text(row.get(cli_col)) if cli_col else None
        preconciled_cliente_nombre = _as_clean_text(row.get(cli_nombre_col)) if cli_nombre_col else None
        preconciled_fecha_recibo = _coerce_export_date_str(row.get(fecha_rec_col)) if fecha_rec_col else None
        preconciled_medio_pago = _as_clean_text(row.get(medio_rec_col)) if medio_rec_col else None
        preconciled_importe_recibo = _parse_amount(row.get(imp_rec_col)) if imp_rec_col else None
        txns.append(
            _txn(
                txn_id=f"BBVA:{sheet_name}:{int(idx)}",
                origen="BBVA",
                sheet_name=sheet_name,
                record_key=record_key,
                row_index=int(idx) + 2,
                fecha=d,
                hora=None,
                importe=amt,
                texto_ref=texto,
                was_preconciled=was_preconciled,
                preconciled_recibo=preconciled_recibo,
                preconciled_nro_cliente=preconciled_nro_cliente,
                preconciled_cliente_nombre=preconciled_cliente_nombre,
                preconciled_fecha_recibo=preconciled_fecha_recibo,
                preconciled_medio_pago=preconciled_medio_pago,
                preconciled_importe_recibo=preconciled_importe_recibo,
                cuit=cuit,
                parse_error=None if d else "fecha invalida",
            )
        )
    return txns


def _load_legacy_bbva_sheet(xls: pd.ExcelFile, *, sheet_name: str, record_key: str | None) -> List[BankTxn]:
    txns: List[BankTxn] = []
    try:
        raw = xls.parse(sheet_name, header=None, dtype=str)
    except Exception:
        return txns
    header_row = None
    for r in range(min(25, len(raw))):
        row_vals = [str(x) for x in raw.iloc[r].tolist() if x is not None and str(x) != "nan"]
        joined = " ".join(row_vals).lower()
        if ("numero documento" in joined or "número documento" in joined) and "importe" in joined:
            header_row = r
            break
    if header_row is None:
        return txns
    # Reutilizamos el DataFrame ya cargado en memoria: evitamos releer el archivo.
    try:
        df = raw.iloc[header_row + 1:].copy()
        df.columns = raw.iloc[header_row].tolist()
        df = df.reset_index(drop=True)
    except Exception:
        return txns

    date_col = df.columns[0]
    importe_col = _find_col(df, ["Importe", "Credito", "Crédito"])
    if importe_col is None and len(df.columns) >= 5:
        importe_col = df.columns[4]
    ref_col = df.columns[1] if len(df.columns) > 1 else df.columns[0]
    detalle_col = _find_col(df, ["Detalle"])
    ok_col = _find_col(df, ["ok", "ok?", "recibio", "recibio?"])
    rec_col = _find_col(df, ["recibo", "nro recibo", "nro_recibo", "nro. recibo"])
    cli_col = _find_col(df, ["cliente", "nro cliente", "nro_cliente", "nro. cliente"])
    cli_nombre_col = _find_col(df, ["cliente nombre", "nombre cliente", "cliente_nombre"])
    fecha_rec_col = _find_col(df, ["fecha recibo", "fecha_recibo"])
    medio_rec_col = _find_col(df, ["medio de pago", "medio_pago"])
    imp_rec_col = _find_col(df, ["importe recibo", "importe_recibo"])

    for idx, row in df.iterrows():
        amt = _parse_amount(row.get(importe_col))
        if amt is None or amt <= 0:
            continue
        d = _parse_date(row.get(date_col))
        texto = _join_parts(row.get(ref_col), row.get(detalle_col))
        if _is_excluded_bbva_concept(texto):
            continue
        cuit = _extract_cuit11(texto)
        excel_row = int(idx) + int(header_row) + 2
        was_preconciled = _is_ok_marker(row.get(ok_col)) if ok_col else False
        preconciled_recibo = _as_clean_text(row.get(rec_col)) if rec_col else None
        preconciled_nro_cliente = _as_clean_text(row.get(cli_col)) if cli_col else None
        preconciled_cliente_nombre = _as_clean_text(row.get(cli_nombre_col)) if cli_nombre_col else None
        preconciled_fecha_recibo = _coerce_export_date_str(row.get(fecha_rec_col)) if fecha_rec_col else None
        preconciled_medio_pago = _as_clean_text(row.get(medio_rec_col)) if medio_rec_col else None
        preconciled_importe_recibo = _parse_amount(row.get(imp_rec_col)) if imp_rec_col else None
        txns.append(
            _txn(
                txn_id=f"BBVA:{sheet_name}:{int(idx)}",
                origen="BBVA",
                sheet_name=sheet_name,
                record_key=record_key,
                row_index=excel_row,
                fecha=d,
                hora=None,
                importe=amt,
                texto_ref=texto,
                was_preconciled=was_preconciled,
                preconciled_recibo=preconciled_recibo,
                preconciled_nro_cliente=preconciled_nro_cliente,
                preconciled_cliente_nombre=preconciled_cliente_nombre,
                preconciled_fecha_recibo=preconciled_fecha_recibo,
                preconciled_medio_pago=preconciled_medio_pago,
                preconciled_importe_recibo=preconciled_importe_recibo,
                cuit=cuit,
                parse_error=None if d else "fecha invalida",
            )
        )
    return txns


def load_bank_txns(excel_path: str, *, record_key: str | None = None) -> List[BankTxn]:
    txns: List[BankTxn] = []
    handled_sheets: set[str] = set()

    try:
        # Abrimos el archivo una sola vez; todos los loaders reutilizan este objeto
        # en lugar de reabrir el ZIP por cada hoja.
        xls = pd.ExcelFile(excel_path)
        sheet_names = list(xls.sheet_names)
    except Exception:
        sheet_names = []
        xls = None  # type: ignore[assignment]

    if xls is None:
        return txns

    for sheet_name in sheet_names:
        try:
            df = xls.parse(sheet_name)
        except Exception:
            continue
        headers = {_norm_text(c) for c in df.columns}
        parsed: List[BankTxn] = []
        if "fecha de pago" in headers:
            parsed = _load_mp_df(df, sheet_name=sheet_name, record_key=record_key)
        elif "fecha" in headers and "concepto" in headers and ("credito" in headers or "crédito" in headers):
            parsed = _load_gba_bbva_df(df, sheet_name=sheet_name, record_key=record_key)
        elif "fecha" in headers and (
            "importe" in headers or "creditos" in headers or "créditos" in headers
        ) and (
            "razon social" in headers
            or "razón social" in headers
            or "cuit" in headers
            or "leyendas adicionales 1" in headers
        ):
            parsed = _load_galicia_df(df, sheet_name=sheet_name, record_key=record_key)
        if parsed:
            handled_sheets.add(sheet_name)
            txns.extend(parsed)

    for sheet_name in sheet_names:
        if sheet_name in handled_sheets:
            continue
        txns.extend(_load_legacy_bbva_sheet(xls, sheet_name=sheet_name, record_key=record_key))

    return txns
