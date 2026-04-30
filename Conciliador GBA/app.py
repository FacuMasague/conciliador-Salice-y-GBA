from __future__ import annotations

import os
import json
import shutil
import tempfile
import uuid
import zipfile
from typing import Optional

from fastapi import FastAPI, File, Form, HTTPException, Query, UploadFile
from fastapi.middleware.cors import CORSMiddleware
from fastapi.staticfiles import StaticFiles
from fastapi.responses import FileResponse

from src.conciliador.env_loader import load_project_env
load_project_env()

from src.conciliador.pipeline import compare_excel_pdfs
from src.conciliador.exporter import (
    export_xlsx,
    export_filled_bank_excel,
    export_filled_generic_excel,
    export_zip_csv,
    export_no_encontrados_xlsx,
)
from src.conciliador.raw_bank_ingest import build_runtime_workbook_from_raw, detect_raw_bank_kind
from src.conciliador.external.errors import ExternalConfigError, ExternalProviderError, ExternalSchemaError, ExternalTimeoutError


# Versión visible en UI y en /docs
APP_VERSION = "1.2.15"
app = FastAPI(title="Conciliador de Recibos e Ingresos", version=APP_VERSION)

# Para desarrollo web (frontend local) sin fricción.
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=False,
    allow_methods=["*"],
    allow_headers=["*"],
)

# UI mínima (una sola página) para subir archivos y ver resultados.
WEB_DIR = os.path.join(os.path.dirname(__file__), "web")
if os.path.isdir(WEB_DIR):
    app.mount("/static", StaticFiles(directory=WEB_DIR), name="static")


@app.get("/")
def index():
    index_path = os.path.join(WEB_DIR, "index.html")
    if not os.path.exists(index_path):
        raise HTTPException(status_code=404, detail="UI no instalada")
    return FileResponse(index_path, headers={"Cache-Control":"no-store"})


def _suffix(filename: str, default: str) -> str:
    _, ext = os.path.splitext(filename or "")
    return ext if ext else default


async def _save_upload(upload: UploadFile, out_path: str) -> None:
    with open(out_path, "wb") as f:
        f.write(await upload.read())


def _parse_json_list(raw: str | None) -> list[dict] | None:
    if not raw:
        return None
    try:
        parsed = json.loads(raw)
        if isinstance(parsed, list):
            return parsed
    except Exception:
        return None
    return None


def _download_name_from_excel_filename(excel_filename: str | None, suffix: str) -> str:
    try:
        base = os.path.splitext(excel_filename or "ingresos.xlsx")[0]
        return f"{base}_{suffix}.xlsx"
    except Exception:
        return f"ingresos_{suffix}.xlsx"


def _iso_min(values: list[str | None]) -> str | None:
    clean = [str(v).strip() for v in values if str(v or "").strip()]
    return min(clean) if clean else None


def _iso_max(values: list[str | None]) -> str | None:
    clean = [str(v).strip() for v in values if str(v or "").strip()]
    return max(clean) if clean else None


class InputMode:
    """Constantes para el campo input_mode de los resultados de /compare y /export."""
    LEGACY = "legacy_excel"
    V4_RAW = "v4_raw_plus_record"
    V5_SPLIT = "v5_split_records"


def _empty_ingest_meta(record_filename: str) -> dict:
    """Retorna un meta de ingestión vacío (cuando no hay archivos crudos)."""
    return {
        "record_excel_filename": record_filename,
        "raw_bank_files": [],
        "raw_ingestion_summary": {},
        "raw_total_input_rows": 0,
        "raw_total_appended_rows": 0,
        "raw_min_date_all": None,
        "raw_min_date": None,
        "raw_max_date": None,
        "raw_recent_files_max_gap_days": None,
        "raw_stale_files_ignored_for_receipts_start_date": [],
    }


async def _build_single_record(
    *,
    tmp_dir: str,
    rid: str,
    record_upload: UploadFile,
    raw_paths: list[str],
    key: str,
    origins: list[str],
) -> tuple[dict, dict]:
    """Guarda el archivo subido, construye el workbook runtime y devuelve (record, ingest_meta).

    Si raw_paths está vacío, copia el record tal cual (sin ingestión de crudos).
    """
    record_ext = _suffix(record_upload.filename or "", default=".xlsx")
    record_path = os.path.join(tmp_dir, f"{rid}_record_{key}{record_ext}")
    await _save_upload(record_upload, record_path)
    runtime_path = os.path.join(tmp_dir, f"{rid}_runtime_{key}.xlsx")
    if raw_paths:
        ingest_meta = build_runtime_workbook_from_raw(
            record_excel_path=record_path,
            raw_bank_paths=raw_paths,
            out_excel_path=runtime_path,
        )
    else:
        shutil.copy2(record_path, runtime_path)
        ingest_meta = _empty_ingest_meta(os.path.basename(record_path))
    record = {
        "key": key,
        "working_excel_path": runtime_path,
        "base_excel_filename": record_upload.filename,
        "origins": origins,
        "export_mode": "generic",
    }
    return record, ingest_meta


def _merge_ingestion_meta(metas: list[dict]) -> dict:
    summary = {
        "BBVA": {"input": 0, "appended": 0, "duplicates_skipped": 0, "sheet": None},
        "GALICIA": {"input": 0, "appended": 0, "duplicates_skipped": 0, "sheet": None},
        "MERCADOPAGO": {"input": 0, "appended": 0, "duplicates_skipped": 0, "sheet": None},
    }
    sheets_by_bank: dict[str, list[str]] = {"BBVA": [], "GALICIA": [], "MERCADOPAGO": []}
    raw_bank_files: list[dict] = []
    stale_files: list[str] = []
    raw_min_all: list[str | None] = []
    raw_min_recent: list[str | None] = []
    raw_max_recent: list[str | None] = []
    recent_gap = None
    total_input = 0
    total_appended = 0

    for meta in metas:
        raw_bank_files.extend([x for x in (meta.get("raw_bank_files") or []) if isinstance(x, dict)])
        stale_files.extend([str(x) for x in (meta.get("raw_stale_files_ignored_for_receipts_start_date") or []) if str(x).strip()])
        raw_min_all.append(meta.get("raw_min_date_all"))
        raw_min_recent.append(meta.get("raw_min_date"))
        raw_max_recent.append(meta.get("raw_max_date"))
        if recent_gap is None and meta.get("raw_recent_files_max_gap_days") is not None:
            recent_gap = meta.get("raw_recent_files_max_gap_days")
        total_input += int(meta.get("raw_total_input_rows") or 0)
        total_appended += int(meta.get("raw_total_appended_rows") or 0)
        meta_summary = meta.get("raw_ingestion_summary") or {}
        if not isinstance(meta_summary, dict):
            continue
        for bank in summary.keys():
            data = meta_summary.get(bank) or {}
            if not isinstance(data, dict):
                continue
            summary[bank]["input"] += int(data.get("input") or 0)
            summary[bank]["appended"] += int(data.get("appended") or 0)
            summary[bank]["duplicates_skipped"] += int(data.get("duplicates_skipped") or 0)
            sheet = str(data.get("sheet") or "").strip()
            if sheet and sheet not in sheets_by_bank[bank]:
                sheets_by_bank[bank].append(sheet)

    for bank, sheets in sheets_by_bank.items():
        if sheets:
            summary[bank]["sheet"] = ", ".join(sheets)

    return {
        "raw_bank_files": raw_bank_files,
        "raw_ingestion_summary": summary,
        "raw_total_input_rows": total_input,
        "raw_total_appended_rows": total_appended,
        "raw_min_date_all": _iso_min(raw_min_all),
        "raw_min_date": _iso_min(raw_min_recent),
        "raw_max_date": _iso_max(raw_max_recent),
        "raw_recent_files_max_gap_days": recent_gap,
        "raw_stale_files_ignored_for_receipts_start_date": stale_files,
    }


async def _prepare_excel_for_run(
    *,
    tmp_dir: str,
    rid: str,
    excel: UploadFile | None,
    record_excel: UploadFile | None,
    record_excel_bank: UploadFile | None,
    record_excel_mp: UploadFile | None,
    raw_bank_files: list[UploadFile] | None,
) -> dict:
    """Prepara el Excel de trabajo.

    Modos:
      - v4: record_excel + raw_bank_files => construye workbook runtime (record + crudos).
      - legacy: excel => usa directamente el Excel subido.
    """
    raw_files = [f for f in (raw_bank_files or []) if f is not None]

    # GBA: dos records separados (bancos y Mercado Pago).
    if record_excel_bank is not None or record_excel_mp is not None:
        if not raw_files:
            raise HTTPException(status_code=400, detail="Tenés que subir al menos 1 archivo crudo bancario.")

        raw_paths: list[str] = []
        raw_names: list[str] = []
        bank_raw_paths: list[str] = []
        mp_raw_paths: list[str] = []
        for i, up in enumerate(raw_files):
            ext = _suffix(up.filename or "", default=".xlsx")
            p = os.path.join(tmp_dir, f"{rid}_raw_{i+1}{ext}")
            await _save_upload(up, p)
            raw_paths.append(p)
            raw_names.append(up.filename or os.path.basename(p))
            try:
                kind = detect_raw_bank_kind(p)
            except Exception as e:
                raise HTTPException(status_code=400, detail=str(e))
            if kind == "MERCADOPAGO":
                mp_raw_paths.append(p)
            else:
                bank_raw_paths.append(p)

        if bank_raw_paths and record_excel_bank is None:
            raise HTTPException(status_code=400, detail="Falta el record de movimientos bancarios para los extractos BBVA/Galicia.")
        if mp_raw_paths and record_excel_mp is None:
            raise HTTPException(status_code=400, detail="Falta el record de Mercado Pago para los extractos de MP.")

        records: list[dict] = []
        metas: list[dict] = []

        if record_excel_bank is not None:
            rec, meta = await _build_single_record(
                tmp_dir=tmp_dir,
                rid=rid,
                record_upload=record_excel_bank,
                raw_paths=bank_raw_paths,
                key="bank",
                origins=["BBVA", "GALICIA"],
            )
            records.append(rec)
            metas.append(meta)

        if record_excel_mp is not None:
            rec, meta = await _build_single_record(
                tmp_dir=tmp_dir,
                rid=rid,
                record_upload=record_excel_mp,
                raw_paths=mp_raw_paths,
                key="mp",
                origins=["MERCADOPAGO"],
            )
            records.append(rec)
            metas.append(meta)

        merged_meta = _merge_ingestion_meta(metas)
        return {
            "working_excel_paths": [str(r["working_excel_path"]) for r in records],
            "records": records,
            "excel_record_map": {str(r["working_excel_path"]): str(r["key"]) for r in records},
            "base_excel_filename": None,
            "input_mode": InputMode.V5_SPLIT,
            "raw_bank_filenames": raw_names,
            "raw_ingestion_meta": merged_meta,
        }

    # V4 (preferido): consolidado + crudos.
    if record_excel is not None or raw_files:
        if record_excel is None:
            raise HTTPException(status_code=400, detail="Falta el Excel de record consolidado.")
        if not raw_files:
            raise HTTPException(status_code=400, detail="Tenés que subir al menos 1 archivo crudo bancario.")

        record_ext = _suffix(record_excel.filename or "", default=".xlsx")
        record_path = os.path.join(tmp_dir, f"{rid}_record{record_ext}")
        await _save_upload(record_excel, record_path)

        raw_paths: list[str] = []
        raw_names: list[str] = []
        for i, up in enumerate(raw_files):
            ext = _suffix(up.filename or "", default=".xlsx")
            p = os.path.join(tmp_dir, f"{rid}_raw_{i+1}{ext}")
            await _save_upload(up, p)
            raw_paths.append(p)
            raw_names.append(up.filename or os.path.basename(p))

        runtime_path = os.path.join(tmp_dir, f"{rid}_runtime_record.xlsx")
        ingest_meta = build_runtime_workbook_from_raw(
            record_excel_path=record_path,
            raw_bank_paths=raw_paths,
            out_excel_path=runtime_path,
        )
        return {
            "working_excel_path": runtime_path,
            "working_excel_paths": [runtime_path],
            "records": [
                {
                    "key": "default",
                    "working_excel_path": runtime_path,
                    "base_excel_filename": record_excel.filename,
                    "origins": ["BBVA", "GALICIA", "MERCADOPAGO"],
                    "export_mode": "legacy",
                }
            ],
            "excel_record_map": {runtime_path: "default"},
            "base_excel_filename": record_excel.filename,
            "input_mode": InputMode.V4_RAW,
            "raw_bank_filenames": raw_names,
            "raw_ingestion_meta": ingest_meta,
        }

    # Legacy
    if excel is None:
        raise HTTPException(
            status_code=400,
            detail="Tenés que subir el Excel consolidado de record y al menos 1 archivo crudo bancario.",
        )

    excel_ext = _suffix(excel.filename or "", default=".xlsx")
    excel_path = os.path.join(tmp_dir, f"{rid}_ingresos{excel_ext}")
    await _save_upload(excel, excel_path)
    return {
        "working_excel_path": excel_path,
        "working_excel_paths": [excel_path],
        "records": [
            {
                "key": "default",
                "working_excel_path": excel_path,
                "base_excel_filename": excel.filename,
                "origins": ["BBVA", "GALICIA", "MERCADOPAGO"],
                "export_mode": "legacy",
            }
        ],
        "excel_record_map": {excel_path: "default"},
        "base_excel_filename": excel.filename,
        "input_mode": InputMode.LEGACY,
        "raw_bank_filenames": [],
        "raw_ingestion_meta": None,
    }


@app.get("/health")
def health() -> dict:
    return {"ok": True}


@app.get("/version")
def version() -> dict:
    return {"app": "Conciliador", "version": APP_VERSION}


@app.post("/compare")
async def compare(
    excel: UploadFile | None = File(None, description="(LEGACY) Excel ya consolidado/procesado."),
    record_excel: UploadFile | None = File(None, description="Excel consolidado (record) a actualizar."),
    record_excel_bank: UploadFile | None = File(None, description="Record de movimientos bancarios (BBVA/Galicia)."),
    record_excel_mp: UploadFile | None = File(None, description="Record de Mercado Pago."),
    raw_bank_files: list[UploadFile] | None = File(None, description="Archivos crudos de bancos (BBVA/Galicia/MercadoPago)."),
    pdf_salice: UploadFile | None = File(None, description="PDF de recibos SALICE (opcional)"),
    pdf_alarcon: UploadFile | None = File(None, description="PDF de recibos Alarcón (opcional)"),
    force_validations: str | None = Form(None, description="Lista JSON para promover Dudosos a Validados (opcional)"),
    drop_dudosos: str | None = Form(None, description="Lista JSON para quitar casos de Dudosos (opcional)"),
    margin_days: int = Query(5, ge=0, le=31),
    tolerance_days_suspect: int = Query(7, ge=0, le=31),
    # V3.5: multiplicador de días separado por signo.
    day_weight_bank_before: float = Query(20.0, ge=0.0, le=10000.0, description="Multiplicador de días si el ingreso bancario es anterior (o mismo día) que el recibo"),
    day_weight_bank_after: float = Query(35.0, ge=0.0, le=10000.0, description="Multiplicador de días si el ingreso bancario es posterior al recibo (delay del banco)"),
    # Compatibilidad (deprecado): si se envía, pisa ambos day_weight_*.
    day_weight: float | None = Query(None, ge=0.0, le=10000.0, description="(DEPRECADO) Multiplicador de días único"),
    valid_max_peso: float = Query(260.0, ge=0.0, le=100000.0),
    dudoso_max_peso: float = Query(3500.0, ge=0.0, le=100000.0),
    mp_mismatch_penalty: float = Query(35.0, ge=0.0, le=100000.0, description="Penalización cuando medio MP↔origen banco está cruzado"),
    preconciled_penalty: float = Query(150.0, ge=0.0, le=100000.0, description="Penalización si el ingreso ya estaba conciliado (columna ok = ok)"),
    penalty_salice_to_galicia: float = Query(45.0, ge=0.0, le=100000.0, description="Penalización si recibo SALICE matchea contra banco GALICIA"),
    penalty_alarcon_to_bbva: float = Query(45.0, ge=0.0, le=100000.0, description="Penalización si recibo ALARCON matchea contra banco BBVA"),
    cliente_cuit_mismatch_penalty: float = Query(0.0, ge=0.0, le=100000.0, description="Sin efecto en GBA: el CUIT se usa para restringir candidatos plausibles, no para sumar peso"),
    alternatives_cost_delta: float = Query(35.0, ge=0.0, le=100000.0, description="Alternativos si peso <= principal + delta"),
    max_options: int = Query(3, ge=1, le=4, description="Máx. opciones por caso dudoso (1 principal + alternativos)"),
    stage2_candidate_top_k: int = Query(120, ge=0, le=2000, description="Prefiltro de candidatos por recibo para Hungarian (0 = sin límite)"),
    mem_debug: bool = Query(False, description="Activa métricas de memoria por etapa en meta"),
    show_peso: bool = Query(False, description="Mostrar la columna Peso en los resultados (por defecto oculto)"),
    show_cuit: bool = Query(False, description="Mostrar columnas de CUIT en los resultados (por defecto oculto)"),
    receipts_source: str = Query("pdf", pattern="^(pdf|api)$", description="Fuente de recibos: pdf (legacy) o api (v5)"),
    api_receipts_days: int = Query(4, ge=1, le=15, description="Fallback de días si no se especifica rango manual de recibos por API"),
    api_start_date: str | None = Query(None, description="Fecha desde para recibos API (YYYY-MM-DD)"),
    api_end_date: str | None = Query(None, description="Fecha hasta para recibos API (YYYY-MM-DD)"),
    api_empresa_filter: str | None = Query(None, description="Filtro opcional de empresa para APIs"),
    request_id: Optional[str] = Query(None, description="Id opcional para trazabilidad"),
) -> dict:
    """Concilia un Excel de ingresos contra recibos (PDF legacy o API v5).

    Devuelve un JSON con:
      - validados: matches claros
      - dudosos: casos que requieren revisión
      - no_encontrados: incluye BANCO_SIN_RECIBO y RECIBO_SIN_BANCO
      - meta: info de la corrida
    """

    source = (receipts_source or "pdf").strip().lower()
    if source == "pdf" and pdf_salice is None and pdf_alarcon is None:
        raise HTTPException(status_code=400, detail="Tenés que subir al menos 1 PDF (SALICE o Alarcón) o usar receipts_source=api.")

    tmp_dir = tempfile.mkdtemp(prefix="conciliador_")
    rid = request_id or uuid.uuid4().hex
    pdfs: list[tuple[str, str | None, str]] = []  # (path, empresa_override, original_filename)
    if source == "pdf" and pdf_salice is not None:
        pdf_ext = _suffix(pdf_salice.filename or "", default=".pdf")
        pdf_path = os.path.join(tmp_dir, f"{rid}_salice{pdf_ext}")
        pdfs.append((pdf_path, "SALICE", pdf_salice.filename or ""))
    if source == "pdf" and pdf_alarcon is not None:
        pdf_ext = _suffix(pdf_alarcon.filename or "", default=".pdf")
        pdf_path = os.path.join(tmp_dir, f"{rid}_alarcon{pdf_ext}")
        pdfs.append((pdf_path, "ALARCON", pdf_alarcon.filename or ""))

    try:
        prepared = await _prepare_excel_for_run(
            tmp_dir=tmp_dir,
            rid=rid,
            excel=excel,
            record_excel=record_excel,
            record_excel_bank=record_excel_bank,
            record_excel_mp=record_excel_mp,
            raw_bank_files=raw_bank_files,
        )
        working_excel_input = prepared.get("working_excel_paths") or prepared.get("working_excel_path")

        if source == "pdf":
            for path, _empresa, _fn in pdfs:
                up = pdf_salice if _empresa == "SALICE" else pdf_alarcon
                if up is None:
                    continue
                await _save_upload(up, path)

        # Compat: si llega day_weight (viejo), lo aplicamos a ambos multiplicadores.
        if day_weight is not None:
            day_weight_bank_before = float(day_weight)
            day_weight_bank_after = float(day_weight)

        force_list = _parse_json_list(force_validations)
        drop_list = _parse_json_list(drop_dudosos)

        result = compare_excel_pdfs(
            working_excel_input,
            [(p, emp) for (p, emp, _fn) in pdfs],
            margin_days=margin_days,
            tolerance_days_suspect=tolerance_days_suspect,
            day_weight_bank_before=day_weight_bank_before,
            day_weight_bank_after=day_weight_bank_after,
            valid_max_peso=valid_max_peso,
            dudoso_max_peso=dudoso_max_peso,
            mp_mismatch_penalty=mp_mismatch_penalty,
            preconciled_penalty=preconciled_penalty,
            penalty_salice_to_galicia=penalty_salice_to_galicia,
            penalty_alarcon_to_bbva=penalty_alarcon_to_bbva,
            cliente_cuit_mismatch_penalty=cliente_cuit_mismatch_penalty,
            alternatives_cost_delta=alternatives_cost_delta,
            max_options=max_options,
            stage2_candidate_top_k=stage2_candidate_top_k,
            mem_debug=mem_debug,
            show_peso=show_peso,
            show_cuit=show_cuit,
            receipts_source=source,
            api_receipts_days=api_receipts_days,
            api_empresa_filter=api_empresa_filter,
            api_start_date_override=(str(api_start_date).strip() or None) if api_start_date else None,
            api_end_date_override=(
                (str(api_end_date).strip() or None)
                if api_end_date
                else (str((prepared.get("raw_ingestion_meta") or {}).get("raw_max_date") or "").strip() or None)
            ),
            force_validations=force_list,
            drop_dudosos=drop_list,
            excel_record_map=prepared.get("excel_record_map"),
        )
        # Trazabilidad extra
        meta = result.setdefault("meta", {})
        meta["request_id"] = rid
        meta["app_version"] = APP_VERSION
        meta["excel_filename"] = prepared.get("base_excel_filename")
        meta["record_excel_filename"] = prepared.get("base_excel_filename")
        meta["record_excel_filenames"] = [r.get("base_excel_filename") for r in (prepared.get("records") or [])]
        meta["pdfs_filenames"] = [fn for (_p, _e, fn) in pdfs]
        meta["input_mode"] = prepared.get("input_mode")
        meta["raw_bank_filenames"] = prepared.get("raw_bank_filenames")
        if isinstance(prepared.get("raw_ingestion_meta"), dict):
            meta.update(prepared["raw_ingestion_meta"])
        return result

    except ExternalSchemaError as e:
        raise HTTPException(status_code=424, detail=str(e))
    except (ExternalProviderError, ExternalTimeoutError) as e:
        detail = str(e)
        req_id = getattr(e, "request_id", None)
        if req_id:
            detail += f" (request_id={req_id})"
        raise HTTPException(status_code=502, detail=detail)
    except (ExternalConfigError, ValueError) as e:
        raise HTTPException(status_code=400, detail=str(e))
    except Exception as e:
        # En producción, acá iría logging.
        raise HTTPException(status_code=400, detail=f"No se pudo procesar los archivos: {e}")

    finally:
        shutil.rmtree(tmp_dir, ignore_errors=True)


@app.post("/export")
async def export(
    excel: UploadFile | None = File(None, description="(LEGACY) Excel ya consolidado/procesado."),
    record_excel: UploadFile | None = File(None, description="Excel consolidado (record) a actualizar."),
    record_excel_bank: UploadFile | None = File(None, description="Record de movimientos bancarios (BBVA/Galicia)."),
    record_excel_mp: UploadFile | None = File(None, description="Record de Mercado Pago."),
    raw_bank_files: list[UploadFile] | None = File(None, description="Archivos crudos de bancos (BBVA/Galicia/MercadoPago)."),
    pdf_salice: UploadFile | None = File(None, description="PDF de recibos SALICE (opcional)"),
    pdf_alarcon: UploadFile | None = File(None, description="PDF de recibos Alarcón (opcional)"),
    force_validations: str | None = Form(None, description="JSON para promover Dudosos a Validados (opcional)"),
    drop_dudosos: str | None = Form(None, description="Lista JSON para quitar casos de Dudosos (opcional)"),
    format: str = Query("xlsx", pattern="^(xlsx|noencontradosxlsx|devxlsx|zipcsv)$"),
    margin_days: int = Query(5, ge=0, le=31),
    tolerance_days_suspect: int = Query(7, ge=0, le=31),
    # V3.5: multiplicador de días separado por signo.
    day_weight_bank_before: float = Query(20.0, ge=0.0, le=10000.0),
    day_weight_bank_after: float = Query(35.0, ge=0.0, le=10000.0),
    # Compatibilidad (deprecado): si se envía, pisa ambos day_weight_*.
    day_weight: float | None = Query(None, ge=0.0, le=10000.0),
    valid_max_peso: float = Query(260.0, ge=0.0, le=100000.0),
    dudoso_max_peso: float = Query(3500.0, ge=0.0, le=100000.0),
    mp_mismatch_penalty: float = Query(35.0, ge=0.0, le=100000.0),
    preconciled_penalty: float = Query(150.0, ge=0.0, le=100000.0, description="Penalización si el ingreso ya estaba conciliado (columna ok = ok)"),
    penalty_salice_to_galicia: float = Query(45.0, ge=0.0, le=100000.0),
    penalty_alarcon_to_bbva: float = Query(45.0, ge=0.0, le=100000.0),
    cliente_cuit_mismatch_penalty: float = Query(0.0, ge=0.0, le=100000.0, description="Sin efecto en GBA: el CUIT se usa para restringir candidatos plausibles, no para sumar peso"),
    alternatives_cost_delta: float = Query(35.0, ge=0.0, le=100000.0),
    max_options: int = Query(3, ge=1, le=4, description="Máx. opciones por caso dudoso (1 principal + alternativos)"),
    stage2_candidate_top_k: int = Query(120, ge=0, le=2000, description="Prefiltro de candidatos por recibo para Hungarian (0 = sin límite)"),
    mem_debug: bool = Query(False, description="Activa métricas de memoria por etapa en meta"),
    show_peso: bool = Query(False, description="Mostrar la columna Peso en los resultados (por defecto oculto)"),
    show_cuit: bool = Query(False, description="Mostrar columnas de CUIT en los resultados (por defecto oculto)"),
    receipts_source: str = Query("pdf", pattern="^(pdf|api)$", description="Fuente de recibos: pdf (legacy) o api (v5)"),
    api_receipts_days: int = Query(4, ge=1, le=15, description="Fallback de días si no se especifica rango manual de recibos por API"),
    api_start_date: str | None = Query(None, description="Fecha desde para recibos API (YYYY-MM-DD)"),
    api_end_date: str | None = Query(None, description="Fecha hasta para recibos API (YYYY-MM-DD)"),
    api_empresa_filter: str | None = Query(None, description="Filtro opcional de empresa para APIs"),
    request_id: Optional[str] = Query(None, description="Id opcional para trazabilidad"),
):
    """Procesa y devuelve un archivo descargable.

    format:
      - xlsx       -> Excel de ingresos original, completado con VALIDADOS (permite sobrescribir)
      - noencontradosxlsx -> Excel con 4 hojas de no encontrados (BBVA, Mercado Pago, Galicia, Recibos sin banco)
      - devxlsx -> Excel técnico con hojas (Validados/Dudosos/No encontrados/Meta)
      - zipcsv  -> (legacy) zip con CSV
    """
    source = (receipts_source or "pdf").strip().lower()
    if source == "pdf" and pdf_salice is None and pdf_alarcon is None:
        raise HTTPException(status_code=400, detail="Tenés que subir al menos 1 PDF (SALICE o Alarcón) o usar receipts_source=api.")

    tmp_dir = tempfile.mkdtemp(prefix="conciliador_export_")
    rid = request_id or uuid.uuid4().hex
    pdfs: list[tuple[str, str | None, str]] = []  # (path, empresa_override, original_filename)
    if source == "pdf" and pdf_salice is not None:
        pdf_ext = _suffix(pdf_salice.filename or "", default=".pdf")
        pdf_path = os.path.join(tmp_dir, f"{rid}_salice{pdf_ext}")
        pdfs.append((pdf_path, "SALICE", pdf_salice.filename or ""))
    if source == "pdf" and pdf_alarcon is not None:
        pdf_ext = _suffix(pdf_alarcon.filename or "", default=".pdf")
        pdf_path = os.path.join(tmp_dir, f"{rid}_alarcon{pdf_ext}")
        pdfs.append((pdf_path, "ALARCON", pdf_alarcon.filename or ""))

    try:
        prepared = await _prepare_excel_for_run(
            tmp_dir=tmp_dir,
            rid=rid,
            excel=excel,
            record_excel=record_excel,
            record_excel_bank=record_excel_bank,
            record_excel_mp=record_excel_mp,
            raw_bank_files=raw_bank_files,
        )
        working_excel_input = prepared.get("working_excel_paths") or prepared.get("working_excel_path")
        base_excel_filename = prepared.get("base_excel_filename")

        if source == "pdf":
            for path, _empresa, _fn in pdfs:
                up = pdf_salice if _empresa == "SALICE" else pdf_alarcon
                if up is None:
                    continue
                await _save_upload(up, path)

        if day_weight is not None:
            day_weight_bank_before = float(day_weight)
            day_weight_bank_after = float(day_weight)

        force_list = _parse_json_list(force_validations)
        drop_list = _parse_json_list(drop_dudosos)

        result = compare_excel_pdfs(
            working_excel_input,
            [(p, emp) for (p, emp, _fn) in pdfs],
            margin_days=margin_days,
            tolerance_days_suspect=tolerance_days_suspect,
            day_weight_bank_before=day_weight_bank_before,
            day_weight_bank_after=day_weight_bank_after,
            valid_max_peso=valid_max_peso,
            dudoso_max_peso=dudoso_max_peso,
            mp_mismatch_penalty=mp_mismatch_penalty,
            preconciled_penalty=preconciled_penalty,
            penalty_salice_to_galicia=penalty_salice_to_galicia,
            penalty_alarcon_to_bbva=penalty_alarcon_to_bbva,
            cliente_cuit_mismatch_penalty=cliente_cuit_mismatch_penalty,
            alternatives_cost_delta=alternatives_cost_delta,
            max_options=max_options,
            stage2_candidate_top_k=stage2_candidate_top_k,
            mem_debug=mem_debug,
            show_peso=show_peso,
            show_cuit=show_cuit,
            receipts_source=source,
            api_receipts_days=api_receipts_days,
            api_empresa_filter=api_empresa_filter,
            api_start_date_override=(str(api_start_date).strip() or None) if api_start_date else None,
            api_end_date_override=(
                (str(api_end_date).strip() or None)
                if api_end_date
                else (str((prepared.get("raw_ingestion_meta") or {}).get("raw_max_date") or "").strip() or None)
            ),
            force_validations=force_list,
            drop_dudosos=drop_list,
            excel_record_map=prepared.get("excel_record_map"),
        )
        result.setdefault("meta", {})["request_id"] = rid
        result["meta"]["app_version"] = APP_VERSION
        result["meta"]["excel_filename"] = base_excel_filename
        result["meta"]["record_excel_filename"] = base_excel_filename
        result["meta"]["record_excel_filenames"] = [r.get("base_excel_filename") for r in (prepared.get("records") or [])]
        result["meta"]["pdfs_filenames"] = [fn for (_p, _e, fn) in pdfs]
        result["meta"]["input_mode"] = prepared.get("input_mode")
        result["meta"]["raw_bank_filenames"] = prepared.get("raw_bank_filenames")
        if isinstance(prepared.get("raw_ingestion_meta"), dict):
            result["meta"].update(prepared["raw_ingestion_meta"])

        if format == "devxlsx":
            out_path = os.path.join(tmp_dir, f"{rid}_resultado_dev.xlsx")
            export_xlsx(result, out_path)
            # Safari a veces no descarga correctamente si el content-type queda genérico.
            return FileResponse(
                out_path,
                filename="resultado_conciliacion_dev.xlsx",
                media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )

        if format == "xlsx":
            records = prepared.get("records") or []
            if len(records) > 1:
                zip_path = os.path.join(tmp_dir, f"{rid}_records_conciliados.zip")
                with zipfile.ZipFile(zip_path, "w", compression=zipfile.ZIP_DEFLATED) as zf:
                    for rec in records:
                        out_name = _download_name_from_excel_filename(rec.get("base_excel_filename"), "conciliado")
                        out_path = os.path.join(tmp_dir, f"{rid}_{rec.get('key')}_conciliado.xlsx")
                        export_filled_generic_excel(
                            str(rec["working_excel_path"]),
                            result,
                            out_path,
                            allowed_origins=set(rec.get("origins") or []),
                            record_key=str(rec.get("key") or ""),
                        )
                        zf.write(out_path, arcname=out_name)
                return FileResponse(
                    zip_path,
                    filename="records_conciliados.zip",
                    media_type="application/zip",
                )

            default_empresa = None
            if len(pdfs) == 1:
                default_empresa = pdfs[0][1]
            rec = records[0] if records else {}
            out_path = os.path.join(tmp_dir, f"{rid}_ingresos_conciliados.xlsx")
            if rec.get("export_mode") == "generic":
                export_filled_generic_excel(
                    str(rec.get("working_excel_path") or ""),
                    result,
                    out_path,
                    allowed_origins=set(rec.get("origins") or []),
                    record_key=str(rec.get("key") or ""),
                )
            else:
                export_filled_bank_excel(str(rec.get("working_excel_path") or ""), result, out_path, default_empresa=default_empresa)
            dl_name = _download_name_from_excel_filename(rec.get("base_excel_filename") or base_excel_filename, "conciliado")
            return FileResponse(
                out_path,
                filename=dl_name,
                media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )

        if format == "noencontradosxlsx":
            out_path = os.path.join(tmp_dir, f"{rid}_no_encontrados.xlsx")
            export_no_encontrados_xlsx(result, out_path)
            dl_name = _download_name_from_excel_filename(base_excel_filename, "no_encontrados")
            return FileResponse(
                out_path,
                filename=dl_name,
                media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )

        out_path = os.path.join(tmp_dir, f"{rid}_resultado.zip")
        export_zip_csv(result, out_path)
        return FileResponse(
            out_path,
            filename="resultado_conciliacion.zip",
            media_type="application/zip",
        )

    except ExternalSchemaError as e:
        raise HTTPException(status_code=424, detail=str(e))
    except (ExternalProviderError, ExternalTimeoutError) as e:
        detail = str(e)
        req_id = getattr(e, "request_id", None)
        if req_id:
            detail += f" (request_id={req_id})"
        raise HTTPException(status_code=502, detail=detail)
    except (ExternalConfigError, ValueError) as e:
        raise HTTPException(status_code=400, detail=str(e))
    except Exception as e:
        raise HTTPException(status_code=400, detail=f"No se pudo procesar los archivos: {e}")

    finally:
        # Limpieza best-effort (FileResponse needs the file to exist while sending;
        # in dev this is fine. In prod we'd use a background task or persistent storage.)
        pass
