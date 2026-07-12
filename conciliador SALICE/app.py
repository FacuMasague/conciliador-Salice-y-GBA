from __future__ import annotations

import os
import json
import shutil
import tempfile
import uuid
from typing import Optional

from dotenv import load_dotenv
import pathlib
# Carga .env usando la ruta absoluta del directorio donde está app.py,
# independientemente del directorio de trabajo desde donde se lanza uvicorn.
load_dotenv(pathlib.Path(__file__).parent / ".env", override=True)

from fastapi import FastAPI, File, Form, HTTPException, Query, UploadFile
from fastapi.middleware.cors import CORSMiddleware
from fastapi.staticfiles import StaticFiles
from fastapi.responses import FileResponse

from src.conciliador.pipeline import compare_excel_pdfs
from src.conciliador.exporter import export_xlsx, export_filled_bank_excel, export_zip_csv, export_no_encontrados_xlsx
from src.conciliador.raw_bank_ingest import build_runtime_workbook_from_raw
from src.conciliador.external.errors import ExternalConfigError, ExternalProviderError, ExternalSchemaError, ExternalTimeoutError


# Versión visible en UI y en /docs
APP_VERSION = "5.2.0"
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


async def _prepare_excel_for_run(
    *,
    tmp_dir: str,
    rid: str,
    excel: UploadFile | None,
    record_excel: UploadFile | None,
    raw_bank_files: list[UploadFile] | None,
) -> dict:
    """Prepara el Excel de trabajo.

    Modos:
      - v4: record_excel + raw_bank_files => construye workbook runtime (record + crudos).
      - legacy: excel => usa directamente el Excel subido.
    """
    raw_files = [f for f in (raw_bank_files or []) if f is not None]

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
            "base_excel_filename": record_excel.filename,
            "input_mode": "v4_raw_plus_record",
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
        "base_excel_filename": excel.filename,
        "input_mode": "legacy_excel",
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
    raw_bank_files: list[UploadFile] | None = File(None, description="Archivos crudos de bancos (BBVA/Galicia/MercadoPago)."),
    pdf_salice: UploadFile | None = File(None, description="PDF de recibos SALICE (opcional)"),
    pdf_alarcon: UploadFile | None = File(None, description="PDF de recibos Alarcón (opcional)"),
    force_validations: str | None = Form(None, description="Lista JSON para promover Dudosos a Validados (opcional)"),
    drop_dudosos: str | None = Form(None, description="Lista JSON para quitar casos de Dudosos (opcional)"),
    margin_days: int = Query(5, ge=0, le=31),
    tolerance_days_suspect: int = Query(7, ge=0, le=31),
    # V3.5: multiplicador de días separado por signo.
    day_weight_bank_before: float = Query(40.0, ge=0.0, le=10000.0, description="Multiplicador de días si el ingreso bancario es anterior (o mismo día) que el recibo"),
    day_weight_bank_after: float = Query(50.0, ge=0.0, le=10000.0, description="Multiplicador de días si el ingreso bancario es posterior al recibo (delay del banco)"),
    # Compatibilidad (deprecado): si se envía, pisa ambos day_weight_*.
    day_weight: float | None = Query(None, ge=0.0, le=10000.0, description="(DEPRECADO) Multiplicador de días único"),
    valid_max_peso: float = Query(150.0, ge=0.0, le=100000.0),
    dudoso_max_peso: float = Query(3500.0, ge=0.0, le=100000.0),
    mp_mismatch_penalty: float = Query(35.0, ge=0.0, le=100000.0, description="Penalización cuando medio MP↔origen banco está cruzado"),
    preconciled_penalty: float = Query(150.0, ge=0.0, le=100000.0, description="Penalización si el ingreso ya estaba conciliado (columna ok = ok)"),
    alternatives_cost_delta: float = Query(50.0, ge=0.0, le=100000.0, description="Alternativos si peso <= principal + delta"),
    max_options: int = Query(4, ge=1, le=4, description="Máx. opciones por caso dudoso (1 principal + alternativos)"),
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
            raw_bank_files=raw_bank_files,
        )
        working_excel_path = str(prepared["working_excel_path"])

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
            working_excel_path,
            [(p, emp) for (p, emp, _fn) in pdfs],
            margin_days=margin_days,
            tolerance_days_suspect=tolerance_days_suspect,
            day_weight_bank_before=day_weight_bank_before,
            day_weight_bank_after=day_weight_bank_after,
            valid_max_peso=valid_max_peso,
            dudoso_max_peso=dudoso_max_peso,
            mp_mismatch_penalty=mp_mismatch_penalty,
            preconciled_penalty=preconciled_penalty,
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
        )
        # Trazabilidad extra
        meta = result.setdefault("meta", {})
        meta["request_id"] = rid
        meta["app_version"] = APP_VERSION
        meta["excel_filename"] = prepared.get("base_excel_filename")
        meta["record_excel_filename"] = prepared.get("base_excel_filename")
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
    raw_bank_files: list[UploadFile] | None = File(None, description="Archivos crudos de bancos (BBVA/Galicia/MercadoPago)."),
    pdf_salice: UploadFile | None = File(None, description="PDF de recibos SALICE (opcional)"),
    pdf_alarcon: UploadFile | None = File(None, description="PDF de recibos Alarcón (opcional)"),
    force_validations: str | None = Form(None, description="JSON para promover Dudosos a Validados (opcional)"),
    drop_dudosos: str | None = Form(None, description="Lista JSON para quitar casos de Dudosos (opcional)"),
    format: str = Query("xlsx", pattern="^(xlsx|dudososxlsx|noencontradosxlsx|devxlsx|zipcsv)$"),
    margin_days: int = Query(5, ge=0, le=31),
    tolerance_days_suspect: int = Query(7, ge=0, le=31),
    # V3.5: multiplicador de días separado por signo.
    day_weight_bank_before: float = Query(40.0, ge=0.0, le=10000.0),
    day_weight_bank_after: float = Query(50.0, ge=0.0, le=10000.0),
    # Compatibilidad (deprecado): si se envía, pisa ambos day_weight_*.
    day_weight: float | None = Query(None, ge=0.0, le=10000.0),
    valid_max_peso: float = Query(150.0, ge=0.0, le=100000.0),
    dudoso_max_peso: float = Query(3500.0, ge=0.0, le=100000.0),
    mp_mismatch_penalty: float = Query(35.0, ge=0.0, le=100000.0),
    preconciled_penalty: float = Query(150.0, ge=0.0, le=100000.0, description="Penalización si el ingreso ya estaba conciliado (columna ok = ok)"),
    alternatives_cost_delta: float = Query(50.0, ge=0.0, le=100000.0),
    max_options: int = Query(4, ge=1, le=4, description="Máx. opciones por caso dudoso (1 principal + alternativos)"),
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
      - dudososxlsx -> Excel de ingresos original, completado solamente con DUDOSOS principales
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
            raw_bank_files=raw_bank_files,
        )
        working_excel_path = str(prepared["working_excel_path"])
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
            working_excel_path,
            [(p, emp) for (p, emp, _fn) in pdfs],
            margin_days=margin_days,
            tolerance_days_suspect=tolerance_days_suspect,
            day_weight_bank_before=day_weight_bank_before,
            day_weight_bank_after=day_weight_bank_after,
            valid_max_peso=valid_max_peso,
            dudoso_max_peso=dudoso_max_peso,
            mp_mismatch_penalty=mp_mismatch_penalty,
            preconciled_penalty=preconciled_penalty,
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
        )
        result.setdefault("meta", {})["request_id"] = rid
        result["meta"]["app_version"] = APP_VERSION
        result["meta"]["excel_filename"] = base_excel_filename
        result["meta"]["record_excel_filename"] = base_excel_filename
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
            # User-facing: devolvemos el mismo Excel de ingresos, pero completado con validados.
            # Si es modo SIMPLE (1 PDF), inferimos Empresa por el PDF subido.
            default_empresa = None
            if len(pdfs) == 1:
                default_empresa = pdfs[0][1]
            out_path = os.path.join(tmp_dir, f"{rid}_ingresos_conciliados.xlsx")
            export_filled_bank_excel(working_excel_path, result, out_path, default_empresa=default_empresa)
            dl_name = _download_name_from_excel_filename(base_excel_filename, "conciliado")
            return FileResponse(
                out_path,
                filename=dl_name,
                media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )

        if format == "dudososxlsx":
            out_path = os.path.join(tmp_dir, f"{rid}_ingresos_dudosos.xlsx")
            export_filled_bank_excel(
                working_excel_path,
                result,
                out_path,
                row_source="dudosos",
                only_ranking_1=True,
                write_cliente_nombre_col=True,
                clear_existing_assignments=True,
                write_ok_marker=False,
                compact_only_source_rows=True,
            )
            dl_name = _download_name_from_excel_filename(base_excel_filename, "dudosos")
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
