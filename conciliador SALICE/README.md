# Conciliador - Tests con datos reales

Estos tests usan como **base de prueba** los archivos reales que subiste:

- `/mnt/data/Movimientos bancarios 2026.xlsx`
- `/mnt/data/reporte salice.pdf`
- `/mnt/data/reporte alarcon.pdf`

## Estructura
- `src/conciliador/` : parser de PDF, loader de Excel, matcher y pipeline (mínimo para poder testear)
- `tests/` : suite de tests (contrato, parser, loader, end-to-end)

## Correr los tests
Desde la carpeta del proyecto:

```bash
cd /mnt/data/conciliador
pytest
```

## Levantar la API (para la web)

Este repo incluye un backend FastAPI y una UI mínima.

Desde la carpeta del proyecto:

```bash
cd /mnt/data/conciliador
uvicorn app:app --reload --host 0.0.0.0 --port 8000
```

- UI: http://localhost:8000/
- Swagger: http://localhost:8000/docs

Endpoint principal:
- `POST /compare`
  - V5.2.3: multipart `record_excel` + `raw_bank_files[]`; recibos y cobradores se resuelven automáticamente.
  - Legacy: multipart `excel` + `pdf_salice/pdf_alarcon`
- `POST /export`
  - V5.2.3: multipart `record_excel` + `raw_bank_files[]`; recibos y cobradores se resuelven automáticamente.
  - Legacy: multipart `excel` + `pdf_salice/pdf_alarcon`


## Notas
- `nro_recibo` se trata siempre como **texto**.
- Los **raros** se clasifican dentro de `dudosos` con `is_raro=true`.
- `no_encontrados` incluye dos tipos:
  - `BANCO_SIN_RECIBO` (movimiento en Excel sin recibo)
  - `RECIBO_SIN_BANCO` (recibo en PDF sin movimiento en Excel)
- Optimización de memoria (incremental):
  - En etapa Dudosos, el matcher usa prefiltrado de candidatos por recibo (`stage2_candidate_top_k`, default `120`) y resuelve Hungarian por componentes conectados para evitar una sola matriz gigante.
  - El parseo de PDF se reutiliza por archivo (se evita releer el mismo PDF múltiples veces para parse/warnings/rango).
  - Se agregó debug opcional de memoria por etapa (`mem_debug=1` en `/compare` o `/export`, o variable `CONCILIADOR_MEM_DEBUG=1`), visible en `meta.mem_stages`.
- V3.9.0:
  - Si una fila del Excel ya tiene columna `ok` con valor `ok`, el matcher aplica una penalización configurable (`preconciled_penalty`, default `150`) y agrega motivo en Dudosos indicando el recibo previamente conciliado (si está disponible en columna `recibo`).
  - El export de XLSX de usuario ahora permite sobrescribir `ok/cliente/recibo` para los validados.
- V3.9.3:
  - El botón principal pasa a llamarse `Descargar validados`.
  - Se agrega `Descargar dudosos`, que genera `<nombre_original>_dudosos.xlsx` usando el mismo Excel base y completándolo con las opciones dudosas.
  - En ese export de dudosos se agrega columna `cliente nombre` a la derecha de `recibo`.
- V4.0.0:
  - Ya no es obligatorio subir un Excel de ingresos previamente procesado.
  - Se pueden subir archivos crudos de bancos (`BBVA`, `GALICIA`, `MERCADOPAGO`) y el sistema detecta automáticamente el tipo por formato.
  - El backend fusiona los crudos en el Excel consolidado (record) antes de conciliar.
  - El export de `Descargar validados` devuelve el record consolidado actualizado con los ingresos nuevos + conciliación.
- V5.2.2:
  - Los recibos, importes y medios de pago siguen obteniéndose desde la API de GESI.
  - El `fletero/cobrador` se completa por número de recibo desde el reporte PDF de Pedidos Móviles, que informa explícitamente `[ID - nombre]`.
  - Las fojas de reparto y el vendedor comercial del cliente ya no se usan como cobrador.
  - La interfaz solicita el PDF de control también en modo API y muestra la cobertura obtenida en la metadata.
  - La UI incorpora ordenamiento desde los encabezados, divisor visual banco/recibo y acciones de revisión alineadas con GBA.
- V5.2.3:
  - Se eliminó de la interfaz el control y la carga de PDF de cobradores.
  - La web funciona siempre con los recibos de GESI y completa el cobrador desde una fuente interna administrada por el sistema.
  - El operador sólo sube los extractos bancarios y el record consolidado.
  - El catálogo interno predeterminado está en `data/cobradores.json`; puede reemplazarse del lado del servidor con `CONCILIADOR_COBRADORES_PATH`, sin exponer esa tarea en la interfaz.

Si en algún entorno cambiás rutas o nombres de archivos, ajustá `tests/conftest.py`.
