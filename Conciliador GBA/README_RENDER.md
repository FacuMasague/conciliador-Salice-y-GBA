# Deploy en Render

## Build
`pip install --upgrade pip && pip install -r requirements.txt`

## Start
`uvicorn app:app --host 0.0.0.0 --port $PORT`

## Variables mínimas
- `API_MODE_ENABLED=true`
- `RECEIPTS_API_BASE_URL=https://m5gba.grupoesi.com.ar`
- `PADRON_API_BASE_URL=https://m5gba.grupoesi.com.ar`
- `RECEIPTS_API_EMPRESA_IDS=2`
- `PADRON_API_EMPRESA_ID=2`
- `GESI_API_USERNAME=...`
- `GESI_API_PASSWORD=...`

## Variables recomendadas
- `RECEIPTS_API_USERNAME=...`
- `RECEIPTS_API_PASSWORD=...`
- `PADRON_API_USERNAME=...`
- `PADRON_API_PASSWORD=...`
- `RECEIPTS_API_PAGE_SIZE=100`
- `RECEIPTS_API_PAGE_SIZE_FALLBACKS=50`
- `RECEIPTS_API_WINDOW_DAYS=1`
- `RECEIPTS_API_TIMEOUT_SECONDS=60`
- `PADRON_API_TIMEOUT_SECONDS=60`
- `PADRON_API_GETITEM_CONCURRENCY=4`

## Nota
No subas `.env` a Render. Cargá las credenciales como variables del servicio.
