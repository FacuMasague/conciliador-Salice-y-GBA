# Render setup

Este repo esta preparado para correr Salice y GBA en un solo Web Service de Render, usando rutas.

La raiz tiene un `router.py` que:

- levanta `Conciliador GBA/app.py` en un puerto interno local;
- levanta `conciliador SALICE/app.py` en otro puerto interno local;
- envia `/gba/...` a GBA;
- envia `/salice/...` a Salice;
- redirige `/` a `/gba` por default.

El codigo interno de cada app queda separado y no se mezcla.

## URLs

Cuando el servicio este deployado:

- GBA: `https://conciliador-salice-y-gba.onrender.com/gba`
- Salice: `https://conciliador-salice-y-gba.onrender.com/salice`

## Servicio unico

| Servicio Render | Ejecuta | Rutas |
| --- | --- | --- |
| `conciliador-salice-y-gba` | `router.py` | `/gba`, `/salice` |

## Crear en Render

Opcion recomendada:

1. En Render, crear o abrir el Project.
2. Usar `New +` -> `Blueprint`.
3. Conectar el repo `FacuMasague/conciliador-Salice-y-GBA`.
4. Render leera el `render.yaml` de la raiz y creara un solo servicio: `conciliador-salice-y-gba`.
5. Completar manualmente los secretos marcados como `sync: false`.

Si lo haces manualmente como Web Service:

- `Name`: `conciliador-salice-y-gba`
- `Root Directory`: dejar vacio
- `Build Command`: `pip install --upgrade pip && pip install -r requirements.txt && pip install -r "Conciliador GBA/requirements.txt" && pip install -r "conciliador SALICE/requirements.txt"`
- `Start Command`: `uvicorn router:app --host 0.0.0.0 --port $PORT`

## Secretos

Como ambos programas antes usaban nombres de variables iguales, en el servicio unico se usan prefijos. El router los traduce para cada app.

Para GBA completar:

- `GBA_GESI_API_USERNAME`
- `GBA_GESI_API_PASSWORD`
- `GBA_RECEIPTS_API_USERNAME`
- `GBA_RECEIPTS_API_PASSWORD`
- `GBA_PADRON_API_USERNAME`
- `GBA_PADRON_API_PASSWORD`

Para Salice completar:

- `SALICE_GESI_API_USERNAME`
- `SALICE_GESI_API_PASSWORD`

## Variables importantes

GBA usa las variables sin prefijo porque son las defaults del servicio:

- `RECEIPTS_API_BASE_URL=https://m5gba.grupoesi.com.ar`
- `PADRON_API_BASE_URL=https://m5gba.grupoesi.com.ar`
- `RECEIPTS_API_EMPRESA_IDS=2`
- `PADRON_API_EMPRESA_ID=2`

Salice necesita overrides prefijados para no heredar el tenant de GBA:

- `SALICE_RECEIPTS_API_BASE_URL=https://m5mdp.grupoesi.com.ar`
- `SALICE_PADRON_API_BASE_URL=https://m5mdp.grupoesi.com.ar`
- `SALICE_RECEIPTS_API_EMPRESA_IDS=3,6`
- `SALICE_RECEIPTS_API_PAGE_SIZE=500`

El router les saca el prefijo `SALICE_` solo al proceso de Salice.

## Variables de ruteo

Default incluido en `render.yaml`:

- `DEFAULT_APP=gba`

Si queres que `/` abra Salice, cambia `DEFAULT_APP=salice`.

## Dominios

No hace falta dominio propio para usar rutas. Si mas adelante compras un dominio, podes agregarlo al mismo Web Service y seguir usando `/gba` y `/salice`.
