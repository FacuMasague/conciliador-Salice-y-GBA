# Render setup

Este repo esta preparado para correr Salice y GBA en un solo Web Service de Render.

La raiz tiene un `router.py` que:

- levanta `Conciliador GBA/app.py` en un puerto interno local;
- levanta `conciliador SALICE/app.py` en otro puerto interno local;
- decide a que programa enviar cada request segun el subdominio/host.

El codigo interno de cada app queda separado y no se mezcla.

## Servicio unico

| Servicio Render | Ejecuta | Enruta por host |
| --- | --- | --- |
| `conciliador-salice-y-gba` | `router.py` | `gba` -> GBA, `salice` -> Salice |

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

## Dominios

Agregar ambos dominios al mismo Web Service `conciliador-salice-y-gba` desde `Settings` -> `Custom Domains`:

- `gba.tudominio.com`
- `salice.tudominio.com`

Despues crear los registros DNS que indique Render. Normalmente seran CNAME desde cada subdominio al dominio `onrender.com` del mismo servicio.

## Variables de ruteo

Defaults incluidos en `render.yaml`:

- `GBA_HOSTS=gba`
- `SALICE_HOSTS=salice`
- `DEFAULT_APP=gba`

Si tus dominios no contienen las palabras `gba` o `salice`, cambia `GBA_HOSTS` y `SALICE_HOSTS` por los hosts reales separados por coma.
