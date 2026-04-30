# Render setup

Este repo esta preparado como monorepo para correr dos Web Services independientes dentro del mismo Render Project.

## Servicios

| Servicio | Root Directory | App |
| --- | --- | --- |
| `conciliador-gba` | `Conciliador GBA` | GBA |
| `conciliador-salice` | `conciliador SALICE` | Salice |

Cada servicio conserva su propio codigo, `app.py`, `requirements.txt`, UI y variables. No comparten proceso.

## Crear en Render

1. En Render, crear un Project, por ejemplo `conciliadores`.
2. Usar `New +` -> `Blueprint`.
3. Conectar el repo `FacuMasague/conciliador-Salice-y-GBA`.
4. Render leera el `render.yaml` de la raiz y creara los dos servicios.
5. Completar manualmente los secretos marcados como `sync: false`.

## Secretos

Para `conciliador-gba` completar:

- `GESI_API_USERNAME`
- `GESI_API_PASSWORD`
- `RECEIPTS_API_USERNAME`
- `RECEIPTS_API_PASSWORD`
- `PADRON_API_USERNAME`
- `PADRON_API_PASSWORD`

Para `conciliador-salice` completar:

- `GESI_API_USERNAME`
- `GESI_API_PASSWORD`

## Dominios sugeridos

Agregar los dominios desde `Settings` -> `Custom Domains` en cada servicio:

- `gba.tudominio.com` -> `conciliador-gba`
- `salice.tudominio.com` -> `conciliador-salice`

Despues crear los registros DNS que indique Render, normalmente CNAME desde cada subdominio al dominio `onrender.com` del servicio correspondiente.
