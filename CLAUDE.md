# download-data-inactive-client — Guía para Claude

Script Python que exporta la data de un gym de FlowPass a **un Excel con 4 pestañas de data + "Cómo usar"**.
Se usa cuando un cliente (normalmente uno que se va / inactivo) pide "su data".

## Cómo se corre

```bash
source .venv/bin/activate
python export.py   # lee .env: SUPABASE_URL, SUPABASE_SERVICE_ROLE_KEY, GYM_ID (uno o varios con coma)
```

Sale `output/<Nombre del gym>.xlsx`. `.env` y `output/` están en `.gitignore`:
tienen la service role key y datos personales de alumnos, **nunca se suben**.

Mismas variables que FlowPassAPI. Para probar sin crear `.env` (apunta a dev):

```bash
(set -a; source <(grep -E "^SUPABASE_(URL|SERVICE_ROLE_KEY)=" ../FlowPassAPI/.env); set +a; GYM_ID=<id> .venv/bin/python export.py)
```

## Las 4 pestañas

| Pestaña | Tabla base | Una fila = |
|---|---|---|
| Alumnos | `members` + `members_groups → groups` | un alumno |
| Catálogo de paquetes | `planes` (incluye eliminados, columna `activo`) | cada paquete que vende/vendió el gym |
| Paquetes y pagos | `historico` + `charges → payments` | cada paquete comprado o renovado |
| Asistencias | `ingresos` + `planes`, `member_plans`, `groups` | cada asistencia |

**Cruce entre pestañas** (se lo explicamos así al cliente, en la pestaña "Cómo usar" = `GUIA` en `export.py`; si cambian columnas, actualizarla):
- `id_alumno` = `members.id`, está en Alumnos, Paquetes y pagos, Asistencias.
- `id_plan` = `planes.id`. Une catálogo ↔ compras ↔ asistencias.
- `id_compra` = `pago_id`. Une compra ↔ asistencias (y es la llave hacia `charges`).

## Reglas de negocio (no romper)

- **Solo alumnos no eliminados** (`members.deleted_at is null`) en las pestañas de alumnos/compras/asistencias. También se excluyen filas con `deleted_at` propio.
- **Paquetes salen de `historico`, no de `member_plans`**: `member_plans` guarda solo el ciclo
  actual; cada renovación queda como fila nueva en `historico`. Sin `historico` se pierde cuánto pagó en cada renovación.
- **`charges` no tiene FK con `historico`**: se cruzan en Python por `pago_id`. Se ignoran charges con `deleted_at` o `estado = cancelado`.
- **`pagado`** = suma de `payments` con `estado = confirmed`. Draft/cancelled no cuentan.
- **`pagado` y `debe` vacíos** = paquete antiguo, de antes del módulo de cobros. Ahí `estado_pago` viene de `historico.estado_pago`.
- **Asistencias excluyen `tipo = UPDATE`** (ajustes internos, no asistencias reales).
- **`fecha_hora` va en la zona horaria del gym** (`gyms.timezone`), sin tz porque Excel no la guarda.
- `clases_tomadas/limite_clases` en asistencias son foto del momento del check-in, no estado actual.
- Medio de pago se traduce con `MEDIOS` (espejo de `PAYMENT_METHOD_LABELS` en FlowPassAPI).

## Detalles técnicos

- Supabase devuelve máx. 1000 filas por request → `fetch_all` pagina con `.range()` ordenando por `id`.
- Embeds PostgREST: `members!inner(...)` + filtro `members.deleted_at is null` hace el "solo no eliminados".
- Verificado en dev (gym "Demo Grupos"): los conteos coinciden con las queries SQL equivalentes (474 alumnos / 173 paquetes / 617 asistencias).
- Schema de referencia: ver `CLAUDE.md` de FlowPassAPI.
