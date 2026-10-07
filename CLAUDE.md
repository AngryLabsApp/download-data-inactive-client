# download-data-inactive-client

## Propósito

Cuando un cliente de FlowPass **se da de baja** y nos pide **toda su data**, este repositorio la
descarga en un solo archivo Excel que le entregamos como **backup**.

Dos prioridades, en este orden:

1. **La data está completa y es correcta**: alumnos, paquetes, compras, pagos y asistencias.
2. **El cliente la entiende solo**: el Excel abre en una guía con ejemplos sencillos, y todas las
   pestañas se cruzan con tres códigos (`id_alumno`, `id_plan`, `id_compra`).

Todo cambio debe respetar esas dos prioridades. Ante la duda, gana lo más simple para el cliente.

> **⚠️ REGLA DE ORO: solo lectura.** Este proyecto usa **100 % consultas de lectura** (`.select()`).
> **Nunca** se agrega `insert`, `update`, `upsert`, `delete`, `rpc` ni nada que edite o cambie datos.
> Corre con la **service role key**, que se salta RLS y escribe en producción sin restricciones:
> un solo error de escritura afectaría la data real del cliente. Si una tarea pide cambiar datos,
> no va en este repo.

---

## 1. Cómo se corre

```bash
source .venv/bin/activate
python export.py
```

Lee del `.env`:

| Variable | Descripción |
|---|---|
| `SUPABASE_URL` | URL del proyecto Supabase (mismas variables que FlowPassAPI) |
| `SUPABASE_SERVICE_ROLE_KEY` | Service role key |
| `GYM_ID` | Uno o varios gyms separados por coma. Genera un Excel por gym |

Salida: `output/<Nombre del gym>.xlsx`.

> **Seguridad:** `.env` y `output/` están en `.gitignore`. Contienen la service role key y datos
> personales de alumnos. **Nunca se suben al repo.**

Para probar sin crear `.env` (apunta a **dev**):

```bash
(set -a; source <(grep -E "^SUPABASE_(URL|SERVICE_ROLE_KEY)=" ../FlowPassAPI/.env); set +a; GYM_ID=<id> .venv/bin/python export.py)
```

---

## 2. Estructura del Excel

| # | Pestaña | Fuente | Una fila = |
|---|---|---|---|
| 1 | **Cómo usar** | `GUIA` en `export.py` (texto fijo) | — Guía para el cliente |
| 2 | **Alumnos** | `members` + `members_groups → groups` | Un alumno |
| 3 | **Catálogo de paquetes** | `planes` (incluye eliminados) | Un paquete que el gym vende o vendió |
| 4 | **Paquetes y pagos** | `historico` + `charges → payments` | Una compra o renovación de un alumno |
| 5 | **Asistencias** | `ingresos` + `planes`, `member_plans`, `groups` | Una asistencia |

### Columnas por pestaña

| Pestaña | Columnas |
|---|---|
| Alumnos | `id_alumno`, `apellidos`, `nombre`, `codigo_ingreso`, `tipo_documento`, `identificacion`, `email`, `telefono`, `cumpleanos`, `direccion`, `fecha_alta`, `grupos` |
| Catálogo de paquetes | `id_plan`, `paquete`, `precio`, `duracion_dias`, `limite_clases`, `compartido`, `gratis`, `activo` |
| Paquetes y pagos | `id_alumno`, `id_compra`, `apellidos`, `nombre`, `id_plan`, `paquete`, `inicio`, `vence`, `precio`, `pagado`, `debe`, `estado_pago`, `medio_de_pago`, `fecha_ultimo_pago` |
| Asistencias | `id_alumno`, `id_compra`, `apellidos`, `nombre`, `fecha_hora`, `id_plan`, `paquete`, `grupo`, `clases_tomadas`, `limite_clases`, `nota` |

---

## 3. Cómo se relacionan las pestañas

```
Alumnos ──── id_alumno ────┐
                           ├──── Paquetes y pagos ──── id_compra ──── Asistencias
Catálogo ─── id_plan ──────┘
              (id_alumno e id_plan también están en Asistencias)
```

| Código | Origen | Qué identifica | Está en |
|---|---|---|---|
| `id_alumno` | `members.id` | Un alumno | Alumnos, Paquetes y pagos, Asistencias |
| `id_plan` | `planes.id` | Un **tipo** de paquete (el del catálogo) | Catálogo, Paquetes y pagos, Asistencias |
| `id_compra` | `pago_id` | Una **compra concreta** (cada renovación es una nueva) | Paquetes y pagos, Asistencias |

### `id_plan` vs `id_compra` (la duda más común del cliente)

- `id_plan` dice **qué** paquete es. Es el mismo para todos los alumnos que compraron ese paquete.
- `id_compra` dice **cuál** compra fue. Si un alumno renueva el mismo paquete 3 veces, tiene 3 `id_compra` distintos y el mismo `id_plan`.

| Alumno | Paquete | `id_plan` | `id_compra` |
|---|---|---|---|
| Juan | Plan Mensual | `aaa` | `111` (enero) |
| Juan | Plan Mensual | `aaa` | `222` (renovación de febrero) |
| María | Plan Mensual | `aaa` | `333` |
| María | Plan Ilimitado | `bbb` | `444` |

Respuesta corta para el cliente:
> "`id_plan` te dice **qué paquete** es. `id_compra` te dice **cuál compra** fue, porque un alumno puede comprar o renovar el mismo paquete varias veces."

---

## 4. Pestaña "Cómo usar" (guía para el cliente)

Una sola pestaña, la primera del archivo, para que el cliente no tenga que buscar en varios lugares.
Contenido en la constante `GUIA` de `export.py`, renderizado por `add_guide()`:

1. **Título y firma**: "Creado por el equipo de desarrollo de FlowPass" (una sola celda, bajo el título).
2. **Pestañas**: qué es cada una.
3. **Cómo se relacionan**: tabla de `id_alumno` / `id_plan` / `id_compra`.
4. **Ejemplo de fórmula**: `BUSCARV` (y `VLOOKUP` para Excel en inglés), guardada como **texto** (`data_type = "s"`) para que se lea y no se calcule.
5. **Notas**: paquetes antiguos sin `pagado`/`debe`, clases como foto del momento, zona horaria.
6. **Ejemplos** (datos inventados: Juan, María):
   - `id_plan` vs `id_compra`.
   - Cómo leer una fila de "Paquetes y pagos" y de "Asistencias", con la frase "→ Se lee: …".
   - Preguntas típicas ("¿Cuánto debe un alumno?", "¿Cuántas clases usó de una compra?", …) y cómo responderlas.

**Convención de `GUIA`** (`add_guide()` la usa para dar formato):

| Fila | Formato |
|---|---|
| Primera fila | Título principal |
| Segunda fila | Firma |
| Tupla de 1 elemento `("…",)` | Barra de sección |
| Tupla de varios elementos | Encabezado de tabla; las filas siguientes llevan borde hasta la próxima fila vacía |
| Lista `[...]` | Texto normal |
| Empieza con `=` | Fórmula de ejemplo (texto, fuente monoespaciada) |
| Empieza con `→` | Lectura de ejemplo (cursiva) |

> **Mantenimiento:** si cambian columnas, nombres de pestaña o el orden de columnas, actualizar `GUIA`
> (en especial la referencia "columna F = id_plan" de la fórmula de ejemplo).

---

## 5. Estilo visual (marca FlowPass)

Colores tomados de `FlowPass-new/src/app.css`, definidos al inicio de `export.py`:

| Constante | Color | Uso |
|---|---|---|
| `VERDE` | `#00F28B` (brand) | Encabezados de las pestañas de data; texto de las barras de sección |
| `OSCURO` | `#13131A` (dark-100) | Fondo de las barras de sección; texto del título y de los encabezados |
| `VERDE_SUAVE` | `#E8FFF2` (brand-50) | Encabezados de las tablas de la guía |
| `GRIS` | `#F5F5F5` (light-200) | Fondo de la fórmula de ejemplo |
| `BORDE_GRIS` | `#CBD5E1` (neutral-300) | Bordes de las tablas |

Pestañas de data: encabezado fijo (`freeze_panes`), autofiltro y ancho de columna 20.
Guía: sin líneas de cuadrícula y con texto ajustado.

---

## 6. Reglas de negocio (no romper)

- **Solo alumnos no eliminados** (`members.deleted_at is null`) en Alumnos, Paquetes y pagos y Asistencias. También se excluyen filas con `deleted_at` propio.
- **El catálogo sí incluye paquetes eliminados** (`activo = No`), porque las compras y asistencias antiguas los referencian. Orden: activos primero, luego por `order`.
- **Las compras salen de `historico`, no de `member_plans`.** `member_plans` guarda solo el ciclo actual; cada renovación queda como fila nueva en `historico`. Sin `historico` se pierde cuánto pagó en cada renovación.
- **`charges` no tiene FK con `historico`**: se cruzan en Python por `pago_id`. Se ignoran charges con `deleted_at` o `estado = cancelado`.
- **`pagado`** = suma de `payments` con `estado = confirmed`. Draft y cancelled no cuentan.
- **`pagado` y `debe` vacíos** = paquete antiguo, anterior al módulo de cobros. En ese caso `estado_pago` viene de `historico.estado_pago`.
- **`estado_pago`** toma los valores `pagado`, `parcial` o `pendiente`.
- **Asistencias excluyen `tipo = UPDATE`** (ajustes internos, no asistencias reales).
- **`fecha_hora` está en la zona horaria del gym** (`gyms.timezone`), sin tz porque Excel no la guarda.
- **`clases_tomadas` / `limite_clases`** en Asistencias son una foto del momento del check-in, no el estado actual. `limite_clases = 0` = ilimitado.
- **Catálogo**: `limite_clases = "Ilimitado"` si `planes.ilimitado`; `duracion_dias` vacío si `neverExpires`.
- **Medio de pago** se traduce con `MEDIOS` (espejo de `PAYMENT_METHOD_LABELS` en FlowPassAPI).

---

## 7. Detalles técnicos

- Supabase devuelve máx. 1000 filas por request → `fetch_all` pagina con `.range()` ordenando por `id`.
- Embeds PostgREST: `members!inner(...)` + filtro `members.deleted_at is null` implementa el "solo no eliminados".
- Schema de referencia: ver `CLAUDE.md` de FlowPassAPI.

### Verificación (dev, gym "Demo Grupos")

| Pestaña | Filas |
|---|---|
| Alumnos | 474 |
| Catálogo de paquetes | 18 |
| Paquetes y pagos | 173 |
| Asistencias | 617 |

Los conteos coinciden con las queries SQL equivalentes. Todo `id_plan` de Paquetes y pagos y de Asistencias existe en el Catálogo.
