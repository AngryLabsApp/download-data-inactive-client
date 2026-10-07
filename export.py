import os
import re
import sys
from datetime import date, datetime
from pathlib import Path
from zoneinfo import ZoneInfo

from dotenv import load_dotenv
from openpyxl import Workbook
from openpyxl.styles import Font
from supabase import create_client

PAGE = 1000  # tope de filas por request de Supabase

MEDIOS = {
    "cash": "Efectivo", "transfer": "Transferencia",
    "card-credit": "Tarjeta de crédito", "card-debit": "Tarjeta de débito",
    "yape": "Yape", "plin": "Plin", "codi": "CoDi",
}

load_dotenv()
url, key = os.getenv("SUPABASE_URL"), os.getenv("SUPABASE_SERVICE_ROLE_KEY")
gym_ids = [g.strip() for g in os.getenv("GYM_ID", "").split(",") if g.strip()]
if not url or not key or not gym_ids:
    sys.exit("Falta SUPABASE_URL, SUPABASE_SERVICE_ROLE_KEY o GYM_ID en .env")

sb = create_client(url, key)


def fetch_all(table, select, gym_id, **filters):
    """Trae todas las filas del gym paginando de a PAGE."""
    rows, start = [], 0
    while True:
        q = sb.table(table).select(select).eq("gym_id", gym_id)
        for col in filters.get("is_null", []):
            q = q.is_(col, "null")
        for col, val in filters.get("neq", {}).items():
            q = q.neq(col, val)
        batch = q.order("id").range(start, start + PAGE - 1).execute().data
        rows += batch
        if len(batch) < PAGE:
            return rows
        start += PAGE


def to_date(s):
    return date.fromisoformat(s[:10]) if s else None


def by_alumno(rows, fecha_col):
    # alumno (apellidos, nombre) y luego por fecha
    return sorted(rows, key=lambda r: ((r[2] or "").lower(), (r[3] or "").lower(), r[fecha_col] or date.min))


def alumnos(gym_id):
    rows = fetch_all(
        "members",
        "id,apellidos,nombre,codigo_ingreso,tipo_documento,identificacion,email,telefono,"
        "cumpleanos,direccion,fecha_alta,members_groups(groups(title))",
        gym_id, is_null=["deleted_at"],
    )
    out = [
        (m["id"], m["apellidos"], m["nombre"], m["codigo_ingreso"], m["tipo_documento"],
         m["identificacion"], m["email"], m["telefono"], to_date(m["cumpleanos"]),
         m["direccion"], to_date(m["fecha_alta"]),
         ", ".join(sorted(mg["groups"]["title"] for mg in m["members_groups"] if mg["groups"])))
        for m in rows
    ]
    out.sort(key=lambda r: ((r[1] or "").lower(), (r[2] or "").lower()))
    headers = ["id_alumno", "apellidos", "nombre", "codigo_ingreso", "tipo_documento",
               "identificacion", "email", "telefono", "cumpleanos", "direccion", "fecha_alta", "grupos"]
    return headers, out


def catalogo(gym_id):
    # incluye eliminados: compras y asistencias viejas los referencian
    rows = fetch_all(
        "planes",
        "id,label,amount,dias_mes,ilimitado,limite_clases,partners,is_free,neverExpires,order,deleted_at",
        gym_id,
    )
    rows.sort(key=lambda p: (p["deleted_at"] is not None, p["order"] or 0, (p["label"] or "").lower()))
    out = [
        (p["id"], (p["label"] or "").strip(), p["amount"],
         None if p["neverExpires"] else p["dias_mes"],
         "Ilimitado" if p["ilimitado"] else p["limite_clases"],
         "Sí" if p["partners"] else "No", "Sí" if p["is_free"] else "No",
         "No" if p["deleted_at"] else "Sí")
        for p in rows
    ]
    headers = ["id_plan", "paquete", "precio", "duracion_dias", "limite_clases",
               "compartido", "gratis", "activo"]
    return headers, out


def paquetes(gym_id):
    compras = fetch_all(
        "historico",
        "id,member_id,pago_id,plan_id,plan,fecha_inicio_plan,proxima_fecha_pago,monto,estado_pago,"
        "members!inner(apellidos,nombre),planes(label)",
        gym_id, is_null=["deleted_at", "members.deleted_at"],
    )
    # charges no tiene FK con historico: se cruzan por pago_id
    cargos = {
        c["pago_id"]: c
        for c in fetch_all(
            "charges",
            "id,pago_id,monto_total,monto_pendiente,estado,payments(monto_total,medio_de_pago,paid_at,estado)",
            gym_id, is_null=["deleted_at"], neq={"estado": "cancelado"},
        )
    }
    out = []
    for h in compras:
        c = cargos.get(h["pago_id"])
        pagos = [p for p in (c["payments"] if c else []) if p["estado"] == "confirmed"]
        out.append((
            h["member_id"], h["pago_id"], h["members"]["apellidos"], h["members"]["nombre"],
            h["plan_id"],
            (h["planes"] or {}).get("label") or h["plan"],
            to_date(h["fecha_inicio_plan"]),
            to_date(h["proxima_fecha_pago"]),
            c["monto_total"] if c else h["monto"],
            sum(p["monto_total"] for p in pagos) if c else None,
            c["monto_pendiente"] if c else None,
            c["estado"] if c else (h["estado_pago"] or "").lower() or None,
            ", ".join(sorted({MEDIOS.get(p["medio_de_pago"], p["medio_de_pago"]) for p in pagos if p["medio_de_pago"]})) or None,
            max((to_date(p["paid_at"]) for p in pagos if p["paid_at"]), default=None),
        ))
    headers = ["id_alumno", "id_compra", "apellidos", "nombre", "id_plan", "paquete", "inicio", "vence",
               "precio", "pagado", "debe", "estado_pago", "medio_de_pago", "fecha_ultimo_pago"]
    return headers, by_alumno(out, 6)


def asistencias(gym_id, tz):
    rows = fetch_all(
        "ingresos",
        "id,member_id,pago_id,plan_id,check_in,clases_tomadas,limite_clases,nota,"
        "members!inner(apellidos,nombre),planes(label),member_plans(plan),groups(title)",
        gym_id, is_null=["deleted_at", "members.deleted_at"], neq={"tipo": "UPDATE"},
    )
    out = [
        (i["member_id"], i["pago_id"], i["members"]["apellidos"], i["members"]["nombre"],
         # hora local del gym; Excel no guarda zona horaria
         datetime.fromisoformat(i["check_in"]).astimezone(tz).replace(tzinfo=None) if i["check_in"] else None,
         i["plan_id"],
         (i["planes"] or {}).get("label") or (i["member_plans"] or {}).get("plan"),
         (i["groups"] or {}).get("title"),
         i["clases_tomadas"], i["limite_clases"], i["nota"])
        for i in rows
    ]
    headers = ["id_alumno", "id_compra", "apellidos", "nombre", "fecha_hora", "id_plan", "paquete", "grupo",
               "clases_tomadas", "limite_clases", "nota"]
    return headers, by_alumno(out, 4)


def add_sheet(book, title, headers, rows):
    sheet = book.create_sheet(title)
    sheet.append(headers)
    for r in rows:
        sheet.append(r)
    for c in sheet[1]:
        c.font = Font(bold=True)
    sheet.freeze_panes = "A2"
    sheet.auto_filter.ref = sheet.dimensions
    for col in sheet.columns:
        sheet.column_dimensions[col[0].column_letter].width = 20


Path("output").mkdir(exist_ok=True)
for gym_id in gym_ids:
    gym = sb.table("gyms").select("name,timezone").eq("id", gym_id).maybe_single().execute()
    if not gym or not gym.data:
        print(f"GYM_ID {gym_id} no existe, lo salto")
        continue
    name, tz = gym.data["name"], ZoneInfo(gym.data["timezone"] or "America/Lima")

    book = Workbook()
    book.remove(book.active)
    for title, (headers, rows) in [
        ("Alumnos", alumnos(gym_id)),
        ("Catálogo de paquetes", catalogo(gym_id)),
        ("Paquetes y pagos", paquetes(gym_id)),
        ("Asistencias", asistencias(gym_id, tz)),
    ]:
        add_sheet(book, title, headers, rows)
        print(f"{name} · {title}: {len(rows)} filas")

    out = Path("output") / f"{re.sub(r'[^\w\- ]', '', name).strip()}.xlsx"
    book.save(out)
    print(f"→ {out}")
