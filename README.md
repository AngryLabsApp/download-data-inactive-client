# download-data-inactive-client

Descarga la data de un cliente de FlowPass (alumnos, paquetes y pagos, asistencias) en un Excel con 3 pestañas por gym.

```bash
python3 -m venv .venv && source .venv/bin/activate
pip install -r requirements.txt
cp .env.example .env   # SUPABASE_URL, SUPABASE_SERVICE_ROLE_KEY y GYM_ID
python export.py       # → output/<Nombre del gym>.xlsx
```

- `GYM_ID` acepta varios separados por coma: sale un Excel por gym.
- Solo alumnos no eliminados. Asistencias sin `tipo = UPDATE`.
- Cruce entre pestañas: `id_alumno` (las 3) e `id_paquete` (paquetes ↔ asistencias).
