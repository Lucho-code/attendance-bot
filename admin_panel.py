import os
import sys
import io
from datetime import date, datetime, timedelta

import streamlit as st
import pandas as pd
import pytz
from dotenv import load_dotenv

load_dotenv(os.path.join(os.path.dirname(__file__), ".env"))
sys.path.insert(0, os.path.dirname(__file__))

from database   import Database
from reports    import build_xlsx, MESES_ES, ABSENCE_TYPES, calcular_horas
from pdf_report import build_pdf

TIMEZONE = pytz.timezone("America/Argentina/Buenos_Aires")

def ahora():
    return datetime.now(TIMEZONE)

@st.cache_resource
def get_db():
    return Database()

db = get_db()

# ── Configuración de página ───────────────────────────────────────────────────
st.set_page_config(
    page_title="Panel de Asistencia",
    page_icon="📋",
    layout="wide",
)

st.markdown("""
<style>
  .stDeployButton { display: none !important; }
  #MainMenu { visibility: hidden; }
  footer    { visibility: hidden; }

  /* ── Métricas tipo Dashcube ── */
  div[data-testid="stMetric"] {
    background: #333330;
    border-radius: 12px;
    padding: 20px 24px;
    border: 1px solid rgba(232,232,0,0.12);
  }
  div[data-testid="stMetric"] label {
    color: #ECEFCA !important;
    opacity: 0.5;
    font-size: 11px;
    text-transform: uppercase;
    letter-spacing: .06em;
  }
  div[data-testid="stMetric"] [data-testid="stMetricValue"] {
    color: #ECEFCA !important;
    font-size: 32px;
    font-weight: 700;
    line-height: 1.1;
  }

  /* ── Tabs ── */
  .stTabs [data-baseweb="tab-list"] {
    background: transparent;
    border-bottom: 1px solid rgba(236,239,202,0.1);
    gap: 8px;
    padding-bottom: 0;
  }
  .stTabs [data-baseweb="tab"] {
    color: #ECEFCA !important;
    opacity: 0.45;
    font-size: 13px;
    padding-bottom: 10px;
    border-bottom: 2px solid transparent;
    border-radius: 0;
    background: transparent !important;
  }
  .stTabs [aria-selected="true"] {
    opacity: 1 !important;
    color: #E8E800 !important;
    border-bottom: 2px solid #E8E800 !important;
    background: transparent !important;
    font-weight: 600;
  }

  /* ── Expanders ── */
  .streamlit-expanderHeader {
    background: #333330 !important;
    border: 1px solid rgba(236,239,202,0.08) !important;
    border-radius: 10px !important;
  }

  /* ── DataFrames ── */
  .stDataFrame { border-radius: 10px; overflow: hidden; }

  /* ── Inputs y selects ── */
  .stTextInput input, .stSelectbox select {
    background: #333330 !important;
    border: 1px solid rgba(236,239,202,0.15) !important;
    border-radius: 8px !important;
  }

  /* ── Botones ── */
  .stButton > button {
    border-radius: 8px !important;
    font-weight: 600;
    letter-spacing: .02em;
  }

  /* ── Header 2H ── */
  .header-2h {
    display: flex;
    align-items: center;
    gap: 16px;
    padding: 10px 0 18px;
    border-bottom: 1px solid rgba(232,232,0,0.25);
    margin-bottom: 24px;
  }
  .header-title {
    font-size: 20px;
    font-weight: 700;
    color: #ECEFCA;
    letter-spacing: -0.3px;
  }
  .header-sub {
    font-size: 11px;
    color: #ECEFCA;
    opacity: 0.4;
    margin-top: 3px;
    letter-spacing: .03em;
  }
</style>
""", unsafe_allow_html=True)

# Header con logo 2H real
import base64, pathlib
_logo_path = pathlib.Path(__file__).parent / "logo_2h.png"
_logo_b64  = base64.b64encode(_logo_path.read_bytes()).decode() if _logo_path.exists() else ""
_logo_html = (f'<img src="data:image/png;base64,{_logo_b64}" width="56" height="56" '
              f'style="border-radius:50%;">' if _logo_b64 else "")

st.markdown(f"""
<div class="header-2h">
  {_logo_html}
  <div>
    <div class="header-title">2H Mov. Suelos &nbsp;<span style="color:#E8E800;">|</span>&nbsp; Panel de Control</div>
    <div class="header-sub">Actualizado: {ahora().strftime('%d/%m/%Y  %H:%M')}</div>
  </div>
</div>
""", unsafe_allow_html=True)

tab_hoy, tab_emp, tab_rep, tab_obras, tab_normas = st.tabs(["📅 Hoy", "👥 Empleados", "📊 Reportes", "🏗️ Obras", "📋 Normas"])

# ── TAB 1: Estado de hoy (auto-refresca cada 30 seg) ─────────────────────────
@st.fragment(run_every=30)
def tab_hoy_content():
    hoy       = ahora().date()
    empleados = db.list_employees()

    rows = []
    presentes = ausentes = sin_salida = salieron = 0

    for emp in empleados:
        st_hoy = db.get_today_status(emp["telegram_id"], hoy)
        shift  = db.get_shift(emp["telegram_id"])
        turno  = f"{shift[0]:02d}:{shift[1]:02d} – {shift[2]:02d}:{shift[3]:02d}"

        if st_hoy and st_hoy.get("entry_time") and not st_hoy.get("exit_time"):
            estado    = "🟢 Entró"
            presentes += 1
        elif st_hoy and st_hoy.get("exit_time"):
            estado   = "🔵 Salió"
            salieron += 1
        elif not st_hoy:
            estado   = "🔴 Sin registro"
            ausentes += 1
        else:
            estado     = "🟡 Sin salida"
            sin_salida += 1

        ent = st_hoy["entry_time"].strftime("%H:%M") if st_hoy and st_hoy.get("entry_time") else "–"
        sal = st_hoy["exit_time"].strftime("%H:%M")  if st_hoy and st_hoy.get("exit_time")  else "–"

        # Calcular horas con tolerancia aplicada (por sesión, sumando todas)
        n_tot = e50_tot = e100_tot = 0.0
        if st_hoy:
            is_holiday = bool(db.is_holiday(hoy))
            for s in db.get_daily_sessions(emp["telegram_id"], hoy):
                if s.get("entry_time") and s.get("exit_time"):
                    sn, se50, se100 = calcular_horas(
                        s["entry_time"], s["exit_time"], hoy.weekday(), is_holiday)
                    n_tot   += sn
                    e50_tot += se50
                    e100_tot += se100

        hs   = f"{n_tot:.2f}"   if n_tot   > 0 else "–"
        e50  = f"{e50_tot:.2f}" if e50_tot  > 0 else "–"
        e100 = f"{e100_tot:.2f}"if e100_tot > 0 else "–"

        rows.append({
            "Empleado 2H Mov. Suelos": emp["name"],
            "Turno":        turno,
            "Estado":       estado,
            "Entrada":      ent,
            "Salida":       sal,
            "Hs. Normales": hs,
            "Extra 50%":    e50,
            "Extra 100%":   e100,
        })

    c1, c2, c3, c4, c5 = st.columns(5)
    c1.metric("Total",            len(empleados))
    c2.metric("🟢 Entraron",      presentes)
    c3.metric("🔵 Salieron",      salieron)
    c4.metric("🔴 Sin registrar", ausentes)
    c5.metric("🟡 Sin salida",    sin_salida)

    st.divider()

    # Mapa nombre -> categoria para separar secciones
    cat_map = {e["name"]: e.get("categoria", "empleado") for e in empleados}

    secciones = [
        ("empleado",          "👷 Empleados"),
        ("administracion",    "💼 Administración"),
        ("direccion_tecnica", "⚙️ Dirección Técnica"),
    ]

    col_key = "Empleado 2H Mov. Suelos"
    cols_empleados = ["Empleado 2H Mov. Suelos", "Turno", "Estado",
                      "Entrada", "Salida", "Hs. Normales", "Extra 50%", "Extra 100%"]
    cols_otros     = ["Empleado 2H Mov. Suelos", "Turno", "Estado", "Entrada", "Salida"]

    for cat_key, cat_label in secciones:
        filas = [r for r in rows if cat_map.get(r[col_key]) == cat_key]
        if not filas:
            continue
        st.subheader(cat_label)
        cols = cols_empleados if cat_key == "empleado" else cols_otros
        st.dataframe(
            pd.DataFrame(filas)[cols],
            use_container_width=True,
            hide_index=True,
            column_config={"Estado": st.column_config.TextColumn(width="medium")},
        )

        # Detalle de sesiones para Administración y Dirección Técnica
        if cat_key in ("administracion", "direccion_tecnica"):
            for emp in [e for e in empleados if e.get("categoria", "empleado") == cat_key]:
                sesiones = db.get_daily_sessions(emp["telegram_id"], hoy)
                if len(sesiones) > 1:
                    with st.expander(f"  {emp['name']} — {len(sesiones)} sesiones hoy"):
                        ses_rows = []
                        for i, s in enumerate(sesiones, 1):
                            ent = s["entry_time"].strftime("%H:%M") if s.get("entry_time") else "–"
                            sal = s["exit_time"].strftime("%H:%M")  if s.get("exit_time")  else "Abierta"
                            if s.get("entry_time") and s.get("exit_time"):
                                is_h = bool(db.is_holiday(hoy))
                                sn, se50, se100 = calcular_horas(s["entry_time"], s["exit_time"], hoy.weekday(), is_h)
                                hs = f"{sn + se50 + se100:.2f}"
                            else:
                                hs = "–"
                            ses_rows.append({
                                "Sesión": f"#{i}",
                                "Entrada": ent,
                                "Salida":  sal,
                                "Horas":   hs,
                            })
                        st.dataframe(pd.DataFrame(ses_rows), use_container_width=True, hide_index=True)

    st.caption(f"Actualizado: {ahora().strftime('%H:%M:%S')} · refresca cada 30 seg")

with tab_hoy:
    tab_hoy_content()

# ── TAB 2: Empleados ──────────────────────────────────────────────────────────
with tab_emp:
    empleados = db.list_employees()
    if not empleados:
        st.info("No hay empleados registrados aún.")
    else:
        secciones = [
            ("empleado",          "👷 Empleados"),
            ("administracion",    "💼 Administración"),
            ("direccion_tecnica", "⚙️ Dirección Técnica"),
        ]

        col_cfg = {
            "Nombre":     st.column_config.TextColumn(width="large"),
            "Turno":      st.column_config.TextColumn(width="small"),
            "Registrado": st.column_config.TextColumn(width="small"),
        }
        for cat_key, cat_label in secciones:
            grupo = [e for e in empleados if e.get("categoria", "empleado") == cat_key]
            if not grupo:
                continue
            st.subheader(cat_label)
            filas = []
            for emp in grupo:
                shift = db.get_shift(emp["telegram_id"])
                turno = f"{shift[0]:02d}:{shift[1]:02d} – {shift[2]:02d}:{shift[3]:02d}"
                filas.append({
                    "Nombre":     emp["name"],
                    "Turno":      turno,
                    "Registrado": emp["registered_at"][:10],
                })
            st.dataframe(
                pd.DataFrame(filas),
                use_container_width=True,
                hide_index=True,
                column_config=col_cfg,
            )

        st.divider()
        st.subheader("Últimos fichajes por empleado")
        emp_sel = st.selectbox(
            "Seleccioná un empleado",
            options=[e["name"] for e in empleados],
        )
        if emp_sel:
            emp_obj = next(e for e in empleados if e["name"] == emp_sel)
            records = db.get_employee_records(emp_obj["telegram_id"], limit=15)
            es_empleado = emp_obj.get("categoria", "empleado") == "empleado"
            if records:
                rows_rec = []
                for r in records:
                    ent = r["entry_time"].strftime("%H:%M") if r.get("entry_time") else "–"
                    sal = r["exit_time"].strftime("%H:%M")  if r.get("exit_time")  else "Sin salida"
                    rn = re50 = re100 = 0.0
                    if r.get("entry_time") and r.get("exit_time"):
                        d = date.fromisoformat(r["date"])
                        is_h = bool(db.is_holiday(d))
                        rn, re50, re100 = calcular_horas(r["entry_time"], r["exit_time"], d.weekday(), is_h)
                    hs = f"{rn + re50 + re100:.2f}" if (rn + re50 + re100) > 0 else "–"
                    row = {"Fecha": r["date"], "Entrada": ent, "Salida": sal}
                    if es_empleado:
                        row["Hs. Normales"] = f"{rn:.2f}" if rn > 0 else "–"
                        row["Extra 50%"]    = f"{re50:.2f}" if re50 > 0 else "–"
                        row["Extra 100%"]   = f"{re100:.2f}" if re100 > 0 else "–"
                    else:
                        row["Horas"] = hs
                    rows_rec.append(row)
                st.dataframe(pd.DataFrame(rows_rec), use_container_width=True,
                             hide_index=True)
            else:
                st.info("Sin registros.")

# ── TAB 3: Reportes ───────────────────────────────────────────────────────────
with tab_rep:
    st.subheader("Descargar reporte XLSX")

    col_a, col_b = st.columns(2)
    with col_a:
        modo = st.radio("Período", ["Mes actual", "Año completo", "Rango personalizado"])
    with col_b:
        hoy = ahora().date()
        if modo == "Mes actual":
            start_r = hoy.replace(day=1)
            end_r   = hoy
            label_r = f"{MESES_ES[hoy.month]} {hoy.year}"
        elif modo == "Año completo":
            start_r = date(hoy.year, 1, 1)
            end_r   = date(hoy.year, 12, 31)
            label_r = f"Año {hoy.year}"
        else:
            start_r = st.date_input("Desde", value=hoy.replace(day=1))
            end_r   = st.date_input("Hasta", value=hoy)
            label_r = f"{start_r.strftime('%d/%m')} al {end_r.strftime('%d/%m/%Y')}"

    col_xlsx, col_pdf = st.columns(2)

    with col_xlsx:
        if st.button("📥 Generar Excel"):
            with st.spinner("Generando Excel..."):
                records  = db.get_records_by_period(start_r, end_r)
                holidays = db.get_holidays(start_r, end_r)
                absences = db.get_absences(start_r, end_r)
                buf = build_xlsx(
                    db, records, f"Reporte {label_r}",
                    start_date=start_r, end_date=end_r,
                    holidays=holidays, absences=absences,
                )
            filename = f"asistencia_{start_r.strftime('%Y%m%d')}_{end_r.strftime('%Y%m%d')}.xlsx"
            st.download_button(
                label="⬇️ Descargar Excel",
                data=buf,
                file_name=filename,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )

    with col_pdf:
        if st.button("📄 Generar PDF"):
            with st.spinner("Generando PDF..."):
                buf_pdf = build_pdf(db, start_r, end_r, f"Reporte {label_r}")
            filename_pdf = f"asistencia_{start_r.strftime('%Y%m%d')}_{end_r.strftime('%Y%m%d')}.pdf"
            st.download_button(
                label="⬇️ Descargar PDF",
                data=buf_pdf,
                file_name=filename_pdf,
                mime="application/pdf",
            )

    st.divider()
    st.subheader("Estadísticas rápidas del mes")

    hoy   = ahora().date()
    start = hoy.replace(day=1)
    recs  = db.get_records_by_period(start, hoy)

    if recs:
        filas_mes = []
        for r in recs:
            hs = 0.0
            for s in r.get("sessions", [r]):
                if s.get("entry_time") and s.get("exit_time"):
                    d = date.fromisoformat(r["date"])
                    is_h = bool(db.is_holiday(d))
                    n, e50, e100 = calcular_horas(s["entry_time"], s["exit_time"], d.weekday(), is_h)
                    hs += n + e50 + e100
            if hs > 0:
                filas_mes.append({"Empleado": r["name"], "Horas": round(hs, 2)})

        if filas_mes:
            resumen = pd.DataFrame(filas_mes).groupby("Empleado")["Horas"].sum().reset_index()
            resumen.columns = ["Empleado", "Total Horas del Mes"]
            st.dataframe(resumen, use_container_width=True, hide_index=True)
    else:
        st.info("Sin registros para el mes actual.")

# ── TAB 4: Obras ──────────────────────────────────────────────────────────────
with tab_obras:
    hoy    = ahora().date()
    start  = hoy.replace(day=1)
    obras  = db.list_obras(active_only=False)

    if not obras:
        st.info("No hay obras registradas. Usá /obra Nombre en el bot para agregar una.")
    else:
        obras_activas = [o for o in obras if o["active"]]
        obras_cerradas = [o for o in obras if not o["active"]]

        def _render_obra(obra, activa: bool):
            registros = db.get_obra_hours(obra["id"], start, hoy)
            total_hs  = sum(r.get("total_hours") or 0 for r in registros)
            en_obra   = []
            if activa:
                for emp in db.list_employees():
                    st_hoy = db.get_today_status(emp["telegram_id"], hoy)
                    if (st_hoy and st_hoy.get("obra_id") == obra["id"]
                            and st_hoy.get("entry_time") and not st_hoy.get("exit_time")):
                        en_obra.append(emp["name"])

            estado = "🟢" if activa else "🔴"
            header = (f"{estado} **{obra['name']}** — {total_hs:.1f} hs este mes"
                      + (f" · {', '.join(en_obra)}" if en_obra else ""))

            with st.expander(header):
                if registros:
                    filas = []
                    for r in registros:
                        ent = r["entry_time"].strftime("%H:%M") if r.get("entry_time") else "–"
                        sal = r["exit_time"].strftime("%H:%M")  if r.get("exit_time")  else "Abierta"
                        if r.get("entry_time") and r.get("exit_time"):
                            d = date.fromisoformat(r["date"])
                            is_h = bool(db.is_holiday(d))
                            on, oe50, oe100 = calcular_horas(r["entry_time"], r["exit_time"], d.weekday(), is_h)
                            hs = f"{on + oe50 + oe100:.2f}"
                        else:
                            hs = "–"
                        filas.append({
                            "Empleado": r["name"],
                            "Fecha":    r["date"],
                            "Entrada":  ent,
                            "Salida":   sal,
                            "Horas":    hs,
                        })
                    st.dataframe(pd.DataFrame(filas), use_container_width=True, hide_index=True)
                else:
                    st.caption("Sin registros este mes.")

                if activa:
                    if st.button(f"🔴 Cerrar obra", key=f"cerrar_{obra['id']}"):
                        db.close_obra(obra["name"])
                        st.success(f"Obra cerrada: {obra['name']}")
                        st.rerun()

        if obras_activas:
            st.subheader("Obras activas")
            for obra in obras_activas:
                _render_obra(obra, activa=True)

        if obras_cerradas:
            st.subheader("Obras cerradas este mes")
            for obra in obras_cerradas:
                _render_obra(obra, activa=False)

    st.divider()
    st.subheader("Gestión de obras")

    col_add, col_close = st.columns(2)

    with col_add:
        st.markdown("**Agregar obra**")
        if "nueva_obra_input" not in st.session_state:
            st.session_state["nueva_obra_input"] = ""
        nueva_obra = st.text_input(
            "Nombre de la nueva obra",
            placeholder="Ej: C02 - NUEVO CLIENTE",
            key="nueva_obra_input",
        )
        if st.button("➕ Agregar obra"):
            if nueva_obra.strip():
                db.create_obra(nueva_obra.strip())
                st.success(f"Obra agregada: {nueva_obra.strip()}")
                st.session_state["nueva_obra_input"] = ""
                st.rerun()
            else:
                st.warning("Escribí el nombre de la obra.")

    with col_close:
        st.markdown("**Cerrar obra**")
        obras_act = db.list_obras(active_only=True)
        if obras_act:
            obra_cerrar = st.selectbox(
                "Seleccioná la obra a cerrar",
                options=[o["name"] for o in obras_act],
                key="selectbox_cerrar"
            )
            if st.button("🔴 Cerrar obra"):
                cerradas = db.close_obra(obra_cerrar)
                if cerradas:
                    st.success(f"Obra cerrada: {obra_cerrar}")
                    st.rerun()
        else:
            st.info("No hay obras activas para cerrar.")

# ── TAB 5: Normas operativas ──────────────────────────────────────────────────
with tab_normas:
    st.subheader("Horario de trabajo")
    st.dataframe(pd.DataFrame([
        {"Día": "Lunes a viernes", "Franja": "07:00 – 16:00",                "Tipo": "Horas normales"},
        {"Día": "Lunes a viernes", "Franja": "Antes 07:00 / Después 16:00",  "Tipo": "Extra 50%"},
        {"Día": "Sábado",          "Franja": "07:00 – 13:00",                "Tipo": "Extra 50%"},
        {"Día": "Sábado",          "Franja": "13:00 en adelante",            "Tipo": "Extra 100%"},
        {"Día": "Domingo",         "Franja": "Todo el día",                  "Tipo": "Extra 100%"},
        {"Día": "Feriados",        "Franja": "Todo el día",                  "Tipo": "Extra 100%"},
    ]), use_container_width=True, hide_index=True)

    st.divider()
    st.subheader("Límites mensuales de horas extra — Empleados")
    col1, col2 = st.columns(2)
    col1.metric("Máximo extra 50%",  "20 horas / mes")
    col2.metric("Máximo extra 100%", "10 horas / mes")
    st.caption("Cuando un empleado alcanza estos límites, el sistema envía una alerta automática al administrador por Telegram.")

    st.divider()
    st.subheader("Tolerancias de fichaje")
    st.dataframe(pd.DataFrame([
        {"Situación": "Llegada hasta 15 min tarde",       "Resultado": "Se cuenta como entrada a horario"},
        {"Situación": "Salida hasta 15 min antes",        "Resultado": "Se cuenta como salida a horario"},
        {"Situación": "Sábado: salida hasta 10:53",       "Resultado": "Se cuenta como salida a las 11:00"},
        {"Situación": "Llegada más de 15 min tarde",      "Resultado": "Se descuentan los minutos de atraso"},
    ]), use_container_width=True, hide_index=True)

    st.divider()
    st.subheader("Reportes automáticos")
    st.dataframe(pd.DataFrame([
        {"Cuándo": "Día 15 del mes a las 18:00",    "Qué": "Reporte 1ra quincena (días 1–15) por Telegram y email"},
        {"Cuándo": "Último día del mes a las 18:00","Qué": "Reporte 2da quincena (días 16–fin) por Telegram y email"},
        {"Cuándo": "Todos los días a las 23:00",    "Qué": "Backup automático de la base de datos (PC + OneDrive + Google Drive)"},
        {"Cuándo": "09:00 a 12:00 (días hábiles)",  "Qué": "Alerta si algún empleado no registró entrada"},
        {"Cuándo": "18:30 (días hábiles)",          "Qué": "Recordatorio a empleados que no registraron salida"},
    ]), use_container_width=True, hide_index=True)

    st.divider()
    st.subheader("Cómo fichar — galpón")
    col1, col2 = st.columns(2)
    with col1:
        st.markdown("**Entrada al galpón** — escribir o audio:")
        st.code("llegué · llegue · llegamos · ya llegué\npresente · acá estoy · ya estoy · estoy\narranqué · arranco · arrancamos · empecé\nentrada · entro · entré · entre\nbuen día · buenos días · buenas · check in")
    with col2:
        st.markdown("**Salida del galpón** — escribir o audio:")
        st.code("me voy · nos vamos · salgo · salí · salimos\nchau · hasta mañana · hasta luego · nos vemos\nterminé · terminamos · listo · ya está\nfin · finalizamos · me retiro · check out")
    st.caption("El bot de Telegram acepta texto y mensajes de voz. Funcionan igual.")

    st.divider()
    st.subheader("Cómo registrar horas en obra")
    col1, col2 = st.columns(2)
    with col1:
        st.markdown("**Llegada a la obra** — escribir o audio:")
        st.code("en obra · llegué a la obra · llegué obra\nestoy en obra · arranco obra · arranqué obra\nentro a la obra · empiezo obra\nya llegué a la obra")
        st.caption("El bot muestra los botones de obras activas para seleccionar.")
    with col2:
        st.markdown("**Salida de la obra** — escribir o audio:")
        st.code("salgo de obra · salgo de la obra\nme voy de obra · fin de obra\nterminé obra · listo en obra\nsaliendo de la obra")
    st.info("El registro de obra es independiente del fichaje del galpón. Un empleado puede hacer ambos el mismo día.")

    st.divider()
    st.subheader("Geolocalización — solo Empleados")
    st.dataframe(pd.DataFrame([
        {"Quiénes": "👷 Demetrio, Gustavo, Matias", "Requisito": "Deben compartir ubicación para fichar entrada y salida"},
        {"Quiénes": "💼 Administración / ⚙️ Dirección Técnica", "Requisito": "Fichan directamente con texto o voz, sin ubicación"},
    ]), use_container_width=True, hide_index=True)
    st.caption("Cómo compartir la ubicación en Telegram: tocá el clip 📎 → Ubicación → Enviar mi ubicación actual.")

    st.divider()
    st.subheader("¿Olvidaste fichar o hay un error?")
    st.dataframe(pd.DataFrame([
        {"Situación": "Olvidé registrar entrada",      "Solución": "Avisarle al administrador → /corregir Nombre entrada HH:MM DD/MM"},
        {"Situación": "Olvidé registrar salida",       "Solución": "Avisarle al administrador → /corregir Nombre salida HH:MM DD/MM"},
        {"Situación": "Registré hora incorrecta",      "Solución": "Avisarle al administrador → /corregir Nombre entrada/salida HH:MM DD/MM"},
        {"Situación": "Quiero ver mis últimos fichajes","Solución": "Escribir /estado en el bot"},
    ]), use_container_width=True, hide_index=True)
