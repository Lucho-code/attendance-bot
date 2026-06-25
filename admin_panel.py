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

from database import Database
from reports  import build_xlsx, MESES_ES, ABSENCE_TYPES

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

st.title("Panel de Asistencia")
st.caption(f"Actualizado: {ahora().strftime('%d/%m/%Y  %H:%M')}")

tab_hoy, tab_emp, tab_rep = st.tabs(["📅 Hoy", "👥 Empleados", "📊 Reportes"])

# ── TAB 1: Estado de hoy ──────────────────────────────────────────────────────
with tab_hoy:
    hoy       = ahora().date()
    empleados = db.list_employees()

    rows = []
    presentes = ausentes = sin_salida = 0

    for emp in empleados:
        st_hoy = db.get_today_status(emp["telegram_id"], hoy)
        shift  = db.get_shift(emp["telegram_id"])
        turno  = f"{shift[0]:02d}:{shift[1]:02d} – {shift[2]:02d}:{shift[3]:02d}"

        if st_hoy and st_hoy.get("entry_time") and not st_hoy.get("exit_time"):
            estado    = "🟢 Dentro"
            presentes += 1
        elif st_hoy and st_hoy.get("exit_time"):
            estado    = "🔵 Salió"
            sin_salida = sin_salida  # no change
        elif not st_hoy:
            estado   = "🔴 Sin registro"
            ausentes += 1
        else:
            estado = "🟡 Sin salida"
            sin_salida += 1

        ent = st_hoy["entry_time"].strftime("%H:%M") if st_hoy and st_hoy.get("entry_time") else "–"
        sal = st_hoy["exit_time"].strftime("%H:%M")  if st_hoy and st_hoy.get("exit_time")  else "–"
        hs  = f"{st_hoy['total_hours']:.1f} hs"     if st_hoy and st_hoy.get("total_hours") else "–"

        rows.append({
            "Empleado": emp["name"],
            "Turno":    turno,
            "Estado":   estado,
            "Entrada":  ent,
            "Salida":   sal,
            "Horas":    hs,
        })

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Total empleados",  len(empleados))
    c2.metric("🟢 Dentro ahora",  presentes)
    c3.metric("🔴 Sin registrar", ausentes)
    c4.metric("🟡 Sin salida",    sin_salida)

    st.divider()
    st.dataframe(
        pd.DataFrame(rows),
        use_container_width=True,
        hide_index=True,
        column_config={
            "Estado": st.column_config.TextColumn(width="medium"),
        },
    )

    if st.button("🔄 Actualizar"):
        st.cache_resource.clear()
        st.rerun()

# ── TAB 2: Empleados ──────────────────────────────────────────────────────────
with tab_emp:
    empleados = db.list_employees()
    if not empleados:
        st.info("No hay empleados registrados aún.")
    else:
        rows_emp = []
        for emp in empleados:
            shift = db.get_shift(emp["telegram_id"])
            rows_emp.append({
                "Empleado":    emp["name"],
                "Turno entr.": f"{shift[0]:02d}:{shift[1]:02d}",
                "Turno sal.":  f"{shift[2]:02d}:{shift[3]:02d}",
                "ID Telegram": emp["telegram_id"],
                "Registrado":  emp["registered_at"][:10],
            })
        st.dataframe(
            pd.DataFrame(rows_emp),
            use_container_width=True,
            hide_index=True,
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
            if records:
                rows_rec = []
                for r in records:
                    ent = r["entry_time"].strftime("%H:%M") if r.get("entry_time") else "–"
                    sal = r["exit_time"].strftime("%H:%M")  if r.get("exit_time")  else "Sin salida"
                    hs  = f"{r['total_hours']:.2f}"         if r.get("total_hours") else "–"
                    rows_rec.append({"Fecha": r["date"], "Entrada": ent,
                                     "Salida": sal, "Horas": hs})
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

    if st.button("📥 Generar y descargar"):
        with st.spinner("Generando reporte..."):
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
        st.success(f"Reporte generado: {len(records)} registros")

    st.divider()
    st.subheader("Estadísticas rápidas del mes")

    hoy   = ahora().date()
    start = hoy.replace(day=1)
    recs  = db.get_records_by_period(start, hoy)

    if recs:
        df = pd.DataFrame([{
            "Empleado": r["name"],
            "Horas":    round(r["total_hours"] or 0, 2),
        } for r in recs if r.get("total_hours")])

        if not df.empty:
            resumen = df.groupby("Empleado")["Horas"].sum().reset_index()
            resumen.columns = ["Empleado", "Total Horas del Mes"]
            st.dataframe(resumen, use_container_width=True, hide_index=True)
    else:
        st.info("Sin registros para el mes actual.")
