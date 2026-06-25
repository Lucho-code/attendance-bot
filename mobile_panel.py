"""
Panel móvil de FichaYA — accesible desde el celular via Tailscale.
Corre en el puerto 5050.
"""
import os
from datetime import datetime
from flask import Flask, jsonify
import pytz
from dotenv import load_dotenv

load_dotenv()

from database import Database
from reports  import calcular_horas

TIMEZONE = pytz.timezone("America/Argentina/Buenos_Aires")
app      = Flask(__name__)
db       = Database()


def ahora():
    return datetime.now(TIMEZONE)


def get_estado_hoy():
    hoy      = ahora().date()
    emps     = db.list_employees()
    is_hday  = bool(db.is_holiday(hoy))
    rows     = []

    for emp in emps:
        st   = db.get_today_status(emp["telegram_id"], hoy)
        shift = db.get_shift(emp["telegram_id"])

        if st and st.get("entry_time") and not st.get("exit_time"):
            estado = "dentro"
            badge  = "🟢"
        elif st and st.get("exit_time"):
            estado = "salio"
            badge  = "🔵"
        elif not st:
            estado = "ausente"
            badge  = "🔴"
        else:
            estado = "sin_salida"
            badge  = "🟡"

        ent  = st["entry_time"].strftime("%H:%M") if st and st.get("entry_time") else "—"
        sal  = st["exit_time"].strftime("%H:%M")  if st and st.get("exit_time")  else "—"
        norm = e50 = e100 = "—"

        if st and st.get("entry_time") and st.get("exit_time"):
            n, ex50, ex100 = calcular_horas(
                st["entry_time"], st["exit_time"], hoy.weekday(), is_hday)
            norm = f"{n:.2f}"    if n    > 0 else "—"
            e50  = f"{ex50:.2f}" if ex50 > 0 else "—"
            e100 = f"{ex100:.2f}"if ex100> 0 else "—"

        rows.append({
            "nombre": emp["name"],
            "estado": estado,
            "badge":  badge,
            "entrada": ent,
            "salida":  sal,
            "norm":    norm,
            "e50":     e50,
            "e100":    e100,
        })

    presentes  = sum(1 for r in rows if r["estado"] == "dentro")
    ausentes   = sum(1 for r in rows if r["estado"] == "ausente")
    sin_salida = sum(1 for r in rows if r["estado"] == "sin_salida")
    salieron   = sum(1 for r in rows if r["estado"] == "salio")

    return {
        "hora":       ahora().strftime("%H:%M"),
        "fecha":      ahora().strftime("%A %d/%m/%Y").capitalize(),
        "total":      len(emps),
        "presentes":  presentes,
        "ausentes":   ausentes,
        "sin_salida": sin_salida,
        "salieron":   salieron,
        "empleados":  rows,
        "es_feriado": is_hday,
    }


@app.route("/api/status")
def api_status():
    return jsonify(get_estado_hoy())


@app.route("/")
def index():
    html = """<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1, maximum-scale=1">
<meta name="theme-color" content="#0B1F38">
<meta name="apple-mobile-web-app-capable" content="yes">
<title>FichaYA · Panel</title>
<style>
  * { box-sizing: border-box; margin: 0; padding: 0; }
  body { font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', sans-serif;
         background: #0B1F38; color: #fff; min-height: 100vh; }

  header {
    background: #0B1F38; padding: 16px 20px 12px;
    display: flex; align-items: center; justify-content: space-between;
    position: sticky; top: 0; z-index: 10; border-bottom: 1px solid rgba(255,255,255,.1);
  }
  .logo { display: flex; align-items: center; gap: 10px; }
  .logo-icon { width: 32px; height: 32px; }
  .logo-text { font-size: 18px; font-weight: 300; }
  .logo-text strong { font-weight: 800; color: #85B7EB; }
  .hora { font-size: 13px; color: #7CAED4; text-align: right; line-height: 1.3; }

  .metrics {
    display: grid; grid-template-columns: 1fr 1fr;
    gap: 10px; padding: 16px;
  }
  .metric {
    background: rgba(255,255,255,.07); border-radius: 14px;
    padding: 16px; text-align: center;
  }
  .metric-num { font-size: 36px; font-weight: 800; line-height: 1; }
  .metric-lbl { font-size: 12px; margin-top: 4px; opacity: .7; }
  .m-verde  .metric-num { color: #5DCAA5; }
  .m-rojo   .metric-num { color: #F09595; }
  .m-azul   .metric-num { color: #85B7EB; }
  .m-amarillo .metric-num { color: #FAC775; }

  .section-title {
    font-size: 11px; font-weight: 600; letter-spacing: .08em;
    color: #7CAED4; text-transform: uppercase;
    padding: 0 16px 8px;
  }

  .emp-list { padding: 0 16px 24px; display: flex; flex-direction: column; gap: 10px; }

  .emp-card {
    background: rgba(255,255,255,.07); border-radius: 14px;
    padding: 14px 16px;
  }
  .emp-card.dentro    { border-left: 3px solid #5DCAA5; }
  .emp-card.salio     { border-left: 3px solid #85B7EB; }
  .emp-card.ausente   { border-left: 3px solid #F09595; }
  .emp-card.sin_salida{ border-left: 3px solid #FAC775; }

  .emp-header { display: flex; align-items: center; justify-content: space-between; margin-bottom: 10px; }
  .emp-nombre { font-size: 15px; font-weight: 600; }
  .emp-badge  { font-size: 20px; }

  .emp-data { display: grid; grid-template-columns: 1fr 1fr 1fr; gap: 6px; }
  .emp-cell { background: rgba(0,0,0,.2); border-radius: 8px; padding: 8px; text-align: center; }
  .emp-cell-lbl { font-size: 10px; color: #7CAED4; margin-bottom: 2px; }
  .emp-cell-val { font-size: 14px; font-weight: 600; }

  .emp-hs { display: grid; grid-template-columns: 1fr 1fr 1fr; gap: 6px; margin-top: 6px; }
  .hs-cell { background: rgba(0,0,0,.15); border-radius: 8px; padding: 6px; text-align: center; }
  .hs-lbl  { font-size: 9px; color: #7CAED4; margin-bottom: 1px; }
  .hs-val  { font-size: 13px; font-weight: 600; }
  .hs-extra50  .hs-val { color: #FAC775; }
  .hs-extra100 .hs-val { color: #F09595; }

  .refresh-bar {
    text-align: center; padding: 12px; font-size: 12px; color: #7CAED4;
  }
  #countdown { font-weight: 600; color: #85B7EB; }

  .feriado-banner {
    background: #EF9F27; color: #412402;
    text-align: center; padding: 8px 16px; font-size: 13px; font-weight: 600;
  }
</style>
</head>
<body>

<header>
  <div class="logo">
    <svg class="logo-icon" viewBox="0 0 80 80">
      <rect width="80" height="80" rx="18" fill="#185FA5"/>
      <path d="M 20 42 L 33 56 L 60 28" stroke="white" stroke-width="7"
            stroke-linecap="round" stroke-linejoin="round" fill="none"/>
    </svg>
    <div class="logo-text">Ficha<strong>YA</strong></div>
  </div>
  <div class="hora" id="hora-fecha">—</div>
</header>

<div id="feriado-banner" style="display:none" class="feriado-banner">Hoy es feriado nacional</div>

<div class="metrics">
  <div class="metric m-verde">
    <div class="metric-num" id="m-presentes">—</div>
    <div class="metric-lbl">Dentro ahora</div>
  </div>
  <div class="metric m-rojo">
    <div class="metric-num" id="m-ausentes">—</div>
    <div class="metric-lbl">Sin registrar</div>
  </div>
  <div class="metric m-azul">
    <div class="metric-num" id="m-salieron">—</div>
    <div class="metric-lbl">Salieron</div>
  </div>
  <div class="metric m-amarillo">
    <div class="metric-num" id="m-sin-salida">—</div>
    <div class="metric-lbl">Sin salida</div>
  </div>
</div>

<div class="section-title">Empleados — 2H Mov. Suelos</div>
<div class="emp-list" id="emp-list"></div>

<div class="refresh-bar">Actualiza en <span id="countdown">30</span>s</div>

<script>
let secs = 30;

function render(data) {
  document.getElementById('hora-fecha').innerHTML =
    `<b>${data.hora}</b><br>${data.fecha}`;
  document.getElementById('m-presentes').textContent  = data.presentes;
  document.getElementById('m-ausentes').textContent   = data.ausentes;
  document.getElementById('m-salieron').textContent   = data.salieron;
  document.getElementById('m-sin-salida').textContent = data.sin_salida;

  const banner = document.getElementById('feriado-banner');
  banner.style.display = data.es_feriado ? 'block' : 'none';

  const list = document.getElementById('emp-list');
  list.innerHTML = data.empleados.map(e => `
    <div class="emp-card ${e.estado}">
      <div class="emp-header">
        <div class="emp-nombre">${e.nombre}</div>
        <div class="emp-badge">${e.badge}</div>
      </div>
      <div class="emp-data">
        <div class="emp-cell">
          <div class="emp-cell-lbl">Entrada</div>
          <div class="emp-cell-val">${e.entrada}</div>
        </div>
        <div class="emp-cell">
          <div class="emp-cell-lbl">Salida</div>
          <div class="emp-cell-val">${e.salida}</div>
        </div>
        <div class="emp-cell">
          <div class="emp-cell-lbl">Hs. Norm.</div>
          <div class="emp-cell-val">${e.norm}</div>
        </div>
      </div>
      ${(e.e50 !== '—' || e.e100 !== '—') ? `
      <div class="emp-hs">
        <div class="hs-cell hs-extra50">
          <div class="hs-lbl">Extra 50%</div>
          <div class="hs-val">${e.e50}</div>
        </div>
        <div class="hs-cell hs-extra100">
          <div class="hs-lbl">Extra 100%</div>
          <div class="hs-val">${e.e100}</div>
        </div>
        <div class="hs-cell"></div>
      </div>` : ''}
    </div>
  `).join('');
}

async function refresh() {
  try {
    const res  = await fetch('/api/status');
    const data = await res.json();
    render(data);
    secs = 30;
  } catch(e) { console.error(e); }
}

setInterval(() => {
  secs--;
  document.getElementById('countdown').textContent = secs;
  if (secs <= 0) refresh();
}, 1000);

refresh();
</script>
</body>
</html>"""
    return html


if __name__ == "__main__":
    print("Panel móvil iniciado en http://0.0.0.0:5050")
    print(f"Desde el celular (misma red): http://100.124.31.88:5050")
    app.run(host="0.0.0.0", port=5050, debug=False)
