"""
Generación de reportes PDF para 2H Movimiento de Suelos.
"""
import io
from datetime import date, datetime
from fpdf import FPDF
import pytz

TIMEZONE = pytz.timezone("America/Argentina/Buenos_Aires")

MESES_ES = {
    1:"Enero",2:"Febrero",3:"Marzo",4:"Abril",5:"Mayo",6:"Junio",
    7:"Julio",8:"Agosto",9:"Septiembre",10:"Octubre",11:"Noviembre",12:"Diciembre",
}

CATEGORIAS_LABEL = {
    "empleado":          "Empleados",
    "administracion":    "Administración",
    "direccion_tecnica": "Dirección Técnica",
}


class PDF(FPDF):
    def header(self):
        self.set_font("Helvetica", "B", 14)
        self.set_text_color(24, 95, 165)
        self.cell(0, 8, "2H Movimiento de Suelos", ln=True, align="C")
        self.set_font("Helvetica", "", 9)
        self.set_text_color(100, 100, 100)
        self.cell(0, 5, "Panel de Control — Reporte de Asistencia", ln=True, align="C")
        self.ln(3)
        self.set_draw_color(24, 95, 165)
        self.set_line_width(0.5)
        self.line(10, self.get_y(), 200, self.get_y())
        self.ln(4)

    def footer(self):
        self.set_y(-12)
        self.set_font("Helvetica", "I", 8)
        self.set_text_color(150, 150, 150)
        now = datetime.now(TIMEZONE).strftime("%d/%m/%Y %H:%M")
        self.cell(0, 5, f"Generado el {now}  —  Pág. {self.page_no()}", align="C")

    def section_title(self, title: str):
        self.set_font("Helvetica", "B", 10)
        self.set_fill_color(24, 95, 165)
        self.set_text_color(255, 255, 255)
        self.cell(0, 7, f"  {title}", ln=True, fill=True)
        self.set_text_color(0, 0, 0)
        self.ln(1)

    def table_header(self, cols: list, widths: list):
        self.set_font("Helvetica", "B", 8)
        self.set_fill_color(220, 230, 245)
        self.set_text_color(0, 0, 0)
        for col, w in zip(cols, widths):
            self.cell(w, 6, col, border=1, fill=True, align="C")
        self.ln()

    def table_row(self, values: list, widths: list, aligns: list = None, shade: bool = False):
        self.set_font("Helvetica", "", 8)
        if shade:
            self.set_fill_color(245, 248, 252)
        else:
            self.set_fill_color(255, 255, 255)
        aligns = aligns or ["L"] * len(values)
        for val, w, aln in zip(values, widths, aligns):
            self.cell(w, 5, str(val), border=1, fill=True, align=aln)
        self.ln()


def build_pdf(db, start: date, end: date, titulo: str) -> io.BytesIO:
    pdf = PDF()
    pdf.set_auto_page_break(auto=True, margin=15)
    pdf.add_page()

    periodo = f"{start.strftime('%d/%m/%Y')} al {end.strftime('%d/%m/%Y')}"
    pdf.set_font("Helvetica", "B", 11)
    pdf.set_text_color(50, 50, 50)
    pdf.cell(0, 7, titulo, ln=True, align="C")
    pdf.set_font("Helvetica", "", 9)
    pdf.set_text_color(100, 100, 100)
    pdf.cell(0, 5, f"Período: {periodo}", ln=True, align="C")
    pdf.ln(5)

    # ── Resumen de asistencia por empleado ───────────────────────────────────
    records   = db.get_records_by_period(start, end)
    holidays  = db.get_holidays(start, end)
    empleados = db.list_employees()

    from reports import calcular_horas

    resumen = {}
    for emp in empleados:
        resumen[emp["telegram_id"]] = {
            "name": emp["name"],
            "cat":  emp.get("categoria", "empleado"),
            "norm": 0.0, "e50": 0.0, "e100": 0.0,
        }

    for rec in records:
        tid = rec.get("telegram_id")
        if tid not in resumen:
            continue
        sessions = rec.get("sessions", [rec])
        for s in sessions:
            if not (s.get("entry_time") and s.get("exit_time")):
                continue
            is_hday = bool(holidays.get(rec["date"]))
            weekday = date.fromisoformat(rec["date"]).weekday()
            n, e50, e100 = calcular_horas(s["entry_time"], s["exit_time"], weekday, is_hday)
            resumen[tid]["norm"] += n
            resumen[tid]["e50"]  += e50
            resumen[tid]["e100"] += e100

    for cat_key, cat_label in CATEGORIAS_LABEL.items():
        grupo = [v for v in resumen.values() if v["cat"] == cat_key]
        if not grupo:
            continue

        pdf.section_title(cat_label)
        cols   = ["Empleado", "Hs. Normales", "Extra 50%", "Extra 100%", "Total"]
        widths = [80, 28, 28, 28, 26]
        aligns = ["L", "R", "R", "R", "R"]
        pdf.table_header(cols, widths)

        tot_n = tot_50 = tot_100 = 0.0
        for i, emp in enumerate(sorted(grupo, key=lambda x: x["name"])):
            total = emp["norm"] + emp["e50"] + emp["e100"]
            pdf.table_row([
                emp["name"],
                f"{emp['norm']:.2f}",
                f"{emp['e50']:.2f}",
                f"{emp['e100']:.2f}",
                f"{total:.2f}",
            ], widths, aligns, shade=(i % 2 == 0))
            tot_n   += emp["norm"]
            tot_50  += emp["e50"]
            tot_100 += emp["e100"]

        # Fila de totales
        pdf.set_font("Helvetica", "B", 8)
        pdf.set_fill_color(24, 95, 165)
        pdf.set_text_color(255, 255, 255)
        total_gral = tot_n + tot_50 + tot_100
        for val, w, aln in zip(
            ["TOTAL", f"{tot_n:.2f}", f"{tot_50:.2f}", f"{tot_100:.2f}", f"{total_gral:.2f}"],
            widths, aligns
        ):
            pdf.cell(w, 6, val, border=1, fill=True, align=aln)
        pdf.ln()
        pdf.set_text_color(0, 0, 0)
        pdf.ln(4)

    # ── Resumen de obras ─────────────────────────────────────────────────────
    obras = db.list_obras(active_only=False)
    obra_data = []
    for obra in obras:
        rows = db.get_obra_hours(obra["id"], start, end)
        if rows:
            total_hs = sum(r.get("total_hours") or 0 for r in rows)
            obra_data.append((obra, rows, total_hs))

    if obra_data:
        if pdf.get_y() > 230:
            pdf.add_page()
        pdf.section_title("Horas por obra")
        cols   = ["Obra", "Empleado", "Fecha", "Entrada", "Salida", "Horas"]
        widths = [55, 40, 20, 18, 18, 18]
        aligns = ["L", "L", "C", "C", "C", "R"]
        pdf.table_header(cols, widths)

        shade = False
        for obra, rows, total_hs in obra_data:
            for r in rows:
                ent = r["entry_time"].strftime("%H:%M") if r.get("entry_time") else "–"
                sal = r["exit_time"].strftime("%H:%M")  if r.get("exit_time")  else "Abierta"
                hs  = f"{r['total_hours']:.2f}"         if r.get("total_hours") else "–"
                pdf.table_row([obra["name"], r["name"], r["date"], ent, sal, hs],
                              widths, aligns, shade=shade)
                shade = not shade

            # Subtotal por obra
            pdf.set_font("Helvetica", "B", 8)
            pdf.set_fill_color(220, 230, 245)
            pdf.set_text_color(0, 0, 80)
            for val, w, aln in zip(
                [f"  Total {obra['name']}", "", "", "", "", f"{total_hs:.2f}"],
                widths, aligns
            ):
                pdf.cell(w, 5, val, border=1, fill=True, align=aln)
            pdf.ln()
            pdf.set_text_color(0, 0, 0)

    buf = io.BytesIO()
    pdf.output(buf)
    buf.seek(0)
    return buf
