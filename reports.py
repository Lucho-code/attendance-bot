"""
Generación de reportes XLSX y envío de email.
Compartido entre bot.py y admin_panel.py.
"""
import io
import os
import smtplib
from datetime import date, datetime, timedelta
from itertools import groupby
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email import encoders

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
import pytz

TIMEZONE = pytz.timezone("America/Argentina/Buenos_Aires")

DIAS_ES = ["Lunes", "Martes", "Miércoles", "Jueves", "Viernes", "Sábado", "Domingo"]

MESES_ES = {
    1: "Enero", 2: "Febrero", 3: "Marzo", 4: "Abril",
    5: "Mayo", 6: "Junio", 7: "Julio", 8: "Agosto",
    9: "Septiembre", 10: "Octubre", 11: "Noviembre", 12: "Diciembre",
}

ABSENCE_TYPES = {
    "vacacion": "Vacación", "vacación": "Vacación",
    "enfermedad": "Enfermedad",
    "licencia": "Licencia",
    "justificada": "Justificada",
    "injustificada": "Injustificada",
}

_COLOR = {
    "weekend":       "D9D9D9",
    "holiday":       "FFE699",
    "vacacion":      "BDD7EE",
    "enfermedad":    "C6EFCE",
    "licencia":      "FCE4D6",
    "justificada":   "E2EFDA",
    "injustificada": "FFC7CE",
    "no_exit":       "FFEB9C",
    "missing":       "FFC7CE",
    "overtime":      "FFD966",
}


def _row_fill(key: str) -> PatternFill:
    return PatternFill("solid", fgColor=_COLOR.get(key, "FFFFFF"))


def _sheet_name(name: str) -> str:
    for ch in r"/\?*[]:":
        name = name.replace(ch, "")
    return name[:31]


def _scheduled_hours(shift: tuple) -> float:
    """Horas de jornada según turno (eh, em, sh, sm)."""
    eh, em, sh, sm = shift
    return (sh * 60 + sm - eh * 60 - em) / 60


def build_xlsx(db, records, titulo: str,
               start_date: date = None, end_date: date = None,
               holidays: dict = None, absences: dict = None) -> io.BytesIO:
    """
    Genera el XLSX de asistencia.
    - Si start_date/end_date están presentes: modo calendario completo con colores.
    - Si no: modo simple (solo registros existentes).
    Incluye columna de horas extra basada en el turno de cada empleado.
    """
    holidays = holidays or {}
    absences = absences or {}
    cal_mode = start_date is not None and end_date is not None
    today    = datetime.now(TIMEZONE).date()

    if cal_mode:
        headers   = ["Fecha", "Día", "Estado", "Entrada", "Salida",
                     "Horas Trab.", "Hs. Extra"]
        col_entry = 4; col_exit = 5; col_hours = 6; col_extra = 7; ncols = 7
    else:
        headers   = ["Fecha", "Día", "Entrada", "Salida",
                     "Horas Trab.", "Hs. Extra"]
        col_entry = 3; col_exit = 4; col_hours = 5; col_extra = 6; ncols = 6

    hdr_fill  = PatternFill("solid", fgColor="1F4E79")
    hdr_font  = Font(bold=True, color="FFFFFF")
    tot_fill  = PatternFill("solid", fgColor="1F4E79")
    tot_font  = Font(bold=True, color="FFFFFF")
    name_fill = PatternFill("solid", fgColor="2E75B6")
    name_font = Font(bold=True, color="FFFFFF", size=12)

    def style_hdr(ws, row):
        for c in range(1, ncols + 1):
            cell = ws.cell(row=row, column=c)
            cell.fill = hdr_fill; cell.font = hdr_font
            cell.alignment = Alignment(horizontal="center")

    def style_tot(ws, row):
        for c in range(1, ncols + 1):
            ws.cell(row=row, column=c).fill = tot_fill
            ws.cell(row=row, column=c).font = tot_font

    def color_row(ws, row, key):
        fill = _row_fill(key)
        for c in range(1, ncols + 1):
            ws.cell(row=row, column=c).fill = fill

    wb = openpyxl.Workbook()

    # ── Hoja Resumen ──────────────────────────────────────────────────────────
    ws_res = wb.active
    ws_res.title = "Resumen"
    ws_res.merge_cells("A1:C1")
    t = ws_res.cell(1, 1, titulo)
    t.font = Font(bold=True, size=13)
    t.alignment = Alignment(horizontal="center")

    for col, h in enumerate(["Empleado", "Total Horas", "Hs. Extra"], start=1):
        c = ws_res.cell(2, col, h)
        c.fill = hdr_fill; c.font = hdr_font
        c.alignment = Alignment(horizontal="center")

    ws_res.column_dimensions["A"].width = 26
    ws_res.column_dimensions["B"].width = 14
    ws_res.column_dimensions["C"].width = 12

    records_sorted = sorted(records, key=lambda r: (r["name"], r["date"]))
    grouped = {n: list(g)
               for n, g in groupby(records_sorted, key=lambda r: r["name"])}

    if cal_mode:
        for e in db.list_employees():
            if e["name"] not in grouped:
                grouped[e["name"]] = []

    summary_row  = 3
    summary_refs = {}
    detail_info  = {}   # name -> (sheet_name, total_row, extra_row)

    for name in sorted(grouped):
        ws_res.cell(summary_row, 1, name)
        summary_refs[name] = summary_row
        summary_row += 1

    grand_row = summary_row
    style_tot(ws_res, grand_row)
    for c in range(1, 4):
        ws_res.cell(grand_row, c).fill = tot_fill
    ws_res.cell(grand_row, 1, "TOTAL GENERAL").font = tot_font
    ws_res.cell(grand_row, 1).fill = tot_fill

    # Leyenda
    leg_row = grand_row + 2
    ws_res.cell(leg_row, 1, "Leyenda:").font = Font(bold=True)
    leyenda = [
        ("Fin de semana", "weekend"), ("Feriado", "holiday"),
        ("Vacación", "vacacion"),     ("Enfermedad", "enfermedad"),
        ("Licencia", "licencia"),     ("Sin salida", "no_exit"),
        ("Ausente", "missing"),       ("Horas extra", "overtime"),
    ]
    for i, (label, key) in enumerate(leyenda):
        r = leg_row + 1 + i
        ws_res.cell(r, 1, label).fill = _row_fill(key)

    # ── Una hoja por empleado ─────────────────────────────────────────────────
    for name, emp_records in sorted(grouped.items()):
        sname = _sheet_name(name)
        ws    = wb.create_sheet(title=sname)

        by_date  = {r["date"]: r for r in emp_records}
        emp_list = db.find_employee_by_name(name)
        emp_tid  = emp_list[0]["telegram_id"] if emp_list else None
        emp_abs  = absences.get(emp_tid, {}) if emp_tid else {}
        shift    = db.get_shift(emp_tid) if emp_tid else (9, 0, 18, 0)
        sched_h  = _scheduled_hours(shift)

        ws.merge_cells(f"A1:{chr(64+ncols)}1")
        nc = ws.cell(1, 1, name)
        nc.fill = name_fill; nc.font = name_font
        nc.alignment = Alignment(horizontal="center")

        for col, h in enumerate(headers, start=1):
            ws.cell(2, col, h)
        style_hdr(ws, 2)

        data_start = 3
        cur_row    = data_start

        if cal_mode:
            current_day = start_date
            while current_day <= end_date:
                d_str   = current_day.isoformat()
                weekday = current_day.weekday()
                is_we   = weekday >= 5
                holiday = holidays.get(d_str)
                absence = emp_abs.get(d_str)
                record  = by_date.get(d_str)

                ws.cell(cur_row, 1, current_day.strftime("%d/%m/%Y"))
                ws.cell(cur_row, 2, DIAS_ES[weekday])

                if is_we:
                    ws.cell(cur_row, 3, "Fin de semana")
                    color_row(ws, cur_row, "weekend")

                elif holiday:
                    ws.cell(cur_row, 3, f"Feriado: {holiday}")
                    color_row(ws, cur_row, "holiday")

                elif absence:
                    ws.cell(cur_row, 3, ABSENCE_TYPES.get(absence, absence.capitalize()))
                    color_row(ws, cur_row, absence)

                elif record and record.get("entry_time") and record.get("exit_time"):
                    worked = record["total_hours"] or 0
                    extra  = max(0.0, round(worked - sched_h, 2))
                    ws.cell(cur_row, 3, "Trabajó")
                    c = ws.cell(cur_row, col_entry,
                                record["entry_time"].replace(tzinfo=None).time())
                    c.number_format = "HH:MM"
                    c = ws.cell(cur_row, col_exit,
                                record["exit_time"].replace(tzinfo=None).time())
                    c.number_format = "HH:MM"
                    c = ws.cell(cur_row, col_hours, round(worked, 2))
                    c.number_format = "0.00"
                    if extra > 0:
                        c = ws.cell(cur_row, col_extra, extra)
                        c.number_format = "0.00"
                        c.alignment = Alignment(horizontal="center")
                        color_row(ws, cur_row, "overtime")

                elif record and record.get("entry_time"):
                    ws.cell(cur_row, 3, "Sin salida")
                    c = ws.cell(cur_row, col_entry,
                                record["entry_time"].replace(tzinfo=None).time())
                    c.number_format = "HH:MM"
                    color_row(ws, cur_row, "no_exit")

                elif current_day <= today:
                    ws.cell(cur_row, 3, "Ausente")
                    color_row(ws, cur_row, "missing")

                cur_row     += 1
                current_day += timedelta(days=1)

        else:
            for r in emp_records:
                d = date.fromisoformat(r["date"])
                ws.cell(cur_row, 1, r["date"])
                ws.cell(cur_row, 2, DIAS_ES[d.weekday()])
                if r.get("entry_time"):
                    c = ws.cell(cur_row, col_entry,
                                r["entry_time"].replace(tzinfo=None).time())
                    c.number_format = "HH:MM"
                if r.get("exit_time"):
                    c = ws.cell(cur_row, col_exit,
                                r["exit_time"].replace(tzinfo=None).time())
                    c.number_format = "HH:MM"
                worked = r.get("total_hours") or 0
                extra  = max(0.0, round(worked - sched_h, 2))
                if worked:
                    c = ws.cell(cur_row, col_hours, round(worked, 2))
                    c.number_format = "0.00"
                if extra > 0:
                    c = ws.cell(cur_row, col_extra, extra)
                    c.number_format = "0.00"
                    color_row(ws, cur_row, "overtime")
                elif not r.get("exit_time"):
                    color_row(ws, cur_row, "no_exit")
                cur_row += 1

        data_end  = cur_row - 1
        total_row = cur_row
        style_tot(ws, total_row)

        lbl_col = col_hours - 1
        lbl = ws.cell(total_row, lbl_col, "TOTAL")
        lbl.font = tot_font; lbl.fill = tot_fill
        lbl.alignment = Alignment(horizontal="right")

        hc = ws.cell(total_row, col_hours,
                     f"=SUM({chr(64+col_hours)}{data_start}:{chr(64+col_hours)}{data_end})")
        hc.number_format = "0.00"; hc.font = tot_font
        hc.fill = tot_fill; hc.alignment = Alignment(horizontal="center")

        ec = ws.cell(total_row, col_extra,
                     f"=SUM({chr(64+col_extra)}{data_start}:{chr(64+col_extra)}{data_end})")
        ec.number_format = "0.00"; ec.font = tot_font
        ec.fill = tot_fill; ec.alignment = Alignment(horizontal="center")

        detail_info[name] = (sname, total_row)

        ws.column_dimensions["A"].width = 14
        ws.column_dimensions["B"].width = 12
        if cal_mode:
            ws.column_dimensions["C"].width = 20
            ws.column_dimensions["D"].width = 10
            ws.column_dimensions["E"].width = 10
            ws.column_dimensions["F"].width = 13
            ws.column_dimensions["G"].width = 12
        else:
            ws.column_dimensions["C"].width = 10
            ws.column_dimensions["D"].width = 10
            ws.column_dimensions["E"].width = 13
            ws.column_dimensions["F"].width = 12

    # ── Completar Resumen ─────────────────────────────────────────────────────
    for name, srow in summary_refs.items():
        if name not in detail_info:
            continue
        sname, trow = detail_info[name]
        hcol = chr(64 + col_hours)
        ecol = chr(64 + col_extra)
        c = ws_res.cell(srow, 2, f"='{sname}'!{hcol}{trow}")
        c.number_format = "0.00"; c.alignment = Alignment(horizontal="center")
        c = ws_res.cell(srow, 3, f"='{sname}'!{ecol}{trow}")
        c.number_format = "0.00"; c.alignment = Alignment(horizontal="center")

    if summary_refs:
        h_refs = ",".join(f"B{r}" for r in summary_refs.values())
        e_refs = ",".join(f"C{r}" for r in summary_refs.values())
        gc = ws_res.cell(grand_row, 2, f"=SUM({h_refs})")
        gc.number_format = "0.00"; gc.font = tot_font
        gc.fill = tot_fill; gc.alignment = Alignment(horizontal="center")
        ge = ws_res.cell(grand_row, 3, f"=SUM({e_refs})")
        ge.number_format = "0.00"; ge.font = tot_font
        ge.fill = tot_fill; ge.alignment = Alignment(horizontal="center")
    else:
        for c in (2, 3):
            ws_res.cell(grand_row, c, 0).font = tot_font
            ws_res.cell(grand_row, c).fill = tot_fill

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf


def send_email(buf: io.BytesIO, filename: str, subject: str) -> bool:
    email_from = os.getenv("EMAIL_FROM")
    email_pass = os.getenv("EMAIL_PASSWORD")
    email_to   = os.getenv("EMAIL_TO")
    if not all([email_from, email_pass, email_to]):
        return False
    msg = MIMEMultipart()
    msg["From"] = email_from
    msg["To"]   = email_to
    msg["Subject"] = subject
    msg.attach(MIMEText(
        f"Adjunto el reporte de asistencia: {filename}\n\n"
        "Enviado automáticamente por el sistema.",
        "plain", "utf-8",
    ))
    part = MIMEBase("application", "octet-stream")
    part.set_payload(buf.read())
    encoders.encode_base64(part)
    part.add_header("Content-Disposition", f'attachment; filename="{filename}"')
    msg.attach(part)
    with smtplib.SMTP_SSL("smtp.gmail.com", 465) as s:
        s.login(email_from, email_pass)
        s.sendmail(email_from, email_to.split(","), msg.as_string())
    return True
