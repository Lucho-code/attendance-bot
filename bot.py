import os
import io
import calendar
from datetime import datetime, time as dt_time
from itertools import groupby

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment

import pytz
from dotenv import load_dotenv
from telegram import Update
from telegram.ext import (
    Application,
    CommandHandler,
    MessageHandler,
    filters,
    ContextTypes,
    CallbackContext,
)

from database import Database

load_dotenv()

TOKEN = os.getenv("TELEGRAM_TOKEN")
ADMIN_IDS = [int(x) for x in os.getenv("ADMIN_IDS", "").split(",") if x.strip()]
TIMEZONE = pytz.timezone("America/Argentina/Buenos_Aires")

db = Database()

PALABRAS_ENTRADA = [
    "entro", "entré", "entre", "llegué", "llegue",
    "buenos días", "buenos dias", "buen dia", "buen día",
    "inicio", "empiezo",
]
PALABRAS_SALIDA = [
    "salgo", "me voy", "hasta mañana", "hasta manana",
    "chau", "me retiro", "salida", "termino", "terminé",
]

AWAITING_NAME = "awaiting_name"


def ahora() -> datetime:
    return datetime.now(TIMEZONE)


def es_admin(user_id: int) -> bool:
    return user_id in ADMIN_IDS


async def _pedir_nombre(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Pide el nombre al usuario y activa el flag de espera."""
    context.user_data[AWAITING_NAME] = True
    await update.message.reply_text(
        "Hola! Para registrarte necesito tu nombre completo.\n"
        "¿Cómo te llamás?"
    )


async def _guardar_nombre(update: Update, context: ContextTypes.DEFAULT_TYPE) -> bool:
    """Si hay un nombre pendiente, lo guarda y devuelve True."""
    if not context.user_data.get(AWAITING_NAME):
        return False
    name = update.message.text.strip()
    db.register_employee(update.effective_user.id, name)
    context.user_data[AWAITING_NAME] = False
    await update.message.reply_text(
        f"*Bienvenido/a, {name}!*\n"
        f"Ya estás registrado en el sistema de asistencia.\n"
        f"\n"
        f"*Cómo usarlo:*\n"
        f"\n"
        f"Al llegar, escribí cualquiera de estas frases:\n"
        f"  › llegué\n"
        f"  › entro\n"
        f"  › buenos días\n"
        f"  › /entro\n"
        f"\n"
        f"Al irte, escribí cualquiera de estas:\n"
        f"  › me voy\n"
        f"  › salgo\n"
        f"  › chau\n"
        f"  › /salgo\n"
        f"\n"
        f"Para ver cómo vas hoy:\n"
        f"  › /estado\n"
        f"\n"
        f"_Eso es todo. Sin formularios, sin papel._\n"
        f"_Cualquier duda hablá con tu administrador._",
        parse_mode="Markdown",
    )
    return True


# ---------- handlers ----------

async def cmd_start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    employee = db.get_employee(update.effective_user.id)
    if employee:
        await update.message.reply_text(
            f"Ya estás registrado como *{employee['name']}*\n\n"
            "  /entro  - Registrar entrada\n"
            "  /salgo  - Registrar salida\n"
            "  /estado - Ver tu estado hoy",
            parse_mode="Markdown",
        )
        return
    await _pedir_nombre(update, context)


async def _hacer_entro(update: Update, context: ContextTypes.DEFAULT_TYPE = None):
    user = update.effective_user
    employee = db.get_employee(user.id)
    if not employee:
        if context:
            await _pedir_nombre(update, context)
        else:
            await update.message.reply_text("Primero registrate enviando /start")
        return

    ts = ahora()
    result = db.register_entry(user.id, ts)

    if result == "already_in":
        status = db.get_today_status(user.id, ts.date())
        hora = status["entry_time"].strftime("%H:%M") if status else "?"
        await update.message.reply_text(
            f"Ya registraste entrada hoy a las {hora}."
        )
        return

    await update.message.reply_text(
        f"*Entrada registrada*\n"
        f"Nombre: {employee['name']}\n"
        f"Hora:   {ts.strftime('%H:%M')}\n"
        f"Fecha:  {ts.strftime('%d/%m/%Y')}",
        parse_mode="Markdown",
    )


async def _hacer_salgo(update: Update, context: ContextTypes.DEFAULT_TYPE = None):
    user = update.effective_user
    employee = db.get_employee(user.id)
    if not employee:
        if context:
            await _pedir_nombre(update, context)
        else:
            await update.message.reply_text("Primero registrate enviando /start")
        return

    ts = ahora()
    result = db.register_exit(user.id, ts)

    if result is None:
        await update.message.reply_text(
            "No tenés entrada registrada hoy. Usá /entro primero."
        )
        return

    horas = result["total_hours"]
    await update.message.reply_text(
        f"*Salida registrada*\n"
        f"Nombre:  {employee['name']}\n"
        f"Hora:    {ts.strftime('%H:%M')}\n"
        f"Trabajaste {horas:.1f} hs hoy.",
        parse_mode="Markdown",
    )


async def cmd_entro(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await _hacer_entro(update, context)


async def cmd_salgo(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await _hacer_salgo(update, context)


async def cmd_estado(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user = update.effective_user
    employee = db.get_employee(user.id)
    if not employee:
        await _pedir_nombre(update, context)
        return

    ts = ahora()
    status = db.get_today_status(user.id, ts.date())

    if not status:
        msg = f"*{employee['name']}*\nHoy todavía no registraste entrada."
    elif status["exit_time"] is None:
        msg = (
            f"*{employee['name']}*\n"
            f"Dentro desde las {status['entry_time'].strftime('%H:%M')}"
        )
    else:
        msg = (
            f"*{employee['name']}*\n"
            f"Jornada cerrada - {status['total_hours']:.1f} hs trabajadas\n"
            f"Entrada: {status['entry_time'].strftime('%H:%M')} | "
            f"Salida: {status['exit_time'].strftime('%H:%M')}"
        )

    await update.message.reply_text(msg, parse_mode="Markdown")


async def cmd_reporte(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user = update.effective_user
    if not es_admin(user.id):
        await update.message.reply_text(
            "Solo los administradores pueden generar reportes."
        )
        return

    records = db.get_all_records()

    buffer = _build_xlsx(records, "Reporte completo")
    filename = f"asistencia_{ahora().strftime('%Y%m%d_%H%M')}.xlsx"
    await update.message.reply_document(
        document=buffer,
        filename=filename,
        caption=f"Reporte de asistencia - {len(records)} registros",
    )


async def cmd_empleados(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user = update.effective_user
    if not es_admin(user.id):
        await update.message.reply_text(
            "Solo los administradores pueden ver la lista de empleados."
        )
        return

    empleados = db.list_employees()
    if not empleados:
        await update.message.reply_text("No hay empleados registrados aún.")
        return

    lines = [f"*Empleados registrados ({len(empleados)}):*"]
    for e in empleados:
        lines.append(f"  - {e['name']} (ID: {e['telegram_id']})")

    await update.message.reply_text("\n".join(lines), parse_mode="Markdown")


async def handle_text(update: Update, context: ContextTypes.DEFAULT_TYPE):
    # Prioridad 1: si estamos esperando el nombre, guardarlo
    if await _guardar_nombre(update, context):
        return

    # Prioridad 2: si no está registrado, pedirle el nombre
    if not db.get_employee(update.effective_user.id):
        await _pedir_nombre(update, context)
        return

    # Prioridad 3: detectar entrada/salida por palabras clave
    text = update.message.text.lower().strip()
    if any(p in text for p in PALABRAS_ENTRADA):
        await _hacer_entro(update, context)
    elif any(p in text for p in PALABRAS_SALIDA):
        await _hacer_salgo(update, context)
    else:
        await update.message.reply_text(
            "No entendí. Comandos disponibles:\n"
            "  /entro  - Registrar entrada\n"
            "  /salgo  - Registrar salida\n"
            "  /estado - Ver tu estado hoy"
        )


# ---------- reporte quincena ----------

MESES_ES = {
    1: "Enero", 2: "Febrero", 3: "Marzo", 4: "Abril",
    5: "Mayo", 6: "Junio", 7: "Julio", 8: "Agosto",
    9: "Septiembre", 10: "Octubre", 11: "Noviembre", 12: "Diciembre",
}


def _sheet_name(name: str) -> str:
    """Nombre de hoja válido para Excel: max 31 chars, sin caracteres especiales."""
    for ch in r"/\?*[]:":
        name = name.replace(ch, "")
    return name[:31]


def _build_xlsx(records, titulo: str) -> io.BytesIO:
    # Estilos reutilizables
    hdr_fill = PatternFill("solid", fgColor="1F4E79")
    hdr_font = Font(bold=True, color="FFFFFF")
    tot_fill = PatternFill("solid", fgColor="1F4E79")
    tot_font = Font(bold=True, color="FFFFFF")
    name_fill = PatternFill("solid", fgColor="2E75B6")
    name_font = Font(bold=True, color="FFFFFF", size=12)

    def style_header(ws, row, cols):
        for col in range(1, cols + 1):
            c = ws.cell(row=row, column=col)
            c.fill = hdr_fill
            c.font = hdr_font
            c.alignment = Alignment(horizontal="center")

    def style_total(ws, row, cols):
        for col in range(1, cols + 1):
            c = ws.cell(row=row, column=col)
            c.fill = tot_fill
            c.font = tot_font

    wb = openpyxl.Workbook()

    # ── Hoja 1: Resumen ──────────────────────────────────────────────────────
    ws_res = wb.active
    ws_res.title = "Resumen"

    # Título del período
    ws_res.merge_cells("A1:B1")
    t = ws_res.cell(row=1, column=1, value=titulo)
    t.font = Font(bold=True, size=13)
    t.alignment = Alignment(horizontal="center")

    # Encabezados resumen
    ws_res.cell(row=2, column=1, value="Empleado")
    ws_res.cell(row=2, column=2, value="Total Horas")
    style_header(ws_res, 2, 2)

    ws_res.column_dimensions["A"].width = 26
    ws_res.column_dimensions["B"].width = 16

    records_sorted = sorted(records, key=lambda r: (r["name"], r["date"]))
    grouped = {name: list(g) for name, g in groupby(records_sorted, key=lambda r: r["name"])}

    summary_row = 3
    summary_refs = {}   # name -> fila en resumen
    detail_total = {}   # name -> (sheet_name, fila_total)

    for name in sorted(grouped.keys()):
        ws_res.cell(row=summary_row, column=1, value=name)
        summary_refs[name] = summary_row
        summary_row += 1

    # Fila TOTAL GENERAL en resumen (se completa después)
    grand_row = summary_row
    style_total(ws_res, grand_row, 2)
    ws_res.cell(row=grand_row, column=1, value="TOTAL GENERAL").font = tot_font
    ws_res.cell(row=grand_row, column=1).fill = tot_fill
    ws_res.cell(row=grand_row, column=1).alignment = Alignment(horizontal="left")

    # ── Una hoja por empleado ─────────────────────────────────────────────────
    for name, emp_records in sorted(grouped.items()):
        sname = _sheet_name(name)
        ws = wb.create_sheet(title=sname)

        # Fila 1: nombre del empleado (título)
        ws.merge_cells("A1:D1")
        nc = ws.cell(row=1, column=1, value=name)
        nc.fill = name_fill
        nc.font = name_font
        nc.alignment = Alignment(horizontal="center")

        # Fila 2: encabezados
        for col, h in enumerate(["Fecha", "Entrada", "Salida", "Horas Trabajadas"], start=1):
            ws.cell(row=2, column=col, value=h)
        style_header(ws, 2, 4)

        # Filas de datos (desde fila 3)
        data_start = 3
        row = data_start
        for r in emp_records:
            ws.cell(row=row, column=1, value=r["date"])

            if r["entry_time"]:
                c = ws.cell(row=row, column=2, value=r["entry_time"].replace(tzinfo=None).time())
                c.number_format = "HH:MM"
            else:
                ws.cell(row=row, column=2, value="")

            if r["exit_time"]:
                c = ws.cell(row=row, column=3, value=r["exit_time"].replace(tzinfo=None).time())
                c.number_format = "HH:MM"
            else:
                ws.cell(row=row, column=3, value="Sin salida")

            c = ws.cell(row=row, column=4,
                        value=f'=IF(C{row}="Sin salida","Sin salida",(C{row}-B{row})*24)')
            c.number_format = "0.00"
            c.alignment = Alignment(horizontal="center")
            row += 1

        data_end = row - 1
        total_row = row

        # Fila TOTAL del empleado
        style_total(ws, total_row, 4)
        ws.cell(row=total_row, column=3, value="TOTAL").font = tot_font
        ws.cell(row=total_row, column=3).fill = tot_fill
        ws.cell(row=total_row, column=3).alignment = Alignment(horizontal="right")
        tc = ws.cell(row=total_row, column=4,
                     value=f'=SUMIF(D{data_start}:D{data_end},"<>Sin salida",D{data_start}:D{data_end})')
        tc.number_format = "0.00"
        tc.font = tot_font
        tc.fill = tot_fill
        tc.alignment = Alignment(horizontal="center")

        detail_total[name] = (sname, total_row)

        ws.column_dimensions["A"].width = 14
        ws.column_dimensions["B"].width = 10
        ws.column_dimensions["C"].width = 12
        ws.column_dimensions["D"].width = 18

    # ── Completar referencias en Resumen ─────────────────────────────────────
    for name, srow in summary_refs.items():
        sname, trow = detail_total[name]
        c = ws_res.cell(row=srow, column=2, value=f"='{sname}'!D{trow}")
        c.number_format = "0.00"
        c.alignment = Alignment(horizontal="center")

    grand_refs = ",".join([f"B{r}" for r in summary_refs.values()])
    gc = ws_res.cell(row=grand_row, column=2, value=f"=SUM({grand_refs})")
    gc.number_format = "0.00"
    gc.font = tot_font
    gc.fill = tot_fill
    gc.alignment = Alignment(horizontal="center")

    buffer = io.BytesIO()
    wb.save(buffer)
    buffer.seek(0)
    return buffer


async def job_quincena(context: CallbackContext):
    hoy = ahora().date()
    ultimo_dia = calendar.monthrange(hoy.year, hoy.month)[1]
    mes = MESES_ES[hoy.month]

    if hoy.day == 15:
        start = hoy.replace(day=1)
        end = hoy
        label = f"1ra quincena {mes} {hoy.year} (1-15)"
        filename = f"asistencia_{hoy.year}{hoy.month:02d}_q1.xlsx"
    elif hoy.day == ultimo_dia:
        start = hoy.replace(day=16)
        end = hoy
        label = f"2da quincena {mes} {hoy.year} (16-{ultimo_dia})"
        filename = f"asistencia_{hoy.year}{hoy.month:02d}_q2.xlsx"
    else:
        return

    records = db.get_records_by_period(start, end)
    buffer = _build_xlsx(records, label)

    for admin_id in ADMIN_IDS:
        await context.bot.send_document(
            chat_id=admin_id,
            document=buffer,
            filename=filename,
            caption=f"Reporte automático - {label}\n{len(records)} registros",
        )
        buffer.seek(0)


# ---------- main ----------

def main():
    import asyncio
    asyncio.set_event_loop(asyncio.new_event_loop())

    if not TOKEN:
        raise ValueError("Falta la variable de entorno TELEGRAM_TOKEN")
    if not ADMIN_IDS:
        print("ADVERTENCIA: ADMIN_IDS no configurado. Nadie podrá descargar reportes.")

    app = Application.builder().token(TOKEN).build()

    app.add_handler(CommandHandler("start", cmd_start))
    app.add_handler(CommandHandler("entro", cmd_entro))
    app.add_handler(CommandHandler("salgo", cmd_salgo))
    app.add_handler(CommandHandler("estado", cmd_estado))
    app.add_handler(CommandHandler("reporte", cmd_reporte))
    app.add_handler(CommandHandler("empleados", cmd_empleados))
    app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, handle_text))

    # Reporte automático de quincena: todos los días a las 18:00 (verifica si es día 15 o último del mes)
    app.job_queue.run_daily(
        job_quincena,
        time=dt_time(hour=18, minute=0, tzinfo=TIMEZONE),
    )

    print("Bot iniciado. Esperando mensajes...")
    app.run_polling(drop_pending_updates=True)


if __name__ == "__main__":
    main()
