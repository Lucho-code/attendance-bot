import os
import io
from datetime import datetime

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


def ahora() -> datetime:
    return datetime.now(TIMEZONE)


def es_admin(user_id: int) -> bool:
    return user_id in ADMIN_IDS


# ---------- handlers ----------

async def cmd_start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user = update.effective_user
    name = f"{user.first_name} {user.last_name or ''}".strip()
    if context.args:
        name = " ".join(context.args)

    db.register_employee(user.id, name)
    await update.message.reply_text(
        f"Registrado como *{name}*\n\n"
        "Comandos disponibles:\n"
        "  /entro   - Registrar entrada\n"
        "  /salgo   - Registrar salida\n"
        "  /estado  - Ver tu estado hoy\n\n"
        "También podés escribir frases como \"llegué\" o \"me voy\".",
        parse_mode="Markdown",
    )


async def _hacer_entro(update: Update):
    user = update.effective_user
    employee = db.get_employee(user.id)
    if not employee:
        await update.message.reply_text(
            "No estás registrado. Enviá /start primero."
        )
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


async def _hacer_salgo(update: Update):
    user = update.effective_user
    employee = db.get_employee(user.id)
    if not employee:
        await update.message.reply_text(
            "No estás registrado. Enviá /start primero."
        )
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
    await _hacer_entro(update)


async def cmd_salgo(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await _hacer_salgo(update)


async def cmd_estado(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user = update.effective_user
    employee = db.get_employee(user.id)
    if not employee:
        await update.message.reply_text(
            "No estás registrado. Enviá /start primero."
        )
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

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Asistencia"

    # Encabezados con formato
    headers = ["Nombre", "Fecha", "Entrada", "Salida", "Horas Trabajadas"]
    ws.append(headers)
    header_fill = PatternFill("solid", fgColor="1F4E79")
    header_font = Font(bold=True, color="FFFFFF")
    for col, _ in enumerate(headers, start=1):
        cell = ws.cell(row=1, column=col)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center")

    # Filas de datos
    for i, r in enumerate(records, start=2):
        ws.cell(row=i, column=1, value=r["name"])
        ws.cell(row=i, column=2, value=r["date"])

        # Columna C: Entrada como valor de hora real
        if r["entry_time"]:
            cell_c = ws.cell(row=i, column=3, value=r["entry_time"].replace(tzinfo=None).time())
            cell_c.number_format = "HH:MM"
        else:
            ws.cell(row=i, column=3, value="")

        # Columna D: Salida como valor de hora real
        if r["exit_time"]:
            cell_d = ws.cell(row=i, column=4, value=r["exit_time"].replace(tzinfo=None).time())
            cell_d.number_format = "HH:MM"
        else:
            ws.cell(row=i, column=4, value="Sin salida")

        # Columna E: Fórmula automática de horas
        cell_e = ws.cell(row=i, column=5, value=f'=IF(D{i}="Sin salida","Sin salida",(D{i}-C{i})*24)')
        cell_e.number_format = "0.00"
        cell_e.alignment = Alignment(horizontal="center")

    # Ancho de columnas
    ws.column_dimensions["A"].width = 22
    ws.column_dimensions["B"].width = 14
    ws.column_dimensions["C"].width = 10
    ws.column_dimensions["D"].width = 10
    ws.column_dimensions["E"].width = 18

    buffer = io.BytesIO()
    wb.save(buffer)
    buffer.seek(0)

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
    text = update.message.text.lower().strip()
    if any(p in text for p in PALABRAS_ENTRADA):
        await _hacer_entro(update)
    elif any(p in text for p in PALABRAS_SALIDA):
        await _hacer_salgo(update)
    else:
        await update.message.reply_text(
            "No entendí. Comandos disponibles:\n"
            "  /entro  - Registrar entrada\n"
            "  /salgo  - Registrar salida\n"
            "  /estado - Ver tu estado hoy"
        )


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

    print("Bot iniciado. Esperando mensajes...")
    app.run_polling(drop_pending_updates=True)


if __name__ == "__main__":
    main()
