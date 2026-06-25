import os
import csv
import io
from datetime import datetime

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

    output = io.StringIO()
    writer = csv.writer(output)
    writer.writerow(["Nombre", "Fecha", "Entrada", "Salida", "Horas_Trabajadas"])
    for r in records:
        writer.writerow([
            r["name"],
            r["date"],
            r["entry_time"].strftime("%H:%M") if r["entry_time"] else "",
            r["exit_time"].strftime("%H:%M") if r["exit_time"] else "Sin salida",
            f"{r['total_hours']:.2f}" if r["total_hours"] is not None else "",
        ])

    output.seek(0)
    filename = f"asistencia_{ahora().strftime('%Y%m%d_%H%M')}.csv"
    data = output.getvalue().encode("utf-8-sig")  # BOM para abrir bien en Excel

    await update.message.reply_document(
        document=data,
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
