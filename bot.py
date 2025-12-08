import os
import sqlite3
from datetime import datetime, time
from math import radians, sin, cos, sqrt, atan2

from telegram import Update, KeyboardButton, ReplyKeyboardMarkup
from telegram.ext import (
    ApplicationBuilder, CommandHandler, MessageHandler, filters, ContextTypes
)
from openpyxl import Workbook


# --------------------------------------------------------
# CONFIGURAZIONE
# --------------------------------------------------------

TOKEN = os.getenv("TOKEN")  # <-- Render leggerà il token da Environment Variables

# Posizione della sede
SEDE_LAT = 41.8619944089824
SEDE_LON = 12.840221654723194
RAGGIO_MAX = 150  # metri consentiti


# --------------------------------------------------------
# DATABASE
# --------------------------------------------------------

conn = sqlite3.connect("presenze.db", check_same_thread=False)
c = conn.cursor()

c.execute("""
CREATE TABLE IF NOT EXISTS presenze (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    user_id INTEGER,
    nome TEXT,
    azione TEXT,
    timestamp TEXT,
    lat REAL,
    lon REAL
)
""")
conn.commit()


# --------------------------------------------------------
# FUNZIONI DI SERVIZIO
# --------------------------------------------------------

def distanza_m(lat1, lon1, lat2, lon2):
    """Calcola distanza in metri (Haversine)."""
    R = 6371000
    dlat = radians(lat2 - lat1)
    dlon = radians(lon2 - lon1)
    a = sin(dlat/2)**2 + cos(radians(lat1)) * cos(radians(lat2)) * sin(dlon/2)**2
    cang = 2 * atan2(sqrt(a), sqrt(1 - a))
    return R * cang


def registra_presenza(user_id, nome, azione, lat, lon):
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    c.execute(
        "INSERT INTO presenze (user_id, nome, azione, timestamp, lat, lon) VALUES (?, ?, ?, ?, ?, ?)",
        (user_id, nome, azione, timestamp, lat, lon)
    )
    conn.commit()


# --------------------------------------------------------
# COMANDI BOT
# --------------------------------------------------------

async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    button = KeyboardButton("Invia posizione 📍", request_location=True)
    keyboard = ReplyKeyboardMarkup([[button]], resize_keyboard=True)

    await update.message.reply_text(
        "👋 Ciao! Sono il bot presenze.\n\n"
        "Usa:\n"
        "✔ /entra per registrare entrata\n"
        "✔ /esci per registrare uscita\n"
        "✔ /ore per vedere ore lavorate\n"
        "✔ /export per esportare in Excel\n\n"
        "Per registrare una timbratura devi inviare la posizione GPS.",
        reply_markup=keyboard
    )


async def entra(update: Update, context: ContextTypes.DEFAULT_TYPE):
    context.user_data["azione"] = "ENTRATA"
    await update.message.reply_text("📍 Invia la posizione per registrare l'ENTRATA.")


async def esci(update: Update, context: ContextTypes.DEFAULT_TYPE):
    context.user_data["azione"] = "USCITA"
    await update.message.reply_text("📍 Invia la posizione per registrare l'USCITA.")


async def posizione(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user = update.effective_user
    azione = context.user_data.get("azione")

    if not azione:
        return await update.message.reply_text("❗ Usa prima /entra o /esci.")

    lat = update.message.location.latitude
    lon = update.message.location.longitude

    distanza = distanza_m(lat, lon, SEDE_LAT, SEDE_LON)

    if distanza > RAGGIO_MAX:
        return await update.message.reply_text(
            f"❌ Troppo distante dalla sede: {int(distanza)} metri.\n"
            "Timbratura annullata."
        )

    registra_presenza(user.id, user.full_name, azione, lat, lon)
    context.user_data["azione"] = None

    await update.message.reply_text(f"✔ {azione} registrata correttamente!")


# --------------------------------------------------------
# CALCOLO ORE
# --------------------------------------------------------

def fascia_normale(dt):
    if dt.weekday() < 5:
        return time(7, 0), time(16, 30)
    else:
        return time(7, 30), time(13, 0)


async def ore(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user = update.effective_user

    c.execute(
        "SELECT azione, timestamp FROM presenze WHERE user_id=? ORDER BY timestamp",
        (user.id,)
    )
    rows = c.fetchall()

    entrate, uscite = [], []
    for azione, ts in rows:
        dt = datetime.strptime(ts, "%Y-%m-%d %H:%M:%S")
        if azione == "ENTRATA":
            entrate.append(dt)
        else:
            uscite.append(dt)

    ore_normali = ore_extra = 0

    for i in range(min(len(entrate), len(uscite))):
        start = entrate[i]
        end = uscite[i]

        giorno = start.date()
        inizio_std, fine_std = fascia_normale(start)
        inizio = datetime.combine(giorno, inizio_std)
        fine = datetime.combine(giorno, fine_std)

        # Normali
        normale = max(0, (min(end, fine) - max(start, inizio)).total_seconds())
        ore_normali += normale

        # Extra prima
        if start < inizio:
            ore_extra += (inizio - start).total_seconds()

        # Extra dopo
        if end > fine:
            ore_extra += (end - fine).total_seconds()

    await update.message.reply_text(
        f"⏳ Ore normali: {int(ore_normali//3600)}h {int((ore_normali%3600)//60)}m\n"
        f"🔥 Straordinari: {int(ore_extra//3600)}h {int((ore_extra%3600)//60)}m"
    )


# --------------------------------------------------------
# EXPORT EXCEL
# --------------------------------------------------------

async def export(update: Update, context: ContextTypes.DEFAULT_TYPE):
    wb = Workbook()
    ws = wb.active
    ws.append(["ID", "User", "Nome", "Azione", "Ora", "Lat", "Lon"])

    c.execute("SELECT * FROM presenze ORDER BY id ASC")
    for row in c:
        ws.append(row)

    filename = "presenze.xlsx"
    wb.save(filename)

    await update.message.reply_document(open(filename, "rb"), caption="Ecco il file Excel 📄")


# --------------------------------------------------------
# MAIN
# --------------------------------------------------------

def main():
    app = ApplicationBuilder().token(TOKEN).build()

    app.add_handler(CommandHandler("start", start))
    app.add_handler(CommandHandler("entra", entra))
    app.add_handler(CommandHandler("esci", esci))
    app.add_handler(CommandHandler("ore", ore))
    app.add_handler(CommandHandler("export", export))

    app.add_handler(MessageHandler(filters.LOCATION, posizione))

    print("BOT ONLINE ✔")
    app.run_polling()


if __name__ == "__main__":
    main()
