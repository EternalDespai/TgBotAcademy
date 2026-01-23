from telegram import Update
from telegram.ext import ApplicationBuilder, CommandHandler, MessageHandler, ContextTypes, filters
from config import BOT_TOKEN
from excel_parser import process_excel_file
from utils import send_long_message

async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text("👋 Привет! Пришли мне .xlsx файл (расписание или темы).")

async def on_document(update: Update, context: ContextTypes.DEFAULT_TYPE):
    doc = update.message.document
    await update.message.reply_text("📥 Анализирую файл...")

    try:
        tg_file = await doc.get_file()
        data = await tg_file.download_as_bytearray()

        report_text = process_excel_file(data)

        await send_long_message(update, report_text)

    except Exception as e:
        await update.message.reply_text(f"❌ Критическая ошибка бота: {e}")

def main():
    if not BOT_TOKEN or "PASTE" in BOT_TOKEN:
        print("Ошибка: Укажи токен в config.py!")
        return

    app = ApplicationBuilder().token(BOT_TOKEN).build()
    app.add_handler(CommandHandler("start", start))
    app.add_handler(MessageHandler(filters.Document.ALL, on_document))

    print("Бот запущен...")
    app.run_polling()


if __name__ == "__main__":
    main()
