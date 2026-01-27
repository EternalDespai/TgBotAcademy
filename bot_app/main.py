from telegram import Update, InlineKeyboardButton, InlineKeyboardMarkup
from telegram.ext import (
    ApplicationBuilder,
    CommandHandler,
    MessageHandler,
    CallbackQueryHandler,
    ContextTypes,
    filters,
)
from config import BOT_TOKEN
from excel_parser import (
    process_excel_file,
    is_students_reports_3_or_6,
    process_students_bad_grades_from_bytes,
    process_students_hw_completion_from_bytes,
)
from utils import send_long_message


async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text("👋 Привет! Пришли мне .xlsx файл.")


async def on_document(update: Update, context: ContextTypes.DEFAULT_TYPE):
    doc = update.message.document
    await update.message.reply_text("📥 Анализирую файл...")

    try:
        tg_file = await doc.get_file()
        data = await tg_file.download_as_bytearray()
        data_bytes = bytes(data)

        # Сохраняем последний файл пользователя для кнопок
        context.user_data["last_xlsx_bytes"] = data_bytes

        if is_students_reports_3_or_6(data_bytes):
            keyboard = [
                [InlineKeyboardButton("📌 Отчёт по студентам (ДЗ=1, КР<3)", callback_data="rep:3")],
                [InlineKeyboardButton("📌 % выполненных ДЗ (<70%)", callback_data="rep:6")],
            ]
            await update.message.reply_text(
                "Выберите отчет:",
                reply_markup=InlineKeyboardMarkup(keyboard),
            )
            return

        report_text = process_excel_file(data_bytes)
        await send_long_message(update, report_text)

    except Exception as e:
        await update.message.reply_text(f"❌ Критическая ошибка бота: {e}")


async def on_choose_report(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()

    data_bytes = context.user_data.get("last_xlsx_bytes")
    if not data_bytes:
        await query.edit_message_text("❌ Файл не найден. Пришлите .xlsx заново.")
        return

    try:
        if query.data == "rep:3":
            await query.edit_message_text("📥 Готовлю отчёт по студентам (ДЗ=1, КР<3)...")
            report_text = process_students_bad_grades_from_bytes(data_bytes)
        elif query.data == "rep:6":
            await query.edit_message_text("📥 Готовлю отчёт по % выполненных ДЗ...")
            report_text = process_students_hw_completion_from_bytes(data_bytes)
        else:
            await query.edit_message_text("❌ Неизвестный выбор.")
            return

        await send_long_message(Update(update.update_id, message=query.message), report_text)

    except Exception as e:
        await query.edit_message_text(f"❌ Ошибка при формировании отчёта: {e}")


def main():
    if not BOT_TOKEN or "PASTE" in BOT_TOKEN:
        print("Ошибка: Укажи токен в config.py!")
        return

    app = ApplicationBuilder().token(BOT_TOKEN).build()
    app.add_handler(CommandHandler("start", start))
    app.add_handler(MessageHandler(filters.Document.ALL, on_document))
    app.add_handler(CallbackQueryHandler(on_choose_report, pattern=r"^rep:(3|6)$"))

    print("Бот запущен...")
    app.run_polling()


if __name__ == "__main__":
    main()
