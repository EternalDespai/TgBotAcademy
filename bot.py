import io
from collections import Counter

from telegram import Update
from telegram.ext import ApplicationBuilder, CommandHandler, MessageHandler, ContextTypes, filters

import openpyxl
import xlrd


BOT_TOKEN = "PASTE_YOUR_TOKEN_HERE"

# защита от переименования magic bytes
def detect_excel_type(data: bytes) -> str:
    if len(data) >= 2 and data[0:2] == b"PK":
        return "xlsx"
    if len(data) >= 8 and data[0:8] == bytes.fromhex("D0CF11E0A1B11AE1"):
        return "xls"
    return "unknown"


def count_subjects_from_text(text: str, counter: Counter):
    for line in str(text).splitlines():
        line = line.strip()
        if line.startswith("Предмет:"):
            subj = line.replace("Предмет:", "", 1).strip()
            if subj:
                counter[subj] += 1


def parse_xlsx(data: bytes) -> Counter:
    counter = Counter()
    wb = openpyxl.load_workbook(io.BytesIO(data), data_only=True)
    ws = wb.worksheets[0]  # первый лист

    for row in ws.iter_rows(values_only=True):
        for cell in row:
            if isinstance(cell, str) and "Предмет:" in cell:
                count_subjects_from_text(cell, counter)

    return counter


def parse_xls(data: bytes) -> Counter:
    counter = Counter()
    book = xlrd.open_workbook(file_contents=data)
    sheet = book.sheet_by_index(0)  # первый лист

    for r in range(sheet.nrows):
        for c in range(sheet.ncols):
            cell = sheet.cell_value(r, c)
            if isinstance(cell, str) and "Предмет:" in cell:
                count_subjects_from_text(cell, counter)

    return counter


async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text(
        "Привет! Пришли файл расписания (.xlsx или .xls). "
        "Я посчитаю количество пар по предметам"
    )


async def on_document(update: Update, context: ContextTypes.DEFAULT_TYPE):
    doc = update.message.document
    await update.message.reply_text("Файл получил, читаю...")

    tg_file = await doc.get_file()
    data = await tg_file.download_as_bytearray()

    file_type = detect_excel_type(data)

    if file_type == "xlsx":
        counter = parse_xlsx(bytes(data))
    elif file_type == "xls":
        counter = parse_xls(bytes(data))
    else:
        await update.message.reply_text("Не понял формат файла. Нужен .xlsx или .xls.")
        return

    if not counter:
        await update.message.reply_text("Не нашел строк 'Предмет:' в таблице.")
        return

    lines = ["Количество пар по предметам:"]
    for name, cnt in counter.most_common():
        lines.append(f"- {name}: {cnt}")

    await update.message.reply_text("\n".join(lines))


def main():
    app = ApplicationBuilder().token(BOT_TOKEN).build()
    app.add_handler(CommandHandler("start", start))
    app.add_handler(MessageHandler(filters.Document.ALL, on_document))
    print("Bot started...")
    app.run_polling()


if __name__ == "__main__":
    main()