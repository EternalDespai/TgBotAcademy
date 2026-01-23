from telegram import Update
from telegram.constants import ParseMode

async def send_long_message(update: Update, text: str):
    LIMIT = 4050

    if len(text) <= LIMIT:
        await update.message.reply_text(text, parse_mode=ParseMode.HTML)
        return

    buffer = ""
    for line in text.splitlines(keepends=True):
        if len(buffer) + len(line) > LIMIT:
            await update.message.reply_text(buffer, parse_mode=ParseMode.HTML)
            buffer = ""
        buffer += line

    if buffer:
        await update.message.reply_text(buffer, parse_mode=ParseMode.HTML)
