# bot.py

import os
from telegram import Update, InputMediaDocument
from telegram.ext import (
    ApplicationBuilder, CommandHandler, MessageHandler, ContextTypes, ConversationHandler, filters,
)
from gpt_integration import improve_text, create_combat_report_from_text
from grades import collect_grades_via_bot
from dotenv import load_dotenv

load_dotenv()

# Constants for conversation states
INPUT_TEXT, COLLECT_GRADES = range(2)

async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text(
        "ברוכים הבאים לבוט יצירת דוח סיכום קרב!\n"
        "שלח לי את הטקסט שברצונך לשפר ולכלול בדוח."
    )
    return INPUT_TEXT

async def receive_text(update: Update, context: ContextTypes.DEFAULT_TYPE):
    input_text = update.message.text
    context.user_data['input_text'] = input_text

    # Improve the text using GPT-4
    await update.message.reply_text("משפר את הטקסט, אנא המתן...")
    improved_text = improve_text(input_text)
    context.user_data['improved_text'] = improved_text

    await update.message.reply_text(
        "הטקסט שופר בהצלחה.\n"
        "עכשיו נתחיל באיסוף ציונים.\n"
        "הזן ציון בין 1 ל-10 עבור כל פריט."
    )

    # Start collecting grades
    context.user_data['grades_data'] = {}
    return await collect_grades_via_bot(update, context)

async def cancel(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text("הפעולה בוטלה. ניתן להתחיל מחדש עם /start.")
    return ConversationHandler.END

def main():
    # Get the bot token from environment variables
    bot_token = os.getenv("TELEGRAM_BOT_TOKEN")
    application = ApplicationBuilder().token(bot_token).build()

    conv_handler = ConversationHandler(
        entry_points=[CommandHandler('start', start)],
        states={
            INPUT_TEXT: [MessageHandler(filters.TEXT & ~filters.COMMAND, receive_text)],
            COLLECT_GRADES: [MessageHandler(filters.TEXT & ~filters.COMMAND, collect_grades_via_bot)],
        },
        fallbacks=[CommandHandler('cancel', cancel)],
    )

    application.add_handler(conv_handler)
    application.run_polling()

if __name__ == '__main__':
    main()
