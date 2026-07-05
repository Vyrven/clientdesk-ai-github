from telegram import Update, InlineKeyboardButton, InlineKeyboardMarkup
from telegram.ext import (
    Application, CommandHandler, CallbackQueryHandler,
    MessageHandler, filters, ContextTypes
)
from telegram.error import BadRequest
from dotenv import load_dotenv

from bot.db import init_db, get_business_by_code, save_session
from bot.handlers import (
    handle_entrepreneur, handle_registration, handle_niche,
    handle_menu, handle_order_detail, handle_order_status,
    handle_client_start, handle_client_menu, handle_urgency,
    handle_select_business, handle_work_format, handle_work_mode,
    show_start_menu, build_client_portal_text, build_client_portal_keyboard
)
from bot.help import handle_help_command, handle_help_callback
from bot.admin import register_admin_handlers
from bot.notify import start_link_reminder_worker

import os
import traceback
from pathlib import Path


load_dotenv(Path(__file__).resolve().parent / ".env")
BOT_TOKEN = os.getenv("BOT_TOKEN")


async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    args = context.args

    if args and args[0].startswith("biz_"):
        code = args[0].replace("biz_", "")
        business = get_business_by_code(code)

        if business and business.get("name"):
            await handle_client_start(update, context, business, from_link=True)
            return

        await update.message.reply_text(
            "❌ Бізнес не знайдено. Можливо посилання застаріло."
        )
        return

    await show_start_menu(update, context)


async def handle_back_to_start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    await show_start_menu(update, context, edit=True)


async def handle_role_client(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()

    await query.edit_message_text(
        build_client_portal_text(),
        parse_mode="HTML",
        reply_markup=build_client_portal_keyboard()
    )


async def post_init(application):
    start_link_reminder_worker(application)


async def handle_bot_error(update: object, context: ContextTypes.DEFAULT_TYPE):
    error = context.error

    if isinstance(error, BadRequest) and "Message is not modified" in str(error):
        return

    print("Unhandled bot error:")
    traceback.print_exception(type(error), error, error.__traceback__)


def main():
    if not BOT_TOKEN:
        raise RuntimeError("BOT_TOKEN не знайдено в .env")

    init_db()

    app = Application.builder().token(BOT_TOKEN).post_init(post_init).build()
    app.add_error_handler(handle_bot_error)

    app.add_handler(CommandHandler("start", start))
    app.add_handler(CommandHandler("help", handle_help_command))

    # Реєструємо до текстового MessageHandler, інакше адмін-команди його не дістануться.
    register_admin_handlers(app)

    # Ставимо вище широких меню, щоб help_ точно ловився
    app.add_handler(CallbackQueryHandler(handle_help_callback, pattern="^help_"))

    app.add_handler(CallbackQueryHandler(handle_entrepreneur, pattern="^role_entrepreneur$"))
    app.add_handler(CallbackQueryHandler(handle_role_client, pattern="^role_client$"))
    app.add_handler(CallbackQueryHandler(handle_back_to_start, pattern="^back_to_start$"))

    app.add_handler(CallbackQueryHandler(handle_niche, pattern="^niche_"))
    app.add_handler(CallbackQueryHandler(handle_work_mode, pattern="^work_mode_"))

    app.add_handler(CallbackQueryHandler(handle_menu, pattern="^menu_"))

    app.add_handler(CallbackQueryHandler(handle_order_detail, pattern=r"^order_\d+$"))
    app.add_handler(CallbackQueryHandler(handle_order_status, pattern="^status_"))

    app.add_handler(CallbackQueryHandler(handle_client_menu, pattern="^client_"))
    app.add_handler(CallbackQueryHandler(handle_work_format, pattern="^workformat_"))
    app.add_handler(CallbackQueryHandler(handle_urgency, pattern="^urgency_"))
    app.add_handler(CallbackQueryHandler(handle_select_business, pattern="^select_biz_"))

    app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, handle_registration))

    print("Bot started.")
    app.run_polling()


if __name__ == "__main__":
    main()
