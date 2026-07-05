import os
import html

from telegram import InlineKeyboardButton, InlineKeyboardMarkup, Update
from telegram.ext import ContextTypes

from bot.db import get_business, get_session, get_business_by_id


SUPPORT_USERNAME = (os.getenv("SUPPORT_USERNAME") or "Alexxplorerr").replace("@", "").strip()


def h(value) -> str:
    if value is None:
        return ""
    return html.escape(str(value))


def support_url() -> str:
    return f"https://t.me/{SUPPORT_USERNAME}"


def build_support_button() -> InlineKeyboardButton:
    return InlineKeyboardButton("💬 Написати підтримці", url=support_url())


def build_help_home_keyboard() -> InlineKeyboardMarkup:
    return InlineKeyboardMarkup([
        [InlineKeyboardButton("👔 Допомога для бізнесу", callback_data="help_business")],
        [InlineKeyboardButton("👤 Допомога для клієнта", callback_data="help_client")],
        [build_support_button()],
        [InlineKeyboardButton("🏠 На початок", callback_data="back_to_start")],
    ])


def build_business_help_keyboard(has_business: bool) -> InlineKeyboardMarkup:
    if has_business:
        return InlineKeyboardMarkup([
            [build_support_button()],
            [InlineKeyboardButton("📋 Заявки", callback_data="menu_orders")],
            [InlineKeyboardButton("🔗 Моє посилання", callback_data="menu_link")],
            [InlineKeyboardButton("📚 AI-база знань", callback_data="menu_knowledge")],
            [InlineKeyboardButton("🔙 Назад до допомоги", callback_data="help_home")],
        ])

    return InlineKeyboardMarkup([
        [build_support_button()],
        [InlineKeyboardButton("👔 Зареєструвати бізнес", callback_data="role_entrepreneur")],
        [InlineKeyboardButton("🔙 Назад до допомоги", callback_data="help_home")],
    ])


def build_client_help_keyboard(has_client_business: bool) -> InlineKeyboardMarkup:
    if has_client_business:
        return InlineKeyboardMarkup([
            [InlineKeyboardButton("📝 Залишити заявку", callback_data="client_order")],
            [InlineKeyboardButton("💰 Дізнатись ціну", callback_data="client_price")],
            [InlineKeyboardButton("🕐 Графік роботи", callback_data="client_schedule")],
            [InlineKeyboardButton("📍 Адреса", callback_data="client_address")],
            [build_support_button()],
            [InlineKeyboardButton("🔙 Назад до допомоги", callback_data="help_home")],
        ])

    return InlineKeyboardMarkup([
        [InlineKeyboardButton("🔎 Знайти бізнес", callback_data="role_client")],
        [build_support_button()],
        [InlineKeyboardButton("🔙 Назад до допомоги", callback_data="help_home")],
    ])


def build_help_home_text() -> str:
    return (
        "🆘 <b>Допомога ClientDesk AI</b>\n\n"
        "Оберіть, для кого потрібна допомога.\n\n"
        "👔 <b>Для бізнесу</b> - керування AI-адміністратором, заявками, базою знань, посиланням і тарифом.\n\n"
        "👤 <b>Для клієнта</b> - як залишити заявку, знайти бізнес, дізнатись ціну, адресу або графік.\n\n"
        "Якщо щось не працює або потрібне підключення - натисніть <b>Написати підтримці</b>."
    )


def build_business_help_text(user_id: int) -> tuple[str, bool]:
    business = get_business(user_id)

    if not business or not business.get("name"):
        text = (
            "👔 <b>Допомога для бізнесу</b>\n\n"
            "ClientDesk AI допомагає бізнесу приймати заявки 24/7, уточнювати деталі у клієнтів "
            "і передавати власнику вже готові ліди.\n\n"
            "<b>Щоб почати:</b>\n"
            "- зареєструйте бізнес\n"
            "- вкажіть нішу, місто та послуги\n"
            "- заповніть AI-базу знань\n"
            "- скопіюйте посилання для клієнтів\n"
            "- розмістіть його в Instagram, Telegram, Google Maps або на сайті\n\n"
            "Після цього клієнти зможуть залишати заявки, а ви будете отримувати їх прямо в Telegram."
        )
        return text, False

    name = business.get("name") or "-"
    city = business.get("city") or "-"
    niche = business.get("niche") or "-"
    tariff = str(business.get("tariff") or "free").upper()

    text = (
        "👔 <b>Допомога для бізнесу</b>\n\n"
        f"🏢 Бізнес: <b>{h(name)}</b>\n"
        f"📍 Місто: <b>{h(city)}</b>\n"
        f"🏷️ Ніша: <b>{h(niche)}</b>\n"
        f"💼 Тариф: <b>{h(tariff)}</b>\n\n"
        "<b>Що важливо зробити:</b>\n"
        "- заповнити AI-базу знань\n"
        "- скопіювати посилання для клієнтів\n"
        "- розмістити посилання в Instagram / Telegram / Google Maps\n"
        "- перевіряти заявки і змінювати статуси\n\n"
        "<b>Як працює AI-база знань:</b>\n"
        "Туди потрібно додати графік, ціни, адресу, умови, гарантію, часті питання і те, що AI не має обіцяти клієнтам.\n"
        "Після цього бот зможе точніше відповідати клієнтам.\n\n"
        "<b>Як обробляти заявки:</b>\n"
        "Відкрийте <b>Заявки</b>, перегляньте звернення і змініть статус: в роботі, зв'язались, успішно або відмова."
    )
    return text, True


def build_client_help_text(user_id: int) -> tuple[str, bool]:
    session = get_session(user_id)
    business = None

    if session and session.get("data"):
        business_id = session["data"].get("business_id")
        if business_id:
            business = get_business_by_id(business_id)

    if business and business.get("name"):
        text = (
            "👤 <b>Допомога для клієнта</b>\n\n"
            f"Ви зараз у бізнесі: <b>{h(business.get('name'))}</b>\n"
            f"📍 Місто: <b>{h(business.get('city') or '-')}</b>\n\n"
            "<b>Що ви можете зробити:</b>\n"
            "- залишити заявку на послугу\n"
            "- дізнатись ціну\n"
            "- подивитись графік роботи\n"
            "- отримати адресу\n\n"
            "<b>Як краще залишати заявку:</b>\n"
            "Напишіть одним повідомленням усе, що знаєте. Наприклад: пристрій, проблема, терміновість, ім'я і телефон.\n\n"
            "AI сам розбере деталі і поставить лише потрібні уточнюючі питання."
        )
        return text, True

    text = (
        "👤 <b>Допомога для клієнта</b>\n\n"
        "Щоб залишити заявку, спочатку потрібно знайти бізнес.\n\n"
        "<b>Як це працює:</b>\n"
        "- натисніть <b>Знайти бізнес</b>\n"
        "- напишіть місто\n"
        "- введіть назву компанії, код бізнесу або напрямок послуги\n"
        "- відкрийте потрібний бізнес\n"
        "- залиште заявку\n\n"
        "Якщо у вас є пряме посилання від бізнесу - просто відкрийте його, і бот одразу покаже потрібну компанію."
    )
    return text, False


async def handle_help_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.effective_message.reply_text(
        build_help_home_text(),
        parse_mode="HTML",
        reply_markup=build_help_home_keyboard()
    )


async def handle_help_callback(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()

    user_id = query.from_user.id
    action = query.data

    if action == "help_home":
        await query.edit_message_text(
            build_help_home_text(),
            parse_mode="HTML",
            reply_markup=build_help_home_keyboard()
        )
        return

    if action == "help_business":
        text, has_business = build_business_help_text(user_id)
        await query.edit_message_text(
            text,
            parse_mode="HTML",
            reply_markup=build_business_help_keyboard(has_business)
        )
        return

    if action == "help_client":
        text, has_client_business = build_client_help_text(user_id)
        await query.edit_message_text(
            text,
            parse_mode="HTML",
            reply_markup=build_client_help_keyboard(has_client_business)
        )
        return

    await query.edit_message_text(
        build_help_home_text(),
        parse_mode="HTML",
        reply_markup=build_help_home_keyboard()
    )
