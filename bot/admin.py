import os
import re
import html
from typing import Dict, List, Optional, Tuple

from telegram import InlineKeyboardButton, InlineKeyboardMarkup, Update
from telegram.ext import CommandHandler, CallbackQueryHandler, ContextTypes

from bot.db import (
    build_subscription_period,
    display_datetime,
    get_connection,
    get_tariff_usage,
    is_paid_tariff,
)

try:
    from config import TARIFF_LIMITS, TARIFF_PRICES
except Exception:
    TARIFF_LIMITS = {
        "free": 5,
        "start": 50,
        "business": 300,
        "vip": 1000,
    }
    TARIFF_PRICES = {
        "free": "Пробний тариф - до 5 заявок",
        "start": "2990 грн/міс - до 50 заявок",
        "business": "5990 грн/міс - до 300 заявок",
        "vip": "12990 грн/міс - до 1000 заявок",
    }


# Конфіг адмінки

TARIFF_ALIASES = {
    "free": "free",
    "test": "free",
    "trial": "free",

    # /settariff CODE basic still works as an old alias for START.
    "basic": "start",
    "base": "start",
    "start": "start",
    "starter": "start",

    "business": "business",
    "biz": "business",
    "pro": "vip",
    "premium": "vip",

    "vip": "vip",
    "max": "vip",
    "maximum": "vip",
}

TARIFF_PUBLIC_NAMES = {
    "free": "FREE",
    "start": "START",
    "business": "BUSINESS",
    "vip": "VIP",
}


# Базові хелпери

def h(value) -> str:
    if value is None:
        return ""
    return html.escape(str(value))


def get_admin_ids() -> List[int]:
    raw = os.getenv("ADMIN_IDS", "")
    ids = []

    for part in re.split(r"[,;\s]+", raw.strip()):
        if not part:
            continue
        try:
            ids.append(int(part))
        except ValueError:
            pass

    return ids


def is_admin_id(user_id: Optional[int]) -> bool:
    if user_id is None:
        return False
    return int(user_id) in get_admin_ids()


def normalize_code(code: str) -> str:
    return str(code or "").strip().upper()


def normalize_tariff(tariff: str) -> Optional[str]:
    key = str(tariff or "").strip().lower()
    return TARIFF_ALIASES.get(key)


def tariff_name(tariff: str) -> str:
    tariff = str(tariff or "free").lower()
    return TARIFF_PUBLIC_NAMES.get(tariff, tariff.upper())


def tariff_limit(tariff: str) -> int:
    return int(TARIFF_LIMITS.get(str(tariff or "free").lower(), 5))


def tariff_price(tariff: str) -> str:
    return str(TARIFF_PRICES.get(str(tariff or "free").lower(), "-"))


def admin_keyboard() -> InlineKeyboardMarkup:
    return InlineKeyboardMarkup([
        [InlineKeyboardButton("🏢 Бізнеси", callback_data="admin_businesses")],
        [InlineKeyboardButton("📊 Статистика", callback_data="admin_stats")],
        [InlineKeyboardButton("💼 Тарифи", callback_data="admin_tariffs")],
    ])


def back_admin_keyboard() -> InlineKeyboardMarkup:
    return InlineKeyboardMarkup([
        [InlineKeyboardButton("🔙 Адмін-панель", callback_data="admin_home")]
    ])


async def deny_access(update: Update) -> None:
    message = update.effective_message
    if message:
        await message.reply_text(
            "⛔ <b>Доступ закрито</b>\n\n"
            "Ця команда доступна тільки власнику ClientDesk AI.",
            parse_mode="HTML"
        )


async def answer_denied_callback(update: Update) -> None:
    query = update.callback_query
    if query:
        await query.answer("⛔ Немає доступу", show_alert=True)


# Хелпери для бази даних

def get_business_by_code_admin(code: str) -> Optional[Dict]:
    code = normalize_code(code)
    conn = get_connection()
    cur = conn.cursor()
    cur.execute("SELECT * FROM businesses WHERE UPPER(link_code) = ?", (code,))
    row = cur.fetchone()
    conn.close()
    return dict(row) if row else None


def get_business_by_id_admin(business_id: int) -> Optional[Dict]:
    conn = get_connection()
    cur = conn.cursor()
    cur.execute("SELECT * FROM businesses WHERE id = ?", (business_id,))
    row = cur.fetchone()
    conn.close()
    return dict(row) if row else None


def list_businesses_admin(limit: int = 50) -> List[Dict]:
    conn = get_connection()
    cur = conn.cursor()
    cur.execute(
        """
        SELECT *
        FROM businesses
        WHERE name IS NOT NULL AND TRIM(name) != ''
        ORDER BY created_at DESC, id DESC
        LIMIT ?
        """,
        (limit,)
    )
    rows = cur.fetchall()
    conn.close()
    return [dict(row) for row in rows]


def count_business_orders(business_id: int) -> Dict[str, int]:
    conn = get_connection()
    cur = conn.cursor()

    def one(query: str, params: tuple = ()) -> int:
        cur.execute(query, params)
        row = cur.fetchone()
        return int(row[0] or 0)

    data = {
        "total": one("SELECT COUNT(*) FROM orders WHERE business_id = ?", (business_id,)),
        "new": one("SELECT COUNT(*) FROM orders WHERE business_id = ? AND status = 'new'", (business_id,)),
        "hot": one("SELECT COUNT(*) FROM orders WHERE business_id = ? AND is_hot = 1", (business_id,)),
        "success": one("SELECT COUNT(*) FROM orders WHERE business_id = ? AND status = 'success'", (business_id,)),
        "cancelled": one("SELECT COUNT(*) FROM orders WHERE business_id = ? AND status IN ('cancelled', 'cancelled_by_client')", (business_id,)),
    }

    conn.close()
    return data


def update_business_tariff(code: str, tariff: str) -> Optional[Dict]:
    business = get_business_by_code_admin(code)
    if not business:
        return None

    conn = get_connection()
    cur = conn.cursor()
    if is_paid_tariff(tariff):
        started_at, expires_at = build_subscription_period()
        cur.execute(
            """
            UPDATE businesses
            SET tariff = ?, subscription_started_at = ?, subscription_expires_at = ?
            WHERE id = ?
            """,
            (tariff, started_at, expires_at, business["id"])
        )
    else:
        cur.execute(
            """
            UPDATE businesses
            SET tariff = ?, subscription_started_at = NULL, subscription_expires_at = NULL
            WHERE id = ?
            """,
            (tariff, business["id"])
        )
    conn.commit()
    conn.close()

    return get_business_by_code_admin(code)


def delete_business_by_code(code: str) -> Tuple[bool, Optional[Dict]]:
    business = get_business_by_code_admin(code)
    if not business:
        return False, None

    conn = get_connection()
    cur = conn.cursor()

    business_id = business["id"]
    owner_id = business.get("owner_id")

    cur.execute("DELETE FROM orders WHERE business_id = ?", (business_id,))
    cur.execute("DELETE FROM sessions WHERE business_id = ?", (business_id,))
    if owner_id:
        cur.execute("DELETE FROM sessions WHERE user_id = ?", (owner_id,))
    cur.execute("DELETE FROM businesses WHERE id = ?", (business_id,))

    conn.commit()
    conn.close()
    return True, business


def get_global_stats() -> Dict[str, int]:
    conn = get_connection()
    cur = conn.cursor()

    def one(query: str) -> int:
        cur.execute(query)
        row = cur.fetchone()
        return int(row[0] or 0)

    data = {
        "businesses": one("SELECT COUNT(*) FROM businesses WHERE name IS NOT NULL AND TRIM(name) != ''"),
        "orders": one("SELECT COUNT(*) FROM orders"),
        "new": one("SELECT COUNT(*) FROM orders WHERE status = 'new'"),
        "hot": one("SELECT COUNT(*) FROM orders WHERE is_hot = 1"),
        "success": one("SELECT COUNT(*) FROM orders WHERE status = 'success'"),
        "cancelled": one("SELECT COUNT(*) FROM orders WHERE status IN ('cancelled', 'cancelled_by_client')"),
    }

    conn.close()
    return data


def tariff_distribution() -> Dict[str, int]:
    conn = get_connection()
    cur = conn.cursor()
    cur.execute(
        """
        SELECT tariff, COUNT(*) AS cnt
        FROM businesses
        WHERE name IS NOT NULL AND TRIM(name) != ''
        GROUP BY tariff
        """
    )
    rows = cur.fetchall()
    conn.close()
    return {str(row["tariff"] or "free").lower(): int(row["cnt"] or 0) for row in rows}


# Побудова текстів

def build_admin_home_text(admin_id: int) -> str:
    stats = get_global_stats()
    return (
        "🧑‍💻 <b>Адмін-панель ClientDesk AI</b>\n\n"
        f"Ваш ID: <code>{h(admin_id)}</code>\n\n"
        "<b>Огляд проекту</b>\n"
        f"🏢 Бізнесів: <b>{stats['businesses']}</b>\n"
        f"📋 Заявок: <b>{stats['orders']}</b>\n"
        f"🆕 Нових: <b>{stats['new']}</b>\n"
        f"🔥 Гарячих: <b>{stats['hot']}</b>\n"
        f"✅ Успішних: <b>{stats['success']}</b>\n\n"
        "<b>Команди</b>\n"
        "<code>/businesses</code> - список бізнесів\n"
        "<code>/biz КОД</code> - картка бізнесу\n"
        "<code>/settariff КОД start</code> - змінити тариф\n"
        "<code>/banbiz КОД</code> - видалити бізнес\n"
        "<code>/statsall</code> - загальна статистика"
    )


def build_business_card(business: Dict) -> str:
    orders = count_business_orders(int(business["id"]))
    usage = get_tariff_usage(business)
    tariff = usage["stored_tariff"]
    limit = usage["limit"]
    used = usage["used"]
    left = usage["remaining"]
    subscription_line = ""
    if usage["is_paid"]:
        if usage["active"]:
            subscription_line = f"<i>Активний до: <b>{h(display_datetime(usage.get('expires_at')) or '-')}</b></i>\n"
        else:
            subscription_line = f"<i>Підписка закінчилась: <b>{h(display_datetime(usage.get('expires_at')) or '-')}</b></i>\n"

    return (
        "🏢 <b>Картка бізнесу</b>\n\n"
        f"<b>{h(business.get('name') or '-')}</b>\n"
        f"Код: <code>{h(business.get('link_code') or '-')}</code>\n"
        f"Owner ID: <code>{h(business.get('owner_id') or '-')}</code>\n\n"
        f"Місто: <b>{h(business.get('city') or '-')}</b>\n"
        f"Ніша: <b>{h(business.get('niche') or '-')}</b>\n"
        f"Послуги: {h(business.get('services') or '-')}\n\n"
        f"Тариф: <b>{h(tariff_name(tariff))}</b>\n"
        f"{subscription_line}"
        f"Ліміт: <b>{used}/{limit}</b> заявок\n"
        f"Залишилось: <b>{left}</b>\n\n"
        "<b>Заявки</b>\n"
        f"Всього: <b>{orders['total']}</b>\n"
        f"Нові: <b>{orders['new']}</b>\n"
        f"Гарячі: <b>{orders['hot']}</b>\n"
        f"Успішні: <b>{orders['success']}</b>\n"
        f"Скасовані: <b>{orders['cancelled']}</b>\n\n"
        "<b>Швидкі команди</b>\n"
        f"<code>/settariff {h(business.get('link_code'))} start</code>\n"
        f"<code>/settariff {h(business.get('link_code'))} business</code>\n"
        f"<code>/banbiz {h(business.get('link_code'))}</code>"
    )


def build_businesses_list_text() -> str:
    businesses = list_businesses_admin(50)

    if not businesses:
        return (
            "🏢 <b>Бізнеси</b>\n\n"
            "Поки немає зареєстрованих бізнесів."
        )

    text = "🏢 <b>Бізнеси ClientDesk AI</b>\n\n"

    for business in businesses:
        name = business.get("name") or "-"
        code = business.get("link_code") or "-"
        city = business.get("city") or "-"
        usage = get_tariff_usage(business)
        tariff = tariff_name(usage["stored_tariff"])
        if usage["expired"] and usage["is_paid"]:
            tariff += " (закінчився)"
        used = usage["used"]
        limit = usage["limit"]

        text += (
            f"<b>{h(name)}</b>\n"
            f"Код: <code>{h(code)}</code> | {h(city)} | {h(tariff)} | {used}/{limit}\n"
            f"<code>/biz {h(code)}</code>\n\n"
        )

    if len(text) > 3900:
        text = text[:3800] + "\n\n...список обрізано. Використовуйте <code>/biz КОД</code>."

    return text


def build_stats_text() -> str:
    stats = get_global_stats()
    dist = tariff_distribution()

    tariff_lines = []
    for key in ["free", "start", "business", "vip"]:
        tariff_lines.append(f"{tariff_name(key)}: <b>{dist.get(key, 0)}</b>")

    return (
        "📊 <b>Статистика ClientDesk AI</b>\n\n"
        "<b>Проект</b>\n"
        f"Бізнесів: <b>{stats['businesses']}</b>\n"
        f"Заявок всього: <b>{stats['orders']}</b>\n"
        f"Нових заявок: <b>{stats['new']}</b>\n"
        f"Гарячих лідів: <b>{stats['hot']}</b>\n"
        f"Успішних: <b>{stats['success']}</b>\n"
        f"Скасованих: <b>{stats['cancelled']}</b>\n\n"
        "<b>Тарифи</b>\n"
        + "\n".join(tariff_lines)
    )


def build_tariffs_text() -> str:
    return (
        "💼 <b>Тарифи</b>\n\n"
        f"FREE - {h(tariff_price('free'))}\n"
        f"START - {h(tariff_price('start'))}\n"
        f"BUSINESS - {h(tariff_price('business'))}\n"
        f"VIP - {h(tariff_price('vip'))}\n\n"
        "<b>Змінити тариф бізнесу:</b>\n"
        "<code>/settariff КОД start</code>\n"
        "<code>/settariff КОД business</code>\n"
        "<code>/settariff КОД vip</code>\n"
        "<code>/settariff КОД free</code>"
    )


def build_owner_tariff_message(business: Dict, old_tariff: str, new_tariff: str) -> str:
    new_name = tariff_name(new_tariff)
    limit = tariff_limit(new_tariff)
    price = tariff_price(new_tariff)
    usage = get_tariff_usage(business)
    subscription_line = ""
    if usage["is_paid"]:
        subscription_line = f"<i>Підписка активна до: <b>{h(display_datetime(usage.get('expires_at')) or '-')}</b></i>\n"

    if new_tariff == "free":
        title = "💼 <b>Тариф оновлено</b>"
        congrats = "Ваш тариф змінено на безкоштовний план."
    else:
        title = "🎉 <b>Вітаємо з покупкою!</b>"
        congrats = "Ваш тариф успішно активовано. Тепер ClientDesk AI працює для вашого бізнесу з новими можливостями."

    return (
        f"{title}\n\n"
        f"🏢 Бізнес: <b>{h(business.get('name') or '-')}</b>\n"
        f"📍 Місто: <b>{h(business.get('city') or '-')}</b>\n\n"
        f"Попередній тариф: <b>{h(tariff_name(old_tariff))}</b>\n"
        f"Новий тариф: <b>{h(new_name)}</b>\n"
        f"Ліміт заявок: <b>{limit}</b>\n"
        f"{subscription_line}"
        f"Умови: {h(price)}\n\n"
        f"{congrats}\n\n"
        "Щоб отримувати максимум користі, перевірте AI-базу знань, посилання для клієнтів і заявки в кабінеті бізнесу."
    )


# Команди

async def myid_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user = update.effective_user
    user_id = user.id if user else None

    if is_admin_id(user_id):
        await update.effective_message.reply_text(
            f"🧑‍💻 <b>Ваш Telegram ID:</b> <code>{h(user_id)}</code>\n\n"
            "✅ Ви маєте доступ до адмін-панелі.\n"
            "Команда: <code>/admin</code>",
            parse_mode="HTML"
        )
        return

    await update.effective_message.reply_text(
        f"<b>Ваш Telegram ID:</b> <code>{h(user_id)}</code>",
        parse_mode="HTML"
    )


async def admin_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id if update.effective_user else None

    if not is_admin_id(user_id):
        await deny_access(update)
        return

    await update.effective_message.reply_text(
        build_admin_home_text(user_id),
        parse_mode="HTML",
        reply_markup=admin_keyboard()
    )


async def businesses_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id if update.effective_user else None

    if not is_admin_id(user_id):
        await deny_access(update)
        return

    await update.effective_message.reply_text(
        build_businesses_list_text(),
        parse_mode="HTML",
        reply_markup=back_admin_keyboard()
    )


async def biz_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id if update.effective_user else None

    if not is_admin_id(user_id):
        await deny_access(update)
        return

    if not context.args:
        await update.effective_message.reply_text(
            "Напишіть код бізнесу:\n<code>/biz КОД</code>",
            parse_mode="HTML"
        )
        return

    code = normalize_code(context.args[0])
    business = get_business_by_code_admin(code)

    if not business:
        await update.effective_message.reply_text(
            f"❌ Бізнес з кодом <code>{h(code)}</code> не знайдено.",
            parse_mode="HTML"
        )
        return

    await update.effective_message.reply_text(
        build_business_card(business),
        parse_mode="HTML",
        reply_markup=back_admin_keyboard()
    )


async def statsall_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id if update.effective_user else None

    if not is_admin_id(user_id):
        await deny_access(update)
        return

    await update.effective_message.reply_text(
        build_stats_text(),
        parse_mode="HTML",
        reply_markup=back_admin_keyboard()
    )


async def settariff_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id if update.effective_user else None

    if not is_admin_id(user_id):
        await deny_access(update)
        return

    if len(context.args) < 2:
        await update.effective_message.reply_text(
            "💼 <b>Зміна тарифу</b>\n\n"
            "Формат:\n"
            "<code>/settariff КОД start</code>\n"
            "<code>/settariff КОД business</code>\n"
            "<code>/settariff КОД vip</code>\n"
            "<code>/settariff КОД free</code>",
            parse_mode="HTML"
        )
        return

    code = normalize_code(context.args[0])
    new_tariff = normalize_tariff(context.args[1])

    if not new_tariff:
        await update.effective_message.reply_text(
            "❌ Невідомий тариф.\n\n"
            "Доступні: <code>free</code>, <code>start</code>, <code>business</code>, <code>vip</code>.",
            parse_mode="HTML"
        )
        return

    old_business = get_business_by_code_admin(code)
    if not old_business:
        await update.effective_message.reply_text(
            f"❌ Бізнес з кодом <code>{h(code)}</code> не знайдено.",
            parse_mode="HTML"
        )
        return

    old_tariff = str(old_business.get("tariff") or "free").lower()
    updated_business = update_business_tariff(code, new_tariff)

    if not updated_business:
        await update.effective_message.reply_text(
            "❌ Не вдалося змінити тариф. Спробуйте ще раз.",
            parse_mode="HTML"
        )
        return

    owner_id = updated_business.get("owner_id")
    owner_notified = False

    if owner_id:
        try:
            await context.bot.send_message(
                chat_id=int(owner_id),
                text=build_owner_tariff_message(updated_business, old_tariff, new_tariff),
                parse_mode="HTML"
            )
            owner_notified = True
        except Exception:
            owner_notified = False

    confirmation_text = (
        "✅ <b>Тариф успішно змінено</b>\n\n"
        f"🏢 Бізнес: <b>{h(updated_business.get('name') or '-')}</b>\n"
        f"Код: <code>{h(updated_business.get('link_code') or '-')}</code>\n\n"
        f"Було: <b>{h(tariff_name(old_tariff))}</b>\n"
        f"Стало: <b>{h(tariff_name(new_tariff))}</b>\n"
        f"Ліміт: <b>{tariff_limit(new_tariff)}</b> заявок\n\n"
    )

    if is_paid_tariff(new_tariff):
        confirmation_text += f"<i>Активний до: <b>{h(display_datetime(updated_business.get('subscription_expires_at')) or '-')}</b></i>\n\n"

    confirmation_text += f"Повідомлення власнику: <b>{'надіслано ✅' if owner_notified else 'не надіслано ⚠️'}</b>"

    await update.effective_message.reply_text(
        confirmation_text,
        parse_mode="HTML",
        reply_markup=back_admin_keyboard()
    )


async def banbiz_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id if update.effective_user else None

    if not is_admin_id(user_id):
        await deny_access(update)
        return

    if not context.args:
        await update.effective_message.reply_text(
            "Напишіть код бізнесу:\n<code>/banbiz КОД</code>",
            parse_mode="HTML"
        )
        return

    code = normalize_code(context.args[0])
    business = get_business_by_code_admin(code)

    if not business:
        await update.effective_message.reply_text(
            f"❌ Бізнес з кодом <code>{h(code)}</code> не знайдено.",
            parse_mode="HTML"
        )
        return

    keyboard = InlineKeyboardMarkup([
        [InlineKeyboardButton("🗑 Так, видалити бізнес", callback_data=f"admin_ban_confirm_{code}")],
        [InlineKeyboardButton("↩️ Скасувати", callback_data="admin_home")]
    ])

    await update.effective_message.reply_text(
        "⚠️ <b>Підтвердіть видалення бізнесу</b>\n\n"
        f"Бізнес: <b>{h(business.get('name') or '-')}</b>\n"
        f"Код: <code>{h(code)}</code>\n"
        f"Owner ID: <code>{h(business.get('owner_id') or '-')}</code>\n\n"
        "Буде видалено бізнес, його заявки та сесії.\n"
        "Цю дію не можна швидко відкотити без backup.",
        parse_mode="HTML",
        reply_markup=keyboard
    )


# Колбеки

async def admin_callback(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    user_id = query.from_user.id if query and query.from_user else None

    if not is_admin_id(user_id):
        await answer_denied_callback(update)
        return

    await query.answer()
    data = query.data

    if data == "admin_home":
        await query.edit_message_text(
            build_admin_home_text(user_id),
            parse_mode="HTML",
            reply_markup=admin_keyboard()
        )
        return

    if data == "admin_businesses":
        await query.edit_message_text(
            build_businesses_list_text(),
            parse_mode="HTML",
            reply_markup=back_admin_keyboard()
        )
        return

    if data == "admin_stats":
        await query.edit_message_text(
            build_stats_text(),
            parse_mode="HTML",
            reply_markup=back_admin_keyboard()
        )
        return

    if data == "admin_tariffs":
        await query.edit_message_text(
            build_tariffs_text(),
            parse_mode="HTML",
            reply_markup=back_admin_keyboard()
        )
        return

    if data.startswith("admin_ban_confirm_"):
        code = normalize_code(data.replace("admin_ban_confirm_", ""))
        ok, business = delete_business_by_code(code)

        if not ok or not business:
            await query.edit_message_text(
                f"❌ Бізнес з кодом <code>{h(code)}</code> не знайдено або вже видалений.",
                parse_mode="HTML",
                reply_markup=back_admin_keyboard()
            )
            return

        owner_id = business.get("owner_id")
        if owner_id:
            try:
                await context.bot.send_message(
                    chat_id=int(owner_id),
                    text=(
                        "⚠️ <b>Бізнес відключено</b>\n\n"
                        f"Бізнес: <b>{h(business.get('name') or '-')}</b>\n\n"
                        "Доступ до кабінету ClientDesk AI для цього бізнесу більше не активний."
                    ),
                    parse_mode="HTML"
                )
            except Exception:
                pass

        await query.edit_message_text(
            "🗑 <b>Бізнес видалено</b>\n\n"
            f"Назва: <b>{h(business.get('name') or '-')}</b>\n"
            f"Код: <code>{h(code)}</code>",
            parse_mode="HTML",
            reply_markup=back_admin_keyboard()
        )
        return

    await query.edit_message_text(
        build_admin_home_text(user_id),
        parse_mode="HTML",
        reply_markup=admin_keyboard()
    )


# Реєстрація хендлерів

def register_admin_handlers(app) -> None:
    app.add_handler(CommandHandler("myid", myid_command))
    app.add_handler(CommandHandler("admin", admin_command))
    app.add_handler(CommandHandler("businesses", businesses_command))
    app.add_handler(CommandHandler("biz", biz_command))
    app.add_handler(CommandHandler("settariff", settariff_command))
    app.add_handler(CommandHandler("banbiz", banbiz_command))
    app.add_handler(CommandHandler("statsall", statsall_command))
    app.add_handler(CallbackQueryHandler(admin_callback, pattern="^admin_"))
