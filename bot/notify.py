import html
import asyncio

from telegram import InlineKeyboardButton, InlineKeyboardMarkup
from telegram.error import BadRequest, Forbidden

from bot.db import (
    display_datetime,
    get_business_by_id,
    get_due_link_reminders,
    get_orders,
    get_tariff_usage,
    mark_link_reminder_status,
    parse_datetime,
    schedule_link_reminder_record,
)


STATUS_LABELS = {
    "new": "🆕 Нова",
    "in_progress": "🔄 В роботі",
    "contacted": "📞 Зв'язались",
    "success": "✅ Успішно",
    "rejected": "❌ Відмова",
    "cancelled": "🚫 Скасовано клієнтом",
    "cancelled_by_client": "🚫 Скасовано клієнтом",
}

CLIENT_STATUS_TEXTS = {
    "in_progress": (
        "🔄 <b>Ваша заявка вже в роботі</b>\n\n"
        "Менеджер переглянув звернення і скоро зв'яжеться з вами."
    ),
    "contacted": (
        "📞 <b>Статус заявки оновлено</b>\n\n"
        "Менеджер позначив, що з вами вже зв'язались."
    ),
    "success": (
        "✅ <b>Заявку успішно оброблено</b>\n\n"
        "Дякуємо за звернення! 🙏"
    ),
    "rejected": (
        "❌ <b>Заявку відхилено</b>\n\n"
        "На жаль, бізнес не зможе виконати це звернення."
    ),
}

FIELD_CONFIRM_LABELS = {
    "device": "Пристрій",
    "problem": "Проблему",
    "car": "Авто",
    "service": "Послугу",
    "format": "Формат",
    "shipping_city": "Населений пункт відправки",
    "district": "Локацію",
    "urgency": "Терміновість",
    "name": "Ім'я",
    "phone": "Телефон",
}


def h(value) -> str:
    if value is None:
        return ""
    return html.escape(str(value))


def get_order_client_chat_id(order: dict):
    possible_keys = [
        "client_id", "client_user_id", "client_telegram_id",
        "telegram_id", "user_id", "client_chat_id", "chat_id"
    ]

    for key in possible_keys:
        value = order.get(key)
        if value:
            try:
                return int(value)
            except Exception:
                continue

    return None


LINK_REMINDER_DELAY_SECONDS = 60 * 60
LINK_REMINDER_WORKER_INTERVAL_SECONDS = 60
ACTIVE_ORDER_STATUSES = {"new", "in_progress", "contacted"}


def build_client_requests_paused_text(business: dict) -> str:
    return (
        f"ℹ️ <b>{h(business.get('name') or 'Бізнес')} зараз тимчасово приймає заявки через менеджера.</b>\n\n"
        "Спробуйте, будь ласка, трохи пізніше або зв'яжіться з бізнесом напряму."
    )


def _order_matches_client(order: dict, user_id: int, username: str = "") -> bool:
    client_chat_id = get_order_client_chat_id(order)
    if client_chat_id and int(client_chat_id) == int(user_id):
        return True

    if username:
        order_username = str(order.get("client_username") or "").lstrip("@").casefold()
        return order_username == str(username).lstrip("@").casefold()

    return False


def _client_has_order_after_entry(business_id: int, user_id: int, username: str, entered_at: str) -> bool:
    entered_dt = parse_datetime(entered_at)

    try:
        orders = get_orders(business_id)
    except Exception:
        return False

    for order in orders:
        if not _order_matches_client(order, user_id, username):
            continue

        status = str(order.get("status") or "new").lower()
        if status in ACTIVE_ORDER_STATUSES:
            return True

        created_at = parse_datetime(order.get("created_at"))
        if entered_dt and created_at and created_at >= entered_dt:
            return True

    return False


def schedule_client_link_reminder(context, chat_id: int, user_id: int, business: dict, username: str = "") -> bool:
    business_id = int(business["id"])
    try:
        record = schedule_link_reminder_record(
            business_id=business_id,
            chat_id=chat_id,
            user_id=int(user_id),
            username=username or "",
            delay_seconds=LINK_REMINDER_DELAY_SECONDS,
        )
    except Exception as e:
        print(f"Client link reminder schedule failed: {e}")
        return False

    application = getattr(context, "application", None)
    start_link_reminder_worker(application)

    return bool(record)


def start_link_reminder_worker(application) -> bool:
    if not application:
        return False

    bot_data = getattr(application, "bot_data", None)
    if bot_data is not None and bot_data.get("_link_reminder_worker_started"):
        return False

    if bot_data is not None:
        bot_data["_link_reminder_worker_started"] = True

    if getattr(application, "running", False) and hasattr(application, "create_task"):
        application.create_task(client_link_reminder_worker(application.bot))
    else:
        asyncio.create_task(client_link_reminder_worker(application.bot))

    return True


async def client_link_reminder_worker(bot):
    while True:
        try:
            await send_due_client_link_reminders(bot)
        except Exception as e:
            print(f"Client link reminder worker failed: {e}")

        await asyncio.sleep(LINK_REMINDER_WORKER_INTERVAL_SECONDS)


async def send_due_client_link_reminders(bot):
    for reminder in get_due_link_reminders():
        await send_client_link_reminder(bot, reminder)


async def send_client_link_reminder(bot, reminder: dict):
    reminder_id = int(reminder["id"])
    business_id = int(reminder["business_id"])
    chat_id = int(reminder["chat_id"])
    user_id = int(reminder["user_id"])
    username = reminder.get("username") or ""
    entered_at = reminder.get("entered_at") or ""

    if _client_has_order_after_entry(business_id, user_id, username, entered_at):
        mark_link_reminder_status(reminder_id, "cancelled")
        return

    business = get_business_by_id(business_id)
    if not business or not business.get("name"):
        mark_link_reminder_status(reminder_id, "skipped")
        return

    business_name = business.get("name") or "цей бізнес"
    text = (
        f"💡 <b>Хочете, щоб {h(business_name)} швидше з вами зв'язались?</b>\n\n"
        "Залиште заявку: коротко опишіть, що потрібно, і додайте контактний телефон. "
        "AI поставить тільки потрібні уточнюючі питання."
    )

    try:
        await bot.send_message(
            chat_id=chat_id,
            text=text,
            parse_mode="HTML",
            reply_markup=InlineKeyboardMarkup([
                [InlineKeyboardButton("📝 Залишити заявку", callback_data=f"client_order_biz_{business_id}")]
            ])
        )
        mark_link_reminder_status(reminder_id, "sent")
    except Forbidden:
        mark_link_reminder_status(reminder_id, "failed")
    except BadRequest:
        mark_link_reminder_status(reminder_id, "failed")


async def notify_owner_about_order_cancel(context, business: dict, order: dict):
    text = (
        f"🚫 <b>Клієнт скасував заявку - {h(business.get('name'))}</b>\n\n"
        f"📋 Заявка #{h(order.get('id'))}\n"
        f"👤 Клієнт: {h(order.get('client_name') or '-')}\n"
        f"📞 Телефон: {h(order.get('client_phone') or '-')}\n"
        f"💬 Telegram: {h('@' + order.get('client_username') if order.get('client_username') else '-')}\n\n"
        "Статус заявки змінено на: <b>🚫 Скасовано клієнтом</b>"
    )

    keyboard = [
        [InlineKeyboardButton("📋 Відкрити заявку", callback_data=f"order_{order.get('id')}")],
        [InlineKeyboardButton("🏠 Головний екран", callback_data="menu_back")],
    ]

    try:
        await context.bot.send_message(
            chat_id=business["owner_id"],
            text=text,
            parse_mode="HTML",
            reply_markup=InlineKeyboardMarkup(keyboard)
        )
    except Forbidden:
        pass
    except BadRequest:
        pass


async def notify_owner_about_tariff_pause(context, business: dict):
    usage = get_tariff_usage(business)
    reason = "Підписка закінчилась" if usage["expired"] and usage["is_paid"] else "Ліміт заявок використано"
    expires = display_datetime(usage.get("expires_at")) or "-"

    text = (
        "⚠️ <b>Нову заявку не прийнято через тариф</b>\n\n"
        f"Причина: <b>{h(reason)}</b>\n"
        f"Тариф: <b>{h(str(usage.get('stored_tariff') or 'free').upper())}</b>\n"
        f"Використано: <b>{h(usage.get('used'))}/{h(usage.get('limit'))}</b> заявок\n"
        f"<i>Активний до: <b>{h(expires)}</b></i>\n\n"
        "Клієнту показано нейтральне повідомлення без згадки про ліміт."
    )

    try:
        await context.bot.send_message(
            chat_id=business["owner_id"],
            text=text,
            parse_mode="HTML"
        )
    except Forbidden:
        pass
    except BadRequest:
        pass
    except Exception:
        pass


async def notify_owner_about_order_edit(
    context,
    business: dict,
    old_order: dict,
    updated_order: dict,
    changed_keys: list,
    order_detail_text: str = "",
):
    changed_labels = []
    for key in changed_keys:
        label = FIELD_CONFIRM_LABELS.get(key)
        if label:
            changed_labels.append(label)

    changed_text = ", ".join(changed_labels) if changed_labels else "дані заявки"

    text = (
        f"✏️ <b>Клієнт оновив заявку - {h(business.get('name'))}</b>\n\n"
        f"📋 Заявка #{h(updated_order.get('id'))}\n"
        f"🔄 Змінено: <b>{h(changed_text)}</b>\n\n"
        f"{order_detail_text}"
    )

    keyboard = [
        [InlineKeyboardButton("📋 Відкрити заявку", callback_data=f"order_{updated_order.get('id')}")],
        [InlineKeyboardButton("🏠 Головний екран", callback_data="menu_back")],
    ]

    await context.bot.send_message(
        chat_id=business["owner_id"],
        text=text,
        parse_mode="HTML",
        reply_markup=InlineKeyboardMarkup(keyboard)
    )


async def notify_client_about_status(context, order: dict, business: dict, status: str):
    client_chat_id = get_order_client_chat_id(order)

    if not client_chat_id:
        print("Client status notify skipped: no client chat id in order")
        return

    text = CLIENT_STATUS_TEXTS.get(
        status,
        f"📊 <b>Статус вашої заявки оновлено</b>\n\nНовий статус: {h(STATUS_LABELS.get(status, status))}"
    )

    text += f"\n\n🏢 Бізнес: <b>{h(business.get('name'))}</b>"
    text += f"\n📋 Заявка #{h(order.get('id'))}"

    try:
        await context.bot.send_message(
            chat_id=client_chat_id,
            text=text,
            parse_mode="HTML",
            reply_markup=InlineKeyboardMarkup([
                [InlineKeyboardButton("📋 Мої заявки", callback_data=f"client_my_orders_{business.get('id')}")]
            ])
        )
    except Forbidden:
        print("Client status notify failed: bot blocked by client")
    except Exception as e:
        print(f"Client status notify failed: {e}")


async def notify_owner_about_new_order(context, data: dict, username: str, business=None):
    if business is None:
        business = get_business_by_id(data.get("business_id"))

    if not business:
        return

    collected = data.get("collected", {})
    urgency = collected.get("urgency", "-")
    is_hot = bool(urgency and "Сьогодні" in urgency)
    emoji = "🔥" if is_hot else "📋"

    field_labels = {
        "name": "👤 Клієнт",
        "phone": "📞 Телефон",
        "device": "📱 Пристрій",
        "car": "🚗 Авто",
        "service": "🔧 Послуга",
        "problem": "❗ Проблема",
        "format": "💻 Формат",
        "shipping_city": "📦 Населений пункт відправки",
        "urgency": "⏰ Терміновість",
        "district": "📍 Район",
    }

    lines = [f"{emoji} <b>Нова заявка - {h(business.get('name'))}</b>\n"]

    for key, label in field_labels.items():
        val = collected.get(key)
        if val:
            lines.append(f"{h(label)}: {h(val)}")

    if username:
        lines.append(f"💬 Telegram: @{h(username)}")

    lines.append(f"\n🤖 AI-оцінка: {h(data.get('ai_comment', '-'))}")

    if is_hot:
        lines.append("\n⚡️ Рекомендація: зв'яжіться протягом 10 хвилин!")

    keyboard = [
        [InlineKeyboardButton("📋 Відкрити заявки", callback_data="menu_orders")],
        [InlineKeyboardButton("🏠 Головний екран", callback_data="menu_back")],
    ]

    await context.bot.send_message(
        chat_id=business["owner_id"],
        text="\n".join(lines),
        parse_mode="HTML",
        reply_markup=InlineKeyboardMarkup(keyboard)
    )
