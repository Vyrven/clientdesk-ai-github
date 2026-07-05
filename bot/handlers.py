import html
import io
import re
import sqlite3
import zipfile
from datetime import datetime
from difflib import SequenceMatcher

from telegram import Update, InlineKeyboardButton, InlineKeyboardMarkup
from telegram.ext import ContextTypes
from telegram.error import BadRequest

from bot.db import (
    get_business, get_business_by_id, create_business,
    update_business, save_session, get_session,
    get_business_by_code, search_businesses,
    save_order, get_orders, get_recent_client_orders, count_orders, get_orders_page, get_order_for_business, update_order_status,
    check_tariff_limit, display_datetime, get_stats, get_tariff_usage, DB_PATH
)
from bot.help import build_support_button
from bot.niches import get_fields
from bot.ai import qualify_order, analyze_client_request, answer_business_question
from bot.notify import (
    build_client_requests_paused_text,
    notify_client_about_status,
    notify_owner_about_new_order,
    notify_owner_about_order_cancel,
    notify_owner_about_order_edit,
    notify_owner_about_tariff_pause,
    schedule_client_link_reminder,
)
from config import TARIFF_PRICES


NICHES = ["Ремонт техніки", "Автосервіс"]

REPAIR_NICHE = "Ремонт техніки"
WORK_MODE_OPTIONS = {
    "offline": {
        "label": "Тільки офлайн",
        "order_format": "Офлайн",
        "button": "🏢 Тільки офлайн",
    },
    "online": {
        "label": "Тільки онлайн",
        "order_format": "Онлайн",
        "button": "💻 Тільки онлайн",
    },
    "both": {
        "label": "Офлайн і онлайн",
        "order_format": "",
        "button": "🔁 Офлайн і онлайн",
    },
}

NICHE_BUTTON_LABELS = {
    "Ремонт техніки": "🛠 <b>Ремонт техніки</b>",
    "Автосервіс": "🚗 <b>Автосервіс</b>",
}

NICHE_BUTTON_TEXTS = {
    "Ремонт техніки": "🛠 Ремонт техніки",
    "Автосервіс": "🚗 Автосервіс",
}

REGISTRATION_SERVICE_EXAMPLES = {
    "Ремонт техніки": "Ремонт iPhone, Заміна екрану, Чистка після води, Ремонт ноутбуків",
    "Автосервіс": "Діагностика авто, Заміна масла, Ремонт ходової, Шиномонтаж",
}

REGISTRATION_NICHE_HINTS = {
    "Ремонт техніки": "AI буде збирати модель пристрою, проблему, терміновість, ім'я та телефон.",
    "Автосервіс": "AI буде збирати авто, проблему, терміновість, локацію, ім'я та телефон.",
}

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

ACTIVE_ORDER_STATUSES = {"new", "in_progress", "contacted"}
CANCELLED_ORDER_STATUSES = {"cancelled", "cancelled_by_client"}

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

BTN_BACK_START = InlineKeyboardButton("🔙 На початок", callback_data="back_to_start")
BTN_BACK_MENU = InlineKeyboardButton("🔙 Головне меню", callback_data="menu_back")
BTN_BACK_ORDERS = InlineKeyboardButton("🔙 До заявок", callback_data="menu_orders")
ORDERS_PAGE_SIZE = 10
ORDERS_EXPORT_MIN_COUNT = 5
ORDERS_EXPORT_TARIFFS = {"business", "vip"}


# Базові хелпери

def h(value) -> str:
    if value is None:
        return ""
    return html.escape(str(value))


def short_value(value, limit: int = 34) -> str:
    text = str(value or "").strip()
    if len(text) <= limit:
        return text
    return text[: max(0, limit - 1)].rstrip() + "…"


def xml_escape(value) -> str:
    text = re.sub(r"[\x00-\x08\x0B\x0C\x0E-\x1F]", "", str(value or ""))
    return html.escape(text, quote=True)


def excel_column_name(index: int) -> str:
    name = ""
    while index:
        index, remainder = divmod(index - 1, 26)
        name = chr(65 + remainder) + name
    return name


def xlsx_text_cell(row: int, col: int, value, style: int = 0) -> str:
    cell_ref = f"{excel_column_name(col)}{row}"
    style_attr = f' s="{style}"' if style else ""
    text = xml_escape(value)
    return f'<c r="{cell_ref}" t="inlineStr"{style_attr}><is><t>{text}</t></is></c>'


def build_xlsx_sheet_xml(rows: list) -> str:
    sheet_rows = []
    for row_index, values in enumerate(rows, start=1):
        cells = [
            xlsx_text_cell(row_index, col_index, value, style=1 if row_index == 1 else 0)
            for col_index, value in enumerate(values, start=1)
        ]
        sheet_rows.append(f'<row r="{row_index}">{"".join(cells)}</row>')

    last_row = max(1, len(rows))
    last_col = excel_column_name(len(rows[0]) if rows else 1)

    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
        '<sheetViews><sheetView workbookViewId="0">'
        '<pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/>'
        '</sheetView></sheetViews>'
        '<cols>'
        '<col min="1" max="1" width="8" customWidth="1"/>'
        '<col min="2" max="3" width="18" customWidth="1"/>'
        '<col min="4" max="8" width="22" customWidth="1"/>'
        '<col min="9" max="16" width="28" customWidth="1"/>'
        '</cols>'
        f'<sheetData>{"".join(sheet_rows)}</sheetData>'
        f'<autoFilter ref="A1:{last_col}{last_row}"/>'
        '</worksheet>'
    )


def build_orders_xlsx_bytes(business: dict, orders: list) -> bytes:
    headers = [
        "ID", "Дата", "Статус", "Гаряча",
        "Клієнт", "Телефон", "Telegram",
        "Послуга", "Пристрій", "Авто", "Формат",
        "Населений пункт відправки", "Локація",
        "Проблема", "Терміновість", "AI-оцінка",
    ]

    rows = [headers]
    for order in orders:
        status = STATUS_LABELS.get(order.get("status"), order.get("status") or "-")
        username = f"@{order.get('client_username')}" if order.get("client_username") else ""
        rows.append([
            order.get("id") or "",
            display_datetime(order.get("created_at")) or "",
            status,
            "Так" if order.get("is_hot") else "",
            order.get("client_name") or "",
            order.get("client_phone") or "",
            username,
            order.get("service") or "",
            order.get("device") or "",
            order.get("car") or "",
            order.get("format") or "",
            order.get("shipping_city") or "",
            order.get("district") or "",
            order.get("problem") or "",
            order.get("urgency") or "",
            order.get("ai_comment") or "",
        ])

    now_iso = datetime.utcnow().strftime("%Y-%m-%dT%H:%M:%SZ")
    sheet_xml = build_xlsx_sheet_xml(rows)
    title = xml_escape(f"Заявки {business.get('name') or 'ClientDesk AI'}")

    files = {
        "[Content_Types].xml": (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
            '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
            '<Default Extension="xml" ContentType="application/xml"/>'
            '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
            '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
            '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>'
            '<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>'
            '<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>'
            '</Types>'
        ),
        "_rels/.rels": (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>'
            '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>'
            '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>'
            '</Relationships>'
        ),
        "xl/workbook.xml": (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
            'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
            '<sheets><sheet name="Заявки" sheetId="1" r:id="rId1"/></sheets>'
            '</workbook>'
        ),
        "xl/_rels/workbook.xml.rels": (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>'
            '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'
            '</Relationships>'
        ),
        "xl/worksheets/sheet1.xml": sheet_xml,
        "xl/styles.xml": (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
            '<fonts count="2">'
            '<font><sz val="11"/><name val="Calibri"/></font>'
            '<font><b/><sz val="11"/><name val="Calibri"/></font>'
            '</fonts>'
            '<fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>'
            '<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>'
            '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>'
            '<cellXfs count="2">'
            '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>'
            '<xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1"/>'
            '</cellXfs>'
            '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>'
            '</styleSheet>'
        ),
        "docProps/core.xml": (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" '
            'xmlns:dc="http://purl.org/dc/elements/1.1/" '
            'xmlns:dcterms="http://purl.org/dc/terms/" '
            'xmlns:dcmitype="http://purl.org/dc/dcmitype/" '
            'xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">'
            f'<dc:title>{title}</dc:title><dc:creator>ClientDesk AI</dc:creator>'
            f'<dcterms:created xsi:type="dcterms:W3CDTF">{now_iso}</dcterms:created>'
            f'<dcterms:modified xsi:type="dcterms:W3CDTF">{now_iso}</dcterms:modified>'
            '</cp:coreProperties>'
        ),
        "docProps/app.xml": (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" '
            'xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">'
            '<Application>ClientDesk AI</Application>'
            '</Properties>'
        ),
    }

    output = io.BytesIO()
    with zipfile.ZipFile(output, "w", zipfile.ZIP_DEFLATED) as archive:
        for path, content in files.items():
            archive.writestr(path, content)

    return output.getvalue()


def orders_export_available(business: dict, total: int) -> bool:
    usage = get_tariff_usage(business)
    tariff = str(usage.get("stored_tariff") or "free").lower()
    return total > ORDERS_EXPORT_MIN_COUNT and usage.get("active") and tariff in ORDERS_EXPORT_TARIFFS


def build_orders_export_locked_text(business: dict, total: int) -> str:
    usage = get_tariff_usage(business)
    tariff = str(usage.get("stored_tariff") or "free").upper()

    if total <= ORDERS_EXPORT_MIN_COUNT:
        reason = (
            f"Експорт з'явиться, коли буде більше "
            f"<b>{ORDERS_EXPORT_MIN_COUNT}</b> заявок."
        )
    elif usage.get("expired") and usage.get("is_paid"):
        reason = "Підписка закінчилась. Після продовження тарифу експорт знову відкриється."
    else:
        reason = "XLSX-експорт доступний з тарифу <b>BUSINESS</b>."

    return (
        "📥 <b>Експорт заявок у XLSX</b>\n\n"
        f"Заявок у бізнесі: <b>{h(total)}</b>\n"
        f"Поточний тариф: <b>{h(tariff)}</b>\n\n"
        f"{reason}\n\n"
        "Файл містить клієнта, телефон, Telegram, пристрій/авто, проблему, "
        "терміновість, статус та AI-оцінку."
    )


def parse_orders_page(callback_data: str) -> int:
    raw = str(callback_data or "")
    if not raw.startswith("menu_orders_page_"):
        return 1
    try:
        return max(1, int(raw.replace("menu_orders_page_", "", 1)))
    except ValueError:
        return 1


def build_orders_page_text(orders: list, total: int, page: int, total_pages: int) -> str:
    text = (
        "📋 <b>Заявки</b>\n"
        f"Сторінка: <b>{h(page)}/{h(total_pages)}</b> · Всього: <b>{h(total)}</b>\n\n"
    )

    for order in orders:
        status = STATUS_LABELS.get(order.get("status"), order.get("status", "-"))
        name = short_value(order.get("client_name") or "Клієнт", 24)
        service = short_value(order.get("service") or order.get("device") or order.get("car") or "-", 36)
        hot = "🔥 " if order.get("is_hot") else ""
        text += f"{hot}#{h(order.get('id'))} {h(name)} - {h(service)} - {h(status)}\n"

    if total > ORDERS_EXPORT_MIN_COUNT:
        text += "\n📥 XLSX-експорт доступний для тарифу <b>BUSINESS</b> і вище."

    return text


def build_orders_page_keyboard(orders: list, page: int, total_pages: int, business: dict = None, total: int = 0) -> InlineKeyboardMarkup:
    keyboard_rows = []

    for order in orders:
        hot = "🔥 " if order.get("is_hot") else ""
        name = short_value(order.get("client_name") or "Клієнт", 28)
        keyboard_rows.append([
            InlineKeyboardButton(
                f"{hot}#{order.get('id')} - {name}",
                callback_data=f"order_{order.get('id')}"
            )
        ])

    nav_row = []
    if page > 1:
        nav_row.append(InlineKeyboardButton("⬅️ Назад", callback_data=f"menu_orders_page_{page - 1}"))
    if page < total_pages:
        nav_row.append(InlineKeyboardButton("Далі ➡️", callback_data=f"menu_orders_page_{page + 1}"))
    if nav_row:
        keyboard_rows.append(nav_row)

    if total > ORDERS_EXPORT_MIN_COUNT:
        if business and orders_export_available(business, total):
            keyboard_rows.append([InlineKeyboardButton("📥 Завантажити XLSX", callback_data="menu_orders_export")])
        else:
            keyboard_rows.append([InlineKeyboardButton("🔒 XLSX-експорт з BUSINESS", callback_data="menu_orders_export_locked")])

    keyboard_rows.append([BTN_BACK_MENU])
    return InlineKeyboardMarkup(keyboard_rows)


def is_repair_niche(niche: str) -> bool:
    return str(niche or "").strip() == REPAIR_NICHE


def get_work_mode_label(mode: str) -> str:
    return WORK_MODE_OPTIONS.get(str(mode or "").strip(), WORK_MODE_OPTIONS["both"])["label"]


def get_work_mode_order_format(mode: str) -> str:
    return WORK_MODE_OPTIONS.get(str(mode or "").strip(), WORK_MODE_OPTIONS["both"])["order_format"]


def get_business_work_mode(business: dict) -> str:
    if not business or not is_repair_niche(business.get("niche")):
        return ""

    mode = str(business.get("work_mode") or "").strip()
    return mode if mode in WORK_MODE_OPTIONS else "both"


REPAIR_ONLINE_SHIPPING_FIELD = {
    "key": "shipping_city",
    "question": (
        "📦 З якого населеного пункту будете відправляти техніку?\n\n"
        "Наприклад: Київ, село Жашків, смт Бородянка або Луганськ."
    ),
}


def is_repair_online_order(niche: str, collected: dict) -> bool:
    return is_repair_niche(niche) and str((collected or {}).get("format") or "").strip() == "Онлайн"


def has_shipping_city(collected: dict) -> bool:
    return bool(str((collected or {}).get("shipping_city") or "").strip())


def needs_online_shipping_city(business: dict, data: dict) -> bool:
    collected = (data or {}).get("collected") or {}
    return is_repair_online_order((business or {}).get("niche"), collected) and not has_shipping_city(collected)


def get_order_fields(niche: str, collected: dict = None) -> list:
    fields = list(get_fields(niche))

    if is_repair_online_order(niche, collected or {}):
        fields = [REPAIR_ONLINE_SHIPPING_FIELD] + fields

    return fields


def get_next_order_question(niche: str, collected: dict):
    collected = collected or {}

    for field in get_order_fields(niche, collected):
        key = field.get("key")
        if key not in collected or str(collected.get(key) or "").strip() == "":
            return field

    return None


def build_work_mode_keyboard() -> InlineKeyboardMarkup:
    return InlineKeyboardMarkup([
        [InlineKeyboardButton(WORK_MODE_OPTIONS["offline"]["button"], callback_data="work_mode_offline")],
        [InlineKeyboardButton(WORK_MODE_OPTIONS["online"]["button"], callback_data="work_mode_online")],
        [InlineKeyboardButton(WORK_MODE_OPTIONS["both"]["button"], callback_data="work_mode_both")],
        [BTN_BACK_START],
    ])


def build_work_mode_registration_text(business: dict) -> str:
    return (
        "✅ <b>Ніша: Ремонт техніки</b>\n\n"
        "🔧 <b>Крок 3/5 - Як ви працюєте?</b>\n\n"
        "Оберіть формат, щоб AI правильно оформлював заявки клієнтів:\n\n"
        "🏢 <b>Тільки офлайн</b> - клієнт приносить техніку або звертається у сервіс.\n"
        "💻 <b>Тільки онлайн</b> - консультації/діагностика/підтримка онлайн.\n"
        "🔁 <b>Офлайн і онлайн</b> - клієнт сам обере формат під час заявки."
    )


def build_client_work_format_keyboard() -> InlineKeyboardMarkup:
    return InlineKeyboardMarkup([
        [
            InlineKeyboardButton("🏢 Офлайн", callback_data="workformat_offline"),
            InlineKeyboardButton("💻 Онлайн", callback_data="workformat_online"),
        ],
        [InlineKeyboardButton("🔙 Назад", callback_data="client_back")],
    ])


def build_client_work_format_text(business: dict) -> str:
    return (
        f"🔧 <b>{h(business.get('name') or 'Бізнес')}</b>\n\n"
        "Як вам зручніше оформити ремонт?\n\n"
        "🏢 <b>Офлайн</b> - передати техніку у сервіс або домовитись про візит.\n"
        "💻 <b>Онлайн</b> - консультація, діагностика або уточнення дистанційно."
    )


def build_online_shipping_city_text(invalid: bool = False) -> str:
    text = (
        "📦 <b>Для онлайн-ремонту</b> одразу напишіть населений пункт, "
        "з якого будете відправляти техніку."
    )

    if invalid:
        text += (
            "\n\n🔎 Поки не бачу назву населеного пункту.\n"
            "Напишіть місто, село, смт або селище, наприклад: "
            "<b>Луганськ</b>, <b>село Жашків</b>, <b>смт Бородянка</b>."
        )

    return text


def build_online_shipping_city_keyboard() -> InlineKeyboardMarkup:
    return InlineKeyboardMarkup([[InlineKeyboardButton("🔙 Назад", callback_data="client_back")]])


def apply_business_order_defaults(data: dict, business: dict) -> dict:
    if not data:
        data = {}

    mode = get_business_work_mode(business)
    order_format = get_work_mode_order_format(mode)

    if order_format:
        collected = data.get("collected") or {}
        collected["format"] = order_format
        data["collected"] = collected

    return data


def sanitize_telegram_html(text: str) -> str:
    """Робить AI-відповіді безпечними для Telegram HTML.
    OpenAI іноді повертає <br>, <p> або інші HTML-теги, які Telegram не підтримує.
    Через це бот падає з BadRequest: unsupported start tag.
    """
    if text is None:
        return ""

    text = str(text)
    text = re.sub(r"<\s*br\s*/?\s*>", "\n", text, flags=re.IGNORECASE)
    text = re.sub(r"<\s*/\s*p\s*>", "\n\n", text, flags=re.IGNORECASE)
    text = re.sub(r"<\s*p[^>]*>", "", text, flags=re.IGNORECASE)
    text = re.sub(r"<\s*/?\s*div[^>]*>", "\n", text, flags=re.IGNORECASE)

    # Спочатку екрануємо все, потім повертаємо лише прості теги, які Telegram точно підтримує.
    text = html.escape(text, quote=False)
    allowed = ["b", "i", "u", "code"]
    for tag in allowed:
        text = text.replace(f"&lt;{tag}&gt;", f"<{tag}>")
        text = text.replace(f"&lt;/{tag}&gt;", f"</{tag}>")

    # Захист від зайвих порожніх рядків.
    text = re.sub(r"\n{4,}", "\n\n\n", text).strip()
    return text


def digits_only(text: str) -> str:
    return "".join(ch for ch in str(text) if ch.isdigit())


def normalize_phone(text: str) -> str:
    text = str(text).strip()
    digits = digits_only(text)

    if text.startswith("+"):
        return "+" + digits

    return digits


def is_valid_phone(text: str) -> bool:
    raw = str(text).strip()
    digits = digits_only(raw)

    if not digits:
        return False

    if len(digits) < 9 or len(digits) > 13:
        return False

    letters = re.findall(r"[a-zA-Zа-яА-ЯіїєґІЇЄҐ]", raw)
    if len(letters) > 2:
        return False

    if raw.startswith("+"):
        return digits.startswith("380") and len(digits) == 12

    if digits.startswith("380"):
        return len(digits) == 12

    if digits.startswith("0"):
        return len(digits) == 10

    return False


def is_valid_name(text: str) -> bool:
    text = str(text).strip()

    if len(text) < 2 or len(text) > 40:
        return False

    if is_valid_phone(text):
        return False

    if digits_only(text):
        return False

    bad = ["тест", "test", "123", "???", "...", "не знаю", "хз", "hz", "ок", "да", "так"]
    if text.lower() in bad:
        return False

    return True


def is_valid_location(text: str) -> bool:
    text = str(text).strip()

    if len(text) < 2 or len(text) > 80:
        return False

    if is_valid_phone(text):
        return False

    if digits_only(text) and len(digits_only(text)) > 5:
        return False

    bad = ["тест", "test", "123", "???", "...", "не знаю", "хз", "hz"]
    if text.lower() in bad:
        return False

    return True


def is_valid_device(text: str) -> bool:
    text = str(text).strip()
    text_lower = text.lower()

    iphone_match = re.search(r"(?:iphone|айфон)\s*(\d{1,2})", text_lower)
    if iphone_match:
        try:
            model_num = int(iphone_match.group(1))
            if model_num > 20:
                return False
        except ValueError:
            pass

    if len(text) < 2 or len(text) > 80:
        return False

    if is_valid_phone(text):
        return False

    bad = ["тест", "test", "123", "???", "...", "не знаю", "хз", "hz"]
    if text.lower() in bad:
        return False

    return True


def looks_like_bad_problem(text: str) -> bool:
    text = str(text).strip()
    text_lower = text.lower()

    if len(text) < 4 or len(text) > 300:
        return True

    if is_valid_phone(text):
        return True

    bad_exact = [
        "не знаю", "хз", "hz", "???", "...", "123", "тест",
        "ок", "окей", "да", "так", "ні", "нет"
    ]

    if text_lower in bad_exact:
        return True

    if len(text.split()) == 1 and len(text) < 6:
        return True

    return False


def is_help_request(text: str) -> bool:
    text_lower = str(text).lower().strip()

    help_words = [
        "допомож", "помоги", "поможи", "склади", "составь",
        "сформулюй", "сформулируй", "професійно", "профессионально",
        "як написати", "как написать", "як правильно", "как правильно",
        "що писати", "что писать", "напиши за мене", "напиши за меня",
        "не знаю що", "не знаю что"
    ]

    return any(word in text_lower for word in help_words)


def is_question_message(text: str) -> bool:
    text_clean = str(text).strip()
    text_lower = text_clean.lower()

    if "?" in text_clean:
        return True

    question_starts = [
        "як ", "как ", "що ", "что ", "чому ", "почему ",
        "можна ", "можно ", "підкажіть", "подскажите",
        "скільки", "сколько", "де ", "где ", "коли ", "когда "
    ]

    return any(text_lower.startswith(start) for start in question_starts)


def normalize_urgency(text: str) -> str:
    """Нормалізація терміновості без плутанини:
    - сьогодні / терміново = гарячий лід
    - протягом кількох днів = окремий середній пріоритет
    - протягом тижня / до тижня = не терміново
    """
    text_lower = str(text).lower().strip()

    if not text_lower:
        return ""

    normal_words = [
        "не терміново", "не срочно", "коли буде час", "когда будет время",
        "без поспіху", "без спешки", "протягом тижня", "в течение недели",
        "до тижня", "до недели", "на тижні", "на неделе", "через тиждень", "через неделю"
    ]

    today_words = [
        "сьогодні", "сегодня", "терміново", "срочно", "дуже швидко",
        "прямо зараз", "зараз", "как можно быстрее", "якнайшвидше"
    ]

    few_days_words = [
        "протягом кількох днів", "в течение нескольких дней", "через кілька днів",
        "через несколько дней", "найближчі дні", "ближайшие дни", "на днях"
    ]

    tomorrow_words = ["завтра"]

    # Спочатку не терміново, щоб "протягом тижня" не падало у найближчі дні.
    if any(word in text_lower for word in normal_words):
        return "Не терміново"

    if any(word in text_lower for word in today_words):
        return "Сьогодні 🔥"

    if any(word in text_lower for word in few_days_words):
        return "Протягом кількох днів"

    if any(word in text_lower for word in tomorrow_words):
        return "Протягом кількох днів"

    return ""

def get_missing_field_keys(niche: str, collected: dict) -> list:
    missing = []

    for field in get_order_fields(niche, collected):
        key = field.get("key")
        val = collected.get(key)

        if val is None or str(val).strip() == "":
            missing.append(key)

    return missing


def get_question_for_field(niche: str, field_key: str, collected: dict = None) -> str:
    for field in get_order_fields(niche, collected or {}):
        if field.get("key") == field_key:
            return field.get("question", "Уточніть деталі:")

    return "Уточніть деталі:"


def get_field_type(niche: str, field_key: str, collected: dict = None) -> str:
    for field in get_order_fields(niche, collected or {}):
        if field.get("key") == field_key:
            return field.get("type", "text")

    return "text"


def get_step_info(niche: str, collected: dict, next_field_key: str = None) -> tuple:
    fields = get_order_fields(niche, collected)
    total = len(fields)

    if next_field_key:
        for index, field in enumerate(fields, start=1):
            if field.get("key") == next_field_key:
                return index, total

    done = sum(1 for field in fields if field.get("key") in collected and str(collected.get(field.get("key"))).strip())
    return min(done + 1, total), total


def validate_field(field_key: str, value: str, collected: dict) -> tuple:
    if field_key == "phone":
        if not is_valid_phone(value):
            return False, (
                "❌ Це не схоже на коректний номер.\n\n"
                "Напишіть номер у форматі:\n"
                "<b>0961234567</b> або <b>+380961234567</b>"
            )

    elif field_key == "name":
        if not is_valid_name(value):
            return False, (
                "❌ Це не схоже на ім'я.\n\n"
                "Напишіть, як до вас звертатись.\n"
                "Наприклад: <b>Олексій, Влад, Анна</b>"
            )

    elif field_key == "device":
        if not is_valid_device(value):
            return False, (
                "❌ Не зрозумів пристрій.\n\n"
                "Напишіть модель техніки.\n"
                "Наприклад: <b>iPhone 17, Samsung S23, ноутбук Asus</b>"
            )

    elif field_key == "shipping_city":
        if not is_valid_location(value):
            return False, (
                "❌ Не зрозумів населений пункт відправки.\n\n"
                "Напишіть місто, село, смт або селище, з якого будете відправляти техніку.\n"
                "Наприклад: <b>Київ</b>, <b>село Жашків</b>, <b>смт Бородянка</b>"
            )

    elif field_key == "problem":
        if looks_like_bad_problem(value):
            device = collected.get("device", "пристрій")
            return False, (
                f"❌ Опишіть детальніше проблему з <b>{h(device)}</b>.\n\n"
                "Наприклад:\n"
                "- не працює екран\n"
                "- не заряджається\n"
                "- після води не вмикається\n"
                "- завис і не реагує"
            )

    elif field_key in ("district", "address"):
        if not is_valid_location(value):
            return False, (
                "❌ Це не схоже на адресу або район.\n\n"
                "Напишіть місто та район.\n"
                "Наприклад: <b>Київ, Печерськ</b>"
            )

    elif field_key == "urgency":
        urgency = normalize_urgency(value)
        if not urgency:
            return False, (
                "❌ Не зрозумів терміновість.\n\n"
                "Напишіть один з варіантів:\n"
                "- <b>сьогодні</b>\n"
                "- <b>найближчі дні</b>\n"
                "- <b>до тижня</b>\n"
                "- <b>не терміново</b>"
            )

    return True, ""


def clean_extracted_field(key: str, value: str) -> str:
    value = str(value).strip()

    if key == "phone":
        return normalize_phone(value)

    if key == "urgency":
        normalized = normalize_urgency(value)
        return normalized or value

    if key == "format":
        lower = value.lower()
        if any(word in lower for word in ("онлайн", "online", "дистанц", "віддал", "удален")):
            return "Онлайн"
        if any(word in lower for word in ("офлайн", "offline", "сервіс", "сервис", "майстерня", "мастерская")):
            return "Офлайн"

    if key == "shipping_city":
        return re.sub(
            r"(?i)^(населений пункт відправки|населений пункт|локація відправки|локация отправки|місто|город|з міста|із міста|из города|відправляю з|відправлю з|буду відправляти з|отправляю из|отправлю из|буду отправлять из|з|із|из)\s*(:|-)?\s*",
            "",
            value
        ).strip()

    return value


def merge_extracted_fields(niche: str, collected: dict, extracted: dict, field_context: dict = None) -> tuple:
    valid_keys = [f.get("key") for f in get_order_fields(niche, field_context if field_context is not None else collected)]
    saved_keys = []

    for key, value in (extracted or {}).items():
        if key not in valid_keys or value is None:
            continue

        value = clean_extracted_field(key, value)

        if not value:
            continue

        ok, _ = validate_field(key, value, collected)
        if not ok:
            continue

        collected[key] = value
        saved_keys.append(key)

    return collected, saved_keys


def build_saved_prefix(saved_keys: list) -> str:
    if not saved_keys:
        return ""

    labels = []

    for key in saved_keys:
        label = FIELD_CONFIRM_LABELS.get(key)
        if label:
            labels.append(label.lower())

    if not labels:
        return ""

    if len(labels) == 1:
        return f"✅ <b>{FIELD_CONFIRM_LABELS.get(saved_keys[0], 'Дані')} записав.</b>"

    return f"✅ <b>Записав:</b> {h(', '.join(labels))}."


def build_question_warning(current_field: str, niche: str, collected: dict) -> str:
    question = get_question_for_field(niche, current_field, collected) if current_field else "Опишіть, будь ласка, заявку."

    return (
        "🤖 <b>AI допомагає оформити заявку</b>\n\n"
        "Схоже, ви поставили питання, а не дали відповідь для заявки.\n\n"
        "Напишіть відповідь без знаку питання, щоб я міг коректно записати дані.\n\n"
        f"{h(question)}"
    )


def build_order_question_redirect_text() -> str:
    return (
        "💬 <b>Питання по цьому бізнесу краще ставити на головному екрані.</b>\n\n"
        "Під час оформлення заявки я записую тільки дані для звернення, "
        "щоб нічого не переплутати.\n\n"
        "Перейдіть до <b>головного екрану бізнесу</b> і напишіть питання там - "
        "AI відповість по базі знань бізнесу."
    )


def build_order_question_redirect_keyboard() -> InlineKeyboardMarkup:
    return InlineKeyboardMarkup([
        [InlineKeyboardButton("🏠 Головний екран бізнесу", callback_data="client_back")]
    ])


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


def order_belongs_to_user(order: dict, user_id: int, username: str = "") -> bool:
    client_chat_id = get_order_client_chat_id(order)

    if client_chat_id and int(client_chat_id) == int(user_id):
        return True

    if username:
        order_username = str(order.get("client_username") or "").replace("@", "").lower()
        if order_username and order_username == username.replace("@", "").lower():
            return True

    return False


def find_active_order_for_user(business_id: int, user_id: int, username: str = ""):
    try:
        orders = get_orders(business_id)
    except Exception:
        return None

    for order in orders:
        status = order.get("status", "new")

        if status not in ACTIVE_ORDER_STATUSES:
            continue

        if order_belongs_to_user(order, user_id, username):
            return order

    return None


def build_client_order_summary(order: dict, title: str = "📋 Ваша активна заявка") -> str:
    status = STATUS_LABELS.get(order.get("status"), order.get("status", "new"))

    fields = {
        "📊 Статус": status,
        "👤 Ім'я": order.get("client_name"),
        "📞 Телефон": order.get("client_phone"),
        "🔧 Послуга": order.get("service"),
        "📱 Пристрій": order.get("device"),
        "🚗 Авто": order.get("car"),
        "❗ Проблема": order.get("problem"),
        "📦 Населений пункт відправки": order.get("shipping_city"),
        "📍 Локація": order.get("district"),
        "💻 Формат": order.get("format"),
        "⏰ Терміновість": order.get("urgency"),
    }

    text = f"{title}\n<b>Заявка #{h(order.get('id'))}</b>\n\n"
    for label, value in fields.items():
        if not value:
            continue
        if label.startswith("🤖"):
            text += "\n"
        text += f"{h(label)}: {h(value)}\n"
    return text


def build_existing_order_keyboard(order_id: int) -> InlineKeyboardMarkup:
    return InlineKeyboardMarkup([
        [InlineKeyboardButton("✏️ Редагувати заявку", callback_data=f"client_edit_order_{order_id}")],
        [InlineKeyboardButton("📋 Показати заявку", callback_data=f"client_view_order_{order_id}")],
        [InlineKeyboardButton("❌ Скасувати заявку", callback_data=f"client_cancel_order_{order_id}")],
        [InlineKeyboardButton("🔙 Назад", callback_data="client_back")],
    ])


def build_client_view_order_keyboard(order_id: int, is_active: bool = True) -> InlineKeyboardMarkup:
    keyboard = []

    if is_active:
        keyboard.append([InlineKeyboardButton("✏️ Редагувати заявку", callback_data=f"client_edit_order_{order_id}")])
        keyboard.append([InlineKeyboardButton("❌ Скасувати заявку", callback_data=f"client_cancel_order_{order_id}")])

    keyboard.append([InlineKeyboardButton("🔙 Назад", callback_data="client_back")])
    return InlineKeyboardMarkup(keyboard)


def build_existing_order_text(order: dict, business: dict) -> str:
    status = STATUS_LABELS.get(order.get("status"), order.get("status", "new"))

    return (
        f"⚠️ <b>У вас вже є активна заявка в {h(business.get('name'))}</b>\n\n"
        f"📋 Заявка #{h(order.get('id'))}\n"
        f"📊 Статус: {h(status)}\n\n"
        "Щоб не створювати дубль, краще оновити існуючу заявку.\n\n"
        "<b>Для редагування замовлення:</b>\n"
        "- натисніть <b>Редагувати заявку</b>\n"
        "- напишіть одним повідомленням, що змінити\n"
        "- AI сам оновить потрібні поля"
    )


def build_cancel_order_confirm_text(order: dict, business: dict) -> str:
    status = STATUS_LABELS.get(order.get("status"), order.get("status", "new"))
    item = order.get("device") or order.get("service") or order.get("car") or "заявка"

    return (
        f"❌ <b>Скасувати заявку #{h(order.get('id'))}?</b>\n\n"
        f"Бізнес: <b>{h(business.get('name'))}</b>\n"
        f"Звернення: <b>{h(item)}</b>\n"
        f"Поточний статус: {h(status)}\n\n"
        "Після скасування менеджер побачить, що звернення більше не актуальне.\n"
        "Ви зможете створити нову заявку пізніше."
    )


def build_cancel_order_confirm_keyboard(order_id: int) -> InlineKeyboardMarkup:
    return InlineKeyboardMarkup([
        [InlineKeyboardButton("✅ Так, скасувати", callback_data=f"client_confirm_cancel_{order_id}")],
        [InlineKeyboardButton("↩️ Ні, залишити", callback_data=f"client_keep_order_{order_id}")],
    ])


def build_cancelled_order_text(order: dict, business: dict) -> str:
    return (
        f"🚫 <b>Заявку #{h(order.get('id'))} скасовано</b>\n\n"
        f"Бізнес <b>{h(business.get('name'))}</b> отримає повідомлення, що звернення більше не актуальне.\n\n"
        "Якщо питання знову стане актуальним, ви зможете залишити нову заявку."
    )


def build_cancelled_order_keyboard(business_id: int) -> InlineKeyboardMarkup:
    return InlineKeyboardMarkup([
        [InlineKeyboardButton("📋 Мої заявки", callback_data=f"client_my_orders_{business_id}")],
        [InlineKeyboardButton("📝 Нова заявка", callback_data="client_order")],
        [InlineKeyboardButton("🔙 Головне меню", callback_data="client_back")],
    ])


def build_client_my_orders_keyboard(orders: list, business_id: int) -> InlineKeyboardMarkup:
    keyboard = []

    for order in orders[:10]:
        order_id = order.get("id")
        status = order.get("status", "new")
        status_label = STATUS_LABELS.get(status, status)
        title = order.get("device") or order.get("service") or order.get("car") or f"Заявка #{order_id}"

        keyboard.append([
            InlineKeyboardButton(
                f"📋 #{order_id} - {status_label}",
                callback_data=f"client_view_order_{order_id}"
            )
        ])

        if status in ACTIVE_ORDER_STATUSES:
            keyboard.append([
                InlineKeyboardButton(
                    f"✏️ Редагувати #{order_id}",
                    callback_data=f"client_edit_order_{order_id}"
                ),
                InlineKeyboardButton(
                    f"❌ Скасувати #{order_id}",
                    callback_data=f"client_cancel_order_{order_id}"
                )
            ])

    keyboard.append([InlineKeyboardButton("📝 Нова заявка", callback_data="client_order")])
    keyboard.append([InlineKeyboardButton("🔙 Головне меню", callback_data="client_back")])
    return InlineKeyboardMarkup(keyboard)


def build_client_my_orders_text(orders: list, business: dict) -> str:
    if not orders:
        return (
            f"📋 <b>Мої заявки - {h(business.get('name'))}</b>\n\n"
            "У вас ще немає заявок у цьому бізнесі."
        )

    text = f"📋 <b>Мої заявки - {h(business.get('name'))}</b>\n\n"

    for order in orders[:10]:
        order_id = order.get("id")
        status = STATUS_LABELS.get(order.get("status"), order.get("status", "new"))
        item = order.get("device") or order.get("service") or order.get("car") or "Заявка"
        hot = "🔥 " if order.get("is_hot") else ""
        text += f"{hot}<b>#{h(order_id)}</b> - {h(item)} - {h(status)}\n"

    text += "\nОберіть заявку нижче, щоб переглянути або змінити її."
    return text


def build_client_portal_text() -> str:
    return (
        "👤 <b>Кабінет клієнта</b>\n\n"
        "Тут можна швидко знайти потрібний сервіс або повернутися до своїх заявок.\n\n"
        "<blockquote><i><b>Пошук бізнесу:</b> введіть локацію, а потім назву сервісу, код бізнесу або напрямок послуг.</i></blockquote>\n\n"
        "Оберіть дію нижче 👇"
    )


def build_client_portal_keyboard() -> InlineKeyboardMarkup:
    return InlineKeyboardMarkup([
        [InlineKeyboardButton("🔎 Знайти бізнес", callback_data="client_find_business")],
        [InlineKeyboardButton("📋 Мої заявки", callback_data="client_all_orders")],
        [InlineKeyboardButton("🔙 На початок", callback_data="back_to_start")],
    ])


def build_client_global_orders_text(orders: list) -> str:
    if not orders:
        return (
            "📋 <b>Мої заявки</b>\n\n"
            "Поки що тут немає заявок.\n\n"
            "Знайдіть бізнес, відкрийте його екран і залиште першу заявку."
        )

    text = "📋 <b>Мої заявки</b>\n\n"

    for order in orders[:10]:
        order_id = order.get("id")
        business_name = order.get("business_name") or "Бізнес"
        status = STATUS_LABELS.get(order.get("status"), order.get("status", "new"))
        item = order.get("device") or order.get("service") or order.get("car") or "Заявка"
        hot = "🔥 " if order.get("is_hot") else ""
        text += f"{hot}<b>#{h(order_id)}</b> - {h(business_name)}\n{h(item)} - {h(status)}\n\n"

    text += "Щоб відкрити бізнес або оформити нову заявку, натисніть кнопку нижче."
    return text


def build_client_global_orders_keyboard(orders: list) -> InlineKeyboardMarkup:
    keyboard = []
    seen = set()

    for order in orders[:10]:
        business_id = order.get("business_id")
        if not business_id or business_id in seen:
            continue

        seen.add(business_id)
        business_name = order.get("business_name") or "Бізнес"
        city = order.get("business_city") or "-"
        keyboard.append([
            InlineKeyboardButton(
                f"🏢 {business_name} - {city}"[:64],
                callback_data=f"client_open_biz_{business_id}"
            )
        ])

    keyboard.append([InlineKeyboardButton("🔎 Знайти бізнес", callback_data="client_find_business")])
    keyboard.append([InlineKeyboardButton("🔙 Кабінет клієнта", callback_data="role_client")])
    return InlineKeyboardMarkup(keyboard)


def get_orders_for_client(business_id: int, user_id: int, username: str = "") -> list:
    try:
        orders = get_orders(business_id)
    except Exception:
        return []

    result = []
    for order in orders:
        if order_belongs_to_user(order, user_id, username):
            result.append(order)

    return result


def order_to_collected(order: dict) -> dict:
    mapping = {
        "name": "client_name",
        "phone": "client_phone",
        "service": "service",
        "problem": "problem",
        "device": "device",
        "car": "car",
        "format": "format",
        "shipping_city": "shipping_city",
        "urgency": "urgency",
        "district": "district",
    }
    result = {}
    for key, col in mapping.items():
        val = order.get(col)
        if val is not None and str(val).strip():
            result[key] = str(val).strip()
    return result


def append_problem_text(old_problem: str, new_part: str) -> str:
    old_problem = str(old_problem or "").strip()
    new_part = str(new_part or "").strip()

    if not old_problem:
        return new_part

    if not new_part:
        return old_problem

    if new_part.lower() in old_problem.lower():
        return old_problem

    return f"{old_problem}; {new_part}"



def _strip_lead_words(value: str) -> str:
    value = str(value or "").strip()
    value = re.sub(
        r"(?i)^(пристрій|устройство|модель|телефон|номер|терміновість|термін|срок|строк|проблема|ім'я|імя|имя)\s*(:|-)?\s*",
        "",
        value
    ).strip()
    return value


def _looks_like_device_segment(segment: str) -> bool:
    seg = str(segment or "").lower().strip()
    device_markers = [
        "iphone", "айфон", "samsung", "xiaomi", "redmi", "oppo", "realme", "huawei", "honor",
        "ipad", "macbook", "ноут", "ноутбук", "asus", "acer", "lenovo", "hp", "dell", "планшет",
        "телефон", "смартфон", "приставка", "playstation", "ps5", "ps4"
    ]
    return any(marker in seg for marker in device_markers)


def _looks_like_problem_segment(segment: str) -> bool:
    seg = str(segment or "").lower().strip()
    problem_markers = [
        "не працю", "не работ", "не вмика", "не включ", "злам", "слом", "глюч", "завис",
        "екран", "экран", "батар", "заряд", "заряж", "вода", "динамік", "динамик", "камера",
        "кнопка", "дисплей", "скло", "стекло", "тріс", "трес", "лаг", "шум", "гріється", "греется"
    ]
    return any(marker in seg for marker in problem_markers)


KNOWN_CITY_NORMALIZATION = {
    "київ": "Київ", "києва": "Київ", "киев": "Київ", "киева": "Київ",
    "львів": "Львів", "львова": "Львів", "львов": "Львів",
    "одеса": "Одеса", "одеси": "Одеса", "одесса": "Одеса", "одессы": "Одеса",
    "дніпро": "Дніпро", "дніпра": "Дніпро", "днепр": "Дніпро", "днепра": "Дніпро",
    "харків": "Харків", "харкова": "Харків", "харьков": "Харків", "харькова": "Харків",
    "запоріжжя": "Запоріжжя", "запорожье": "Запоріжжя",
    "вінниця": "Вінниця", "вінниці": "Вінниця", "винница": "Вінниця", "винницы": "Вінниця",
    "полтава": "Полтава", "полтави": "Полтава",
    "черкаси": "Черкаси", "черкассы": "Черкаси",
    "чернігів": "Чернігів", "чернігова": "Чернігів", "чернигов": "Чернігів", "чернигова": "Чернігів",
    "суми": "Суми", "сум": "Суми",
    "житомир": "Житомир", "житомира": "Житомир",
    "рівне": "Рівне", "ровно": "Рівне",
    "луцьк": "Луцьк", "луцька": "Луцьк", "луцк": "Луцьк", "луцка": "Луцьк",
    "ужгород": "Ужгород", "ужгорода": "Ужгород",
    "тернопіль": "Тернопіль", "тернополя": "Тернопіль", "тернополь": "Тернопіль",
    "івано-франківськ": "Івано-Франківськ", "івано-франківська": "Івано-Франківськ",
    "ивано-франковск": "Івано-Франківськ", "ивано-франковска": "Івано-Франківськ",
    "чернівці": "Чернівці", "черновцы": "Чернівці",
    "хмельницький": "Хмельницький", "хмельницького": "Хмельницький",
    "хмельницкий": "Хмельницький", "хмельницкого": "Хмельницький",
    "миколаїв": "Миколаїв", "миколаєва": "Миколаїв", "николаев": "Миколаїв", "николаева": "Миколаїв",
    "херсон": "Херсон", "херсона": "Херсон",
    "кропивницький": "Кропивницький", "кропивницького": "Кропивницький",
    "кропивницкий": "Кропивницький", "кропивницкого": "Кропивницький",
    "кременчук": "Кременчук", "кременчука": "Кременчук",
    "кривий ріг": "Кривий Ріг", "кривого рогу": "Кривий Ріг",
    "кривой рог": "Кривий Ріг", "кривого рога": "Кривий Ріг",
    "біла церква": "Біла Церква", "білої церкви": "Біла Церква",
    "белая церковь": "Біла Церква", "белой церкви": "Біла Церква",
    "кам'янець-подільський": "Кам'янець-Подільський", "каменец-подольский": "Кам'янець-Подільський",
    "маріуполь": "Маріуполь", "мариуполь": "Маріуполь",
    "бровари": "Бровари", "бориспіль": "Бориспіль", "борисполя": "Бориспіль", "борисполь": "Бориспіль",
    "луганськ": "Луганськ", "луганська": "Луганськ", "луганск": "Луганськ", "луганска": "Луганськ",
    "ірпінь": "Ірпінь", "ирпень": "Ірпінь", "буча": "Буча",
}


def normalize_known_city(value: str) -> str:
    city = str(value or "").strip(" .,!?:;")
    low = city.lower()
    if low in KNOWN_CITY_NORMALIZATION:
        return KNOWN_CITY_NORMALIZATION[low]

    prefix, core = split_settlement_prefix(city)
    core_low = core.lower()
    normalized_core = KNOWN_CITY_NORMALIZATION.get(core_low, format_settlement_core(core))

    if prefix:
        return f"{prefix} {normalized_core}".strip()

    return normalized_core


def split_settlement_prefix(value: str) -> tuple:
    text = re.sub(r"\s+", " ", str(value or "").strip(" .,!?:;"))
    match = re.match(
        r"(?i)^(селище міського типу|село|села|с\.|смт|пгт|селище|селища|поселок|посёлок|поселка|посёлка|місто|м\.|город)\s+(.+)$",
        text
    )
    if not match:
        return "", text

    raw_prefix = match.group(1).lower().strip(". ")
    core = match.group(2).strip(" .,!?:;")

    if raw_prefix in {"село", "села", "с"}:
        return "село", core
    if raw_prefix in {"смт", "пгт", "селище міського типу"}:
        return "смт", core
    if raw_prefix in {"селище", "селища", "поселок", "посёлок", "поселка", "посёлка"}:
        return "селище", core

    return "", core


def format_settlement_core(value: str) -> str:
    text = re.sub(r"\s+", " ", str(value or "").strip())
    return text.title()


def has_cyrillic_letters(value: str) -> bool:
    return bool(re.search(r"[А-Яа-яІіЇїЄєҐґ]", str(value or "")))


def _looks_like_name_segment(segment: str) -> bool:
    seg = str(segment or "").strip()
    low = seg.lower()
    placeholders = {"ім'я", "імя", "имя", "name", "ваше ім'я", "ваше имя"}
    if low in placeholders:
        return False
    if not is_valid_name(seg):
        return False
    if _looks_like_device_segment(seg) or _looks_like_problem_segment(seg):
        return False
    if normalize_urgency(seg) or is_valid_phone(seg):
        return False
    return True


def _looks_like_city_segment(segment: str) -> bool:
    seg = str(segment or "").strip()

    if not is_valid_location(seg):
        return False

    if _looks_like_device_segment(seg) or _looks_like_problem_segment(seg):
        return False

    if is_valid_phone(seg) or normalize_urgency(seg):
        return False

    return _is_reasonable_city_candidate(seg)


def _is_reasonable_city_candidate(value: str) -> bool:
    city = str(value or "").strip(" .,!?:;")
    low = city.lower()

    if not city or not is_valid_location(city):
        return False

    if "?" in str(value):
        return False

    if re.search(r"\d", city):
        return False

    if not re.fullmatch(r"[A-Za-zА-Яа-яІіЇїЄєҐґ' .-]{2,60}", city):
        return False

    prefix, core = split_settlement_prefix(city)
    core_low = core.lower().strip(" .,-")

    if not core or len(core) < 2:
        return False

    if len(core.split()) > 4:
        return False

    bad_words = [
        "не знаю", "хз", "що писати", "что писать",
        "допомож", "помоги", "телефон", "номер", "ремонт", "екран", "батар",
        "айфон", "iphone", "macbook", "imac", "ноутбук", "смартфон",
        "мене звати", "меня зовут", "ім'я", "имя",
    ]
    if any(word in low for word in bad_words):
        return False

    if low.startswith(("як ", "как ", "де ", "где ", "що ", "что ")):
        return False

    if _looks_like_device_segment(city) or _looks_like_problem_segment(city):
        return False

    if _looks_like_device_segment(core) or _looks_like_problem_segment(core):
        return False

    if low in KNOWN_CITY_NORMALIZATION or core_low in KNOWN_CITY_NORMALIZATION:
        return True

    return has_cyrillic_letters(core)


def extract_initial_order_fields(niche: str, text: str, collected: dict = None) -> dict:
    """Витягує кілька полів з першого повідомлення клієнта без AI.
    Важливо: не зупиняється на телефоні, а намагається забрати device/problem/urgency/name разом.
    Мусорні або неправдоподібні значення не записує.
    """
    text_clean = str(text or "").strip()
    if not text_clean:
        return {}

    valid_keys = [f.get("key") for f in get_order_fields(niche, collected or {})]
    updates = {}

    compact = re.sub(r"[\s\-()]+", "", text_clean)
    phone_match = re.search(r"(\+?380\d{9}|0\d{9})", compact)
    if phone_match and "phone" in valid_keys and is_valid_phone(phone_match.group(1)):
        updates["phone"] = normalize_phone(phone_match.group(1))

    urgency = normalize_urgency(text_clean)
    if urgency and "urgency" in valid_keys:
        updates["urgency"] = urgency

    parts = [p.strip(" .;\n\t") for p in re.split(r"[,;\n]+", text_clean) if p.strip(" .;\n\t")]
    used_problem_parts = []

    for index, part in enumerate(parts):
        part_clean = _strip_lead_words(part)
        if not part_clean:
            continue
        if is_valid_phone(part_clean) or normalize_urgency(part_clean):
            continue
        if "?" in part_clean:
            continue

        if "shipping_city" in valid_keys and "shipping_city" not in updates:
            shipping_city_match = re.search(
                r"(?i)\b(?:населений пункт відправки|населений пункт|локація відправки|локация отправки|місто|город|з міста|из города|відправляю з|отправляю из)\s+([A-Za-zА-Яа-яІіЇїЄєҐґ' .-]{2,60})",
                part_clean
            )
            if shipping_city_match:
                city_candidate = shipping_city_match.group(1).strip(" ,.-")
                if _is_reasonable_city_candidate(city_candidate):
                    updates["shipping_city"] = normalize_known_city(city_candidate)
                    continue

        # Пристрій
        if "device" in valid_keys and "device" not in updates and _looks_like_device_segment(part_clean):
            # Якщо сегмент одночасно містить явні симптоми, спочатку пробуємо відокремити модель.
            device_candidate = part_clean
            symptom_split = re.split(r"(?i)\b(не працю|не работ|не вмика|не включ|злам|слом|глюч|завис|екран|экран|батар|заряд)\b", part_clean, maxsplit=1)
            if len(symptom_split) > 1 and symptom_split[0].strip():
                device_candidate = symptom_split[0].strip(" ,.-")
            if is_valid_device(device_candidate):
                updates["device"] = device_candidate
                # Якщо в тому ж сегменті є проблема після моделі, збережемо її нижче.
                if len(symptom_split) > 1:
                    problem_tail = part_clean.replace(device_candidate, "", 1).strip(" ,.-")
                    if problem_tail and not looks_like_bad_problem(problem_tail):
                        used_problem_parts.append(problem_tail)
                continue

        # Проблема
        if "problem" in valid_keys and _looks_like_problem_segment(part_clean) and not looks_like_bad_problem(part_clean):
            # Не пишемо чисту терміновість/дату в проблему.
            if not normalize_urgency(part_clean):
                used_problem_parts.append(part_clean)
            continue

        if "shipping_city" in valid_keys and "shipping_city" not in updates:
            city_candidate = clean_extracted_field("shipping_city", part_clean)
            if _looks_like_city_segment(city_candidate) or (index == 0 and len(parts) > 1 and _is_reasonable_city_candidate(city_candidate)):
                updates["shipping_city"] = normalize_known_city(city_candidate)
                continue

        # Ім'я
        if "name" in valid_keys and "name" not in updates and _looks_like_name_segment(part_clean):
            updates["name"] = part_clean
            continue

        # Для інших ніш - послуга, якщо вона є в полях і ще не заповнена.
        if niche != "Ремонт техніки" and "service" in valid_keys and "service" not in updates:
            if len(part_clean) >= 3 and not is_valid_phone(part_clean) and not normalize_urgency(part_clean):
                updates["service"] = part_clean

    if used_problem_parts and "problem" in valid_keys:
        problem_value = "; ".join(dict.fromkeys(used_problem_parts))
        if not looks_like_bad_problem(problem_value):
            updates["problem"] = problem_value[:300]

    return updates


def extract_online_shipping_city(text: str) -> str:
    text_clean = str(text or "").strip()
    if not text_clean:
        return ""

    updates = extract_initial_order_fields(REPAIR_NICHE, text_clean, {"format": "Онлайн"})
    city = str(updates.get("shipping_city") or "").strip()
    if city and _is_reasonable_city_candidate(city):
        return normalize_known_city(city)

    cleaned = clean_extracted_field("shipping_city", text_clean)
    candidates = [cleaned]

    candidates.extend(
        part.strip(" .,!?:;")
        for part in re.split(r"[,;\n]+", text_clean)
        if part.strip(" .,!?:;")
    )

    for candidate in candidates:
        candidate = clean_extracted_field("shipping_city", candidate)
        if _is_reasonable_city_candidate(candidate):
            return normalize_known_city(candidate)

    return ""


def extract_direct_edit_fields(niche: str, text: str, current_order: dict) -> dict:
    """Стабільний parser для редагування заявки.
    Пріоритет: телефон -> терміновість -> локація -> ім'я -> пристрій -> проблема.
    Це зменшує помилки AI і не дає датам/термінам потрапляти в проблему.
    """
    text_clean = str(text).strip()
    text_lower = text_clean.lower()
    valid_keys = [f.get("key") for f in get_order_fields(niche, order_to_collected(current_order))]
    extracted = {}

    action_add = text_lower.startswith(("додай", "додайте", "добавь", "добавьте", "ще", "також", "еще"))
    action_replace = text_lower.startswith(("заміни", "замініть", "змініть", "зміни", "измени", "измените", "поменяй", "поміняй"))

    cleaned = re.sub(
        r"(?i)^(додайте|додай|добавь|добавьте|ще|також|еще|заміни|замініть|змініть|зміни|измени|измените|поменяй|поміняй)\s*",
        "",
        text_clean
    ).strip()

    # Телефон
    phone_match = re.search(r"(\+?380\d{9}|0\d{9})", text_clean.replace(" ", ""))
    if phone_match and is_valid_phone(phone_match.group(1)):
        return {"phone": normalize_phone(phone_match.group(1))}
    if is_valid_phone(text_clean):
        return {"phone": normalize_phone(text_clean)}

    # Терміновість - завжди вище проблеми.
    urgency = normalize_urgency(text_clean)
    urgency_markers = [
        "термін", "терміновість", "сроч", "срок", "строк", "сьогодні", "сегодня", "завтра",
        "до тижня", "протягом тижня", "на тижні", "через тиждень", "найближч", "ближайш", "кількох днів", "нескольких дней", "не терміново", "не срочно"
    ]
    if urgency or any(marker in text_lower for marker in urgency_markers):
        if urgency:
            return {"urgency": urgency}
        return {}

    # Локація
    location_markers = ["район", "адрес", "адреса", "місто", "город", "локац", "вул", "улиц", "поділ", "печер", "оболон", "дарниц"]
    shipping_markers = [
        "населений пункт відправки", "населений пункт", "локація відправки", "локация отправки",
        "місто відправки", "город отправки", "відправляю з", "отправляю из", "з міста", "из города"
    ]
    if "shipping_city" in valid_keys and any(word in text_lower for word in shipping_markers):
        value = re.sub(
            r"(?i)^(населений пункт відправки|населений пункт|локація відправки|локация отправки|місто відправки|город отправки|місто|город|з міста|из города|відправляю з|отправляю из)\s*(на|:|-)?\s*",
            "",
            cleaned
        ).strip()
        if is_valid_location(value):
            return {"shipping_city": value}

    if any(word in text_lower for word in location_markers):
        value = re.sub(r"(?i)^(район|адресу|адреса|місто|город|локацію|локация)\s*(на|:|-)?\s*", "", cleaned).strip()
        if is_valid_location(value):
            return {"address" if "address" in valid_keys else "district": value}

    # Ім'я
    name_markers = ["ім'я", "імя", "имя", "звати", "зовут", "називаюсь"]
    if any(word in text_lower for word in name_markers):
        value = re.sub(r"(?i).*(ім'я|імя|имя|звати|зовут|називаюсь)\s*(на|:|-)?\s*", "", text_clean).strip()
        if is_valid_name(value):
            return {"name": value}

    # Пристрій
    device_markers = ["пристр", "устройство", "телефон", "айфон", "iphone", "samsung", "ноут", "asus", "xiaomi", "ipad", "macbook"]
    if "device" in valid_keys and any(word in text_lower for word in device_markers) and not any(word in text_lower for word in ["не працю", "не работ", "злам", "слом", "глюч", "батар", "екран", "заряд"]):
        value = re.sub(r"(?i)^(пристрій|устройство|модель)\s*(на|:|-)?\s*", "", cleaned).strip()
        if is_valid_device(value):
            return {"device": value}

    # Проблема. Якщо написано 'додай/добавь' - дописуємо, якщо 'зміни/замени' - замінюємо.
    problem_markers = [
        "проблем", "злам", "слом", "не працю", "не работ", "глюч", "завис", "тріс", "трес",
        "екран", "экран", "батар", "заряд", "вода", "динамік", "камера", "кнопка", "дисплей"
    ]
    if "problem" in valid_keys and (any(word in text_lower for word in problem_markers) or action_add or action_replace):
        value = re.sub(r"(?i)^(проблема|уточнення|уточни|деталі|детали)\s*(:|-)?\s*", "", cleaned).strip()
        if value and not looks_like_bad_problem(value):
            if action_add:
                return {"problem": append_problem_text(current_order.get("problem"), value)}
            return {"problem": value}

    return extracted


def fallback_extract_edit_fields(niche: str, text: str, current_order: dict) -> dict:
    text_clean = str(text).strip()
    text_lower = text_clean.lower()
    valid_keys = [f.get("key") for f in get_order_fields(niche, order_to_collected(current_order))]

    direct = extract_direct_edit_fields(niche, text_clean, current_order)
    if direct:
        return direct

    # Якщо текст схожий на термін/дату - не записуємо як проблему.
    if normalize_urgency(text_clean) or any(word in text_lower for word in ["термін", "терміновість", "срок", "строк", "дата", "коли", "когда"]):
        return {}

    problem_markers = [
        "проблем", "злам", "слом", "не працю", "не работ", "глюч", "завис", "тріс", "трес",
        "екран", "экран", "батар", "заряд", "вода", "додай", "добавь", "уточни", "ще", "еще"
    ]

    if "problem" in valid_keys and any(word in text_lower for word in problem_markers) and not looks_like_bad_problem(text_clean):
        if text_lower.startswith(("додай", "додайте", "добавь", "добавьте", "ще", "також", "еще")):
            value = re.sub(r"(?i)^(додайте|додай|добавь|добавьте|ще|також|еще)\s*(:|-)?\s*", "", text_clean).strip()
            return {"problem": append_problem_text(current_order.get("problem"), value or text_clean)}
        return {"problem": text_clean}

    return {}


def update_order_fields(order_id: int, updates: dict, ai_comment: str = None):
    if not updates and not ai_comment:
        return

    column_map = {
        "name": "client_name",
        "phone": "client_phone",
        "service": "service",
        "problem": "problem",
        "device": "device",
        "car": "car",
        "format": "format",
        "shipping_city": "shipping_city",
        "urgency": "urgency",
        "district": "district",
    }

    set_parts = []
    values = []

    for key, value in updates.items():
        column = column_map.get(key)
        if not column:
            continue
        set_parts.append(f"{column} = ?")
        values.append(value)

    if "urgency" in updates:
        set_parts.append("is_hot = ?")
        values.append(1 if "Сьогодні" in str(updates.get("urgency")) else 0)

    if ai_comment is not None:
        set_parts.append("ai_comment = ?")
        values.append(ai_comment)

    if not set_parts:
        return

    values.append(order_id)

    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    cursor.execute(
        f"UPDATE orders SET {', '.join(set_parts)} WHERE id = ?",
        values
    )
    conn.commit()
    conn.close()


def get_order_by_id_for_business(business_id: int, order_id: int):
    return get_order_for_business(business_id, order_id)


def build_edit_order_prompt(order: dict, business: dict) -> str:
    return (
        "✏️ <b>Редагування заявки через AI</b>\n\n"
        "Напишіть одним повідомленням, що потрібно змінити або додати.\n"
        "AI оновить тільки потрібні поля і не створить нову заявку.\n\n"
        "<b>Приклади:</b>\n"
        "- додайте: батарея швидко сідає\n"
        "- змініть проблему на не заряджається\n"
        "- номер телефону 0961234567\n"
        "- терміново сьогодні\n\n"
        + build_client_order_summary(order, title="<b>Поточна заявка:</b>")
    )


def build_edit_success_text(updated_order: dict, changed_keys: list) -> str:
    prefix = build_saved_prefix(changed_keys) or "✅ <b>Заявку оновлено.</b>"
    return (
        f"{prefix}\n\n"
        "Менеджер побачить оновлені дані. Нову заявку створювати не потрібно.\n\n"
        + build_client_order_summary(updated_order, title="📋 <b>Оновлена заявка</b>")
    )


async def handle_client_order_edit_text(update: Update, context: ContextTypes.DEFAULT_TYPE, data: dict, text: str):
    user_id = update.effective_user.id
    business = get_business_by_id(data.get("business_id"))

    if not business:
        await update.message.reply_text("❌ Бізнес не знайдено. Перейдіть по посиланню ще раз.")
        return

    order_id = data.get("edit_order_id")
    order = get_order_by_id_for_business(business["id"], order_id)

    if not order or not order_belongs_to_user(order, user_id, update.effective_user.username or ""):
        save_session(user_id, "client_start", {"business_id": business["id"]})
        await update.message.reply_text("❌ Активну заявку не знайдено.")
        return

    if order.get("status") not in ACTIVE_ORDER_STATUSES:
        save_session(user_id, "client_start", {"business_id": business["id"]})
        await update.message.reply_text(
            "ℹ️ Цю заявку вже закрито. Якщо потрібно - створіть нову заявку або перегляньте попередні.",
            reply_markup=InlineKeyboardMarkup([
                [InlineKeyboardButton("📋 Мої заявки", callback_data=f"client_my_orders_{business['id']}")],
                [InlineKeyboardButton("📝 Нова заявка", callback_data="client_order")]
            ])
        )
        return

    text_clean = text.strip()

    if not text_clean:
        await update.message.reply_text("❌ Напишіть, будь ласка, що саме змінити.")
        return

    if is_question_message(text_clean) and not is_help_request(text_clean):
        await update.message.reply_text(
            "🤖 <b>Редагування заявки</b>\n\n"
            "Схоже, це питання, а не зміна заявки.\n"
            "Напишіть конкретно, що змінити. Наприклад:\n"
            "<i>додайте, що батарея швидко сідає</i>",
            parse_mode="HTML"
        )
        return

    if is_help_request(text_clean):
        await update.message.reply_text(
            "🤖 <b>Як редагувати заявку</b>\n\n"
            "Напишіть одне коротке повідомлення з новими даними.\n\n"
            "Наприклад:\n"
            "- <i>додайте, що ноутбук не заряджається</i>\n"
            "- <i>змініть проблему на не працює екран</i>\n"
            "- <i>змініть телефон на 0961234567</i>",
            parse_mode="HTML"
        )
        return

    await update.message.reply_text("🤖 AI оновлює заявку...")

    niche = business.get("niche") or REPAIR_NICHE
    current_collected = order_to_collected(order)
    required_fields = get_order_fields(niche, current_collected)

    # Спочатку deterministic-перевірка: очевидну терміновість, телефон, район не віддаємо AI,
    # щоб AI не записав "терміновість - протягом тижня" у проблему.
    direct_updates = extract_direct_edit_fields(niche, text_clean, order)

    if direct_updates:
        updates = direct_updates
        changed_keys = list(direct_updates.keys())
    else:
        analysis = await analyze_client_request(
            business=business,
            collected=current_collected,
            user_text=text_clean,
            required_fields=required_fields,
            current_field=None
        )

        extracted = analysis.get("extracted") if analysis.get("ok") else {}
        updates, changed_keys = merge_extracted_fields(niche, {}, extracted or {}, field_context=current_collected)

        if "problem" in updates and text_clean.lower().startswith(("додай", "додайте", "добавь", "добавьте", "ще", "також")):
            updates["problem"] = append_problem_text(order.get("problem"), updates.get("problem"))

        # Додатковий запобіжник: якщо AI все одно кинув терміновість/дату у problem - прибираємо.
        if "problem" in updates and (normalize_urgency(text_clean) or any(word in text_clean.lower() for word in ["термін", "терміновість", "срок", "строк", "дата", "коли", "когда"])):
            updates.pop("problem", None)
            changed_keys = [key for key in changed_keys if key != "problem"]

    if not updates:
        updates = fallback_extract_edit_fields(niche, text_clean, order)
        valid_updates = {}
        changed_keys = []
        for key, value in updates.items():
            ok, _ = validate_field(key, value, current_collected)
            if ok:
                valid_updates[key] = clean_extracted_field(key, value)
                changed_keys.append(key)
        updates = valid_updates

    if not updates:
        await update.message.reply_text(
            "❌ Не зміг зрозуміти, що саме змінити.\n\n"
            "Напишіть простіше. Наприклад:\n"
            "- додайте: батарея не тримає заряд\n"
            "- змініть проблему на не працює екран\n"
            "- телефон 0961234567",
            parse_mode="HTML"
        )
        return

    new_collected = dict(current_collected)
    new_collected.update(updates)
    ai_comment = await qualify_order(business, new_collected)

    update_order_fields(order_id, updates, ai_comment=ai_comment)
    updated_order = get_order_by_id_for_business(business["id"], order_id)

    save_session(user_id, "client_start", {"business_id": business["id"]})

    await update.message.reply_text(
        build_edit_success_text(updated_order, changed_keys),
        parse_mode="HTML",
        reply_markup=InlineKeyboardMarkup([
            [InlineKeyboardButton("🔙 Головне меню", callback_data="client_back")]
        ])
    )

    await notify_owner_about_order_edit(
        context,
        business,
        order,
        updated_order,
        changed_keys,
        order_detail_text=build_order_detail_text(updated_order)
    )


# Хелпери для цінності бізнесу і бази знань

KNOWLEDGE_STEPS = [
    {
        "key": "schedule",
        "title": "Графік роботи",
        "question": (
            "📚 <b>База знань - крок 1/7</b>\n\n"
            "Напишіть графік роботи.\n\n"
            "Наприклад: <i>Пн-Пт 09:00-18:00, Сб-Нд 10:00-17:00</i>"
        ),
    },
    {
        "key": "prices",
        "title": "Ціни",
        "question": (
            "📚 <b>База знань - крок 2/7</b>\n\n"
            "Напишіть ціни або як ви їх рахуєте.\n\n"
            "Наприклад: <i>Діагностика безкоштовна. Заміна екрану від 1200 грн. Точна ціна після огляду.</i>"
        ),
    },
    {
        "key": "address",
        "title": "Адреса/філіали",
        "question": (
            "📚 <b>База знань - крок 3/7</b>\n\n"
            "Напишіть точну адресу сервісу або список філіалів.\n"
            "Це буде показуватись клієнту у вкладці <b>Адреса</b>.\n\n"
            "Наприклад: <i>Київ, вул. Хрещатик 10, 2 поверх. Вхід з двору.</i>"
        ),
    },
    {
        "key": "conditions",
        "title": "Умови роботи",
        "question": (
            "📚 <b>База знань - крок 4/7</b>\n\n"
            "Напишіть умови роботи: запис, прийом пристроїв, терміни, передоплата, доставка.\n\n"
            "Наприклад: <i>Працюємо по запису. Клієнт приносить пристрій у сервіс. Термін ремонту залежить від діагностики.</i>"
        ),
    },
    {
        "key": "warranty",
        "title": "Гарантія",
        "question": (
            "📚 <b>База знань - крок 5/7</b>\n\n"
            "Напишіть гарантії або обмеження.\n\n"
            "Наприклад: <i>Гарантія на роботи 30 днів. На пошкодження після падіння гарантія не діє.</i>"
        ),
    },
    {
        "key": "faq",
        "title": "Часті питання",
        "question": (
            "📚 <b>База знань - крок 6/7</b>\n\n"
            "Напишіть часті питання клієнтів і короткі відповіді.\n\n"
            "Наприклад: <i>Чи можна сьогодні? - Так, якщо є вільний слот. Чи можна без запису? - Краще попередньо залишити заявку.</i>"
        ),
    },
    {
        "key": "limits",
        "title": "Що не обіцяти клієнтам",
        "question": (
            "📚 <b>База знань - крок 7/7</b>\n\n"
            "Напишіть, що AI не має обіцяти клієнтам.\n\n"
            "Наприклад: <i>Не називати точну ціну без діагностики. Не гарантувати ремонт за 1 годину.</i>"
        ),
    },
]



NICHE_KNOWLEDGE_EXAMPLES = {
    "Ремонт техніки": {
        "schedule": (
            "Напишіть графік роботи сервісу.",
            "Пн-Пт 09:00-18:00, Сб-Нд 10:00-17:00"
        ),
        "prices": (
            "Напишіть ціни на діагностику, популярні роботи або як формується вартість.",
            "Діагностика безкоштовна. Заміна екрану від 1200 грн. Точна ціна після огляду."
        ),
        "address": (
            "Напишіть точну адресу сервісу або список філіалів.",
            "Київ, вул. Хрещатик 10, 2 поверх. Вхід з двору."
        ),
        "conditions": (
            "Напишіть умови роботи: запис, прийом пристроїв, терміни, передоплата, доставка.",
            "Працюємо по запису. Клієнт приносить пристрій у сервіс. Термін ремонту залежить від діагностики."
        ),
        "warranty": (
            "Напишіть гарантії або обмеження.",
            "Гарантія на роботи 30 днів. На пошкодження після падіння або вологи гарантія не діє."
        ),
        "faq": (
            "Напишіть часті питання клієнтів і короткі відповіді.",
            "Чи можна сьогодні? - Так, якщо є вільний слот. Чи можна без запису? - Краще попередньо залишити заявку."
        ),
        "limits": (
            "Напишіть, що AI не має обіцяти клієнтам.",
            "Не називати точну ціну без діагностики. Не гарантувати ремонт за 1 годину."
        ),
    },
    "Автосервіс": {
        "schedule": (
            "Напишіть графік роботи автосервісу.",
            "Пн-Сб 09:00-19:00, Нд вихідний. Прийом авто бажано за попереднім записом."
        ),
        "prices": (
            "Напишіть ціни на діагностику, базові роботи або як формується вартість.",
            "Діагностика ходової від 400 грн, заміна масла від 300 грн. Запчастини оплачуються окремо."
        ),
        "address": (
            "Напишіть точну адресу СТО або список філіалів.",
            "Київ, вул. Автозаводська 15. В'їзд з боку шиномонтажу."
        ),
        "conditions": (
            "Напишіть умови роботи: запис, прийом авто, терміни, запчастини, оплата.",
            "Працюємо по запису. Термін ремонту залежить від діагностики та наявності запчастин."
        ),
        "warranty": (
            "Напишіть гарантії або обмеження.",
            "Гарантія на виконані роботи 14 днів. На запчастини діє гарантія постачальника."
        ),
        "faq": (
            "Напишіть часті питання клієнтів і короткі відповіді.",
            "Чи можна приїхати сьогодні? - Так, якщо є вільний слот. Чи є свої запчастини? - Можна, але потрібно узгодити з майстром."
        ),
        "limits": (
            "Напишіть, що AI не має обіцяти клієнтам.",
            "Не називати точну вартість без огляду авто. Не гарантувати термін ремонту без діагностики."
        ),
    },
}


CLIENT_QUESTION_HINTS = {
    "Ремонт техніки": "ціна, графік, адреса, гарантія, терміни ремонту або умови роботи",
    "Автосервіс": "ціна, графік, адреса, діагностика, запчастини, запис або гарантія",
}


PRICE_FALLBACKS = {
    "Ремонт техніки": "Точну вартість уточнить менеджер після діагностики.",
    "Автосервіс": "Точну вартість уточнить менеджер після огляду авто або діагностики.",
}


def get_niche_knowledge_prompt(niche: str, key: str) -> tuple:
    niche = str(niche or "").strip()
    key = str(key or "").strip()
    by_niche = NICHE_KNOWLEDGE_EXAMPLES.get(niche) or NICHE_KNOWLEDGE_EXAMPLES.get("Ремонт техніки", {})
    default_by_key = NICHE_KNOWLEDGE_EXAMPLES.get("Ремонт техніки", {})
    return by_niche.get(key) or default_by_key.get(key) or ("Напишіть інформацію для клієнтів.", "Напишіть коротко і зрозуміло.")


def get_niche_knowledge_example(niche: str, key: str) -> str:
    return get_niche_knowledge_prompt(niche, key)[1]


def build_niche_knowledge_question(step_index: int, niche: str = None) -> str:
    step = get_knowledge_step(step_index)
    if not step:
        return ""

    instruction, example = get_niche_knowledge_prompt(niche, step["key"])
    return (
        f"📚 <b>База знань - крок {int(step_index) + 1}/{len(KNOWLEDGE_STEPS)}</b>\n\n"
        f"{instruction}\n\n"
        f"Наприклад: <i>{h(example)}</i>"
    )


def get_client_question_hint(niche: str) -> str:
    return CLIENT_QUESTION_HINTS.get(str(niche or "").strip(), "ціна, графік, адреса, гарантія або умови роботи")


def get_price_fallback_for_niche(niche: str) -> str:
    return PRICE_FALLBACKS.get(str(niche or "").strip(), "Точну вартість уточнить менеджер після деталей заявки.")


def get_knowledge_step(index: int):
    if 0 <= int(index) < len(KNOWLEDGE_STEPS):
        return KNOWLEDGE_STEPS[int(index)]
    return None


def build_knowledge_keyboard(step_index: int = 0) -> InlineKeyboardMarkup:
    keyboard = []
    if step_index < len(KNOWLEDGE_STEPS):
        keyboard.append([InlineKeyboardButton("⏭ Пропустити", callback_data="menu_knowledge_skip")])
    keyboard.append([InlineKeyboardButton("✅ Завершити", callback_data="menu_knowledge_finish")])
    keyboard.append([BTN_BACK_MENU])
    return InlineKeyboardMarkup(keyboard)


def format_schedule_text(value: str) -> str:
    """Зберігає графік красиво: якщо власник пише кожен день з нового рядка - так і показуємо.
    Якщо написав коротко через кому - залишаємо коротко. Якщо написав всі дні в один рядок без ком - розбиваємо.
    """
    value = str(value or "").strip()
    value = value.replace("24\\7", "24/7").replace("24 на 7", "24/7")

    if not value:
        return ""

    # Якщо бізнес уже ввів нормальний багаторядковий графік - не ламаємо формат.
    if "\n" in value:
        lines = [re.sub(r"[ \t]+", " ", line).strip() for line in value.splitlines()]
        return "\n".join(line for line in lines if line)[:900]

    value = re.sub(r"[ \t]+", " ", value).strip()

    # Якщо є коми між групами днів - зберігаємо короткий формат як написав власник.
    if "," in value:
        return value[:900]

    # Якщо всі дні написані одним суцільним рядком - робимо перенос перед кожним днем.
    day_pattern = r"(?i)(понеділок|вівторок|середа|четвер|п['’]?ятниця|субота|неділя|пн|вт|ср|чт|пт|сб|нд)(\s*:)"
    parts = []
    last = 0
    matches = list(re.finditer(day_pattern, value))
    if len(matches) > 1:
        for i, m in enumerate(matches):
            if i > 0:
                part = value[last:m.start()].strip(" ,;\n")
                if part:
                    parts.append(part)
                last = m.start()
        tail = value[last:].strip(" ,;\n")
        if tail:
            parts.append(tail)
        return "\n".join(parts)[:900]

    return value[:900]


def normalize_knowledge_answer(key: str, text: str) -> str:
    value = str(text or "").strip()

    if key == "schedule":
        return format_schedule_text(value)

    # ВАЖЛИВО: не з'їдаємо ентери, які бізнес спеціально поставив у базі знань.
    # Наприклад для цін:
    # 📱 Смартфони
    # ...
    #
    # 💻 Ноутбуки
    # ...
    # Якщо власник написав все одним рядком - ми не додаємо переносів самі.
    if "\n" in value:
        lines = [re.sub(r"[ \t]+", " ", line).strip() for line in value.splitlines()]
        normalized_lines = []
        blank_count = 0

        for line in lines:
            if not line:
                blank_count += 1
                # Залишаємо максимум один порожній рядок підряд, щоб формат був красивий,
                # але без величезних розривів у повідомленні Telegram.
                if blank_count <= 1:
                    normalized_lines.append("")
                continue

            blank_count = 0
            normalized_lines.append(line)

        value = "\n".join(normalized_lines).strip()
    else:
        value = re.sub(r"[ \t]+", " ", value).strip()

    return value[:900]

def build_knowledge_text(kb: dict) -> str:
    blocks = []

    for step in KNOWLEDGE_STEPS:
        key = step["key"]
        title = step["title"]
        value = str(kb.get(key) or "").strip()
        if value:
            blocks.append(f"{title}:\n{value}")

    if not blocks:
        return ""

    return "\n\n".join(blocks)


def build_knowledge_preview(kb: dict) -> str:
    text = build_knowledge_text(kb)
    return text if text else "Ваша база знань поки відсутня"


def parse_knowledge_text(text: str) -> dict:
    """Parse saved knowledge text back into step fields for safe editing."""
    result = {}
    raw = str(text or "").strip()
    if not raw:
        return result

    title_to_key = {step["title"].lower(): step["key"] for step in KNOWLEDGE_STEPS}
    current_key = None
    buffer = []

    def flush():
        nonlocal buffer, current_key
        if current_key and buffer:
            value = "\n".join(buffer).strip()
            if value:
                result[current_key] = value
        buffer = []

    for line in raw.splitlines():
        stripped = line.strip()
        lower = stripped.rstrip(":").lower()
        if lower in title_to_key and stripped.endswith(":"):
            flush()
            current_key = title_to_key[lower]
            continue
        if current_key:
            buffer.append(line)

    flush()
    return result


def build_knowledge_question(step_index: int, kb: dict = None, niche: str = None) -> str:
    step = get_knowledge_step(step_index)
    if not step:
        return ""

    preview = ""
    if kb:
        filled = sum(1 for item in KNOWLEDGE_STEPS if str(kb.get(item["key"]) or "").strip())
        preview = f"\n\n<b>Заповнено:</b> {filled}/{len(KNOWLEDGE_STEPS)}"

    return build_niche_knowledge_question(step_index, niche) + preview


def is_valid_faq_knowledge_answer(text: str) -> bool:
    """
    FAQ intentionally contains client questions.
    Example: "Чи можна сьогодні? - Так, якщо є вільний слот."
    So question marks are allowed here if the owner also provides an answer.
    """
    text_clean = str(text or "").strip()
    text_lower = text_clean.lower()

    if len(text_clean) < 8:
        return False

    answer_markers = [
        " - ", " — ", " – ", ":",
        " так", " ні", " да", " нет",
        "так,", "ні,", "да,", "нет,",
        "можна", "можемо", "працюємо", "є ", "немає"
    ]

    if "?" not in text_clean:
        return True

    return any(marker in text_lower for marker in answer_markers)


def get_knowledge_title(key: str) -> str:
    return next((item["title"] for item in KNOWLEDGE_STEPS if item["key"] == key), str(key))


def build_knowledge_edit_keyboard() -> InlineKeyboardMarkup:
    return InlineKeyboardMarkup([
        [
            InlineKeyboardButton("🕘 Графік", callback_data="menu_kbedit_field_schedule"),
            InlineKeyboardButton("💰 Ціни", callback_data="menu_kbedit_field_prices"),
        ],
        [
            InlineKeyboardButton("📍 Адреса", callback_data="menu_kbedit_field_address"),
            InlineKeyboardButton("📌 Умови", callback_data="menu_kbedit_field_conditions"),
        ],
        [
            InlineKeyboardButton("🛡 Гарантія", callback_data="menu_kbedit_field_warranty"),
            InlineKeyboardButton("❓ FAQ", callback_data="menu_kbedit_field_faq"),
        ],
        [InlineKeyboardButton("🚫 Що не обіцяти", callback_data="menu_kbedit_field_limits")],
        [BTN_BACK_MENU]
    ])


def build_knowledge_ai_edit_prompt(kb: dict) -> str:
    current = build_knowledge_preview(kb)
    return (
        "✏️ <b>Редагування AI-бази знань</b>\n\n"
        "Спочатку виберіть, який розділ потрібно змінити.\n"
        "Можна натиснути кнопку нижче або написати простими словами.\n\n"
        "<b>Приклади:</b>\n"
        "- <i>змінити умови роботи</i>\n"
        "- <i>змінити ціни</i>\n"
        "- <i>оновити адресу</i>\n"
        "- <i>додати гарантію</i>\n"
        "- <i>змінити що не обіцяти клієнтам</i>\n\n"
        "Після цього я попрошу ввести новий текст саме для цього розділу.\n"
        "Так ми не переплутаємо <b>умови роботи</b> з <b>графіком</b> або цінами.\n\n"
        "<b>Поточна інформація:</b>\n"
        f"{h(current)}"
    )


def build_knowledge_field_value_prompt(kb: dict, key: str, niche: str = None) -> str:
    title = get_knowledge_title(key)
    current = str((kb or {}).get(key) or "").strip()
    current_text = h(current) if current else "<i>Поки не заповнено</i>"
    example = get_niche_knowledge_example(niche, key)

    return (
        f"✏️ <b>Змінюємо розділ: {h(title)}</b>\n\n"
        "Напишіть новий текст для цього розділу одним повідомленням.\n"
        "Я збережу його саме сюди і не буду переносити в інший блок.\n\n"
        "<b>Поточне значення:</b>\n"
        f"{current_text}\n\n"
        "<b>Приклад:</b>\n"
        f"<i>{h(example)}</i>"
    )


def classify_knowledge_edit_section(text: str) -> str:
    """Визначає тільки розділ для редагування. Не зберігає сам текст."""
    text_clean = str(text or "").strip().lower()
    if not text_clean:
        return ""

    # важливо: умови роботи перевіряємо окремо і рано, щоб не плутати з графіком роботи
    aliases = [
        ("conditions", [
            "умови роботи", "условия работы", "умови", "условия", "правила роботи",
            "терміни ремонту", "сроки ремонта", "діагностика трива", "майстер починає",
            "передоплата", "прийом пристроїв", "прийом", "прием", "доставка", "запис"
        ]),
        ("schedule", [
            "графік роботи", "график работы", "графік", "график", "розклад", "расписание",
            "коли працю", "когда работа", "години роботи", "часы работы"
        ]),
        ("prices", [
            "ціни", "ціна", "цены", "цена", "вартість", "стоимость", "прайс", "тариф на послуги"
        ]),
        ("address", [
            "адреса", "адрес", "локація", "локация", "філіали", "филиалы", "де знаходит", "де ви", "вул", "улица"
        ]),
        ("warranty", [
            "гарантія", "гарантия", "гарантії", "гарантии"
        ]),
        ("faq", [
            "faq", "часті питання", "частые вопросы", "питання клієнтів", "вопросы клиентов", "чзп"
        ]),
        ("limits", [
            "що не обіцяти", "не обіцяти", "не обещать", "обмеження", "ограничения",
            "заборони", "запреты", "не гарантувати", "не називати точну ціну"
        ]),
    ]

    for key, words in aliases:
        if any(word in text_clean for word in words):
            return key

    return ""


def classify_knowledge_update(text: str) -> dict:
    """
    Старий fallback для сумісності. Нова логіка редагування працює у 2 кроки:
    1) вибір розділу
    2) введення нового тексту для цього розділу.
    """
    key = classify_knowledge_edit_section(text)
    if not key:
        return {}
    return {key: text}


def apply_knowledge_updates(kb: dict, updates: dict) -> tuple:
    changed = []
    updated = dict(kb or {})

    for key, value in (updates or {}).items():
        value = normalize_knowledge_answer(key, value)
        if not value:
            continue
        updated[key] = value
        title = get_knowledge_title(key)
        changed.append(title)

    return updated, changed


async def handle_knowledge_ai_edit_text(update: Update, context: ContextTypes.DEFAULT_TYPE, data: dict, text: str):
    """Крок 1: власник обирає, який розділ AI-бази змінювати."""
    user_id = update.effective_user.id
    business = get_business(user_id)

    if not business:
        await update.message.reply_text("❌ Бізнес не знайдено. Натисніть /start")
        return

    kb = data.get("kb") or parse_knowledge_text(business.get("knowledge") or "")
    text_clean = str(text or "").strip()

    if not text_clean:
        await update.message.reply_text("❌ Напишіть, який розділ потрібно змінити.")
        return

    key = classify_knowledge_edit_section(text_clean)

    if not key:
        await update.message.reply_text(
            "🤖 <b>AI-база знань</b>\n\n"
            "Не зрозумів, який розділ потрібно змінити.\n"
            "Напишіть простіше або оберіть кнопку нижче.\n\n"
            "<b>Приклади:</b>\n"
            "- <i>змінити умови роботи</i>\n"
            "- <i>оновити ціни</i>\n"
            "- <i>змінити адресу</i>\n"
            "- <i>додати гарантію</i>",
            parse_mode="HTML",
            reply_markup=build_knowledge_edit_keyboard()
        )
        return

    save_session(user_id, "knowledge_ai_edit_value", {"kb": kb, "pending_key": key})
    await update.message.reply_text(
        build_knowledge_field_value_prompt(kb, key, business.get("niche")),
        parse_mode="HTML",
        reply_markup=InlineKeyboardMarkup([[BTN_BACK_MENU]])
    )


async def handle_knowledge_ai_edit_value_text(update: Update, context: ContextTypes.DEFAULT_TYPE, data: dict, text: str):
    """Крок 2: власник вводить новий текст для вже вибраного розділу."""
    user_id = update.effective_user.id
    business = get_business(user_id)

    if not business:
        await update.message.reply_text("❌ Бізнес не знайдено. Натисніть /start")
        return

    kb = data.get("kb") or parse_knowledge_text(business.get("knowledge") or "")
    key = data.get("pending_key")

    if not key or not get_knowledge_title(key):
        save_session(user_id, "knowledge_ai_edit", {"kb": kb})
        await update.message.reply_text(
            "❌ Не бачу, який розділ потрібно змінити. Оберіть розділ ще раз.",
            reply_markup=build_knowledge_edit_keyboard()
        )
        return

    text_clean = str(text or "").strip()

    if len(text_clean) < 2:
        await update.message.reply_text("❌ Дуже коротко. Напишіть нову інформацію детальніше.")
        return

    # На цьому кроці ми НЕ класифікуємо текст заново.
    # Якщо власник обрав 'Умови роботи', то будь-який введений текст зберігається саме в 'Умови роботи'.
    value = normalize_knowledge_answer(key, text_clean)
    if not value:
        await update.message.reply_text("❌ Не вдалося зберегти текст. Напишіть інформацію простіше.")
        return

    kb[key] = value
    knowledge_text = build_knowledge_text(kb)
    update_business(user_id, "knowledge", knowledge_text)

    if key == "schedule":
        try:
            update_business(user_id, "schedule", value)
        except Exception:
            pass

    save_session(user_id, "entrepreneur_menu")
    title = get_knowledge_title(key)

    await update.message.reply_text(
        f"✅ <b>AI-базу знань оновлено.</b>\n\n"
        f"🔄 Змінено: <b>{h(title)}</b>\n\n"
        f"<b>Поточна інформація:</b>\n{h(knowledge_text)}",
        parse_mode="HTML",
        reply_markup=InlineKeyboardMarkup([[BTN_BACK_MENU]])
    )
    await show_entrepreneur_menu(update, context, get_business(user_id) or business)

def get_business_orders_safe(business_id: int) -> list:
    try:
        return get_orders(business_id) or []
    except Exception:
        return []


def summarize_business_value_stats(business: dict) -> dict:
    orders = get_business_orders_safe(business["id"])
    total = len(orders)
    hot = sum(1 for o in orders if o.get("is_hot") or "Сьогодні" in str(o.get("urgency") or ""))
    new = sum(1 for o in orders if o.get("status", "new") == "new")
    in_progress = sum(1 for o in orders if o.get("status") == "in_progress")
    contacted = sum(1 for o in orders if o.get("status") == "contacted")
    success = sum(1 for o in orders if o.get("status") == "success")
    rejected = sum(1 for o in orders if o.get("status") == "rejected")
    processed = contacted + success + rejected
    active = new + in_progress + contacted
    success_rate = round((success / processed) * 100) if processed else 0
    hot_rate = round((hot / total) * 100) if total else 0
    estimated_minutes_saved = total * 4 + hot * 2

    return {
        "total": total,
        "hot": hot,
        "new": new,
        "in_progress": in_progress,
        "contacted": contacted,
        "success": success,
        "rejected": rejected,
        "processed": processed,
        "active": active,
        "success_rate": success_rate,
        "hot_rate": hot_rate,
        "estimated_minutes_saved": estimated_minutes_saved,
    }


def build_business_value_analytics_text(business: dict) -> str:
    s = summarize_business_value_stats(business)

    proof_line = "Поки що даних мало - перший кейс з'явиться після 5-10 реальних заявок."
    if s["total"] >= 5:
        proof_line = (
            f"Уже є міні-кейс: бот прийняв {s['total']} заявок, "
            f"з них {s['hot']} гарячих, {s['processed']} оброблено."
        )

    return (
        f"📊 <b>Аналітика користі - {h(business.get('name'))}</b>\n\n"
        f"📋 Всього заявок: <b>{s['total']}</b>\n"
        f"🔥 Гарячих лідів: <b>{s['hot']}</b> ({s['hot_rate']}%)\n"
        f"🆕 Нових: <b>{s['new']}</b>\n"
        f"🔄 В роботі: <b>{s['in_progress']}</b>\n"
        f"📞 Зв'язались: <b>{s['contacted']}</b>\n"
        f"✅ Успішних: <b>{s['success']}</b>\n"
        f"❌ Відмов: <b>{s['rejected']}</b>\n\n"
        f"⚙️ Оброблено менеджером: <b>{s['processed']}</b>\n"
        f"📈 Успішність серед оброблених: <b>{s['success_rate']}%</b>\n"
        f"⏱ Орієнтовно зекономлено часу: <b>{s['estimated_minutes_saved']} хв</b>\n\n"
        f"<b>Що це показує бізнесу:</b>\n"
        f"- скільки звернень не загубились\n"
        f"- скільки клієнтів були терміновими\n"
        f"- скільки заявок реально довели до контакту/успіху\n\n"
        f"🧪 <b>Доказ користі:</b>\n{h(proof_line)}"
    )


# Побудова текстів

def build_ai_order_hint(business: dict, collected: dict = None) -> str:
    niche = business.get("niche") or "послуги"
    services = business.get("services") or ""
    collected = collected or {}
    is_online_repair = is_repair_online_order(niche, collected)

    examples = {
        "Ремонт техніки": "iPhone 17, не працює екран, протягом кількох днів, Ім'я, 0961234567",
        "Автосервіс": "Volkswagen Golf 7, заміна масла і перевірка гальм, протягом кількох днів, Ім'я, 0961234567",
    }

    example = examples.get(
        niche,
        "Опишіть послугу, локацію, терміновість, ім'я і телефон"
    )

    services_text = ""
    if services:
        services_text = f"\n\n<b>Послуги:</b>\n{h(services)}"

    shipping_city = str(collected.get("shipping_city") or "").strip()
    online_text = f"\n\n📦 Населений пункт відправки: <b>{h(shipping_city)}</b>" if is_online_repair and shipping_city else ""
    title = "🤖 <b>AI допоможе оформити заявку швидше</b>"

    if is_online_repair and shipping_city:
        title = "✅ <b>Тепер оформимо заявку швидше</b>"

    return (
        f"{title}\n\n"
        "Напишіть одним повідомленням усе, що знаєте - я сам розберу деталі."
        f"{online_text}"
        f"{services_text}\n\n"
        "<b>Наприклад:</b>\n"
        f"<i>{h(example)}</i>\n\n"
        "Або просто опишіть своїми словами 👇"
    )


def build_problem_help_message(device: str = "пристрій", niche: str = "Ремонт техніки") -> str:
    niche = str(niche or "").strip()

    if niche == "Автосервіс":
        return (
            "🤖 Опишіть, що потрібно зробити з авто або які симптоми є.\n\n"
            "Наприклад:\n"
            "- стукає ходова\n"
            "- потрібна заміна масла\n"
            "- горить помилка на панелі\n"
            "- авто погано заводиться\n\n"
            "Напишіть простими словами - я оформлю це нормально для майстра."
        )

    return (
        f"🤖 Розкажіть, що трапилось з <b>{h(device)}</b>.\n\n"
        "Наприклад:\n"
        "- не працює екран\n"
        "- не заряджається\n"
        "- після води не вмикається\n"
        "- завис і не реагує\n"
        "- тріснуте скло\n\n"
        "Напишіть симптоми простими словами - я оформлю це нормально для майстра."
    )



def is_business_registration_complete(business: dict) -> bool:
    """Реєстрація завершена тільки коли є назва, ніша, місто і послуги."""
    if not business:
        return False

    required = ("name", "niche", "city", "services")
    if not all(str(business.get(field) or "").strip() for field in required):
        return False

    if is_repair_niche(business.get("niche")):
        return str(business.get("work_mode") or "").strip() in WORK_MODE_OPTIONS

    return True


def get_registration_services_example(niche: str) -> str:
    return REGISTRATION_SERVICE_EXAMPLES.get(niche or "", REGISTRATION_SERVICE_EXAMPLES["Ремонт техніки"])


def build_registration_services_text(niche: str) -> str:
    example = get_registration_services_example(niche)
    hint = REGISTRATION_NICHE_HINTS.get(niche or "", "")
    step = 5 if is_repair_niche(niche) else 4
    total = 5 if is_repair_niche(niche) else 4

    text = (
        "✅ <b>Місто збережено!</b>\n\n"
        f"👔 <b>Крок {step}/{total} - Які послуги ви надаєте?</b>\n\n"
        "Напишіть основні послуги через кому. "
        "Цей текст клієнт буде бачити у картці бізнесу, тому краще писати зрозуміло.\n\n"
        f"<b>Приклад для ніші «{h(niche or 'Бізнес')}»:</b>\n"
        f"<i>{h(example)}</i>"
    )

    if hint:
        text += f"\n\n{h(hint)}"

    return text


async def show_registration_step(update, context, business: dict, edit: bool = False):
    """Повертає підприємця на незавершений крок реєстрації, а не в кабінет."""
    user_id = update.effective_user.id
    business = business or {}

    if not str(business.get("name") or "").strip():
        save_session(user_id, "reg_name")
        text = (
            "👔 <b>Реєстрація бізнесу</b>\n\n"
            "Крок 1/4 - Як називається ваш бізнес?\n\n"
            "Напишіть назву:"
        )
        keyboard = InlineKeyboardMarkup([[BTN_BACK_START]])

    elif not str(business.get("niche") or "").strip():
        save_session(user_id, "reg_niche")
        keyboard_rows = [
            [InlineKeyboardButton(NICHE_BUTTON_TEXTS.get(n, n), callback_data=f"niche_{n}")]
            for n in NICHES
        ]
        keyboard_rows.append([BTN_BACK_START])
        text = (
            "👔 <b>Реєстрація бізнесу</b>\n\n"
            f"✅ Назва: <b>{h(business.get('name'))}</b>\n\n"
            "Крок 2/4 - Оберіть нішу бізнесу.\n\n"
            "Виберіть напрямок, під який AI буде збирати заявки:"
        )
        keyboard = InlineKeyboardMarkup(keyboard_rows)

    elif is_repair_niche(business.get("niche")) and str(business.get("work_mode") or "").strip() not in WORK_MODE_OPTIONS:
        save_session(user_id, "reg_work_mode")
        text = build_work_mode_registration_text(business)
        keyboard = build_work_mode_keyboard()

    elif not str(business.get("city") or "").strip():
        save_session(user_id, "reg_city")
        step = 4 if is_repair_niche(business.get("niche")) else 3
        total = 5 if is_repair_niche(business.get("niche")) else 4
        work_mode_line = ""
        if is_repair_niche(business.get("niche")):
            work_mode_line = f"✅ Формат роботи: <b>{h(get_work_mode_label(business.get('work_mode')))}</b>\n"
        text = (
            "👔 <b>Реєстрація бізнесу</b>\n\n"
            f"✅ Назва: <b>{h(business.get('name'))}</b>\n"
            f"✅ Ніша: <b>{h(business.get('niche'))}</b>\n\n"
            f"{work_mode_line}\n"
            f"Крок {step}/{total} - В якому місті працюєте?\n\n"
            "Напишіть місто:"
        )
        keyboard = InlineKeyboardMarkup([[BTN_BACK_START]])

    else:
        save_session(user_id, "reg_services")
        text = build_registration_services_text(business.get("niche"))
        keyboard = InlineKeyboardMarkup([[BTN_BACK_START]])

    if edit and update.callback_query:
        await update.callback_query.edit_message_text(
            text,
            parse_mode="HTML",
            reply_markup=keyboard
        )
    else:
        await update.effective_message.reply_text(
            text,
            parse_mode="HTML",
            reply_markup=keyboard
        )



# Стартове меню

async def show_start_menu(update, context, edit=False):
    keyboard = [[
        InlineKeyboardButton("👔 Я підприємець", callback_data="role_entrepreneur"),
        InlineKeyboardButton("👤 Я клієнт", callback_data="role_client"),
    ]]

    text = (
        "Ласкаво просимо до <b>ClientDesk AI</b> 🤖\n\n"
        "AI-адміністратор для бізнесу в Telegram:\n"
        "- приймає заявки 24/7\n"
        "- кваліфікує ліди через AI\n"
        "- надсилає власнику готові заявки\n\n"
        "Хто ви?"
    )

    markup = InlineKeyboardMarkup(keyboard)

    if edit and update.callback_query:
        await update.callback_query.edit_message_text(
            text,
            parse_mode="HTML",
            reply_markup=markup
        )
    else:
        await update.effective_message.reply_text(
            text,
            parse_mode="HTML",
            reply_markup=markup
        )


# Меню підприємця

async def show_entrepreneur_menu(update, context, business, edit=False):
    usage = get_tariff_usage(business)
    tariff = usage["stored_tariff"]
    tariff_label = tariff.upper()
    limit = usage["limit"]
    used = int(usage["used"] or 0)
    remaining = int(usage["remaining"] or 0)
    subscription_line = ""
    if usage["is_paid"]:
        if usage["active"]:
            subscription_line = f"<i>Активний до: <b>{h(display_datetime(usage.get('expires_at')) or '-')}</b></i>\n"
        else:
            subscription_line = f"<i>Підписка закінчилась: <b>{h(display_datetime(usage.get('expires_at')) or '-')}</b></i>\n"
    else:
        subscription_line = "<i>Активний до: <b>безстроково</b></i>\n"

    try:
        stats = get_stats(business.get("id"))
    except Exception:
        stats = {"today": 0, "month": 0, "hot": 0, "success": 0, "new": 0}

    today = int(stats.get("today") or 0)
    month = int(stats.get("month") or 0)
    hot = int(stats.get("hot") or 0)
    success = int(stats.get("success") or 0)
    new_orders = int(stats.get("new") or 0)

    used_percent = int((used / limit) * 100) if limit else 0
    if usage["expired"] and usage["is_paid"]:
        usage_line = f"<i>Підписка тарифу <b>{h(tariff_label)}</b> закінчилась!</i>"
    elif limit and (remaining <= 0 or used >= limit):
        usage_line = f"<i>Ліміт тарифу <b>{h(tariff_label)}</b> закінчився!</i>"
    elif used_percent >= 90:
        usage_line = f"<i>Ліміт тарифу <b>{h(tariff_label)}</b> майже закінчився!</i>"
    elif used_percent >= 60:
        usage_line = "<i>Ліміт використовується активно.</i>"
    else:
        usage_line = "<i>Ліміт у нормі.</i>"

    if not business.get("knowledge"):
        knowledge_line = "AI-база знань не заповнена"
    else:
        knowledge_line = "AI-база знань активна"

    work_mode_line = ""
    if is_repair_niche(business.get("niche")):
        work_mode_line = f"Формат роботи: <b>{h(get_work_mode_label(business.get('work_mode')))}</b>\n"

    keyboard = [
        [build_support_button()],
        [
            InlineKeyboardButton("📋 Заявки", callback_data="menu_orders"),
            InlineKeyboardButton("🔗 Посилання", callback_data="menu_link"),
        ],
        [
            InlineKeyboardButton("📚 AI-база", callback_data="menu_knowledge"),
            InlineKeyboardButton("💼 Тариф", callback_data="menu_tariff"),
        ],
        [BTN_BACK_START],
    ]

    text = (
        f"<b>Кабінет бізнесу 🧑‍💻</b>\n\n"
        f"🏢 <b>{h(business.get('name') or 'Мій бізнес')}</b>\n"
        f"Місто: <b>{h(business.get('city') or '-')}</b> - {h(business.get('niche') or '-')}\n"
        f"{work_mode_line}"
        f"Тариф: <b>{h(tariff_label)}</b>\n\n"
        f"{subscription_line}"
        f"📋 <b>Огляд заявок</b>\n"
        f"Сьогодні: <b>{h(today)}</b>\n"
        f"Нові: <b>{h(new_orders)}</b>\n"
        f"Гарячі: <b>{h(hot)}</b>\n"
        f"Успішні: <b>{h(success)}</b>\n\n"
        f"📦 <b>Ліміт тарифу</b>\n"
        f"Використано: <b>{h(used)}/{h(limit)}</b> заявок\n"
        f"Залишилось: <b>{h(remaining)}</b>\n"
        f"{usage_line}\n\n"
        f"🤖 <b>AI-адміністратор</b>\n"
        f"{h(knowledge_line)}\n\n"
        f"Оберіть дію:"
    )

    markup = InlineKeyboardMarkup(keyboard)

    if edit and update.callback_query:
        await update.callback_query.edit_message_text(
            text,
            parse_mode="HTML",
            reply_markup=markup
        )
    else:
        await update.effective_message.reply_text(
            text,
            parse_mode="HTML",
            reply_markup=markup
        )


async def handle_entrepreneur(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()

    user_id = query.from_user.id
    business = get_business(user_id)

    if business and is_business_registration_complete(business):
        save_session(user_id, "entrepreneur_menu")
        await show_entrepreneur_menu(update, context, business, edit=True)
        return

    if not business:
        create_business(user_id)
        business = get_business(user_id)

    await show_registration_step(update, context, business, edit=True)


async def handle_niche(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()

    user_id = query.from_user.id
    session = get_session(user_id)

    if not session or session["state"] != "reg_niche":
        return

    niche = query.data.replace("niche_", "")
    update_business(user_id, "niche", niche)

    if is_repair_niche(niche):
        save_session(user_id, "reg_work_mode")
        await query.edit_message_text(
            build_work_mode_registration_text({"niche": niche}),
            parse_mode="HTML",
            reply_markup=build_work_mode_keyboard()
        )
        return

    save_session(user_id, "reg_city")

    await query.edit_message_text(
        f"✅ Ніша: <b>{h(niche)}</b>\n\n"
        "Крок 3/4 - В якому місті працюєте?\n\n"
        "Напишіть місто:",
        parse_mode="HTML",
        reply_markup=InlineKeyboardMarkup([[BTN_BACK_START]])
    )


async def handle_work_mode(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()

    user_id = query.from_user.id
    session = get_session(user_id)

    if not session or session["state"] != "reg_work_mode":
        return

    mode = query.data.replace("work_mode_", "", 1)
    if mode not in WORK_MODE_OPTIONS:
        await query.answer("Оберіть формат роботи", show_alert=True)
        return

    update_business(user_id, "work_mode", mode)

    business = get_business(user_id)
    if business and is_business_registration_complete(business):
        save_session(user_id, "entrepreneur_menu")
        await show_entrepreneur_menu(update, context, business, edit=True)
        return

    save_session(user_id, "reg_city")

    await query.edit_message_text(
        f"✅ Формат роботи: <b>{h(get_work_mode_label(mode))}</b>\n\n"
        "Крок 4/5 - В якому місті працюєте?\n\n"
        "Напишіть місто:",
        parse_mode="HTML",
        reply_markup=InlineKeyboardMarkup([[BTN_BACK_START]])
    )


# Кнопки меню підприємця

async def handle_menu(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()

    user_id = query.from_user.id
    business = get_business(user_id)

    if not business:
        await query.edit_message_text("❌ Бізнес не знайдено. Натисніть /start")
        return

    if query.data in ("knowledge_skip", "menu_knowledge_skip"):
        session = get_session(user_id)
        data = session["data"] if session else {}
        kb = data.get("kb", {})
        step_index = int(data.get("step_index", 0)) + 1

        if step_index >= len(KNOWLEDGE_STEPS):
            knowledge_text = build_knowledge_text(kb)
            if knowledge_text:
                update_business(user_id, "knowledge", knowledge_text)
            save_session(user_id, "entrepreneur_menu")
            await query.edit_message_text(
                "✅ <b>База знань збережена.</b>\n\n" + (h(knowledge_text) if knowledge_text else "<i>Ваша база знань поки відсутня</i>"),
                parse_mode="HTML",
                reply_markup=InlineKeyboardMarkup([[BTN_BACK_MENU]])
            )
            return

        save_session(user_id, "knowledge_collecting", {"kb": kb, "step_index": step_index})
        await query.edit_message_text(
            build_knowledge_question(step_index, kb, business.get("niche")),
            parse_mode="HTML",
            reply_markup=build_knowledge_keyboard(step_index)
        )
        return

    if query.data in ("knowledge_finish", "menu_knowledge_finish"):
        session = get_session(user_id)
        data = session["data"] if session else {}
        kb = data.get("kb", {})
        knowledge_text = build_knowledge_text(kb)
        if knowledge_text:
            update_business(user_id, "knowledge", knowledge_text)
            if kb.get("schedule"):
                try:
                    update_business(user_id, "schedule", kb.get("schedule"))
                except Exception:
                    pass
        save_session(user_id, "entrepreneur_menu")
        await query.edit_message_text(
            "✅ <b>База знань збережена.</b>\n\n" + (h(knowledge_text) if knowledge_text else "<i>Ваша база знань поки відсутня</i>"),
            parse_mode="HTML",
            reply_markup=InlineKeyboardMarkup([[BTN_BACK_MENU]])
        )
        return

    if query.data == "menu_back":
        save_session(user_id, "entrepreneur_menu")
        await show_entrepreneur_menu(update, context, business, edit=True)
        return

    if query.data.startswith("menu_kbedit_field_"):
        key = query.data.replace("menu_kbedit_field_", "", 1)
        allowed_keys = {item["key"] for item in KNOWLEDGE_STEPS}

        if key not in allowed_keys:
            await query.answer("Розділ не знайдено", show_alert=True)
            return

        kb = parse_knowledge_text(business.get("knowledge") or "")
        save_session(user_id, "knowledge_ai_edit_value", {"kb": kb, "pending_key": key})
        await query.edit_message_text(
            build_knowledge_field_value_prompt(kb, key, business.get("niche")),
            parse_mode="HTML",
            reply_markup=InlineKeyboardMarkup([[BTN_BACK_MENU]])
        )
        return

    if query.data == "menu_orders_export_locked":
        total = count_orders(business["id"])
        await query.edit_message_text(
            build_orders_export_locked_text(business, total),
            parse_mode="HTML",
            reply_markup=InlineKeyboardMarkup([
                [build_support_button()],
                [InlineKeyboardButton("💼 Переглянути тариф", callback_data="menu_tariff")],
                [BTN_BACK_ORDERS],
            ])
        )
        return

    if query.data == "menu_orders_export":
        total = count_orders(business["id"])

        if not orders_export_available(business, total):
            await query.edit_message_text(
                build_orders_export_locked_text(business, total),
                parse_mode="HTML",
                reply_markup=InlineKeyboardMarkup([
                    [build_support_button()],
                    [InlineKeyboardButton("💼 Переглянути тариф", callback_data="menu_tariff")],
                    [BTN_BACK_ORDERS],
                ])
            )
            return

        await query.answer("Готую XLSX-файл...")
        orders = get_orders(business["id"])
        xlsx_bytes = build_orders_xlsx_bytes(business, orders)
        file_obj = io.BytesIO(xlsx_bytes)
        file_obj.name = f"clientdesk_orders_{business.get('link_code') or business.get('id')}_{datetime.now().strftime('%Y%m%d')}.xlsx"

        await query.message.reply_document(
            document=file_obj,
            filename=file_obj.name,
            caption=(
                f"📥 <b>Експорт заявок - {h(business.get('name') or 'ClientDesk AI')}</b>\n\n"
                f"У файлі: <b>{h(len(orders))}</b> заявок."
            ),
            parse_mode="HTML"
        )
        return

    if query.data == "menu_orders" or query.data.startswith("menu_orders_page_"):
        total = count_orders(business["id"])

        if total <= 0:
            await query.edit_message_text(
                "📋 <b>Заявки</b>\n\n"
                "Заявок поки немає.\n\n"
                "Поділіться посиланням з клієнтами, щоб отримати першу заявку!",
                parse_mode="HTML",
                reply_markup=InlineKeyboardMarkup([[BTN_BACK_MENU]])
            )
            return

        total_pages = max(1, (total + ORDERS_PAGE_SIZE - 1) // ORDERS_PAGE_SIZE)
        page = min(parse_orders_page(query.data), total_pages)
        orders = get_orders_page(business["id"], page=page, per_page=ORDERS_PAGE_SIZE)

        await query.edit_message_text(
            build_orders_page_text(orders, total, page, total_pages),
            parse_mode="HTML",
            reply_markup=build_orders_page_keyboard(orders, page, total_pages, business=business, total=total)
        )

    elif query.data == "menu_link":
        code = business["link_code"]
        username = context.bot.username
        link = f"https://t.me/{username}?start=biz_{code}"

        await query.edit_message_text(
            f"🔗 <b>Ваше посилання для клієнтів:</b>\n\n"
            f"{h(link)}\n\n"
            f"📱 <b>Код бізнесу:</b> <code>{h(code)}</code>\n\n"
            f"<b>Де розмістити:</b>\n"
            f"- Instagram bio\n"
            f"- Telegram-канал\n"
            f"- Google Maps\n"
            f"- Сайт або QR-код\n\n"
            f"<b>Текст для Instagram:</b>\n"
            f"Записатись / залишити заявку 👇\n"
            f"{h(link)}",
            parse_mode="HTML",
            reply_markup=InlineKeyboardMarkup([[BTN_BACK_MENU]])
        )

    elif query.data == "menu_stats":
        await query.edit_message_text(
            build_business_value_analytics_text(business),
            parse_mode="HTML",
            reply_markup=InlineKeyboardMarkup([[BTN_BACK_MENU]])
        )

    elif query.data == "menu_knowledge":
        current = business.get("knowledge") or ""

        if current.strip():
            save_session(user_id, "entrepreneur_menu")
            await query.edit_message_text(
                "📚 <b>AI-база знань</b>\n\n"
                "AI використовує ці дані, коли відповідає клієнтам про графік, ціни, умови, гарантії та часті питання.\n\n"
                "<b>Поточна інформація:</b>\n"
                f"{h(current)}\n\n"
                "Щоб змінити дані, натисніть кнопку нижче і напишіть одним повідомленням, що потрібно змінити.",
                parse_mode="HTML",
                reply_markup=InlineKeyboardMarkup([
                    [InlineKeyboardButton("✏️ Редагувати з AI", callback_data="menu_knowledge_edit")],
                    [BTN_BACK_MENU]
                ])
            )
            return

        save_session(user_id, "knowledge_collecting", {"kb": {}, "step_index": 0})
        await query.edit_message_text(
            "📚 <b>AI-база знань</b>\n\n"
            "Я заповню її по кроках, як нормальний адміністратор. Потім AI буде використовувати ці дані у відповідях клієнтам.\n\n"
            "<b>Поточна інформація:</b>\n"
            "<i>Ваша база знань поки відсутня</i>\n\n"
            + build_knowledge_question(0, {}, business.get("niche")),
            parse_mode="HTML",
            reply_markup=build_knowledge_keyboard(0)
        )

    elif query.data == "menu_knowledge_edit":
        kb = parse_knowledge_text(business.get("knowledge") or "")
        save_session(user_id, "knowledge_ai_edit", {"kb": kb})
        await query.edit_message_text(
            build_knowledge_ai_edit_prompt(kb),
            parse_mode="HTML",
            reply_markup=build_knowledge_edit_keyboard()
        )

    elif query.data == "menu_tariff":
        usage = get_tariff_usage(business)
        current = usage["stored_tariff"]
        limit = usage["limit"]
        used = usage["used"]
        subscription_line = ""
        if usage["is_paid"]:
            if usage["active"]:
                subscription_line = f"<i>Активний до: <b>{h(display_datetime(usage.get('expires_at')) or '-')}</b></i>\n"
            else:
                subscription_line = f"<i>Підписка закінчилась: <b>{h(display_datetime(usage.get('expires_at')) or '-')}</b></i>\n"
        else:
            subscription_line = "<i>Активний до: <b>безстроково</b></i>\n"

        text = (
            f"💼 <b>Тарифи ClientDesk AI</b>\n\n"
            f"Ваш тариф: <b>{h(str(current).upper())}</b>\n"
            f"Використано: <b>{h(used)}/{h(limit)}</b> заявок\n\n"
            f"{subscription_line}"
        )

        for tariff, price in TARIFF_PRICES.items():
            mark = "✅ " if tariff == current else ""
            text += f"{mark}<b>{h(str(tariff).upper())}</b> - {h(price)}\n"

        text += "\n📩 Для зміни тарифу зверніться до підтримки."

        await query.edit_message_text(
            text,
            parse_mode="HTML",
            reply_markup=InlineKeyboardMarkup([[BTN_BACK_MENU]])
        )


# Деталі заявки

def build_order_detail_text(order: dict) -> str:
    status = STATUS_LABELS.get(order.get("status"), order.get("status", "-"))
    hot_label = "🔥 <b>ГАРЯЧИЙ ЛІД</b>\n" if order.get("is_hot") else ""

    text = f"{hot_label}📋 <b>Заявка #{h(order.get('id'))}</b>\n\n"

    top_fields = [
        ("👤 Клієнт", order.get("client_name")),
        ("📞 Телефон", order.get("client_phone")),
        ("💬 Telegram", f"@{order['client_username']}" if order.get("client_username") else None),
        ("🔧 Послуга", order.get("service")),
        ("📱 Пристрій", order.get("device")),
        ("🚗 Авто", order.get("car")),
        ("📦 Населений пункт відправки", order.get("shipping_city")),
        ("📍 Локація", order.get("district")),
    ]

    for label, val in top_fields:
        if val:
            text += f"{h(label)}: {h(val)}\n"

    problem = order.get("problem")
    if problem:
        text += f"\n❗️ <b>Проблема:</b> {h(problem)}\n\n"

    bottom_fields = [
        ("⏰ Терміновість", order.get("urgency")),
        ("📊 Статус", status),
    ]

    for label, val in bottom_fields:
        if val:
            text += f"{h(label)}: {h(val)}\n"

    ai_comment = order.get("ai_comment")
    if ai_comment:
        text += f"\n🤖 <b>AI-оцінка:</b> {h(ai_comment)}\n"

    return text

def build_order_detail_keyboard(order_id: int, status: str = None) -> InlineKeyboardMarkup:
    if str(status or "").lower() in CANCELLED_ORDER_STATUSES:
        return InlineKeyboardMarkup([[BTN_BACK_ORDERS]])

    return InlineKeyboardMarkup([
        [
            InlineKeyboardButton("🔄 В роботі", callback_data=f"status_{order_id}_in_progress"),
            InlineKeyboardButton("📞 Зв'язались", callback_data=f"status_{order_id}_contacted"),
        ],
        [
            InlineKeyboardButton("✅ Успішно", callback_data=f"status_{order_id}_success"),
            InlineKeyboardButton("❌ Відмова", callback_data=f"status_{order_id}_rejected"),
        ],
        [BTN_BACK_ORDERS]
    ])


async def handle_order_detail(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()

    user_id = query.from_user.id
    business = get_business(user_id)

    if not business:
        await query.edit_message_text("❌ Бізнес не знайдено. Натисніть /start")
        return

    order_id = int(query.data.replace("order_", ""))
    order = get_order_by_id_for_business(business["id"], order_id)

    if not order:
        await query.edit_message_text("❌ Заявку не знайдено.")
        return

    await query.edit_message_text(
        build_order_detail_text(order),
        parse_mode="HTML",
        reply_markup=build_order_detail_keyboard(order_id, order.get("status"))
    )


async def handle_order_status(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    parts = query.data.split("_")

    if len(parts) < 3:
        await query.answer("❌ Некоректний статус", show_alert=True)
        return

    try:
        order_id = int(parts[1])
    except ValueError:
        await query.answer("❌ Некоректний номер заявки", show_alert=True)
        return

    status = "_".join(parts[2:])

    if status not in STATUS_LABELS:
        await query.answer("❌ Невідомий статус", show_alert=True)
        return

    user_id = query.from_user.id
    business = get_business(user_id)

    if not business:
        await query.answer("❌ Бізнес не знайдено", show_alert=True)
        return

    order = get_order_by_id_for_business(business["id"], order_id)

    if not order:
        await query.answer("❌ Заявку не знайдено", show_alert=True)
        return

    if str(order.get("status") or "").lower() in CANCELLED_ORDER_STATUSES:
        await query.answer("🚫 Заявку скасовано клієнтом. Статус змінити не можна.", show_alert=True)
        try:
            await query.edit_message_text(
                build_order_detail_text(order),
                parse_mode="HTML",
                reply_markup=build_order_detail_keyboard(order_id, order.get("status"))
            )
        except BadRequest as e:
            if "Message is not modified" not in str(e):
                raise
        return

    if order.get("status") == status:
        await query.answer(f"ℹ️ Статус вже: {STATUS_LABELS.get(status)}", show_alert=True)
        return

    update_order_status(order_id, status)

    updated_order = get_order_by_id_for_business(business["id"], order_id) or order

    label = STATUS_LABELS.get(status, status)
    await query.answer(f"✅ Статус змінено: {label}", show_alert=True)

    await notify_client_about_status(context, updated_order, business, status)

    try:
        await query.edit_message_text(
            build_order_detail_text(updated_order),
            parse_mode="HTML",
            reply_markup=build_order_detail_keyboard(order_id, updated_order.get("status"))
        )
    except BadRequest as e:
        if "Message is not modified" not in str(e):
            raise


# Старт клієнта

def build_client_main_keyboard() -> InlineKeyboardMarkup:
    return InlineKeyboardMarkup([
        [InlineKeyboardButton("📝 Залишити заявку", callback_data="client_order")],
    ])


def build_client_question_keyboard() -> InlineKeyboardMarkup:
    return InlineKeyboardMarkup([
        [InlineKeyboardButton("📝 Залишити заявку", callback_data="client_order")],
        [InlineKeyboardButton("🔙 Назад", callback_data="client_back")],
    ])


def build_client_main_text(business: dict) -> str:
    name = business.get("name") or "бізнес"
    city = business.get("city") or "-"
    services = (business.get("services") or "").strip()
    question_hint = h(get_client_question_hint(business.get("niche")))
    question_block = (
        "<blockquote>"
        f"<i><b>На цьому екрані можна поставити будь-яке питання по бізнесу:</b> {question_hint}.</i>"
        "</blockquote>"
    )

    services_line = f"\n<b>Послуги:</b> {h(services)}" if services else ""
    work_mode_line = ""
    if is_repair_niche(business.get("niche")):
        work_mode_line = f"\n\n<b>Формат роботи:</b> {h(get_work_mode_label(business.get('work_mode')))}"

    return (
        f"👋 <b>Вітаю! Це {h(name)}</b>\n\n"
        f"📍 <b>Місто:</b> {h(city)}"
        f"{services_line}"
        f"{work_mode_line}\n\n"
        "Я AI-адміністратор цього бізнесу.\n"
        "Відповім на питання по бізнесу або допоможу оформити заявку.\n\n"
        f"{question_block}\n\n"
        "<b>Напишіть, що вас цікавить, або оформіть заявку нижче 👇</b>"
    )


async def handle_client_business_question(update: Update, context: ContextTypes.DEFAULT_TYPE, data: dict, text: str):
    business = get_business_by_id(data.get("business_id"))

    if not business:
        await update.message.reply_text(
            "❌ Бізнес не знайдено. Відкрийте посилання бізнесу ще раз або знайдіть його через пошук.",
            parse_mode="HTML"
        )
        return

    loading_message = await update.message.reply_text("🤖 Перевіряю AI-базу знань бізнесу...")

    result = await answer_business_question(business, text)
    answer_text = result.get("answer") if isinstance(result, dict) else str(result or "")

    if not answer_text:
        answer_text = (
            "ℹ️ <b>У базі знань поки немає точної відповіді на це питання.</b>\n\n"
            "Можете залишити заявку - менеджер побачить звернення і уточнить деталі."
        )

    answer_text = sanitize_telegram_html(answer_text)

    try:
        await loading_message.edit_text(
            answer_text,
            parse_mode="HTML",
            reply_markup=build_client_question_keyboard()
        )
    except BadRequest as e:
        print(f"Telegram HTML answer error: {e}")
        # Абсолютно безпечний fallback без HTML, щоб бот не падав навіть якщо AI повернув дивний текст.
        safe_plain = re.sub(r"<[^>]+>", "", str(answer_text))
        try:
            await loading_message.edit_text(
                safe_plain,
                reply_markup=build_client_question_keyboard()
            )
        except BadRequest:
            await update.message.reply_text(
                safe_plain,
                reply_markup=build_client_question_keyboard()
            )


async def handle_client_start(update: Update, context: ContextTypes.DEFAULT_TYPE, business: dict, from_link: bool = False):
    user_id = update.effective_user.id
    save_session(user_id, "client_start", {"business_id": business["id"]})

    await update.effective_message.reply_text(
        build_client_main_text(business),
        parse_mode="HTML",
        reply_markup=build_client_main_keyboard()
    )

    if from_link:
        schedule_client_link_reminder(
            context=context,
            chat_id=update.effective_chat.id,
            user_id=user_id,
            business=business,
            username=update.effective_user.username or "",
        )


async def show_client_main_menu(update, context, data, edit=False):
    business = get_business_by_id(data.get("business_id"))
    if not business:
        text = (
            "❌ Бізнес не знайдено.\n\n"
            "Можливо, посилання застаріло або бізнес було видалено. "
            "Поверніться на початок і знайдіть бізнес повторно."
        )
        markup = InlineKeyboardMarkup([[InlineKeyboardButton("🔙 На початок", callback_data="back_to_start")]])
    else:
        text = build_client_main_text(business)
        markup = build_client_main_keyboard()

    if edit and update.callback_query:
        await update.callback_query.edit_message_text(
            text,
            parse_mode="HTML",
            reply_markup=markup
        )
    else:
        await update.effective_message.reply_text(
            text,
            parse_mode="HTML",
            reply_markup=markup
        )


# Збір заявки

def build_next_question_keyboard(niche: str, field_key: str, collected: dict = None) -> InlineKeyboardMarkup:
    field_type = get_field_type(niche, field_key, collected or {})

    if field_type == "work_format":
        return build_client_work_format_keyboard()

    if field_type == "urgency":
        return InlineKeyboardMarkup([
            [InlineKeyboardButton("🔥 Сьогодні", callback_data="urgency_today")],
            [InlineKeyboardButton("📅 Протягом кількох днів", callback_data="urgency_soon")],
            [InlineKeyboardButton("🕐 Не терміново", callback_data="urgency_normal")],
            [InlineKeyboardButton("🔙 Назад", callback_data="client_back")],
        ])

    return InlineKeyboardMarkup([[InlineKeyboardButton("🔙 Назад", callback_data="client_back")]])


async def ask_next_question(update, context, session_data, edit=False, confirm_text: str = ""):
    user_id = update.effective_user.id
    business_id = session_data.get("business_id")
    business = get_business_by_id(business_id)

    if not business:
        return

    session_data = apply_business_order_defaults(session_data, business)
    niche = business.get("niche") or REPAIR_NICHE
    collected = session_data.get("collected", {})

    next_field = get_next_order_question(niche, collected)

    if next_field is None:
        await _finish_order(update, context, user_id, session_data, business)
        return

    next_field_key = next_field["key"]
    step, total = get_step_info(niche, collected, next_field_key)

    session_data["current_field"] = next_field_key
    save_session(user_id, "client_collecting", session_data)

    prefix = f"✅ <b>{h(confirm_text)}</b>\n\n" if confirm_text else ""

    text = (
        f"{prefix}"
        f"🤖 <b>AI допомагає оформити заявку</b>\n"
        f"<b>Крок {step}/{total}</b>\n\n"
        f"{next_field['question']}"
    )

    markup = build_next_question_keyboard(niche, next_field_key, collected)

    if edit and update.callback_query:
        await update.callback_query.edit_message_text(
            text,
            parse_mode="HTML",
            reply_markup=markup
        )
    else:
        await update.effective_message.reply_text(
            text,
            parse_mode="HTML",
            reply_markup=markup
        )


async def send_next_field_question(update, context, data: dict, niche: str, next_field_key: str, prefix: str = ""):
    data["current_field"] = next_field_key
    save_session(update.effective_user.id, "client_collecting", data)

    collected = data.get("collected", {})
    step, total = get_step_info(niche, collected, next_field_key)

    message = ""

    if prefix:
        message += prefix + "\n\n"

    message += (
        f"🤖 <b>AI допомагає оформити заявку</b>\n"
        f"<b>Крок {step}/{total}</b>\n\n"
        f"{h(get_question_for_field(niche, next_field_key, collected))}"
    )

    await update.message.reply_text(
        message,
        parse_mode="HTML",
        reply_markup=build_next_question_keyboard(niche, next_field_key, collected)
    )


async def go_next_or_finish(update, context, data: dict, business: dict, saved_keys: list = None, fallback_prefix: str = ""):
    user_id = update.effective_user.id
    data = apply_business_order_defaults(data, business)
    niche = business.get("niche") or REPAIR_NICHE
    collected = data.get("collected", {})

    missing = get_missing_field_keys(niche, collected)

    if not missing:
        await _finish_order(update, context, user_id, data, business)
        return

    prefix = build_saved_prefix(saved_keys or []) or fallback_prefix
    await send_next_field_question(update, context, data, niche, missing[0], prefix=prefix)


# AI-збір даних

async def handle_ai_client_collecting(update, context, data: dict, user_text: str, business: dict):
    user_id = update.effective_user.id
    niche = business.get("niche") or REPAIR_NICHE

    collected = data.get("collected", {})
    required_fields = get_order_fields(niche, collected)
    current_field = data.get("current_field")

    user_text_clean = user_text.strip()

    if not user_text_clean:
        await update.message.reply_text("❌ Напишіть, будь ласка, відповідь текстом.")
        return

    if is_question_message(user_text_clean):
        await update.message.reply_text(
            build_order_question_redirect_text(),
            parse_mode="HTML",
            reply_markup=build_order_question_redirect_keyboard()
        )
        return

    if is_help_request(user_text_clean):
        if current_field == "problem":
            await update.message.reply_text(
                build_problem_help_message(collected.get("device", "пристрій"), niche),
                parse_mode="HTML"
            )
            return

        if current_field == "phone":
            await update.message.reply_text(
                "🤖 Напишіть ваш номер телефону.\n\n"
                "Формат: <b>0961234567</b> або <b>+380961234567</b>",
                parse_mode="HTML"
            )
            return

        if current_field == "name":
            await update.message.reply_text(
                "🤖 Напишіть просто ваше ім'я.\n"
                "Наприклад: <b>Олексій</b>",
                parse_mode="HTML"
            )
            return

        await update.message.reply_text(
            "🤖 Опишіть простими словами, що вам потрібно - я оформлю заявку."
        )
        return

    if current_field:
        if current_field == "shipping_city" and is_repair_online_order(niche, collected):
            city = extract_online_shipping_city(user_text_clean)

            if not city:
                await update.message.reply_text(
                    build_online_shipping_city_text(invalid=True),
                    parse_mode="HTML",
                    reply_markup=build_online_shipping_city_keyboard()
                )
                return

            collected["shipping_city"] = city
            data["collected"] = collected
            data["current_field"] = None
            data["ai_collecting"] = True
            data["fallback_mode"] = False
            save_session(user_id, "client_collecting", data)

            await update.message.reply_text(
                build_ai_order_hint(business, collected),
                parse_mode="HTML",
                reply_markup=InlineKeyboardMarkup([
                    [InlineKeyboardButton("🔙 Назад", callback_data="client_back")]
                ])
            )
            return

        # Якщо клієнт у відповідь на одне питання написав одразу кілька даних,
        # не втрачаємо їх: телефон, терміновість, пристрій, проблему, ім'я.
        if "," in user_text_clean or is_valid_phone(user_text_clean) or normalize_urgency(user_text_clean):
            quick_updates = extract_initial_order_fields(niche, user_text_clean, collected)
            if quick_updates:
                collected.update(quick_updates)
                data["collected"] = collected
                await go_next_or_finish(update, context, data, business, saved_keys=list(quick_updates.keys()))
                return

        ok, error_msg = validate_field(current_field, user_text_clean, collected)

        if not ok:
            await update.message.reply_text(error_msg, parse_mode="HTML")
            return

        value = clean_extracted_field(current_field, user_text_clean)

        if current_field == "urgency":
            value = normalize_urgency(user_text_clean)

        collected[current_field] = value
        data["collected"] = collected

        await go_next_or_finish(update, context, data, business, saved_keys=[current_field])
        return

    # Дешева deterministic-класифікація перед AI. Це зменшує помилки і витрати токенів.
    if not current_field:
        quick_updates = extract_initial_order_fields(niche, user_text_clean, collected)
        if quick_updates:
            collected.update(quick_updates)
            data["collected"] = collected
            await go_next_or_finish(update, context, data, business, saved_keys=list(quick_updates.keys()))
            return

        # Для ремонту техніки: якщо перше повідомлення схоже тільки на модель пристрою - записуємо device.
        if niche == "Ремонт техніки" and is_valid_device(user_text_clean) and len(user_text_clean.split()) <= 5 and not any(x in user_text_clean.lower() for x in ["не ", "злам", "слом", "глюч", "екран", "батар", "заряд", ","]):
            collected["device"] = user_text_clean
            data["collected"] = collected
            await go_next_or_finish(update, context, data, business, saved_keys=["device"])
            return

    analysis = await analyze_client_request(
        business=business,
        collected=collected,
        user_text=user_text_clean,
        required_fields=required_fields,
        current_field=current_field
    )

    if not analysis.get("ok"):
        data["fallback_mode"] = True
        save_session(user_id, "client_collecting", data)

        await update.message.reply_text(
            "🤖 Задам кілька коротких запитань по черзі."
        )

        await ask_next_question(update, context, data)
        return

    intent = analysis.get("intent", "leave_request")
    reply = analysis.get("reply") or ""

    if intent == "assistant_help":
        await update.message.reply_text(
            f"🤖 <b>AI допомагає оформити заявку</b>\n\n{h(reply)}" if reply
            else "🤖 Опишіть, що вам потрібно - я оформлю заявку.",
            parse_mode="HTML"
        )
        return

    if intent == "invalid_input":
        next_question = analysis.get("next_question") or "Уточніть деталі:"
        msg = "🤖 <b>AI допомагає оформити заявку</b>\n\n"

        if reply:
            msg += f"{h(reply)}\n\n"

        msg += h(next_question)

        await update.message.reply_text(msg, parse_mode="HTML")
        return

    extracted = analysis.get("extracted") or {}

    # Захист від AI-галюцинацій: якщо у першому повідомленні клієнт не описав проблему,
    # не дозволяємо AI вигадувати проблему самостійно.
    if niche == "Ремонт техніки" and "problem" in extracted:
        problem_candidate = str(extracted.get("problem") or "").strip()
        if problem_candidate and problem_candidate.lower() not in user_text_clean.lower() and not _looks_like_problem_segment(user_text_clean):
            extracted.pop("problem", None)

    collected, saved_keys = merge_extracted_fields(niche, collected, extracted)

    data["collected"] = collected
    data = apply_business_order_defaults(data, business)

    await go_next_or_finish(update, context, data, business, saved_keys=saved_keys)


# Кнопки меню клієнта

async def handle_client_menu(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()

    user_id = query.from_user.id
    username = query.from_user.username or ""

    if query.data == "client_find_business":
        save_session(user_id, "client_search_city")
        prompt_text, prompt_markup = build_business_search_city_prompt()
        await query.edit_message_text(
            prompt_text,
            parse_mode="HTML",
            reply_markup=prompt_markup
        )
        return

    if query.data == "client_all_orders":
        orders = get_recent_client_orders(user_id, username, limit=10)
        await query.edit_message_text(
            build_client_global_orders_text(orders),
            parse_mode="HTML",
            reply_markup=build_client_global_orders_keyboard(orders)
        )
        return

    if query.data.startswith("client_open_biz_"):
        try:
            business_id = int(query.data.replace("client_open_biz_", ""))
        except ValueError:
            await query.edit_message_text("❌ Бізнес не знайдено.")
            return

        business = get_business_by_id(business_id)
        if not business:
            await query.edit_message_text("❌ Бізнес не знайдено.")
            return

        save_session(user_id, "client_start", {"business_id": business_id})
        await show_client_main_menu(update, context, {"business_id": business_id}, edit=True)
        return

    if query.data.startswith("client_my_orders_"):
        try:
            business_id = int(query.data.replace("client_my_orders_", ""))
        except ValueError:
            await query.edit_message_text("❌ Бізнес не знайдено.")
            return

        business = get_business_by_id(business_id)
        if not business:
            await query.edit_message_text("❌ Бізнес не знайдено.")
            return

        orders = get_orders_for_client(business_id, user_id, username)
        save_session(user_id, "client_start", {"business_id": business_id})
        await query.edit_message_text(
            build_client_my_orders_text(orders, business),
            parse_mode="HTML",
            reply_markup=build_client_my_orders_keyboard(orders, business_id)
        )
        return

    direct_business_id = None
    if query.data.startswith("client_order_biz_"):
        try:
            direct_business_id = int(query.data.replace("client_order_biz_", ""))
        except ValueError:
            await query.edit_message_text("❌ Бізнес не знайдено.")
            return

        business = get_business_by_id(direct_business_id)
        if not business:
            await query.edit_message_text(
                "❌ Бізнес не знайдено.\n\n"
                "Можливо, посилання застаріло. Натисніть /start і знайдіть бізнес через пошук."
            )
            return

        save_session(user_id, "client_start", {"business_id": direct_business_id})

    session = get_session(user_id)

    if not session:
        await query.edit_message_text(
            "❌ Сесія закінчилась.\n\n"
            "Натисніть /start або відкрийте посилання бізнесу ще раз."
        )
        return

    data = session["data"]

    if query.data == "client_back":
        save_session(user_id, "client_start", data)
        await show_client_main_menu(update, context, data, edit=True)
        return

    if query.data == "client_order" or query.data.startswith("client_order_biz_"):
        business = get_business_by_id(data.get("business_id"))

        if not business:
            await query.edit_message_text("❌ Бізнес не знайдено.")
            return

        existing_order = find_active_order_for_user(business["id"], user_id, username)

        if existing_order:
            await query.edit_message_text(
                build_existing_order_text(existing_order, business),
                parse_mode="HTML",
                reply_markup=build_existing_order_keyboard(existing_order.get("id"))
            )
            return

        if not check_tariff_limit(data.get("business_id")):
            await notify_owner_about_tariff_pause(context, business)
            await query.edit_message_text(
                build_client_requests_paused_text(business),
                parse_mode="HTML",
                reply_markup=InlineKeyboardMarkup([
                    [InlineKeyboardButton("🔙 Головне меню", callback_data="client_back")]
                ])
            )
            return

        data["collected"] = {}
        data["ai_collecting"] = True
        data["fallback_mode"] = False
        data["current_field"] = None
        data = apply_business_order_defaults(data, business)

        if is_repair_niche(business.get("niche")) and get_business_work_mode(business) == "both":
            data["awaiting_work_format"] = True
            data["ai_collecting"] = False
            save_session(user_id, "client_collecting", data)
            await query.edit_message_text(
                build_client_work_format_text(business),
                parse_mode="HTML",
                reply_markup=build_client_work_format_keyboard()
            )
            return

        if needs_online_shipping_city(business, data):
            data["awaiting_shipping_city"] = True
            data["ai_collecting"] = False
            save_session(user_id, "client_collecting", data)
            await query.edit_message_text(
                build_online_shipping_city_text(),
                parse_mode="HTML",
                reply_markup=build_online_shipping_city_keyboard()
            )
            return

        save_session(user_id, "client_collecting", data)

        await query.edit_message_text(
            build_ai_order_hint(business, data.get("collected", {})),
            parse_mode="HTML",
            reply_markup=InlineKeyboardMarkup([
                [InlineKeyboardButton("🔙 Назад", callback_data="client_back")]
            ])
        )

    elif query.data.startswith("client_view_order_"):
        business = get_business_by_id(data.get("business_id"))

        if not business:
            await query.edit_message_text("❌ Бізнес не знайдено.")
            return

        try:
            order_id = int(query.data.replace("client_view_order_", ""))
        except ValueError:
            await query.edit_message_text("❌ Заявку не знайдено.")
            return

        order = get_order_by_id_for_business(business["id"], order_id)

        if not order or not order_belongs_to_user(order, user_id, username):
            await query.edit_message_text("❌ Заявку не знайдено.")
            return

        await query.edit_message_text(
            build_client_order_summary(order),
            parse_mode="HTML",
            reply_markup=build_client_view_order_keyboard(order_id, order.get("status") in ACTIVE_ORDER_STATUSES)
        )

    elif query.data.startswith("client_cancel_order_"):
        business = get_business_by_id(data.get("business_id"))

        if not business:
            await query.edit_message_text("❌ Бізнес не знайдено.")
            return

        try:
            order_id = int(query.data.replace("client_cancel_order_", ""))
        except ValueError:
            await query.edit_message_text("❌ Заявку не знайдено.")
            return

        order = get_order_by_id_for_business(business["id"], order_id)

        if not order or not order_belongs_to_user(order, user_id, username):
            await query.edit_message_text("❌ Заявку не знайдено.")
            return

        if order.get("status") not in ACTIVE_ORDER_STATUSES:
            await query.edit_message_text(
                "ℹ️ Цю заявку вже закрито або скасовано.",
                parse_mode="HTML",
                reply_markup=InlineKeyboardMarkup([
                    [InlineKeyboardButton("📋 Мої заявки", callback_data=f"client_my_orders_{business['id']}")],
                    [InlineKeyboardButton("🔙 Назад", callback_data="client_back")],
                ])
            )
            return

        await query.edit_message_text(
            build_cancel_order_confirm_text(order, business),
            parse_mode="HTML",
            reply_markup=build_cancel_order_confirm_keyboard(order_id)
        )

    elif query.data.startswith("client_keep_order_"):
        business = get_business_by_id(data.get("business_id"))

        if not business:
            await query.edit_message_text("❌ Бізнес не знайдено.")
            return

        try:
            order_id = int(query.data.replace("client_keep_order_", ""))
        except ValueError:
            await query.edit_message_text("❌ Заявку не знайдено.")
            return

        order = get_order_by_id_for_business(business["id"], order_id)

        if not order or not order_belongs_to_user(order, user_id, username):
            await query.edit_message_text("❌ Заявку не знайдено.")
            return

        await query.edit_message_text(
            build_client_order_summary(order),
            parse_mode="HTML",
            reply_markup=build_client_view_order_keyboard(order_id, order.get("status") in ACTIVE_ORDER_STATUSES)
        )

    elif query.data.startswith("client_confirm_cancel_"):
        business = get_business_by_id(data.get("business_id"))

        if not business:
            await query.edit_message_text("❌ Бізнес не знайдено.")
            return

        try:
            order_id = int(query.data.replace("client_confirm_cancel_", ""))
        except ValueError:
            await query.edit_message_text("❌ Заявку не знайдено.")
            return

        order = get_order_by_id_for_business(business["id"], order_id)

        if not order or not order_belongs_to_user(order, user_id, username):
            await query.edit_message_text("❌ Заявку не знайдено.")
            return

        if order.get("status") not in ACTIVE_ORDER_STATUSES:
            await query.edit_message_text(
                "ℹ️ Цю заявку вже закрито або скасовано.",
                parse_mode="HTML",
                reply_markup=InlineKeyboardMarkup([
                    [InlineKeyboardButton("📋 Мої заявки", callback_data=f"client_my_orders_{business['id']}")],
                    [InlineKeyboardButton("🔙 Головне меню", callback_data="client_back")],
                ])
            )
            return

        update_order_status(order_id, "cancelled")
        updated_order = get_order_by_id_for_business(business["id"], order_id) or order

        await notify_owner_about_order_cancel(context, business, updated_order)

        save_session(user_id, "client_start", {"business_id": business["id"]})

        await query.edit_message_text(
            build_cancelled_order_text(updated_order, business),
            parse_mode="HTML",
            reply_markup=build_cancelled_order_keyboard(business["id"])
        )

    elif query.data.startswith("client_edit_order_"):
        business = get_business_by_id(data.get("business_id"))

        if not business:
            await query.edit_message_text("❌ Бізнес не знайдено.")
            return

        try:
            order_id = int(query.data.replace("client_edit_order_", ""))
        except ValueError:
            await query.edit_message_text("❌ Заявку не знайдено.")
            return

        order = get_order_by_id_for_business(business["id"], order_id)

        if not order or not order_belongs_to_user(order, user_id, username):
            await query.edit_message_text("❌ Заявку не знайдено.")
            return

        if order.get("status") not in ACTIVE_ORDER_STATUSES:
            await query.edit_message_text(
                "ℹ️ Цю заявку вже закрито. Якщо потрібно - створіть нову заявку або перегляньте попередні.",
                reply_markup=InlineKeyboardMarkup([
                    [InlineKeyboardButton("📋 Мої заявки", callback_data=f"client_my_orders_{business['id']}")],
                    [InlineKeyboardButton("📝 Нова заявка", callback_data="client_order")],
                    [InlineKeyboardButton("🔙 Назад", callback_data="client_back")],
                ])
            )
            return

        save_session(user_id, "client_edit_order", {
            "business_id": business["id"],
            "edit_order_id": order_id
        })

        await query.edit_message_text(
            build_edit_order_prompt(order, business),
            parse_mode="HTML",
            reply_markup=InlineKeyboardMarkup([[InlineKeyboardButton("🔙 Назад", callback_data="client_back")]])
        )

    elif query.data == "client_price":
        business = get_business_by_id(data.get("business_id"))
        knowledge = (business.get("knowledge") or "") if business else ""
        kb = parse_knowledge_text(knowledge)

        price_info = (
            kb.get("prices")
            or (business.get("prices") if business else "")
            or get_price_fallback_for_niche(business.get("niche") if business else "")
        )

        keyboard = [
            [InlineKeyboardButton("📝 Залишити заявку", callback_data="client_order")],
            [InlineKeyboardButton("🔙 Назад", callback_data="client_back")],
        ]

        await query.edit_message_text(
            f"💰 <b>Ціни</b>\n\n{h(price_info)}",
            parse_mode="HTML",
            reply_markup=InlineKeyboardMarkup(keyboard)
        )

    elif query.data == "client_schedule":
        business = get_business_by_id(data.get("business_id"))
        knowledge = (business.get("knowledge") or "") if business else ""
        kb = parse_knowledge_text(knowledge)
        schedule = (business.get("schedule") or "") if business else ""
        schedule_info = schedule or kb.get("schedule") or "Графік роботи уточнюйте у менеджера."

        keyboard = [
            [InlineKeyboardButton("📝 Залишити заявку", callback_data="client_order")],
            [InlineKeyboardButton("🔙 Назад", callback_data="client_back")],
        ]

        await query.edit_message_text(
            f"🕐 <b>Графік роботи</b>\n\n{h(schedule_info)}",
            parse_mode="HTML",
            reply_markup=InlineKeyboardMarkup(keyboard)
        )

    elif query.data == "client_address":
        business = get_business_by_id(data.get("business_id"))
        knowledge = (business.get("knowledge") or "") if business else ""
        kb = parse_knowledge_text(knowledge)
        city = (business.get("city") or "") if business else ""

        addr_text = (
            kb.get("address")
            or "Точну адресу сервісу уточнить менеджер після заявки."
        )

        if city and not kb.get("address"):
            addr_text += f"\n\nМісто: {city}"

        keyboard = [
            [InlineKeyboardButton("📝 Залишити заявку", callback_data="client_order")],
            [InlineKeyboardButton("🔙 Назад", callback_data="client_back")],
        ]

        await query.edit_message_text(
            f"📍 <b>Адреса сервісу</b>\n\n{h(addr_text)}",
            parse_mode="HTML",
            reply_markup=InlineKeyboardMarkup(keyboard)
        )

# Кнопки терміновості

async def handle_work_format(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()

    user_id = query.from_user.id
    session = get_session(user_id)

    if not session:
        return

    data = session["data"]
    business = get_business_by_id(data.get("business_id"))

    if not business:
        return

    format_map = {
        "workformat_offline": "Офлайн",
        "workformat_online": "Онлайн",
    }

    selected_format = format_map.get(query.data)
    if not selected_format:
        await query.answer("Оберіть формат заявки", show_alert=True)
        return

    collected = data.get("collected", {})
    collected["format"] = selected_format
    data["collected"] = collected

    if data.get("awaiting_work_format"):
        data["awaiting_work_format"] = False
        data["ai_collecting"] = True
        data["fallback_mode"] = False
        data["current_field"] = None

        if needs_online_shipping_city(business, data):
            data["awaiting_shipping_city"] = True
            data["ai_collecting"] = False
            save_session(user_id, "client_collecting", data)
            await query.edit_message_text(
                build_online_shipping_city_text(),
                parse_mode="HTML",
                reply_markup=build_online_shipping_city_keyboard()
            )
            return

        save_session(user_id, "client_collecting", data)

        await query.edit_message_text(
            build_ai_order_hint(business, collected),
            parse_mode="HTML",
            reply_markup=InlineKeyboardMarkup([
                [InlineKeyboardButton("🔙 Назад", callback_data="client_back")]
            ])
        )
        return

    save_session(user_id, "client_collecting", data)

    niche = business.get("niche") or REPAIR_NICHE
    missing = get_missing_field_keys(niche, collected)

    if not missing:
        await _finish_order(update, context, user_id, data, business)
        return

    next_field_key = missing[0]
    data["current_field"] = next_field_key
    save_session(user_id, "client_collecting", data)

    step, total = get_step_info(niche, collected, next_field_key)

    await query.edit_message_text(
        "✅ <b>Формат заявки записав.</b>\n\n"
        "🤖 <b>AI допомагає оформити заявку</b>\n"
        f"<b>Крок {step}/{total}</b>\n\n"
        f"{h(get_question_for_field(niche, next_field_key, collected))}",
        parse_mode="HTML",
        reply_markup=build_next_question_keyboard(niche, next_field_key, collected)
    )


async def handle_urgency(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()

    user_id = query.from_user.id
    session = get_session(user_id)

    if not session:
        return

    data = session["data"]
    business = get_business_by_id(data.get("business_id"))

    if not business:
        return

    urgency_map = {
        "urgency_today": "Сьогодні 🔥",
        "urgency_soon": "Протягом кількох днів",
        "urgency_normal": "Не терміново",
    }

    collected = data.get("collected", {})
    current_field = data.get("current_field", "urgency")
    collected[current_field] = urgency_map.get(query.data, "-")

    data["collected"] = collected

    save_session(user_id, "client_collecting", data)

    niche = business.get("niche") or REPAIR_NICHE
    missing = get_missing_field_keys(niche, collected)

    if not missing:
        await _finish_order(update, context, user_id, data, business)
        return

    next_field_key = missing[0]
    data["current_field"] = next_field_key
    save_session(user_id, "client_collecting", data)

    step, total = get_step_info(niche, collected, next_field_key)

    await query.edit_message_text(
        "✅ <b>Терміновість записав.</b>\n\n"
        "🤖 <b>AI допомагає оформити заявку</b>\n"
        f"<b>Крок {step}/{total}</b>\n\n"
        f"{h(get_question_for_field(niche, next_field_key, collected))}",
        parse_mode="HTML",
        reply_markup=build_next_question_keyboard(niche, next_field_key, collected)
    )


# Пошук бізнесу


def normalize_search_text(value: str) -> str:
    value = str(value or "").strip().casefold()
    replacements = {
        "ё": "е",
        "ї": "і",
        "є": "е",
        "ґ": "г",
        "’": "'",
        "`": "'",
        "ʼ": "'",
    }
    for old, new in replacements.items():
        value = value.replace(old, new)
    value = re.sub(r"[^0-9a-zа-яіїєґ'\s]+", " ", value, flags=re.IGNORECASE)
    value = re.sub(r"\s+", " ", value).strip()
    return value


def search_tokens(value: str) -> set:
    normalized = normalize_search_text(value)
    return {token for token in normalized.split() if len(token) >= 3}


def fuzzy_ratio(a: str, b: str) -> float:
    a = normalize_search_text(a)
    b = normalize_search_text(b)
    if not a or not b:
        return 0.0
    return SequenceMatcher(None, a, b).ratio()


def score_business_match(business: dict, query: str, city: str = "") -> float:
    q = normalize_search_text(query)
    if not q:
        return 0.0

    code = str(business.get("link_code") or "").strip().casefold()
    name = normalize_search_text(business.get("name") or "")
    niche = normalize_search_text(business.get("niche") or "")
    services = normalize_search_text(business.get("services") or "")
    biz_city = normalize_search_text(business.get("city") or "")

    if code and q == code:
        return 1.0

    city_norm = normalize_search_text(city)
    city_bonus = 0.0
    if city_norm:
        if city_norm == biz_city:
            city_bonus = 0.12
        elif city_norm in biz_city or biz_city in city_norm:
            city_bonus = 0.08
        else:
            # Якщо місто вказали, але воно зовсім інше - сильно знижуємо релевантність.
            city_bonus = -0.25

    all_text = " ".join([name, niche, services, biz_city])
    q_tokens = search_tokens(q)
    all_tokens = search_tokens(all_text)

    score = 0.0

    if name == q:
        score = max(score, 0.98)
    if q and q in name:
        score = max(score, 0.86)
    if name and name in q and len(name) >= 3:
        score = max(score, 0.84)

    # Пошук за напрямком послуг: ремонт, айфон, клінінг, салон тощо.
    if q and (q in niche or q in services):
        score = max(score, 0.68)

    if q_tokens and all_tokens:
        overlap = len(q_tokens & all_tokens) / max(1, len(q_tokens))
        if overlap:
            score = max(score, 0.55 + overlap * 0.25)

    # Fuzzy для помилок у назві: великі/маленькі літери, одна-дві помилки, е/є/ї тощо.
    name_ratio = fuzzy_ratio(q, name)
    score = max(score, name_ratio * 0.92)

    # Якщо запит складається з кількох слів, порівнюємо також з усіма полями разом.
    all_ratio = fuzzy_ratio(q, all_text)
    score = max(score, all_ratio * 0.62)

    return max(0.0, min(1.0, score + city_bonus))


def is_business_code_query(query: str) -> bool:
    query = str(query or "").strip().upper()
    return bool(re.fullmatch(r"[A-Z0-9]{6}", query))


def search_businesses_smart(query: str, city: str = "", limit: int = 8) -> list:
    """Розумний пошук бізнесу:
    - не залежить від великих/малих літер;
    - знаходить схожі назви з дрібними помилками;
    - шукає по назві, ніші, послугах і коду;
    - не показує зовсім нерелевантний мусор.
    """
    query = str(query or "").strip()
    city = str(city or "").strip()

    if not query:
        return []

    # Точний код - найшвидший шлях.
    if is_business_code_query(query):
        business = get_business_by_code(query.upper())
        return [business] if business else []

    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    cursor = conn.cursor()

    cursor.execute("""
        SELECT * FROM businesses
        WHERE name IS NOT NULL
    """)
    rows = [dict(row) for row in cursor.fetchall()]
    conn.close()

    scored = []
    for business in rows:
        score = score_business_match(business, query, city)

        # Поріг не дуже низький, щоб не показувати рандом.
        # Для коротких запитів типу "ремонт" достатньо збігу по ніші/послугах.
        if score >= 0.42:
            scored.append((score, business))

    scored.sort(
        key=lambda item: (
            item[0],
            int(item[1].get("orders_count") or 0),
            str(item[1].get("created_at") or "")
        ),
        reverse=True
    )

    return [business for _, business in scored[:limit]]


def is_exact_business_match(business: dict, query: str) -> bool:
    q = normalize_search_text(query)
    if not q:
        return False

    code = str(business.get("link_code") or "").strip().casefold()
    name = normalize_search_text(business.get("name") or "")

    # Код або абсолютно точна назва - відкриваємо одразу.
    if q == code or q == name:
        return True

    # Якщо ввели назву майже ідеально, теж відкриваємо одразу.
    # Але не для широких запитів типу "ремонт", бо там краще показати варіанти.
    if len(q) >= 6 and fuzzy_ratio(q, name) >= 0.94:
        return True

    return False


def build_business_search_city_prompt() -> tuple:
    text = (
        "🔎 <b>Пошук бізнесу</b>\n\n"
        "Спочатку напишіть населений пункт, де шукаємо сервіс.\n\n"
        "<blockquote><i><b>Наприклад:</b> Київ, Львів, Одеса, смт Бородянка або село Жашків.</i></blockquote>"
    )
    keyboard = [[InlineKeyboardButton("🔙 Кабінет клієнта", callback_data="role_client")]]
    return text, InlineKeyboardMarkup(keyboard)


def build_business_search_query_prompt(city: str) -> tuple:
    text = (
        "🔎 <b>Пошук бізнесу</b>\n\n"
        f"📍 <b>Локація:</b> {h(city)}\n\n"
        "Тепер напишіть назву сервісу, код бізнесу або що потрібно відремонтувати.\n\n"
        "<blockquote><i><b>Наприклад:</b> iFix, KLV46R, ремонт iPhone або ремонт ноутбука.</i></blockquote>"
    )
    keyboard = [[InlineKeyboardButton("🔙 Кабінет клієнта", callback_data="role_client")]]
    return text, InlineKeyboardMarkup(keyboard)


def build_business_suggestions_text(query: str, city: str, results: list) -> str:
    text = (
        "🔎 <b>Можливі варіанти</b>\n\n"
        f"📍 Локація: <b>{h(city or 'не вказано')}</b>\n"
        f"🔍 Запит: <b>{h(query)}</b>\n\n"
        "Я знайшов схожі бізнеси. Оберіть потрібний зі списку нижче."
    )
    return text


def build_business_suggestions_keyboard(results: list) -> InlineKeyboardMarkup:
    keyboard = []

    for b in results[:8]:
        name = b.get("name") or "Без назви"
        city = b.get("city") or "-"
        niche = b.get("niche") or "-"
        code = b.get("link_code")
        label = f"🏢 {name} - {city}"
        if niche:
            label += f" ({niche})"
        keyboard.append([InlineKeyboardButton(label[:64], callback_data=f"select_biz_{code}")])

    keyboard.append([InlineKeyboardButton("🔎 Шукати заново", callback_data="role_client")])
    keyboard.append([InlineKeyboardButton("🔙 На початок", callback_data="back_to_start")])
    return InlineKeyboardMarkup(keyboard)

async def handle_business_search(update, context):
    user_id = update.effective_user.id
    text = update.message.text.strip()
    session = get_session(user_id)
    data = session["data"] if session else {}
    city = data.get("search_city", "")

    # Якщо клієнт ввів точний код бізнесу - відкриваємо одразу.
    if is_business_code_query(text):
        business = get_business_by_code(text.upper())
        if business:
            await handle_client_start(update, context, business)
            return

    results = search_businesses_smart(text, city=city)

    if not results:
        keyboard = [
            [InlineKeyboardButton("🔎 Шукати заново", callback_data="role_client")],
            [InlineKeyboardButton("🔙 На початок", callback_data="back_to_start")],
        ]
        await update.message.reply_text(
            f"❌ <b>Нічого не знайшов</b>\n\n"
            f"📍 Локація: <b>{h(city or '-')}</b>\n"
            f"🔍 Запит: <b>{h(text)}</b>\n\n"
            f"Спробуйте написати точнішу назву компанії, код бізнесу або напрямок послуг.",
            parse_mode="HTML",
            reply_markup=InlineKeyboardMarkup(keyboard)
        )
        return

    exact_matches = [b for b in results if is_exact_business_match(b, text)]
    if len(exact_matches) == 1:
        await handle_client_start(update, context, exact_matches[0])
        return

    await update.message.reply_text(
        build_business_suggestions_text(text, city, results),
        parse_mode="HTML",
        reply_markup=build_business_suggestions_keyboard(results)
    )


async def handle_select_business(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()

    code = query.data.replace("select_biz_", "")
    business = get_business_by_code(code)

    if business:
        await query.delete_message()
        await handle_client_start(update, context, business)
    else:
        await query.edit_message_text("❌ Бізнес не знайдено.")


# Обробка тексту


async def handle_knowledge_collecting_text(update: Update, context: ContextTypes.DEFAULT_TYPE, data: dict, text: str):
    user_id = update.effective_user.id
    business = get_business(user_id)

    if not business:
        await update.message.reply_text("❌ Бізнес не знайдено. Натисніть /start")
        return

    kb = data.get("kb", {})
    step_index = int(data.get("step_index", 0))
    step = get_knowledge_step(step_index)

    if not step:
        knowledge_text = build_knowledge_text(kb)
        if knowledge_text:
            update_business(user_id, "knowledge", knowledge_text)
        save_session(user_id, "entrepreneur_menu")
        await update.message.reply_text(
            "✅ <b>База знань збережена.</b>\n\n" + (h(knowledge_text) if knowledge_text else "<i>Ваша база знань поки відсутня</i>"),
            parse_mode="HTML"
        )
        await show_entrepreneur_menu(update, context, business)
        return

    text_clean = text.strip()

    if not text_clean:
        await update.message.reply_text("❌ Напишіть відповідь або натисніть Пропустити.")
        return

    # На кроці FAQ знак питання нормальний: власник пише питання клієнта і відповідь.
    # На інших кроках питання не зберігаємо як базу знань, бо AI має отримати готову інформацію.
    if is_question_message(text_clean):
        if step["key"] == "faq" and is_valid_faq_knowledge_answer(text_clean):
            pass
        else:
            await update.message.reply_text(
                "🤖 <b>AI-база знань</b>\n\n"
                "Тут потрібно не питання, а готова інформація для клієнтів.\n"
                "Напишіть відповідь так, ніби це говорить ваш адміністратор.\n\n"
                f"{step['question']}",
                parse_mode="HTML",
                reply_markup=build_knowledge_keyboard(step_index)
            )
            return

    normalized = normalize_knowledge_answer(step["key"], text_clean)

    if len(normalized) < 2:
        await update.message.reply_text("❌ Дуже коротко. Напишіть детальніше або натисніть Пропустити.")
        return

    kb[step["key"]] = normalized

    # Для кнопки графіку роботи у клієнта теж оновлюємо окреме поле, якщо db.py дозволяє.
    if step["key"] == "schedule":
        try:
            update_business(user_id, "schedule", normalized)
        except Exception:
            pass

    step_index += 1

    if step_index >= len(KNOWLEDGE_STEPS):
        knowledge_text = build_knowledge_text(kb)
        update_business(user_id, "knowledge", knowledge_text)
        save_session(user_id, "entrepreneur_menu")

        await update.message.reply_text(
            "✅ <b>База знань готова.</b>\n\n"
            "AI тепер має структуровану інформацію про ваш бізнес:\n\n"
            f"{h(knowledge_text)}",
            parse_mode="HTML"
        )
        await show_entrepreneur_menu(update, context, business)
        return

    save_session(user_id, "knowledge_collecting", {"kb": kb, "step_index": step_index})

    await update.message.reply_text(
        f"✅ <b>{h(step['title'])} збережено.</b>\n\n" + build_knowledge_question(step_index, kb, business.get("niche")),
        parse_mode="HTML",
        reply_markup=build_knowledge_keyboard(step_index)
    )


async def handle_registration(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    text = update.message.text.strip()
    session = get_session(user_id)

    if not session:
        return

    state = session["state"]
    data = session["data"]

    if state == "client_start":
        await handle_client_business_question(update, context, data, text)
        return

    if state == "client_search_city":
        if len(text) < 2:
            await update.message.reply_text("❌ Введіть локацію. Наприклад: <b>Київ</b>", parse_mode="HTML")
            return

        save_session(user_id, "client_search", {"search_city": text})
        prompt_text, prompt_markup = build_business_search_query_prompt(text)
        await update.message.reply_text(
            prompt_text,
            parse_mode="HTML",
            reply_markup=prompt_markup
        )
        return

    if state == "client_search":
        await handle_business_search(update, context)
        return

    if state == "reg_name":
        if len(text) < 2:
            await update.message.reply_text("❌ Назва занадто коротка. Спробуйте ще раз:")
            return

        update_business(user_id, "name", text)
        save_session(user_id, "reg_niche")

        keyboard = [
            [InlineKeyboardButton(NICHE_BUTTON_TEXTS.get(n, n), callback_data=f"niche_{n}")]
            for n in NICHES
        ]
        keyboard.append([BTN_BACK_START])

        await update.message.reply_text(
            "✅ <b>Назва збережена!</b>\n\n"
            "👔 <b>Крок 2/4 - Оберіть нішу бізнесу</b>\n\n"
            "Виберіть напрямок, під який AI буде збирати заявки:",
            parse_mode="HTML",
            reply_markup=InlineKeyboardMarkup(keyboard)
        )

    elif state == "reg_city":
        if len(text) < 2:
            await update.message.reply_text("❌ Введіть коректну назву міста:")
            return

        update_business(user_id, "city", text)
        save_session(user_id, "reg_services")

        business = get_business(user_id)
        await update.message.reply_text(
            build_registration_services_text((business or {}).get("niche")),
            parse_mode="HTML",
            reply_markup=InlineKeyboardMarkup([[BTN_BACK_START]])
        )

    elif state == "reg_work_mode":
        await update.message.reply_text(
            build_work_mode_registration_text({"niche": REPAIR_NICHE}),
            parse_mode="HTML",
            reply_markup=build_work_mode_keyboard()
        )

    elif state == "reg_services":
        if len(text) < 3:
            await update.message.reply_text("❌ Напишіть хоча б одну послугу:")
            return

        update_business(user_id, "services", text)
        save_session(user_id, "entrepreneur_menu")

        business = get_business(user_id)
        username = context.bot.username
        link = f"https://t.me/{username}?start=biz_{business['link_code']}"
        work_mode_line = ""
        if is_repair_niche(business.get("niche")):
            work_mode_line = f"Формат роботи: {h(get_work_mode_label(business.get('work_mode')))}\n"

        await update.message.reply_text(
            f"🎉 <b>Ваш AI-адміністратор готовий!</b>\n\n"
            f"Бізнес: <b>{h(business.get('name'))}</b>\n"
            f"Місто: {h(business.get('city', '-'))}\n"
            f"Ніша: {h(business.get('niche', '-'))}\n\n"
            f"{work_mode_line}"
            f"🔗 <b>Посилання для клієнтів:</b>\n"
            f"{h(link)}\n\n"
            f"📱 <b>Код бізнесу:</b> <code>{h(business.get('link_code'))}</code>\n\n"
            f"Розмістіть посилання в Instagram, Telegram або на сайті.",
            parse_mode="HTML"
        )

        await show_entrepreneur_menu(update, context, business)

    elif state == "knowledge_ai_edit":
        await handle_knowledge_ai_edit_text(update, context, data, text)

    elif state == "knowledge_ai_edit_value":
        await handle_knowledge_ai_edit_value_text(update, context, data, text)

    elif state == "knowledge_collecting":
        await handle_knowledge_collecting_text(update, context, data, text)

    elif state == "knowledge_input":
        # Legacy fallback: якщо стара сесія залишилась після оновлення коду.
        save_session(user_id, "knowledge_collecting", {"kb": {}, "step_index": 0})
        await handle_knowledge_collecting_text(update, context, {"kb": {}, "step_index": 0}, text)

    elif state == "client_collecting":
        business = get_business_by_id(data.get("business_id"))

        if not business:
            await update.message.reply_text("❌ Бізнес не знайдено. Спробуйте перейти по посиланню ще раз.")
            return

        if is_question_message(text):
            await update.message.reply_text(
                build_order_question_redirect_text(),
                parse_mode="HTML",
                reply_markup=build_order_question_redirect_keyboard()
            )
            return

        if data.get("awaiting_work_format"):
            await update.message.reply_text(
                build_client_work_format_text(business),
                parse_mode="HTML",
                reply_markup=build_client_work_format_keyboard()
            )
            return

        if data.get("awaiting_shipping_city"):
            city = extract_online_shipping_city(text)

            if not city:
                await update.message.reply_text(
                    build_online_shipping_city_text(invalid=True),
                    parse_mode="HTML",
                    reply_markup=build_online_shipping_city_keyboard()
                )
                return

            collected = data.get("collected", {})
            collected["shipping_city"] = city
            data["collected"] = collected
            data["awaiting_shipping_city"] = False
            data["current_field"] = None
            data["ai_collecting"] = True
            data["fallback_mode"] = False
            save_session(user_id, "client_collecting", data)

            await update.message.reply_text(
                build_ai_order_hint(business, collected),
                parse_mode="HTML",
                reply_markup=InlineKeyboardMarkup([
                    [InlineKeyboardButton("🔙 Назад", callback_data="client_back")]
                ])
            )
            return

        if data.get("ai_collecting") and not data.get("fallback_mode"):
            await handle_ai_client_collecting(update, context, data, text, business)
            return

        current_field = data.get("current_field")
        niche = business.get("niche") or REPAIR_NICHE

        if not current_field:
            await ask_next_question(update, context, data)
            return

        if current_field == "shipping_city" and is_repair_online_order(niche, data.get("collected", {})):
            city = extract_online_shipping_city(text)

            if not city:
                await update.message.reply_text(
                    build_online_shipping_city_text(invalid=True),
                    parse_mode="HTML",
                    reply_markup=build_online_shipping_city_keyboard()
                )
                return

            collected = data.get("collected", {})
            collected["shipping_city"] = city
            data["collected"] = collected
            data["current_field"] = None
            data["ai_collecting"] = True
            data["fallback_mode"] = False
            save_session(user_id, "client_collecting", data)

            await update.message.reply_text(
                build_ai_order_hint(business, collected),
                parse_mode="HTML",
                reply_markup=InlineKeyboardMarkup([
                    [InlineKeyboardButton("🔙 Назад", callback_data="client_back")]
                ])
            )
            return

        if is_question_message(text):
            await update.message.reply_text(
                build_order_question_redirect_text(),
                parse_mode="HTML",
                reply_markup=build_order_question_redirect_keyboard()
            )
            return

        ok, error_msg = validate_field(current_field, text, data.get("collected", {}))

        if not ok:
            await update.message.reply_text(error_msg, parse_mode="HTML")
            return

        collected = data.get("collected", {})
        collected[current_field] = clean_extracted_field(current_field, text)
        data["collected"] = collected

        await ask_next_question(update, context, data)

    elif state == "client_edit_order":
        await handle_client_order_edit_text(update, context, data, text)

    elif state == "client_phone_only":
        if not is_valid_phone(text):
            await update.message.reply_text(
                "❌ Це не схоже на коректний номер.\n\n"
                "Напишіть номер у форматі:\n"
                "<b>0961234567</b> або <b>+380961234567</b>",
                parse_mode="HTML"
            )
            return

        business = get_business_by_id(data.get("business_id"))

        if business:
            existing_order = find_active_order_for_user(
                business["id"],
                user_id,
                update.effective_user.username or ""
            )

            if existing_order:
                await update.message.reply_text(
                    build_existing_order_text(existing_order, business),
                    parse_mode="HTML",
                    reply_markup=build_existing_order_keyboard(existing_order.get("id"))
                )
                save_session(user_id, "client_start", data)
                return

        collected = data.get("collected", {})
        collected["phone"] = normalize_phone(text)
        data["collected"] = collected

        save_session(user_id, "client_done", data)

        await update.message.reply_text(
            "✅ <b>Дякую!</b>\n\n"
            "Менеджер зв'яжеться з вами найближчим часом. 🙏",
            parse_mode="HTML"
        )

        if business:
            await notify_owner_about_new_order(context, data, update.effective_user.username or "", business)


# Завершення заявки

async def _finish_order(update, context, user_id, data, business=None):
    if business is None:
        business = get_business_by_id(data.get("business_id"))

    if not business:
        return

    data = apply_business_order_defaults(data, business)

    existing_order = find_active_order_for_user(
        business["id"],
        user_id,
        update.effective_user.username or ""
    )

    if existing_order:
        save_session(user_id, "client_start", {"business_id": business["id"]})
        await update.effective_message.reply_text(
            build_existing_order_text(existing_order, business),
            parse_mode="HTML",
            reply_markup=build_existing_order_keyboard(existing_order.get("id"))
        )
        return

    if not check_tariff_limit(business["id"]):
        save_session(user_id, "client_start", {"business_id": business["id"]})
        await notify_owner_about_tariff_pause(context, business)
        await update.effective_message.reply_text(
            build_client_requests_paused_text(business),
            parse_mode="HTML",
            reply_markup=InlineKeyboardMarkup([
                [InlineKeyboardButton("🏠 Головний екран", callback_data="client_back")]
            ])
        )
        return

    username = update.effective_user.username or ""

    await update.effective_message.reply_text("🤖 AI аналізує заявку...")

    ai_comment = await qualify_order(business, data.get("collected", {}))
    data["ai_comment"] = ai_comment

    save_order(business["id"], user_id, username, data)
    save_session(user_id, "client_done", data)

    collected = data.get("collected", {})
    urgency = collected.get("urgency", "")
    is_hot = bool(urgency and "Сьогодні" in urgency)

    hot_text = "🔥 Ваш запит позначено як терміновий!\n\n" if is_hot else ""

    await update.effective_message.reply_text(
        "✅ <b>Заявку прийнято!</b>\n\n"
        + h(hot_text)
        + "Менеджер зв'яжеться з вами найближчим часом.\n\n"
        "Дякуємо за звернення! 🙏",
        parse_mode="HTML",
        reply_markup=InlineKeyboardMarkup([
            [InlineKeyboardButton("🏠 Головний екран", callback_data="client_back")]
        ])
    )

    await notify_owner_about_new_order(context, data, username, business)
