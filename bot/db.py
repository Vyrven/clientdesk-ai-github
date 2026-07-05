import os
import sqlite3
import json
import random
import string
from datetime import datetime, timedelta

DB_PATH = os.getenv("DB_PATH", "clientdesk.db")
DATETIME_FORMAT = "%Y-%m-%d %H:%M:%S"
SUBSCRIPTION_DAYS = 30
PAID_TARIFFS = {"start", "business", "vip"}


def current_time():
    return datetime.now().replace(microsecond=0)


def format_datetime(value: datetime) -> str:
    return value.strftime(DATETIME_FORMAT)


def parse_datetime(value):
    if not value:
        return None

    text = str(value).strip()
    for fmt in (DATETIME_FORMAT, "%Y-%m-%dT%H:%M:%S"):
        try:
            return datetime.strptime(text[:19], fmt)
        except ValueError:
            pass
    return None


def display_datetime(value) -> str:
    parsed = parse_datetime(value)
    if parsed:
        return parsed.strftime("%Y-%m-%d %H:%M")

    text = str(value or "").strip()
    return text[:16] if len(text) >= 16 else text


def is_paid_tariff(tariff: str) -> bool:
    return str(tariff or "").lower() in PAID_TARIFFS


def build_subscription_period(days: int = SUBSCRIPTION_DAYS):
    started_at = current_time()
    expires_at = started_at + timedelta(days=days)
    return format_datetime(started_at), format_datetime(expires_at)


def get_connection():
    db_dir = os.path.dirname(os.path.abspath(DB_PATH))
    if db_dir:
        os.makedirs(db_dir, exist_ok=True)

    conn = sqlite3.connect(DB_PATH)
    conn.execute("PRAGMA busy_timeout = 5000")
    conn.execute("PRAGMA journal_mode = WAL")
    conn.execute("PRAGMA foreign_keys = ON")
    conn.row_factory = sqlite3.Row
    return conn


def init_db():
    conn = get_connection()
    cursor = conn.cursor()

    cursor.execute("""
        CREATE TABLE IF NOT EXISTS businesses (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            owner_id INTEGER UNIQUE,
            name TEXT,
            niche TEXT,
            city TEXT,
            services TEXT,
            schedule TEXT,
            prices TEXT,
            knowledge TEXT,
            ai_style TEXT DEFAULT 'friendly',
            link_code TEXT UNIQUE,
            tariff TEXT DEFAULT 'free',
            orders_count INTEGER DEFAULT 0,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        )
    """)

    cursor.execute("""
        CREATE TABLE IF NOT EXISTS orders (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            business_id INTEGER,
            client_id INTEGER,
            client_name TEXT,
            client_phone TEXT,
            client_username TEXT,
            service TEXT,
            problem TEXT,
            device TEXT,
            car TEXT,
            format TEXT,
            shipping_city TEXT,
            urgency TEXT,
            district TEXT,
            status TEXT DEFAULT 'new',
            ai_comment TEXT,
            is_hot INTEGER DEFAULT 0,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
            FOREIGN KEY (business_id) REFERENCES businesses(id)
        )
    """)

    cursor.execute("""
        CREATE TABLE IF NOT EXISTS sessions (
            user_id INTEGER PRIMARY KEY,
            role TEXT,
            business_id INTEGER,
            state TEXT,
            data TEXT,
            messages TEXT
        )
    """)

    cursor.execute("""
        CREATE TABLE IF NOT EXISTS client_link_reminders (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            business_id INTEGER NOT NULL,
            chat_id INTEGER NOT NULL,
            user_id INTEGER NOT NULL,
            username TEXT,
            entered_at TEXT NOT NULL,
            remind_at TEXT NOT NULL,
            status TEXT DEFAULT 'pending',
            sent_at TEXT,
            created_at TEXT NOT NULL
        )
    """)

    ensure_column(cursor, "businesses", "subscription_started_at", "TEXT")
    ensure_column(cursor, "businesses", "subscription_expires_at", "TEXT")
    ensure_column(cursor, "businesses", "work_mode", "TEXT")
    ensure_column(cursor, "orders", "shipping_city", "TEXT")

    cursor.execute("""
        CREATE INDEX IF NOT EXISTS idx_orders_business_created
        ON orders (business_id, created_at)
    """)
    cursor.execute("""
        CREATE INDEX IF NOT EXISTS idx_orders_client_created
        ON orders (client_id, created_at)
    """)
    cursor.execute("""
        CREATE INDEX IF NOT EXISTS idx_orders_username_created
        ON orders (client_username, created_at)
    """)
    cursor.execute("""
        CREATE INDEX IF NOT EXISTS idx_link_reminders_due
        ON client_link_reminders (status, remind_at)
    """)
    cursor.execute("""
        CREATE INDEX IF NOT EXISTS idx_link_reminders_client
        ON client_link_reminders (business_id, user_id, status)
    """)

    conn.commit()
    conn.close()
    print("Database initialized.")


def ensure_column(cursor, table: str, column: str, definition: str):
    cursor.execute(f"PRAGMA table_info({table})")
    existing = {row[1] for row in cursor.fetchall()}
    if column not in existing:
        cursor.execute(f"ALTER TABLE {table} ADD COLUMN {column} {definition}")


def generate_link_code():
    return ''.join(random.choices(string.ascii_uppercase + string.digits, k=6))


def create_business(owner_id):
    conn = get_connection()
    cursor = conn.cursor()
    code = generate_link_code()
    cursor.execute(
        "INSERT OR IGNORE INTO businesses (owner_id, link_code, created_at) VALUES (?, ?, ?)",
        (owner_id, code, format_datetime(current_time()))
    )
    conn.commit()
    conn.close()


def update_business(owner_id, field, value):
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute(
        f"UPDATE businesses SET {field} = ? WHERE owner_id = ?",
        (value, owner_id)
    )
    conn.commit()
    conn.close()


def get_business(owner_id):
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute("SELECT * FROM businesses WHERE owner_id = ?", (owner_id,))
    row = cursor.fetchone()
    conn.close()
    return dict(row) if row else None


def get_business_by_id(business_id):
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute("SELECT * FROM businesses WHERE id = ?", (business_id,))
    row = cursor.fetchone()
    conn.close()
    return dict(row) if row else None


def get_business_by_code(code):
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute("SELECT * FROM businesses WHERE link_code = ?", (code,))
    row = cursor.fetchone()
    conn.close()
    return dict(row) if row else None


def search_businesses(query: str):
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute("""
        SELECT * FROM businesses
        WHERE (name LIKE ? OR city LIKE ? OR link_code = ?)
        AND name IS NOT NULL
        LIMIT 5
    """, (f"%{query}%", f"%{query}%", query.upper()))
    rows = cursor.fetchall()
    conn.close()
    return [dict(row) for row in rows]


def count_orders_in_window(business_id, started_at=None, expires_at=None) -> int:
    conn = get_connection()
    cursor = conn.cursor()

    query = "SELECT COUNT(*) FROM orders WHERE business_id = ?"
    params = [business_id]

    if started_at:
        query += " AND created_at >= ?"
        params.append(format_datetime(started_at) if isinstance(started_at, datetime) else str(started_at))

    if expires_at:
        query += " AND created_at < ?"
        params.append(format_datetime(expires_at) if isinstance(expires_at, datetime) else str(expires_at))

    cursor.execute(query, tuple(params))
    count = int(cursor.fetchone()[0] or 0)
    conn.close()
    return count


def schedule_link_reminder_record(business_id, chat_id, user_id, username="", delay_seconds=3600) -> dict:
    now = current_time()
    entered_at = format_datetime(now)
    remind_at = format_datetime(now + timedelta(seconds=int(delay_seconds)))

    conn = get_connection()
    cursor = conn.cursor()

    cursor.execute("""
        UPDATE client_link_reminders
        SET chat_id = ?, username = ?, entered_at = ?, remind_at = ?, created_at = ?
        WHERE business_id = ? AND user_id = ? AND status = 'pending'
    """, (
        int(chat_id), username or "", entered_at, remind_at, entered_at,
        int(business_id), int(user_id)
    ))

    if cursor.rowcount == 0:
        cursor.execute("""
            INSERT INTO client_link_reminders (
                business_id, chat_id, user_id, username,
                entered_at, remind_at, status, created_at
            ) VALUES (?, ?, ?, ?, ?, ?, 'pending', ?)
        """, (
            int(business_id), int(chat_id), int(user_id), username or "",
            entered_at, remind_at, entered_at
        ))

    cursor.execute("""
        SELECT * FROM client_link_reminders
        WHERE business_id = ? AND user_id = ? AND status = 'pending'
        ORDER BY id DESC
        LIMIT 1
    """, (int(business_id), int(user_id)))
    row = cursor.fetchone()
    conn.commit()
    conn.close()
    return dict(row) if row else {}


def get_due_link_reminders(limit=50) -> list:
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute("""
        SELECT * FROM client_link_reminders
        WHERE status = 'pending' AND remind_at <= ?
        ORDER BY remind_at ASC
        LIMIT ?
    """, (format_datetime(current_time()), int(limit)))
    rows = cursor.fetchall()
    conn.close()
    return [dict(row) for row in rows]


def mark_link_reminder_status(reminder_id, status: str):
    conn = get_connection()
    cursor = conn.cursor()
    sent_at = format_datetime(current_time()) if status == "sent" else None
    cursor.execute("""
        UPDATE client_link_reminders
        SET status = ?, sent_at = ?
        WHERE id = ? AND status = 'pending'
    """, (status, sent_at, int(reminder_id)))
    conn.commit()
    conn.close()


def cancel_pending_link_reminders(business_id, user_id, username=""):
    conn = get_connection()
    cursor = conn.cursor()
    params = [int(business_id), int(user_id)]
    query = """
        UPDATE client_link_reminders
        SET status = 'cancelled'
        WHERE business_id = ? AND user_id = ? AND status = 'pending'
    """

    if username:
        query = """
            UPDATE client_link_reminders
            SET status = 'cancelled'
            WHERE business_id = ?
              AND status = 'pending'
              AND (user_id = ? OR lower(username) = lower(?))
        """
        params.append(str(username).lstrip("@"))

    cursor.execute(query, tuple(params))
    conn.commit()
    conn.close()


def get_tariff_usage(business: dict) -> dict:
    from config import TARIFF_LIMITS

    if not business:
        return {
            "stored_tariff": "free",
            "active": False,
            "expired": True,
            "is_paid": False,
            "limit": 0,
            "used": 0,
            "remaining": 0,
            "started_at": None,
            "expires_at": None,
        }

    business_id = business.get("id")
    stored_tariff = str(business.get("tariff") or "free").lower()
    paid = is_paid_tariff(stored_tariff)
    started_at = parse_datetime(business.get("subscription_started_at"))
    expires_at = parse_datetime(business.get("subscription_expires_at"))
    now = current_time()

    if paid:
        active = bool(expires_at and expires_at > now)
        expired = not active
        limit = int(TARIFF_LIMITS.get(stored_tariff, 5))
        used = count_orders_in_window(business_id, started_at, expires_at) if business_id and started_at else int(business.get("orders_count") or 0)
    else:
        active = True
        expired = False
        limit = int(TARIFF_LIMITS.get("free", 5))
        used = count_orders_in_window(business_id) if business_id else int(business.get("orders_count") or 0)

    remaining = max(0, limit - used) if active else 0

    return {
        "stored_tariff": stored_tariff,
        "active": active,
        "expired": expired,
        "is_paid": paid,
        "limit": limit,
        "used": used,
        "remaining": remaining,
        "started_at": format_datetime(started_at) if started_at else None,
        "expires_at": format_datetime(expires_at) if expires_at else None,
    }


def check_tariff_limit(business_id) -> bool:
    business = get_business_by_id(business_id)
    if not business:
        return False

    usage = get_tariff_usage(business)
    return usage["active"] and usage["used"] < usage["limit"]


def save_session(user_id, state, data=None, messages=None):
    conn = get_connection()
    cursor = conn.cursor()
    data_json = json.dumps(data or {}, ensure_ascii=False)
    msgs_json = json.dumps(messages or [], ensure_ascii=False)
    cursor.execute("""
        INSERT INTO sessions (user_id, state, data, messages)
        VALUES (?, ?, ?, ?)
        ON CONFLICT(user_id) DO UPDATE SET state=?, data=?, messages=?
    """, (user_id, state, data_json, msgs_json, state, data_json, msgs_json))
    conn.commit()
    conn.close()


def get_session(user_id):
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute("SELECT * FROM sessions WHERE user_id = ?", (user_id,))
    row = cursor.fetchone()
    conn.close()
    if row:
        result = dict(row)
        result['data'] = json.loads(result['data'] or '{}')
        result['messages'] = json.loads(result['messages'] or '[]')
        return result
    return None


def save_order(business_id, client_id, client_username, data):
    conn = get_connection()
    cursor = conn.cursor()
    collected = data.get('collected', {})
    urgency = collected.get('urgency', data.get('urgency', ''))
    is_hot = 1 if urgency and 'Сьогодні' in urgency else 0

    cursor.execute("""
        INSERT INTO orders (
            business_id, client_id, client_username,
            client_name, client_phone,
            service, problem, device, car,
            format, shipping_city,
            urgency, district,
            ai_comment, is_hot, created_at
        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
    """, (
        business_id, client_id, client_username,
        collected.get('name', data.get('client_name')),
        collected.get('phone', data.get('phone')),
        collected.get('service', data.get('service')),
        collected.get('problem', data.get('problem')),
        collected.get('device'),
        collected.get('car'),
        collected.get('format'),
        collected.get('shipping_city'),
        urgency,
        collected.get('district', data.get('district')),
        data.get('ai_comment'),
        is_hot,
        format_datetime(current_time())
    ))
    cursor.execute(
        "UPDATE businesses SET orders_count = orders_count + 1 WHERE id = ?",
        (business_id,)
    )
    conn.commit()
    conn.close()
    try:
        cancel_pending_link_reminders(business_id, client_id, client_username)
    except Exception:
        pass


def get_orders(business_id, status_filter=None):
    conn = get_connection()
    cursor = conn.cursor()
    if status_filter:
        cursor.execute("""
            SELECT * FROM orders WHERE business_id = ? AND status = ?
            ORDER BY created_at DESC
        """, (business_id, status_filter))
    else:
        cursor.execute("""
            SELECT * FROM orders WHERE business_id = ?
            ORDER BY created_at DESC
        """, (business_id,))
    rows = cursor.fetchall()
    conn.close()
    return [dict(row) for row in rows]


def get_recent_client_orders(client_id, username="", limit=10):
    conn = get_connection()
    cursor = conn.cursor()
    username = str(username or "").lstrip("@")

    if username:
        cursor.execute("""
            SELECT
                orders.*,
                businesses.name AS business_name,
                businesses.city AS business_city,
                businesses.niche AS business_niche,
                businesses.link_code AS business_link_code
            FROM orders
            LEFT JOIN businesses ON businesses.id = orders.business_id
            WHERE orders.client_id = ? OR lower(orders.client_username) = lower(?)
            ORDER BY orders.created_at DESC
            LIMIT ?
        """, (client_id, username, int(limit)))
    else:
        cursor.execute("""
            SELECT
                orders.*,
                businesses.name AS business_name,
                businesses.city AS business_city,
                businesses.niche AS business_niche,
                businesses.link_code AS business_link_code
            FROM orders
            LEFT JOIN businesses ON businesses.id = orders.business_id
            WHERE orders.client_id = ?
            ORDER BY orders.created_at DESC
            LIMIT ?
        """, (client_id, int(limit)))

    rows = cursor.fetchall()
    conn.close()
    return [dict(row) for row in rows]


def count_orders(business_id, status_filter=None):
    conn = get_connection()
    cursor = conn.cursor()
    if status_filter:
        cursor.execute(
            "SELECT COUNT(*) FROM orders WHERE business_id = ? AND status = ?",
            (business_id, status_filter)
        )
    else:
        cursor.execute(
            "SELECT COUNT(*) FROM orders WHERE business_id = ?",
            (business_id,)
        )
    total = int(cursor.fetchone()[0] or 0)
    conn.close()
    return total


def get_orders_page(business_id, page=1, per_page=10, status_filter=None):
    page = max(1, int(page or 1))
    per_page = min(25, max(1, int(per_page or 10)))
    offset = (page - 1) * per_page

    conn = get_connection()
    cursor = conn.cursor()
    if status_filter:
        cursor.execute("""
            SELECT * FROM orders
            WHERE business_id = ? AND status = ?
            ORDER BY created_at DESC
            LIMIT ? OFFSET ?
        """, (business_id, status_filter, per_page, offset))
    else:
        cursor.execute("""
            SELECT * FROM orders
            WHERE business_id = ?
            ORDER BY created_at DESC
            LIMIT ? OFFSET ?
        """, (business_id, per_page, offset))
    rows = cursor.fetchall()
    conn.close()
    return [dict(row) for row in rows]


def get_order_for_business(business_id, order_id):
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute("""
        SELECT * FROM orders
        WHERE business_id = ? AND id = ?
        LIMIT 1
    """, (business_id, order_id))
    row = cursor.fetchone()
    conn.close()
    return dict(row) if row else None


def update_order_status(order_id, status):
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute("UPDATE orders SET status = ? WHERE id = ?", (status, order_id))
    conn.commit()
    conn.close()


def get_stats(business_id):
    conn = get_connection()
    cursor = conn.cursor()
    now = current_time()
    today = now.strftime('%Y-%m-%d')
    month = now.strftime('%Y-%m')

    cursor.execute("SELECT COUNT(*) FROM orders WHERE business_id = ?", (business_id,))
    total = cursor.fetchone()[0]

    cursor.execute(
        "SELECT COUNT(*) FROM orders WHERE business_id = ? AND DATE(created_at) = ?",
        (business_id, today)
    )
    today_count = cursor.fetchone()[0]

    cursor.execute(
        "SELECT COUNT(*) FROM orders WHERE business_id = ? AND strftime('%Y-%m', created_at) = ?",
        (business_id, month)
    )
    month_count = cursor.fetchone()[0]

    cursor.execute(
        "SELECT COUNT(*) FROM orders WHERE business_id = ? AND is_hot = 1",
        (business_id,)
    )
    hot = cursor.fetchone()[0]

    cursor.execute(
        "SELECT COUNT(*) FROM orders WHERE business_id = ? AND status = 'success'",
        (business_id,)
    )
    success = cursor.fetchone()[0]

    cursor.execute(
        "SELECT COUNT(*) FROM orders WHERE business_id = ? AND status = 'new'",
        (business_id,)
    )
    new = cursor.fetchone()[0]

    cursor.execute("""
        SELECT service, COUNT(*) as cnt FROM orders
        WHERE business_id = ? AND service IS NOT NULL
        GROUP BY service ORDER BY cnt DESC LIMIT 1
    """, (business_id,))
    top_row = cursor.fetchone()
    top_service = top_row['service'] if top_row else '-'

    conn.close()
    return {
        "total": total,
        "today": today_count,
        "month": month_count,
        "hot": hot,
        "success": success,
        "new": new,
        "top_service": top_service,
    }
