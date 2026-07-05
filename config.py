import os
from pathlib import Path
from dotenv import load_dotenv

load_dotenv(Path(__file__).resolve().parent / ".env")

BOT_TOKEN = os.getenv("BOT_TOKEN")
OPENAI_API_KEY = os.getenv("OPENAI_API_KEY")

TARIFF_LIMITS = {
    "free":     5,
    "start":    50,
    "business": 300,
    "vip":      1000,
}

TARIFF_PRICES = {
    "free":     "Пробний тариф - до 5 заявок",
    "start":    "2990 грн/міс - до 50 заявок",
    "business": "5990 грн/міс - до 300 заявок",
    "vip":      "12990 грн/міс - до 1000 заявок",
}
