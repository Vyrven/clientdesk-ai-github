TECHNICAL_NICHE = "Ремонт техніки"


NICHE_FIELDS = {
    "Ремонт техніки": [
        {"key": "device", "question": "⚙️ Який пристрій потребує ремонту?\n\nНаприклад: iPhone 12, ноутбук Acer"},
        {"key": "problem", "question": "❗ Що сталось? Опишіть проблему:"},
        {"key": "urgency", "question": "⏰ Наскільки терміново?", "type": "urgency"},
        {"key": "name", "question": "👤 Як до вас звертатись?"},
        {"key": "phone", "question": "📞 Залиште номер телефону для зв'язку:"},
    ],
    "Автосервіс": [
        {"key": "car", "question": "🚗 Яка марка і модель автомобіля?"},
        {"key": "problem", "question": "🔧 Що потрібно зробити або яка проблема?"},
        {"key": "urgency", "question": "⏰ Наскільки терміново?", "type": "urgency"},
        {"key": "district", "question": "📍 В якому місті або районі вам зручно звернутись до сервісу?"},
        {"key": "name", "question": "👤 Як до вас звертатись?"},
        {"key": "phone", "question": "📞 Залиште номер телефону:"},
    ],
}


def get_fields(niche: str):
    return NICHE_FIELDS.get(niche, NICHE_FIELDS[TECHNICAL_NICHE])


def get_next_question(niche: str, collected: dict):
    for field in get_fields(niche):
        if field["key"] not in collected:
            return field
    return None


def get_progress(niche: str, collected: dict):
    fields = get_fields(niche)
    done = sum(1 for f in fields if f["key"] in collected)
    return done, len(fields)
