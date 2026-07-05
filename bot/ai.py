import json
import re
import openai

from config import OPENAI_API_KEY

openai.api_key = OPENAI_API_KEY


def _extract_json(text: str) -> dict:
    if not text:
        return {}

    text = text.strip()

    if text.startswith("```"):
        text = re.sub(r"^```json", "", text, flags=re.IGNORECASE).strip()
        text = re.sub(r"^```", "", text).strip()
        text = re.sub(r"```$", "", text).strip()

    try:
        return json.loads(text)
    except Exception:
        pass

    match = re.search(r"\{.*\}", text, re.DOTALL)
    if match:
        try:
            return json.loads(match.group(0))
        except Exception:
            return {}

    return {}


async def analyze_client_request(
    business: dict,
    collected: dict,
    user_text: str,
    required_fields: list,
    current_field: str = None
) -> dict:
    try:
        biz_name = business.get("name", "")
        niche = business.get("niche", "")
        city = business.get("city", "")
        services = business.get("services", "")
        knowledge = business.get("knowledge", "")

        valid_keys = [field.get("key") for field in required_fields if field.get("key")]

        fields_for_prompt = []
        for field in required_fields:
            fields_for_prompt.append({
                "key": field.get("key"),
                "question": field.get("question"),
                "type": field.get("type", "text")
            })

        system_prompt = (
            "Ти AI-адміністратор для малого бізнесу в Telegram.\n"
            "Твоя задача - допомогти клієнту швидко оформити заявку.\n\n"
            "Ти працюєш не як вільний чат, а як розумний помічник для заявки.\n"
            "Потрібно або витягнути дані з повідомлення, або допомогти клієнту правильно відповісти.\n\n"
            "ВАЖЛИВО:\n"
            "- Не вигадуй телефон, ім'я, адресу, дату, ціну або послугу.\n"
            "- Якщо поля немає в повідомленні - не заповнюй його.\n"
            "- Витягуй тільки ті поля, які є у valid_keys.\n"
            "- Не додавай нові ключі.\n"
            "- Якщо клієнт просить допомогти сформулювати проблему, відповідь або деталі - intent = assistant_help.\n"
            "- Якщо клієнт написав не те, що потрібно для поточного поля - intent = invalid_input.\n"
            "- Якщо клієнт дав корисні дані - intent = leave_request.\n"
            "- Якщо клієнт питає ціну, але точної ціни немає в базі знань - не вигадуй ціну.\n"
            "- Відповідай українською мовою, навіть якщо клієнт пише російською.\n"
            "- Поверни тільки валідний JSON без markdown.\n\n"
            "Як поводитись у режимі assistant_help:\n"
            "- Якщо клієнт просить допомогти описати проблему, але не дав симптомів, попроси простими словами написати, що саме не працює.\n"
            "- Якщо клієнт дав симптоми, сформулюй їх професійно і можеш покласти у відповідне поле.\n"
            "- Не повторюй одне й те саме питання тупо. Допоможи клієнту.\n\n"
            "Формат відповіді:\n"
            "{\n"
            '  "ok": true,\n'
            '  "intent": "leave_request | price_question | schedule_question | address_question | human_request | assistant_help | invalid_input | unknown",\n'
            '  "extracted": {},\n'
            '  "missing_fields": [],\n'
            '  "should_finish_order": false,\n'
            '  "next_question": "",\n'
            '  "reply": ""\n'
            "}\n"
        )

        user_prompt = {
            "business": {
                "name": biz_name,
                "niche": niche,
                "city": city,
                "services": services,
                "knowledge": knowledge
            },
            "required_fields": fields_for_prompt,
            "valid_keys": valid_keys,
            "already_collected": collected or {},
            "current_field": current_field,
            "client_message": user_text
        }

        response = openai.chat.completions.create(
            model="gpt-4o-mini",
            temperature=0.2,
            max_tokens=600,
            response_format={"type": "json_object"},
            messages=[
                {"role": "system", "content": system_prompt},
                {"role": "user", "content": json.dumps(user_prompt, ensure_ascii=False)}
            ]
        )

        raw = response.choices[0].message.content
        result = _extract_json(raw)

        if not result:
            return {
                "ok": False, "intent": "unknown",
                "extracted": {}, "missing_fields": [],
                "should_finish_order": False, "next_question": "", "reply": ""
            }

        extracted = result.get("extracted") or {}
        safe_extracted = {
            k: str(v).strip()
            for k, v in extracted.items()
            if k in valid_keys and v is not None and str(v).strip()
        }

        missing_fields = result.get("missing_fields") or []
        safe_missing = [key for key in missing_fields if key in valid_keys]

        return {
            "ok": bool(result.get("ok", True)),
            "intent": result.get("intent", "leave_request"),
            "extracted": safe_extracted,
            "missing_fields": safe_missing,
            "should_finish_order": bool(result.get("should_finish_order", False)),
            "next_question": result.get("next_question", ""),
            "reply": result.get("reply", "")
        }

    except Exception as e:
        print(f"OpenAI analyze error: {e}")
        return {
            "ok": False, "intent": "unknown",
            "extracted": {}, "missing_fields": [],
            "should_finish_order": False, "next_question": "", "reply": ""
        }


async def qualify_order(business: dict, order_data: dict) -> str:
    try:
        field_names = {
            "name": "Ім'я", "phone": "Телефон",
            "device": "Пристрій", "problem": "Проблема",
            "car": "Авто", "service": "Послуга",
            "format": "Формат", "shipping_city": "Населений пункт відправки",
            "urgency": "Терміновість", "district": "Район",
        }
        order_lines = [
            f"{label}: {order_data.get(key)}"
            for key, label in field_names.items()
            if order_data.get(key)
        ]
        summary = "\n".join(order_lines)

        response = openai.chat.completions.create(
            model="gpt-4o-mini",
            temperature=0.3,
            max_tokens=120,
            messages=[
                {
                    "role": "system",
                    "content": (
                        f"Ти AI-помічник для бізнесу '{business.get('name')}' "
                        f"(ніша: {business.get('niche', '')}).\n"
                        "Проаналізуй заявку і дай коротку оцінку у 1-2 речення:\n"
                        "- наскільки клієнт гарячий\n"
                        "- що рекомендуєш менеджеру\n"
                        "Відповідай коротко українською мовою."
                    )
                },
                {"role": "user", "content": f"Заявка:\n{summary}"}
            ]
        )
        return response.choices[0].message.content.strip()

    except Exception as e:
        print(f"OpenAI qualify error: {e}")
        urgency = order_data.get("urgency", "")
        if urgency and "Сьогодні" in urgency:
            return "🔥 Гарячий клієнт - зв'яжіться терміново!"
        return "Звичайна заявка - зв'яжіться найближчим часом."


async def answer_business_question(business: dict, question: str) -> dict:
    try:
        biz_name = business.get("name", "") or "бізнес"
        niche = business.get("niche", "") or ""
        city = business.get("city", "") or ""
        services = business.get("services", "") or ""
        schedule = business.get("schedule", "") or ""
        prices = business.get("prices", "") or ""
        knowledge = business.get("knowledge", "") or ""

        base_text = "\n".join([
            f"Назва бізнесу: {biz_name}",
            f"Місто: {city}",
            f"Ніша: {niche}",
            f"Послуги: {services}",
            f"Графік: {schedule}",
            f"Ціни: {prices}",
            f"AI-база знань:\n{knowledge}",
        ]).strip()

        if not (services or schedule or prices or knowledge or city):
            return {
                "found": False,
                "answer": (
                    "ℹ️ <b>У AI-базі знань цього бізнесу поки майже немає інформації.</b>\n\n"
                    "Я не хочу вигадувати відповідь. Краще залиште заявку - менеджер уточнить деталі і відповість напряму."
                )
            }

        system_prompt = (
            "Ти AI-адміністратор малого бізнесу в Telegram.\n"
            "Відповідай клієнту професійно, коротко і зрозуміло українською мовою.\n\n"
            "КРИТИЧНО:\n"
            "- Використовуй тільки інформацію з бази знань бізнесу.\n"
            "- Не вигадуй ціни, адреси, гарантії, терміни, знижки або умови.\n"
            "- Якщо відповіді немає у базі знань, чесно скажи, що точної інформації поки немає.\n"
            "- Якщо питання схоже на заявку, запропонуй натиснути 'Залишити заявку'.\n"
            "- Якщо інформація є, дай відповідь як нормальний адміністратор бізнесу, не як робот.\n"
            "- Не пиши markdown. Можна використовувати HTML: <b>...</b>.\n"
            "- Не використовуй багато emoji. Максимум 1-2.\n\n"
            "Поверни тільки JSON:\n"
            "{\n"
            '  "found": true/false,\n'
            '  "answer": "готова відповідь клієнту українською з HTML за потреби"\n'
            "}"
        )

        user_payload = {
            "business_base": base_text,
            "client_question": question
        }

        response = openai.chat.completions.create(
            model="gpt-4o-mini",
            temperature=0.15,
            max_tokens=500,
            response_format={"type": "json_object"},
            messages=[
                {"role": "system", "content": system_prompt},
                {"role": "user", "content": json.dumps(user_payload, ensure_ascii=False)}
            ]
        )

        raw = response.choices[0].message.content
        result = _extract_json(raw)

        answer = (result.get("answer") or "").strip()
        found = bool(result.get("found", False))

        if not answer:
            answer = (
                "ℹ️ <b>У базі знань поки немає точної відповіді на це питання.</b>\n\n"
                "Можете залишити заявку - менеджер побачить звернення і уточнить деталі."
            )
            found = False

        return {"found": found, "answer": answer}

    except Exception as e:
        print(f"OpenAI business question error: {e}")
        return {
            "found": False,
            "answer": (
                "ℹ️ <b>Не вдалося отримати відповідь з AI-бази знань.</b>\n\n"
                "Можете залишити заявку - менеджер побачить звернення і відповість напряму."
            )
        }


async def parse_knowledge_update(current_knowledge: str, user_message: str) -> str:
    """
    Отримує поточну базу знань і вільне повідомлення власника.
    Оновлює тільки ті секції, які згадані у повідомленні.
    Повертає повний оновлений текст бази знань.
    """
    try:
        system = (
            "Ти оновлюєш AI-базу знань малого бізнесу для Telegram-бота.\n"
            "База знань — це текст, який AI використовує у відповідях клієнтам.\n\n"
            "Можливі секції бази знань:\n"
            "Графік роботи:\n"
            "Ціни:\n"
            "Адреса/філіали:\n"
            "Умови роботи:\n"
            "Гарантія:\n"
            "Часті питання:\n"
            "Що не обіцяти клієнтам:\n\n"
            "Правила:\n"
            "1. Оновлюй ТІЛЬКИ ті секції, які явно згадані у повідомленні власника.\n"
            "2. Решту секцій залишай точно такими, як вони є у поточній базі.\n"
            "3. Якщо секції ще немає — додай її з правильним заголовком.\n"
            "4. Якщо власник просить видалити секцію — видали лише її.\n"
            "5. Поверни ТІЛЬКИ повний текст бази знань без пояснень, без markdown.\n"
            "6. Формат: кожна секція починається з заголовка і двокрапки, потім текст.\n"
            "7. Між секціями — один порожній рядок."
        )

        response = openai.chat.completions.create(
            model="gpt-4o-mini",
            max_tokens=900,
            temperature=0.1,
            messages=[
                {"role": "system", "content": system},
                {
                    "role": "user",
                    "content": (
                        f"Поточна база знань:\n"
                        f"{current_knowledge.strip() if current_knowledge else '(порожня)'}\n\n"
                        f"Повідомлення власника:\n{user_message}"
                    )
                }
            ]
        )

        return response.choices[0].message.content.strip()

    except Exception as e:
        print(f"parse_knowledge_update error: {e}")
        if current_knowledge and current_knowledge.strip():
            return f"{current_knowledge.strip()}\n\n{user_message}"
        return user_message
