# ClientDeskAI

**English** · [Українська](#clientdeskai-українська)

A Telegram bot that turns free-form customer messages into structured business leads, using an LLM as the extraction and reasoning layer rather than a rigid form.

Built for small service businesses (device repair shops, auto services): a business owner registers via the bot, gets a personal client link, and every incoming message is parsed, validated and routed into an order pipeline — no web app, no forms, just chat.

## Why this project is interesting

The core problem isn't "wrap an OpenAI call in a Telegram handler" — it's making an LLM behave reliably inside a stateful, multi-turn conversation where wrong output has real consequences (a fabricated phone number or price is worse than no answer). The design choices in [`bot/ai.py`](bot/ai.py) and [`bot/handlers.py`](bot/handlers.py) reflect that:

- **Structured output, not free text.** Every LLM call uses `response_format={"type": "json_object"}` with an explicit schema (`intent`, `extracted`, `missing_fields`, `should_finish_order`, `reply`), so the bot's control flow branches on a typed response instead of parsing prose.
- **Intent classification drives the conversation.** A single call classifies the message as `leave_request | price_question | assistant_help | invalid_input | human_request | ...`, letting one model call replace what would otherwise be several branching handlers.
- **Anti-hallucination guardrails baked into the prompt.** The system prompt explicitly forbids inventing phone numbers, prices, or addresses, and instructs the model to only extract fields present in `valid_keys` — the model is a field-extractor first, a conversationalist second.
- **Knowledge grounding (RAG-lite).** Business-specific Q&A (`answer_business_question`) only answers from the business's own knowledge base text and returns `found: false` rather than guessing when the answer isn't there.
- **Hybrid extraction pipeline.** Before spending an API call, [`extract_initial_order_fields`](bot/handlers.py) runs fast regex/heuristic extraction for phone numbers, urgency, and city — the LLM is only invoked for the harder, ambiguous parts of the message. This cuts cost and latency on the common case.
- **Fails safe.** Every OpenAI call is wrapped so that a timeout or malformed response degrades to a safe default (e.g. a rule-based urgency fallback) instead of crashing the conversation.

## Features

- **Business onboarding** — pick a niche (device repair, auto service), work mode (offline/online), get a shareable client link.
- **AI-assisted intake** — clients write naturally; the bot extracts structured fields and asks only for what's missing.
- **Admin panel** — inline-keyboard driven order list, status updates, per-business stats, tariff management.
- **Tariff plans** — monthly request limits (Free / Start / Business / VIP).
- **Background reminders** — a worker notifies business owners about orders needing attention.
- **Built-in help** — separate `/help` flows for business owners and clients.

## Tech stack

- Python, [python-telegram-bot](https://github.com/python-telegram-bot/python-telegram-bot) (async, polling)
- OpenAI API (`gpt-4o-mini`, JSON mode) for extraction, Q&A grounding, and knowledge-base editing
- SQLite for businesses, orders and sessions
- Deployed on [Railway](https://railway.app)

## Project structure

```
.
├── main.py           # entry point, handler registration
├── config.py         # tariffs, pricing, env-based config
├── railway.json       # Railway deploy config
└── bot/
    ├── admin.py       # business admin panel
    ├── ai.py          # OpenAI integration: extraction, Q&A, knowledge updates
    ├── db.py          # SQLite access layer
    ├── handlers.py    # conversation logic + hybrid rule-based extraction
    ├── help.py        # /help flows
    ├── niches.py       # per-niche required order fields
    └── notify.py       # reminders / notifications
```

## Running locally

1. Install dependencies:
   ```bash
   pip install -r requirements.txt
   ```
2. Create a `.env` file based on `.env.example`:
   ```
   BOT_TOKEN=your_botfather_token
   OPENAI_API_KEY=your_openai_key
   ADMIN_IDS=comma_separated_telegram_ids
   SUPPORT_USERNAME=support_username
   ```
3. Run the bot:
   ```bash
   python main.py
   ```

## Deployment

Configured for Railway (`railway.json`, start command `python main.py`, auto-restart on failure).

## License

MIT

---

# ClientDeskAI (Українська)

[English](#clientdeskai) · **Українська**

Telegram-бот, який перетворює довільні повідомлення клієнтів на структуровані заявки, використовуючи LLM як шар витягування та аналізу даних замість жорсткої форми.

Створений для малого сервісного бізнесу (ремонт техніки, автосервіси): власник реєструється через бота, отримує персональне посилання для клієнтів, і кожне вхідне повідомлення розбирається, валідується та потрапляє в конвеєр заявок — без веб-застосунку, без форм, лише чат.

## Чим цей проєкт цікавий

Основна складність не в тому, щоб «обгорнути виклик OpenAI у Telegram-хендлер», а в тому, щоб змусити LLM поводитися надійно в багатокроковому діалозі зі станом, де хибна відповідь має реальні наслідки (вигаданий номер телефону чи ціна — гірше за відсутність відповіді). Архітектурні рішення в [`bot/ai.py`](bot/ai.py) та [`bot/handlers.py`](bot/handlers.py) це відображають:

- **Структурований вивід, а не вільний текст.** Кожен виклик LLM використовує `response_format={"type": "json_object"}` з чіткою схемою (`intent`, `extracted`, `missing_fields`, `should_finish_order`, `reply`), тож логіка бота розгалужується на основі типізованої відповіді, а не парсингу тексту.
- **Класифікація наміру керує діалогом.** Один виклик класифікує повідомлення як `leave_request | price_question | assistant_help | invalid_input | human_request | ...`, замінюючи собою кілька окремих хендлерів.
- **Захист від галюцинацій у промпті.** Системний промпт прямо забороняє вигадувати телефони, ціни чи адреси й дозволяє витягувати лише поля з `valid_keys` — модель насамперед екстрактор полів, а вже потім співрозмовник.
- **Заземлення на базу знань (RAG-lite).** Відповіді на питання про бізнес (`answer_business_question`) базуються лише на тексті бази знань самого бізнесу й повертають `found: false` замість здогадок, коли відповіді немає.
- **Гібридний конвеєр витягування.** Перед витратою виклику API [`extract_initial_order_fields`](bot/handlers.py) робить швидке regex/евристичне витягування телефону, терміновості та міста — LLM викликається лише для складніших, неоднозначних частин. Це зменшує вартість і затримку в типовому випадку.
- **Безпечні відмови.** Кожен виклик OpenAI обгорнутий так, що таймаут чи некоректна відповідь деградує до безпечного значення за замовчуванням (напр. правило-базований fallback терміновості), а не ламає діалог.

## Можливості

- **Онбординг бізнесу** — вибір ніші (ремонт техніки, автосервіс), формату роботи (офлайн/онлайн), генерація посилання для клієнтів.
- **AI-прийом заявок** — клієнт пише природною мовою; бот витягує структуровані поля й запитує лише те, чого бракує.
- **Адмін-панель** — список заявок на inline-кнопках, зміна статусів, статистика по бізнесах, керування тарифами.
- **Тарифні плани** — місячні ліміти заявок (Free / Start / Business / VIP).
- **Фонові нагадування** — воркер сповіщає власників про заявки, що потребують уваги.
- **Вбудована довідка** — окремі `/help`-сценарії для власників бізнесу та для клієнтів.

## Технології

- Python, [python-telegram-bot](https://github.com/python-telegram-bot/python-telegram-bot) (async, polling)
- OpenAI API (`gpt-4o-mini`, JSON-режим) для витягування, відповідей на питання та редагування бази знань
- SQLite для бізнесів, заявок і сесій
- Розгортання на [Railway](https://railway.app)

## Структура проєкту

```
.
├── main.py           # точка входу, реєстрація хендлерів
├── config.py         # тарифи, ціни, конфіг з .env
├── railway.json       # конфіг деплою на Railway
└── bot/
    ├── admin.py       # адмін-панель бізнесу
    ├── ai.py          # інтеграція з OpenAI: витягування, Q&A, оновлення бази знань
    ├── db.py          # шар доступу до SQLite
    ├── handlers.py    # логіка діалогів + гібридне правило-базоване витягування
    ├── help.py        # /help-сценарії
    ├── niches.py       # обов'язкові поля заявки по нішах
    └── notify.py       # нагадування / сповіщення
```

## Запуск локально

1. Встановити залежності:
   ```bash
   pip install -r requirements.txt
   ```
2. Створити файл `.env` за зразком `.env.example`:
   ```
   BOT_TOKEN=токен_з_BotFather
   OPENAI_API_KEY=ключ_OpenAI
   ADMIN_IDS=telegram_id_адмінів_через_кому
   SUPPORT_USERNAME=юзернейм_підтримки
   ```
3. Запустити бота:
   ```bash
   python main.py
   ```

## Розгортання

Налаштовано для Railway (`railway.json`, стартова команда `python main.py`, авто-рестарт при збої).

## Ліцензія

MIT
