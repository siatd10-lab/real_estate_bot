
# Telegram Real Estate Checkup Bot

Бот собирает данные по объекту недвижимости для чекапа и отправляет эксперту структурированную заявку.
Готов к деплою на Render как **Background Worker** (Blueprint).

## Быстрый старт (локально)

```bash
python -m venv venv
source venv/bin/activate  # Windows: venv\Scripts\activate
pip install -r requirements.txt
cp .env.example .env
# отредактируйте .env (BOT_TOKEN, ADMIN_CHAT_ID)
python telegram_real_estate_bot.py
```

## One‑click Deploy на Render (Blueprint)

1. Загрузите этот проект в публичный GitHub-репозиторий.
2. Перейдите по ссылке (замените `<YOUR_REPO_URL>` на URL вашего репозитория):  
   `https://render.com/deploy?repo=<YOUR_REPO_URL>`
3. Render сам подхватит `render.yaml`. Выберите **Free** план и нажмите **Deploy**.
4. На странице сервиса откройте **Environment** → добавьте `BOT_TOKEN` и `ADMIN_CHAT_ID`.
5. Дождитесь статуса **Running** — бот готов.

## Команда отчёта

В чате администратора:
- `/report` — отчёт за 7 дней
- `/report 30` — отчёт за 30 дней

## Структура

- `telegram_real_estate_bot.py` — основной код (FSM, валидации, загрузка файлов, SQLite, отчёты)
- `requirements.txt` — зависимости
- `.env.example` — переменные окружения
- `render.yaml` — конфигурация для Render Blueprint
- `uploads/` — каталог для файлов пользователей
