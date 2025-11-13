"""
Telegram Real Estate Checkup Bot

Сценарий:
- /start → "Создать новый запрос"
- Адрес
- Кадастровый номер (или "нет")
- Кто отправляет запрос
- Загрузка документов (PDF/JPG/PNG) + кнопки "Пропустить" и "Готово"
- Комментарий
- Превью заявки + кнопки "Отправить эксперту" / "Изменить данные"
- Заявка сохраняется в БД и уходит админу (ADMIN_CHAT_ID)

Админ-команды:
- /report <дней> — Excel-отчёт по заявкам
- /whoami — показать свой chat_id (удобно для настройки админа)
"""

import logging
import os
import re
import uuid
from datetime import datetime
from pathlib import Path
import io
import html  # для экранирования текста

from aiogram import Bot, Dispatcher, types
from aiogram.contrib.fsm_storage.memory import MemoryStorage
from aiogram.dispatcher import FSMContext
from aiogram.dispatcher.filters.state import State, StatesGroup
from aiogram.types import ContentType, ParseMode
from aiogram.utils import executor

import aiosqlite
from openpyxl import Workbook
from openpyxl.utils import get_column_letter

# ============================================================
# НАСТРОЙКИ БОТА – УЖЕ ЗАПОЛНЕНЫ ПОД ТЕБЯ
# ============================================================

BOT_TOKEN = "8509916986:AAFuI5YcGsDgRm54n451VrQvKjpG548DULQ"
ADMIN_CHAT_ID = 924325909  # твой Telegram ID

UPLOAD_DIR = Path("./uploads")
DB_PATH = Path("./requests.db")
UPLOAD_DIR.mkdir(exist_ok=True)

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

bot = Bot(token=BOT_TOKEN)
storage = MemoryStorage()
dp = Dispatcher(bot, storage=storage)

# ============================================================
#    FSM STATES
# ============================================================

class CheckUpStates(StatesGroup):
    ADDRESS = State()
    CADASTRAL = State()
    WHO = State()
    WHO_OTHER = State()
    DOCS = State()
    COMMENT = State()
    CONFIRM = State()


# ============================================================
#    HELPERS
# ============================================================

CADASTRAL_RE = re.compile(r"^\d{1,3}:\d{1,3}:\d{1,10}:\d{1,10}$")


async def init_db():
    async with aiosqlite.connect(DB_PATH) as db:
        await db.execute(
            """
            CREATE TABLE IF NOT EXISTS requests (
                id TEXT PRIMARY KEY,
                user_id INTEGER,
                username TEXT,
                address TEXT,
                cadastral TEXT,
                who TEXT,
                comment TEXT,
                files TEXT,
                created_at TEXT
            )
            """
        )
        await db.commit()


def validate_address(text: str) -> bool:
    if not text:
        return False
    parts = text.strip().split()
    return len(parts) >= 2


def validate_cadastral(text: str) -> bool:
    if text.lower() in ("нет", "n", "no"):
        return True
    return bool(CADASTRAL_RE.match(text.strip()))


def kb_cancel_only() -> types.ReplyKeyboardMarkup:
    kb = types.ReplyKeyboardMarkup(resize_keyboard=True)
    kb.add("Отмена")
    return kb


def kb_docs() -> types.ReplyKeyboardMarkup:
    kb = types.ReplyKeyboardMarkup(resize_keyboard=True)
    kb.row("Загрузить документ", "Пропустить", "Готово")
    kb.add("Отмена")
    return kb


def kb_confirm() -> types.ReplyKeyboardMarkup:
    kb = types.ReplyKeyboardMarkup(resize_keyboard=True)
    kb.add("Отправить эксперту", "Изменить данные")
    kb.add("Отмена")
    return kb


def esc(value) -> str:
    """Безопасное экранирование для HTML."""
    if value is None:
        return "-"
    return html.escape(str(value), quote=False)


def fmt_request_message(data: dict) -> str:
    """Форматируем текст заявки для админа и превью."""
    lines = []
    lines.append("<b>Запрос на проверку объекта недвижимости</b>")
    lines.append("")
    lines.append(f"🏠 <b>Адрес:</b> {esc(data.get('address'))}")
    cadastral = data.get("cadastral") or "-"
    lines.append(f"📇 <b>Кадастровый номер:</b> {esc(cadastral)}")
    lines.append(f"👤 <b>Тип заявителя:</b> {esc(data.get('who'))}")
    files = data.get("files") or []
    files_list = "\n".join([f"- {esc(f)}" for f in files]) if files else "-"
    lines.append(f"📎 <b>Приложения:</b>\n{files_list}")
    comment = data.get("comment") or "-"
    lines.append(f"📝 <b>Комментарий:</b> {esc(comment)}")
    lines.append(
        f"\n📅 <b>Дата запроса:</b> {esc(data.get('created_at'))} (UTC)"
    )
    uname = data.get("username") or "-"
    lines.append(
        f"\n🆔 <b>User:</b> {esc(data.get('user_id'))} ({esc(uname)})"
    )
    lines.append(f"🔎 <b>ID заявки:</b> {esc(data.get('id'))}")
    return "\n".join(lines)


async def save_request_to_db(rec: dict):
    async with aiosqlite.connect(DB_PATH) as db:
        await db.execute(
            """
            INSERT INTO requests (
                id, user_id, username, address, cadastral,
                who, comment, files, created_at
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
            """,
            (
                rec["id"],
                rec["user_id"],
                rec["username"],
                rec["address"],
                rec["cadastral"],
                rec["who"],
                rec["comment"],
                "\n".join(rec["files"]),
                rec["created_at"],
            ),
        )
        await db.commit()


# ============================================================
#    БАЗОВЫЕ КОМАНДЫ
# ============================================================

@dp.message_handler(commands=["start", "help"], state="*")
async def cmd_start(message: types.Message, state: FSMContext):
    await state.finish()
    kb = types.ReplyKeyboardMarkup(resize_keyboard=True)
    kb.add("Создать новый запрос")
    kb.add("Отмена")
    await message.answer(
        "Привет! 👋 Я помогу подготовить запрос на проверку объекта недвижимости.\n\n"
        "Нажмите «Создать новый запрос», чтобы начать.",
        reply_markup=kb,
    )


@dp.message_handler(commands=["whoami"], state="*")
async def cmd_whoami(message: types.Message):
    await message.answer(
        f"Ваш chat_id: <code>{message.from_user.id}</code>",
        parse_mode=ParseMode.HTML,
    )


@dp.message_handler(lambda m: m.text == "Отмена", state="*")
async def cmd_cancel(message: types.Message, state: FSMContext):
    await state.finish()
    await message.answer(
        "Операция отменена. Если нужно — начните заново.",
        reply_markup=types.ReplyKeyboardRemove(),
    )


@dp.message_handler(lambda m: m.text == "Создать новый запрос", state="*")
async def start_request(message: types.Message, state: FSMContext):
    await CheckUpStates.ADDRESS.set()
    await message.answer(
        "Введите адрес объекта (улица, дом, город). "
        "Если нет — укажите ориентир или ссылку на карту.",
        reply_markup=kb_cancel_only(),
    )


# ============================================================
#    АДРЕС
# ============================================================

@dp.message_handler(state=CheckUpStates.ADDRESS, content_types=ContentType.TEXT)
async def process_address(message: types.Message, state: FSMContext):
    text = message.text.strip()
    if text.lower() == "отмена":
        return await cmd_cancel(message, state)
    if not validate_address(text):
        await message.reply(
            "Пожалуйста, укажите более точный адрес (минимум улица + дом или город)."
        )
        return
    await state.update_data(address=text)
    await CheckUpStates.CADASTRAL.set()
    await message.answer(
        'Укажите кадастровый номер (пример: 77:01:0004010:1234) или напишите "нет".',
        reply_markup=kb_cancel_only(),
    )


# ============================================================
#    КАДАСТРОВЫЙ НОМЕР
# ============================================================

@dp.message_handler(state=CheckUpStates.CADASTRAL, content_types=ContentType.TEXT)
async def process_cadastral(message: types.Message, state: FSMContext):
    text = message.text.strip()
    if text.lower() == "отмена":
        return await cmd_cancel(message, state)
    if not validate_cadastral(text):
        await message.reply(
            "Неправильный формат кадастрового номера. "
            "Введите в формате 77:01:0004010:1234 или напишите \"нет\"."
        )
        return
    await state.update_data(cadastral=text)
    await CheckUpStates.WHO.set()
    kb = types.ReplyKeyboardMarkup(resize_keyboard=True)
    kb.add("Агент", "Владелец")
    kb.add("Другое", "Отмена")
    await message.answer("Кто отправляет запрос?", reply_markup=kb)


# ============================================================
#    КТО ЗАЯВИТЕЛЬ
# ============================================================

@dp.message_handler(state=Check
