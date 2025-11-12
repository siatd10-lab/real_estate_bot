
"""
Telegram Real Estate Checkup Bot
File: telegram_real_estate_bot.py

Requirements:
  pip install aiogram==2.24 python-dotenv aiosqlite openpyxl

Configuration (.env in project root):
  BOT_TOKEN=<your_bot_token>
  ADMIN_CHAT_ID=<expert_or_crm_incoming_chat_id>
"""
import asyncio
import logging
import os
import re
import uuid
from datetime import datetime
from pathlib import Path

from aiogram import Bot, Dispatcher, types
from aiogram.contrib.fsm_storage.memory import MemoryStorage
from aiogram.dispatcher import FSMContext
from aiogram.dispatcher.filters.state import State, StatesGroup
from aiogram.types import ContentType, ParseMode
from aiogram.utils import executor
from dotenv import load_dotenv
import aiosqlite

# -------------------- Config & Logging --------------------
load_dotenv()
BOT_TOKEN = os.getenv('BOT_TOKEN')
ADMIN_CHAT_ID = int(os.getenv('ADMIN_CHAT_ID', '0'))
UPLOAD_DIR = Path('./uploads')
DB_PATH = Path('./requests.db')
UPLOAD_DIR.mkdir(exist_ok=True)

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

# -------------------- Bot & Storage --------------------
bot = Bot(token=BOT_TOKEN)
storage = MemoryStorage()
dp = Dispatcher(bot, storage=storage)

# -------------------- FSM States --------------------
class CheckUpStates(StatesGroup):
    ADDRESS = State()
    CADASTRAL = State()
    WHO = State()
    WHO_OTHER = State()
    DOCS = State()
    COMMENT = State()
    CONFIRM = State()

# -------------------- Helpers --------------------
CADASTRAL_RE = re.compile(r'^\d{1,3}:\d{1,3}:\d{1,10}:\d{1,10}$')

async def init_db():
    async with aiosqlite.connect(DB_PATH) as db:
        await db.execute('''
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
        ''')
        await db.commit()

def validate_address(text: str) -> bool:
    if not text:
        return False
    parts = text.strip().split()
    return len(parts) >= 2

def validate_cadastral(text: str) -> bool:
    if text.lower() in ('нет', 'n', 'no'):
        return True
    return bool(CADASTRAL_RE.match(text.strip()))

def fmt_request_message(data: dict) -> str:
    lines = []
    lines.append('<b>Запрос на проверку объекта недвижимости</b>')
    lines.append('\n')
    lines.append(f"🏠 <b>Адрес:</b> {types.utils.escape_html(data.get('address','-'))}")
    cadastral = data.get('cadastral') or '-'
    lines.append(f"📇 <b>Кадастровый номер:</b> {types.utils.escape_html(cadastral)}")
    lines.append(f"👤 <b>Тип заявителя:</b> {types.utils.escape_html(data.get('who','-'))}")
    files = data.get('files') or []
    files_list = '\n'.join([f'- {f}' for f in files]) if files else '-'
    lines.append(f"📎 <b>Приложения:</b>\n{files_list}")
    comment = data.get('comment') or '-'
    lines.append(f"📝 <b>Комментарий:</b> {types.utils.escape_html(comment)}")
    lines.append(f"\n📅 <b>Дата запроса:</b> {data.get('created_at')}")
    lines.append('\n')
    uname = data.get('username') or '-'
    lines.append(f"🆔 <b>User:</b> {data.get('user_id')} ({types.utils.escape_html(uname)})")
    lines.append(f"🔎 <b>ID заявки:</b> {data.get('id')}")
    return "\n".join(lines)

async def save_request_to_db(rec: dict):
    async with aiosqlite.connect(DB_PATH) as db:
        await db.execute(
            'INSERT INTO requests (id, user_id, username, address, cadastral, who, comment, files, created_at) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)',
            (
                rec['id'], rec['user_id'], rec['username'], rec['address'], rec['cadastral'],
                rec['who'], rec['comment'], '\n'.join(rec['files']), rec['created_at']
            )
        )
        await db.commit()

# -------------------- Command Handlers --------------------
@dp.message_handler(commands=['start', 'help'])
async def cmd_start(message: types.Message):
    kb = types.ReplyKeyboardMarkup(resize_keyboard=True)
    kb.add('Создать новый запрос')
    kb.add('Отмена')
    await message.answer(
        'Привет! 👋 Я помогу подготовить запрос на проверку объекта недвижимости.\n\n'
        'Нажмите «Создать новый запрос», чтобы начать.', reply_markup=kb
    )

@dp.message_handler(lambda m: m.text == 'Отмена')
async def cmd_cancel(message: types.Message, state: FSMContext):
    await state.finish()
    kb = types.ReplyKeyboardRemove()
    await message.answer('Операция отменена. Если нужно — начните заново.', reply_markup=kb)

@dp.message_handler(lambda m: m.text == 'Создать новый запрос')
async def start_request(message: types.Message):
    await CheckUpStates.ADDRESS.set()
    kb = types.ReplyKeyboardMarkup(resize_keyboard=True)
    kb.add('Отмена')
    await message.answer('Введите адрес объекта (улица, дом, город). Если нет — укажите ориентир или ссылку на карту.', reply_markup=kb)

# Address
@dp.message_handler(state=CheckUpStates.ADDRESS, content_types=ContentType.TEXT)
async def process_address(message: types.Message, state: FSMContext):
    text = message.text.strip()
    if text.lower() == 'отмена':
        return await cmd_cancel(message, state)
    if not validate_address(text):
        await message.reply('Пожалуйста, укажите более точный адрес (минимум улица + дом или город).')
        return
    await state.update_data(address=text)
    await CheckUpStates.CADASTRAL.set()
    await message.answer('Укажите кадастровый номер (пример: 77:01:0004010:1234) или напишите "нет".')

# Cadastral
@dp.message_handler(state=CheckUpStates.CADASTRAL, content_types=ContentType.TEXT)
async def process_cadastral(message: types.Message, state: FSMContext):
    text = message.text.strip()
    if text.lower() == 'отмена':
        return await cmd_cancel(message, state)
    if not validate_cadastral(text):
        await message.reply('Неправильный формат кадастрового номера. Проверьте и введите в формате: 77:01:0004010:1234 или напишите "нет".')
        return
    await state.update_data(cadastral=text)
    await CheckUpStates.WHO.set()
    kb = types.ReplyKeyboardMarkup(resize_keyboard=True)
    kb.add('Агент', 'Владелец')
    kb.add('Другое', 'Отмена')
    await message.answer('Кто отправляет запрос?', reply_markup=kb)

# Who
@dp.message_handler(state=CheckUpStates.WHO, content_types=ContentType.TEXT)
async def process_who(message: types.Message, state: FSMContext):
    text = message.text.strip()
    if text.lower() == 'отмена':
        return await cmd_cancel(message, state)
    if text == 'Другое':
        await CheckUpStates.WHO_OTHER.set()
        await message.answer('Напишите, пожалуйста, кем вы являетесь (например, "юрист покупателя").')
        return
    if text not in ('Агент', 'Владелец'):
        await message.reply('Выберите один из вариантов или "Другое".')
        return
    await state.update_data(who=text)
    await CheckUpStates.DOCS.set()
    kb = types.ReplyKeyboardMarkup(resize_keyboard=True)
    kb.add('Загрузить документ', 'Пропустить')
    kb.add('Отмена')
    await message.answer('Прикрепите документы (PDF, JPG, PNG) или нажмите "Пропустить".', reply_markup=kb)

@dp.message_handler(state=CheckUpStates.WHO_OTHER, content_types=ContentType.TEXT)
async def process_who_other(message: types.Message, state: FSMContext):
    text = message.text.strip()
    if text.lower() == 'отмена':
        return await cmd_cancel(message, state)
    await state.update_data(who=text)
    await CheckUpStates.DOCS.set()
    kb = types.ReplyKeyboardMarkup(resize_keyboard=True)
    kb.row('Загрузить документ', 'Пропустить', 'Готово')
    kb.add('Отмена')
    await message.answer('Прикрепите документы (PDF, JPG, PNG) или нажмите "Пропустить".', reply_markup=kb)

# Files
ALLOWED_DOC_TYPES = ('application/pdf', 'image/jpeg', 'image/png')
MAX_FILE_SIZE = 20 * 1024 * 1024  # 20 MB

@dp.message_handler(state=CheckUpStates.DOCS, content_types=[ContentType.DOCUMENT, ContentType.PHOTO, ContentType.TEXT])
async def process_docs(message: types.Message, state: FSMContext):
    if message.content_type == ContentType.TEXT:
        txt = message.text.strip()
        if txt == 'Пропустить':
            await state.update_data(files=[])
            await CheckUpStates.COMMENT.set()
            await message.answer('Оставьте комментарий к запросу (или напишите "нет").', reply_markup=types.ReplyKeyboardMarkup(resize_keyboard=True).add('Отмена'))
            return
        if txt == 'Загрузить документ':
            await message.answer('Пришлите файл (PDF/JPG/PNG), либо несколько файлов по очереди. Когда закончите — отправьте "Готово".')
            return
        if txt == 'Готово':
            data = await state.get_data()
            files = data.get('files', []) or []
            await state.update_data(files=files)
            await CheckUpStates.COMMENT.set()
            await message.answer('Оставьте комментарий к запросу (или напишите "нет").')
            return
        if txt.lower() == 'отмена':
            return await cmd_cancel(message, state)
        await message.reply('Чтобы пропустить — нажмите "Пропустить". Для загрузки файлов — отправьте их, затем "Готово".')
        return

    file_obj = None
    filename = None
    file_size = None
    mime_type = None
    if message.content_type == ContentType.DOCUMENT:
        doc = message.document
        file_size = doc.file_size or 0
        mime_type = doc.mime_type
        filename = doc.file_name or f'doc_{uuid.uuid4()}.pdf'
        file_obj = await bot.get_file(doc.file_id)
    elif message.content_type == ContentType.PHOTO:
        photo = message.photo[-1]
        file_size = photo.file_size or 0
        mime_type = 'image/jpeg'
        filename = f'photo_{uuid.uuid4()}.jpg'
        file_obj = await bot.get_file(photo.file_id)
    else:
        await message.reply('Неподдерживаемый тип файла.')
        return

    if mime_type not in ALLOWED_DOC_TYPES:
        await message.reply('Неподдерживаемый формат. Допускаются PDF, JPG, PNG.')
        return
    if file_size > MAX_FILE_SIZE:
        await message.reply('Файл слишком большой — ограничение 20 MB.')
        return

    dest = UPLOAD_DIR / f"{datetime.utcnow().strftime('%Y%m%d%H%M%S')}_{filename}"
    try:
        await bot.download_file(file_obj.file_path, destination=dest.open('wb'))
    except Exception:
        if message.content_type == ContentType.DOCUMENT:
            await message.document.download(destination_file=str(dest))
        else:
            await message.photo[-1].download(destination_file=str(dest))

    data = await state.get_data()
    files = data.get('files', []) or []
    files.append(str(dest.name))
    await state.update_data(files=files)
    await message.reply(f'Файл {dest.name} сохранён. Отправьте ещё файлы или "Готово" если закончили.')

# Comment
@dp.message_handler(state=CheckUpStates.COMMENT, content_types=ContentType.TEXT)
async def process_comment(message: types.Message, state: FSMContext):
    text = message.text.strip()
    if text.lower() == 'отмена':
        return await cmd_cancel(message, state)
    if not text:
        text = 'нет'
    await state.update_data(comment=text)
    data = await state.get_data()
    preview = {
        'id': '—',
        'user_id': message.from_user.id,
        'username': message.from_user.username or message.from_user.full_name,
        'address': data.get('address'),
        'cadastral': data.get('cadastral'),
        'who': data.get('who'),
        'comment': data.get('comment'),
        'files': data.get('files', []),
        'created_at': datetime.utcnow().strftime('%Y-%m-%d %H:%M:%S UTC')
    }
    txt = fmt_request_message(preview)
    kb = types.ReplyKeyboardMarkup(resize_keyboard=True)
    kb.add('Отправить эксперту', 'Изменить данные')
    kb.add('Отмена')
    await CheckUpStates.CONFIRM.set()
    await message.answer(txt, parse_mode=ParseMode.HTML, reply_markup=kb)

# Confirm
@dp.message_handler(state=CheckUpStates.CONFIRM, content_types=ContentType.TEXT)
async def process_confirm(message: types.Message, state: FSMContext):
    text = message.text.strip()
    if text == 'Отправить эксперту':
        data = await state.get_data()
        req_id = str(uuid.uuid4())
        rec = {
            'id': req_id,
            'user_id': message.from_user.id,
            'username': message.from_user.username or message.from_user.full_name,
            'address': data.get('address'),
            'cadastral': data.get('cadastral'),
            'who': data.get('who'),
            'comment': data.get('comment'),
            'files': data.get('files', []),
            'created_at': datetime.utcnow().strftime('%Y-%m-%d %H:%M:%S')
        }
        await save_request_to_db(rec)
        txt = fmt_request_message(rec)
        await bot.send_message(ADMIN_CHAT_ID, txt, parse_mode=ParseMode.HTML)
        for fname in rec['files']:
            path = UPLOAD_DIR / fname
            if path.exists():
                try:
                    if path.suffix.lower() == '.pdf':
                        await bot.send_document(ADMIN_CHAT_ID, open(path, 'rb'))
                    else:
                        await bot.send_photo(ADMIN_CHAT_ID, open(path, 'rb'))
                except Exception:
                    logger.exception('Failed to send file %s', path)
        await message.answer('Спасибо! Ваш запрос отправлен эксперту 🧾', reply_markup=types.ReplyKeyboardRemove())
        await state.finish()
        return
    elif text == 'Изменить данные':
        await CheckUpStates.ADDRESS.set()
        await message.answer('Давайте изменим. Введите корректный адрес.')
        return
    elif text == 'Отмена':
        return await cmd_cancel(message, state)
    else:
        await message.reply('Выберите действие: Отправить эксперту / Изменить данные / Отмена.')

# -------------------- Reports --------------------
from openpyxl import Workbook
from openpyxl.utils import get_column_letter
import io

@dp.message_handler(commands=['report'])
async def cmd_report(message: types.Message):
    if message.from_user.id != ADMIN_CHAT_ID:
        await message.answer("⛔ Эта команда доступна только эксперту.")
        return

    args = message.get_args()
    days = 7
    if args and args.isdigit():
        days = int(args)
    elif args:
        await message.answer("Использование: /report <кол-во_дней> (например, /report 30)")
        return

    await message.answer(f"📊 Формирую отчёт за последние {days} дней...")

    query = f"""
        SELECT id, user_id, username, address, cadastral, who, comment, created_at
        FROM requests
        WHERE datetime(created_at) >= datetime('now', '-{days} days')
        ORDER BY created_at DESC
    """

    async with aiosqlite.connect(DB_PATH) as db:
        cursor = await db.execute(query)
        rows = await cursor.fetchall()

    if not rows:
        await message.answer("За указанный период заявок не найдено.")
        return

    wb = Workbook()
    ws = wb.active
    ws.title = "Заявки"

    headers = ["ID заявки", "User ID", "Имя пользователя", "Адрес", "Кадастровый номер", "Тип заявителя", "Комментарий", "Дата"]
    ws.append(headers)

    for row in rows:
        ws.append(row)

    for col_num, col_cells in enumerate(ws.columns, start=1):
        length = max(len(str(cell.value)) for cell in col_cells if cell.value)
        ws.column_dimensions[get_column_letter(col_num)].width = min(length + 2, 60)

    bio = io.BytesIO()
    wb.save(bio)
    bio.seek(0)

    await bot.send_document(
        chat_id=ADMIN_CHAT_ID,
        document=types.InputFile(bio, filename=f"requests_report_{days}d.xlsx"),
        caption=f"📈 Отчёт по заявкам за {days} дней"
    )

# -------------------- Startup --------------------
async def on_startup(dp):
    logger.info('Initializing DB...')
    await init_db()
    logger.info('Bot started')

if __name__ == '__main__':
    executor.start_polling(dp, on_startup=on_startup)
