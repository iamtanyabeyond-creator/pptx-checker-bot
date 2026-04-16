"""
bot.py — Telegram-бот для проверки презентаций
Все действия доступны через кнопки, без ввода команд.
"""

import os
import logging
import tempfile
import zipfile as _zipfile
import shutil
from dotenv import load_dotenv

from telegram import Update, InlineKeyboardButton, InlineKeyboardMarkup, BotCommand
from telegram.error import TimedOut
from telegram.ext import (
    Application,
    CommandHandler,
    MessageHandler,
    CallbackQueryHandler,
    ConversationHandler,
    filters,
    ContextTypes,
)

from analyzer import analyze_presentation, format_report, split_message, extract_all_text
from ai_editor import run_ai_editor, format_ai_report
from pdf_checker import check_pdf_hanging_prepositions, format_pdf_report
from pdf_checker import split_message as pdf_split

# ─────────────────────────────────────────────
# КОНФИГУРАЦИЯ
# ─────────────────────────────────────────────

load_dotenv()
BOT_TOKEN = os.getenv("BOT_TOKEN")

logging.basicConfig(format="%(asctime)s | %(levelname)s | %(message)s", level=logging.INFO)
logger = logging.getLogger(__name__)

MAX_FILE_MB   = 20
MAX_FILE_SIZE = MAX_FILE_MB * 1024 * 1024

# ─────────────────────────────────────────────
# СОСТОЯНИЯ
# ─────────────────────────────────────────────

ASK_MODE      = 0
WAIT_TEMPLATE = 1
WAIT_PPTX     = 2
WAIT_PDF      = 3

# ─────────────────────────────────────────────
# ХРАНИЛИЩЕ
# ─────────────────────────────────────────────

user_store: dict = {}
TEMP_DIR = "/tmp/pptx_bot"
os.makedirs(TEMP_DIR, exist_ok=True)


# ─────────────────────────────────────────────
# ТЕКСТЫ
# ─────────────────────────────────────────────

WELCOME_TEXT = (
    "Привет! Я проверяю презентации .pptx на ошибки.\n\n"
    "Что умею:\n"
    "🔤 Типографику — кавычки, тире, пробелы, сокращения\n"
    "📐 Вёрстку — прыгающие элементы между слайдами\n"
    "📄 Висячие предлоги — точно, через PDF\n"
    "✍️ Стиль текста — через ИИ-редактор"
)

HELP_TEXT = (
    "Как пользоваться:\n\n"
    "1. Нажми «Начать проверку»\n"
    "2. Выбери режим — с шаблоном или без\n"
    "3. Пришли файл .pptx или .zip\n"
    "4. Получи отчёт\n\n"
    "После отчёта появятся кнопки:\n"
    "• ИИ-редактор — канцелярит и стиль\n"
    "• PDF — точная проверка висячих предлогов\n\n"
    "Ограничение файла: до 20 МБ.\n"
    "Если файл больше — жми кнопку «Файл слишком большой»."
)

BIG_FILE_INSTRUCTION = (
    "Файл больше 20 МБ — Telegram не даёт его скачать.\n\n"
    "Способ 1 — сжать онлайн (быстрее всего):\n"
    "👉 https://www.wecompress.com/ru/\n"
    "Загрузи .pptx → нажми «Сжать» → скачай .pptx → пришли боту.\n\n"
    "Способ 2 — через Google Drive:\n"
    "1. Загрузи файл на Google Drive\n"
    "2. Правый клик → «Открыть с помощью» → Google Slides\n"
    "3. Файл → Скачать → Microsoft PowerPoint (.pptx)\n"
    "4. Пришли скачанный файл боту\n"
    "(Google Slides автоматически сжимает изображения)\n\n"
    "Обычно файл уменьшается в 3–10 раз."
)


# ─────────────────────────────────────────────
# КЛАВИАТУРЫ
# ─────────────────────────────────────────────

def main_keyboard():
    """Главное меню — всегда доступно"""
    return InlineKeyboardMarkup([
        [InlineKeyboardButton("📋 Проверить презентацию", callback_data="start_check")],
        [InlineKeyboardButton("✍️ ИИ-редактор", callback_data="ai_edit_direct"),
         InlineKeyboardButton("📄 Висячие предлоги", callback_data="start_pdf")],
        [InlineKeyboardButton("❓ Справка", callback_data="show_help"),
         InlineKeyboardButton("📦 Файл слишком большой", callback_data="show_bigfile")],
    ])

def mode_keyboard():
    return InlineKeyboardMarkup([
        [InlineKeyboardButton("📐 Есть шаблон",  callback_data="mode_template"),
         InlineKeyboardButton("📄 Без шаблона",  callback_data="mode_no_template")],
        [InlineKeyboardButton("← Главное меню",  callback_data="back_to_menu")],
    ])

def wait_template_keyboard():
    return InlineKeyboardMarkup([
        [InlineKeyboardButton("← Назад к выбору режима", callback_data="back_to_mode")],
        [InlineKeyboardButton("✖ Отменить",               callback_data="back_to_menu")],
    ])

def wait_pptx_keyboard(back_callback: str = "back_to_mode"):
    back_text = "← Назад к выбору режима" if back_callback == "back_to_mode" else "← Назад"
    return InlineKeyboardMarkup([
        [InlineKeyboardButton(back_text, callback_data=back_callback)],
        [InlineKeyboardButton("✖ Отменить", callback_data="back_to_menu")],
    ])

def wait_pdf_keyboard():
    return InlineKeyboardMarkup([
        [InlineKeyboardButton("← Назад",   callback_data="back_to_menu")],
        [InlineKeyboardButton("✖ Отменить", callback_data="back_to_menu")],
    ])

def after_report_keyboard():
    return InlineKeyboardMarkup([
        [InlineKeyboardButton("✍️ Проверить текст (ИИ)", callback_data="ai_edit")],
        [InlineKeyboardButton("📄 Висячие предлоги (PDF)", callback_data="start_pdf")],
        [InlineKeyboardButton("🔄 Проверить ещё одну", callback_data="start_check")],
        [InlineKeyboardButton("🏠 Главное меню",         callback_data="back_to_menu")],
    ])

def after_pdf_keyboard():
    return InlineKeyboardMarkup([
        [InlineKeyboardButton("✍️ Проверить текст (ИИ)", callback_data="ai_edit")],
        [InlineKeyboardButton("🔄 Проверить ещё одну",   callback_data="start_check")],
        [InlineKeyboardButton("🏠 Главное меню",          callback_data="back_to_menu")],
    ])

def after_ai_keyboard():
    return InlineKeyboardMarkup([
        [InlineKeyboardButton("📄 Висячие предлоги (PDF)", callback_data="start_pdf")],
        [InlineKeyboardButton("🔄 Проверить ещё одну",     callback_data="start_check")],
        [InlineKeyboardButton("🏠 Главное меню",            callback_data="back_to_menu")],
    ])

def bigfile_keyboard():
    return InlineKeyboardMarkup([
        [InlineKeyboardButton("← Назад", callback_data="back_to_menu")],
    ])


# ─────────────────────────────────────────────
# ВСПОМОГАТЕЛЬНЫЕ
# ─────────────────────────────────────────────

async def download_file(bot, document, dest_path: str) -> tuple[bool, str]:
    """Скачивает файл с увеличенными таймаутами. Возвращает (успех, текст_ошибки)."""
    try:
        tg_file = await bot.get_file(
            document.file_id,
            read_timeout=180,
            write_timeout=180,
            connect_timeout=60,
            pool_timeout=60,
        )
        await tg_file.download_to_drive(
            dest_path,
            read_timeout=180,
            write_timeout=180,
            connect_timeout=60,
            pool_timeout=60,
        )
        return True, ""
    except TimedOut:
        logger.warning("Timed out while downloading file from Telegram")
        return False, (
            "Telegram не успел отдать файл вовремя. "
            "Попробуй отправить его ещё раз или проверь файл чуть позже."
        )
    except Exception as e:
        logger.error(f"Ошибка скачивания: {e}")
        return False, f"Не удалось скачать файл: {e}"

def extract_pptx_from_zip(zip_path: str, dest_dir: str) -> str | None:
    try:
        with _zipfile.ZipFile(zip_path, 'r') as zf:
            pptx_files = [
                n for n in zf.namelist()
                if n.lower().endswith('.pptx') and not n.startswith('__MACOSX')
            ]
            if not pptx_files:
                return None
            target = pptx_files[0]
            dest = os.path.join(dest_dir, os.path.basename(target))
            with zf.open(target) as src, open(dest, 'wb') as dst:
                dst.write(src.read())
            return dest
    except Exception as e:
        logger.error(f"Ошибка разархивирования: {e}")
        return None

def is_too_large(doc) -> bool:
    return bool(doc.file_size and doc.file_size > MAX_FILE_SIZE)

def size_mb(doc) -> float:
    return round((doc.file_size or 0) / 1024 / 1024, 1)

async def show_main_menu(target, is_callback=False, text=None):
    t = text or WELCOME_TEXT
    if is_callback:
        await target.edit_message_text(t, reply_markup=main_keyboard())
    else:
        await target.reply_text(t, reply_markup=main_keyboard())

async def show_mode_question(target, is_callback=False):
    text = (
        "Есть шаблон презентации?\n\n"
        "С шаблоном — бот сравнит расположение всех элементов.\n"
        "Без шаблона — проверит типографику и консистентность."
    )
    if is_callback:
        await target.edit_message_text(text, reply_markup=mode_keyboard())
    else:
        await target.reply_text(text, reply_markup=mode_keyboard())


# ─────────────────────────────────────────────
# ГЛАВНОЕ МЕНЮ И НАВИГАЦИЯ
# ─────────────────────────────────────────────

async def cmd_start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text(WELCOME_TEXT, reply_markup=main_keyboard())
    return ConversationHandler.END

async def btn_back_to_menu(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    user_store.pop(update.effective_user.id, None)
    await show_main_menu(query, is_callback=True)
    return ConversationHandler.END

async def btn_show_help(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    await query.edit_message_text(
        HELP_TEXT,
        reply_markup=InlineKeyboardMarkup([[
            InlineKeyboardButton("← Назад", callback_data="back_to_menu")
        ]])
    )
    return ConversationHandler.END

async def btn_show_bigfile(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    await query.edit_message_text(
        BIG_FILE_INSTRUCTION,
        reply_markup=bigfile_keyboard()
    )
    return ConversationHandler.END


# ─────────────────────────────────────────────
# НАЧАЛО ПРОВЕРКИ
# ─────────────────────────────────────────────

async def btn_start_check(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    user_store.pop(update.effective_user.id, None)
    await show_mode_question(query, is_callback=True)
    return ASK_MODE


async def btn_start_ai_direct(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    user_id = update.effective_user.id
    user_store[user_id] = {"template_path": None, "slides_text": None, "mode": "ai_only"}
    await query.edit_message_text(
        "Пришли презентацию (.pptx или .zip), и я проверю текст через ИИ.",
        reply_markup=wait_pptx_keyboard(back_callback="back_to_menu"),
    )
    return WAIT_PPTX


# ─────────────────────────────────────────────
# ВЫБОР РЕЖИМА И НАВИГАЦИЯ
# ─────────────────────────────────────────────

async def handle_mode_choice(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    user_id = update.effective_user.id
    user_store[user_id] = {"template_path": None, "slides_text": None, "mode": "full"}

    if query.data == "mode_template":
        await query.edit_message_text(
            "Пришли файл шаблона (.pptx или .zip).\n"
            "Я запомню его и сравню с ним вёрстку.",
            reply_markup=wait_template_keyboard(),
        )
        return WAIT_TEMPLATE
    else:
        await query.edit_message_text(
            "Пришли файл презентации (.pptx или .zip).",
            reply_markup=wait_pptx_keyboard(),
        )
        return WAIT_PPTX

async def handle_back_to_mode(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    await show_mode_question(query, is_callback=True)
    return ASK_MODE


# ─────────────────────────────────────────────
# ПРИЁМ ШАБЛОНА
# ─────────────────────────────────────────────

async def handle_template_upload(update: Update, context: ContextTypes.DEFAULT_TYPE):
    doc = update.message.document
    user_id = update.message.from_user.id
    file_name = (doc.file_name or "").lower()

    if not (file_name.endswith(".pptx") or file_name.endswith(".zip")):
        await update.message.reply_text(
            "Неверный формат файла. Для шаблона подходят только .pptx или .zip с .pptx внутри.",
            reply_markup=wait_template_keyboard(),
        )
        return WAIT_TEMPLATE

    if is_too_large(doc):
        await update.message.reply_text(
            f"⚠️ Файл слишком большой: {size_mb(doc)} МБ (максимум {MAX_FILE_MB} МБ)\n\n"
            f"Сожми за 1 минуту — и пришли обратно:\n"
            f"👉 https://www.wecompress.com/ru/\n\n"
            f"Или загрузи на Google Drive → открой в Google Slides → скачай как .pptx",
            reply_markup=InlineKeyboardMarkup([
                [InlineKeyboardButton("📦 Подробнее", callback_data="show_bigfile")],
                [InlineKeyboardButton("← Назад",      callback_data="back_to_mode")],
            ])
        )
        return WAIT_TEMPLATE

    status = await update.message.reply_text("Сохраняю шаблон…")

    with tempfile.TemporaryDirectory() as tmpdir:
        raw_path = os.path.join(tmpdir, doc.file_name or "file")
        ok, err = await download_file(context.bot, doc, raw_path)
        if not ok:
            await status.edit_text(f"Не удалось скачать файл.\n\n{err}", reply_markup=wait_template_keyboard())
            return WAIT_TEMPLATE

        if file_name.endswith(".zip"):
            pptx_path = extract_pptx_from_zip(raw_path, tmpdir)
            if not pptx_path:
                await status.edit_text(
                    "В архиве не найден .pptx файл.",
                    reply_markup=wait_template_keyboard(),
                )
                return WAIT_TEMPLATE
        else:
            pptx_path = raw_path

        dest = os.path.join(TEMP_DIR, f"tmpl_{user_id}.pptx")
        shutil.copy2(pptx_path, dest)

    user_store[user_id]["template_path"] = dest
    await status.edit_text(
        "Шаблон сохранён. Теперь пришли файл презентации (.pptx или .zip).",
        reply_markup=wait_pptx_keyboard(),
    )
    return WAIT_PPTX


# ─────────────────────────────────────────────
# ПРИЁМ ПРЕЗЕНТАЦИИ → АНАЛИЗ
# ─────────────────────────────────────────────

async def handle_presentation_upload(update: Update, context: ContextTypes.DEFAULT_TYPE):
    doc = update.message.document
    user_id = update.message.from_user.id
    file_name = (doc.file_name or "").lower()

    if not (file_name.endswith(".pptx") or file_name.endswith(".zip")):
        await update.message.reply_text(
            "Неверный формат файла. Для проверки подходят только .pptx или .zip с .pptx внутри.",
            reply_markup=wait_pptx_keyboard(),
        )
        return WAIT_PPTX

    if is_too_large(doc):
        await update.message.reply_text(
            f"⚠️ Файл слишком большой: {size_mb(doc)} МБ (максимум {MAX_FILE_MB} МБ)\n\n"
            f"Сожми за 1 минуту и пришли обратно:\n"
            f"👉 https://www.wecompress.com/ru/",
            reply_markup=InlineKeyboardMarkup([
                [InlineKeyboardButton("← Назад", callback_data="back_to_mode")],
            ])
        )
        return WAIT_PPTX

    user_data = user_store.setdefault(user_id, {"template_path": None, "slides_text": None, "mode": "full"})
    template_path = user_data.get("template_path")
    mode = user_data.get("mode", "full")
    status = await update.message.reply_text("Скачиваю файл…")

    with tempfile.TemporaryDirectory() as tmpdir:
        raw_path = os.path.join(tmpdir, doc.file_name or "file")

        ok, err = await download_file(context.bot, doc, raw_path)
        if not ok:
            await status.edit_text(f"Не удалось скачать файл.\n\n{err}", reply_markup=wait_pptx_keyboard())
            return WAIT_PPTX

        if file_name.endswith(".zip"):
            pptx_path = extract_pptx_from_zip(raw_path, tmpdir)
            if not pptx_path:
                await status.edit_text(
                    "В архиве не найден .pptx файл.",
                    reply_markup=wait_pptx_keyboard(),
                )
                return WAIT_PPTX
        else:
            pptx_path = raw_path

        try:
            await status.edit_text("Анализирую…")
            report      = analyze_presentation(pptx_path, template_path=template_path)
            slides_text = extract_all_text(pptx_path)
            if user_id in user_store:
                user_store[user_id]["slides_text"] = slides_text
        except Exception as e:
            logger.error(f"Ошибка анализа: {e}")
            await status.edit_text(f"Ошибка при анализе:\n`{e}`", parse_mode="Markdown")
            return ConversationHandler.END

    try:
        formatted = format_report(report)
        parts = split_message(formatted)
        await status.delete()

        for i, part in enumerate(parts):
            is_last = (i == len(parts) - 1)
            await update.message.reply_text(
                part,
                parse_mode="Markdown",
                reply_markup=after_report_keyboard() if is_last else None,
            )

        if template_path and os.path.exists(template_path):
            os.remove(template_path)

    except TimedOut:
        logger.warning("Timed out while sending report")
        await update.message.reply_text(
            "Telegram не успел отправить отчёт целиком. Попробуй проверить файл ещё раз.",
            reply_markup=main_keyboard(),
        )
    except Exception as e:
        logger.error(f"Ошибка отправки: {e}")
        await update.message.reply_text("Не удалось отправить отчёт.", reply_markup=main_keyboard())

    return ConversationHandler.END


# ─────────────────────────────────────────────
# PDF
# ─────────────────────────────────────────────

async def btn_start_pdf(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    await query.edit_message_reply_markup(reply_markup=None)
    await query.message.reply_text(
        "Пришли PDF-версию презентации.\n\n"
        "Как выгрузить из PowerPoint:\n"
        "Файл → Экспорт → Создать PDF/XPS",
        reply_markup=wait_pdf_keyboard(),
    )
    return WAIT_PDF

async def handle_pdf_upload(update: Update, context: ContextTypes.DEFAULT_TYPE):
    doc = update.message.document
    file_name = (doc.file_name or "").lower()

    if not file_name.endswith(".pdf"):
        await update.message.reply_text(
            "Неверный формат файла. Для этой проверки нужен только .pdf.",
            reply_markup=wait_pdf_keyboard(),
        )
        return WAIT_PDF

    if is_too_large(doc):
        await update.message.reply_text(
            f"⚠️ Файл слишком большой: {size_mb(doc)} МБ\n\n"
            "При экспорте PDF из PowerPoint выбери «Минимальный размер файла».",
            reply_markup=wait_pdf_keyboard(),
        )
        return WAIT_PDF

    status = await update.message.reply_text("Читаю PDF…")

    with tempfile.TemporaryDirectory() as tmpdir:
        pdf_path = os.path.join(tmpdir, "slides.pdf")

        ok, err = await download_file(context.bot, doc, pdf_path)
        if not ok:
            await status.edit_text(f"Не удалось скачать PDF.\n\n{err}", reply_markup=wait_pdf_keyboard())
            return WAIT_PDF

        try:
            import pdfplumber
            with pdfplumber.open(pdf_path) as pdf:
                total_pages = len(pdf.pages)
            results = check_pdf_hanging_prepositions(pdf_path)
        except Exception as e:
            logger.error(f"Ошибка PDF: {e}")
            await status.edit_text(f"Ошибка:\n`{e}`", parse_mode="Markdown")
            return ConversationHandler.END

    try:
        formatted = format_pdf_report(results, total_pages)
        parts = pdf_split(formatted)
        await status.delete()

        for i, part in enumerate(parts):
            is_last = (i == len(parts) - 1)
            await update.message.reply_text(
                part,
                parse_mode="Markdown",
                reply_markup=after_pdf_keyboard() if is_last else None,
            )
    except TimedOut:
        logger.warning("Timed out while sending PDF report")
        await update.message.reply_text(
            "Telegram не успел отправить PDF-отчёт. Попробуй ещё раз.",
            reply_markup=main_keyboard(),
        )
    except Exception as e:
        logger.error(f"Ошибка PDF-отчёта: {e}")
        await update.message.reply_text(
            "Не удалось отправить PDF-отчёт.",
            reply_markup=main_keyboard(),
        )

    return ConversationHandler.END


# ─────────────────────────────────────────────
# ИИ-РЕДАКТОР
# ─────────────────────────────────────────────

async def handle_ai_editor(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    user_id = update.effective_user.id

    slides_text = user_store.get(user_id, {}).get("slides_text")
    if not slides_text:
        await query.message.reply_text(
            "Текст не найден. Пришли презентацию заново или запусти ИИ-редактор из главного меню.",
            reply_markup=main_keyboard()
        )
        return

    await query.edit_message_reply_markup(reply_markup=None)
    status = await query.message.reply_text("ИИ-редактор читает презентацию…")

    ai_response = run_ai_editor(slides_text)
    formatted   = format_ai_report(ai_response)
    await status.delete()

    parts = split_message(formatted)
    for i, part in enumerate(parts):
        is_last = (i == len(parts) - 1)
        await query.message.reply_text(
            part,
            parse_mode="Markdown",
            reply_markup=after_ai_keyboard() if is_last else None,
        )

    user_store.pop(user_id, None)


# ─────────────────────────────────────────────
# ФОЛБЕКИ — любой текст или файл вне диалога
# ─────────────────────────────────────────────

async def handle_stray_message(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text(
        "Выбери действие:",
        reply_markup=main_keyboard()
    )


# ─────────────────────────────────────────────
# ТОЧКА ВХОДА
# ─────────────────────────────────────────────

async def post_init(app: Application):
    """Устанавливает кнопку меню в интерфейсе Telegram"""
    await app.bot.set_my_commands([
        BotCommand("start", "Главное меню"),
    ])


def main():
    if not BOT_TOKEN:
        print("❌ Токен не найден! Проверь .env")
        return

    app = (
        Application.builder()
        .token(BOT_TOKEN)
        .post_init(post_init)
        .read_timeout(180)
        .write_timeout(180)
        .connect_timeout(60)
        .pool_timeout(60)
        .build()
    )

    # Глобальные кнопки-обработчики (вне ConversationHandler)
    app.add_handler(CommandHandler("start", cmd_start))
    app.add_handler(CallbackQueryHandler(btn_show_help,    pattern=r"^show_help$"))
    app.add_handler(CallbackQueryHandler(btn_show_bigfile, pattern=r"^show_bigfile$"))
    app.add_handler(CallbackQueryHandler(btn_back_to_menu, pattern=r"^back_to_menu$"))
    app.add_handler(CallbackQueryHandler(handle_ai_editor, pattern=r"^ai_edit$"))

    # Основной диалог проверки
    conv = ConversationHandler(
        entry_points=[
            CallbackQueryHandler(btn_start_check, pattern=r"^start_check$"),
            CallbackQueryHandler(btn_start_ai_direct, pattern=r"^ai_edit_direct$"),
            CallbackQueryHandler(btn_start_pdf,   pattern=r"^start_pdf$"),
        ],
        states={
            ASK_MODE: [
                CallbackQueryHandler(handle_mode_choice,  pattern=r"^mode_"),
                CallbackQueryHandler(btn_back_to_menu,    pattern=r"^back_to_menu$"),
            ],
            WAIT_TEMPLATE: [
                MessageHandler(filters.Document.ALL, handle_template_upload),
                CallbackQueryHandler(handle_back_to_mode, pattern=r"^back_to_mode$"),
                CallbackQueryHandler(btn_back_to_menu,    pattern=r"^back_to_menu$"),
                CallbackQueryHandler(btn_show_bigfile,    pattern=r"^show_bigfile$"),
            ],
            WAIT_PPTX: [
                MessageHandler(filters.Document.ALL, handle_presentation_upload),
                CallbackQueryHandler(handle_back_to_mode, pattern=r"^back_to_mode$"),
                CallbackQueryHandler(btn_back_to_menu,    pattern=r"^back_to_menu$"),
                CallbackQueryHandler(btn_show_bigfile,    pattern=r"^show_bigfile$"),
            ],
            WAIT_PDF: [
                MessageHandler(filters.Document.ALL, handle_pdf_upload),
                CallbackQueryHandler(btn_back_to_menu,    pattern=r"^back_to_menu$"),
            ],
        },
        fallbacks=[
            CallbackQueryHandler(btn_back_to_menu, pattern=r"^back_to_menu$"),
        ],
        allow_reentry=True,
    )

    app.add_handler(conv)
    app.add_handler(MessageHandler(
        (filters.Document.ALL | filters.TEXT) & ~filters.COMMAND,
        handle_stray_message
    ))

    logger.info("Бот запущен.")
    app.run_polling()


if __name__ == "__main__":
    main()
