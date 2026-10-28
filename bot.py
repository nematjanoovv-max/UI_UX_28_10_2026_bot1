"""Telegram bot: PDF, Excel, CSV, TXT va Markdown fayllarini Word ga o'giradi.

Render.com uchun webhook, lokal ish uchun long-polling avtomatik tanlanadi.
"""
from __future__ import annotations

import asyncio
import hashlib
import html
import logging
import os
import re
import shutil
import tempfile
import time
from pathlib import Path

from aiogram import Bot, Dispatcher, F, Router
from aiogram.client.default import DefaultBotProperties
from aiogram.enums import ParseMode
from aiogram.filters import Command, CommandStart
from aiogram.types import FSInputFile, Message
from aiohttp import web
from dotenv import load_dotenv

from converters import SUPPORTED_EXTENSIONS, ConversionError, convert_file

load_dotenv()

BOT_TOKEN = os.getenv("BOT_TOKEN", "").strip()
MAX_FILE_MB = int(os.getenv("MAX_FILE_MB", "20"))
MAX_FILE_BYTES = MAX_FILE_MB * 1024 * 1024
CONVERT_TIMEOUT = int(os.getenv("CONVERT_TIMEOUT", "300"))
PORT = int(os.getenv("PORT", "8080"))
WEBHOOK_PATH = "/webhook"
WEBHOOK_BASE = (
    os.getenv("WEBHOOK_BASE_URL") or os.getenv("RENDER_EXTERNAL_URL") or ""
).rstrip("/")
WEBHOOK_SECRET = os.getenv("WEBHOOK_SECRET") or hashlib.sha256(
    f"{BOT_TOKEN}:wordbot".encode()
).hexdigest()[:32]

COOLDOWN_SECONDS = 5

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s %(levelname)s %(name)s: %(message)s",
)
logger = logging.getLogger("wordbot")

router = Router()
_last_request: dict[int, float] = {}

HELP_TEXT = """🤖 <b>Word Converter</b>

Men quyidagi fayllarni <b>Word (.docx)</b> ga aylantiraman:

📄 PDF
📊 Excel — .xlsx, .xlsm, .xls
📋 CSV / TSV
📝 TXT / Markdown (.md)

📤 Menga faylni yuboring — tayyor Word nusxasini qaytaraman.

📦 Maksimal hajm: {max_mb} MB
⏱ O'rtacha vaqt: bir necha soniya

<i>Eslatma: skanerlangan PDF uchun OCR ishlatilmaydi; parol bilan
himoyalangan PDF va eski .doc format qo'llab-quvvatlanmaydi.</i>"""


def _safe_stem(name: str) -> str:
    stem = Path(name).stem
    stem = re.sub(r'[\\/:*?"<>|\x00-\x1f]+', "_", stem).strip(" ._")
    return (stem or "hujjat")[:80]


async def _edit(message: Message, text: str) -> None:
    try:
        await message.edit_text(text)
    except Exception:  # xabar o'chirilgan yoki matn o'zgarmagan bo'lishi mumkin
        pass


@router.message(CommandStart())
async def cmd_start(message: Message) -> None:
    await message.answer(HELP_TEXT.format(max_mb=MAX_FILE_MB))


@router.message(Command("help"))
async def cmd_help(message: Message) -> None:
    await message.answer(HELP_TEXT.format(max_mb=MAX_FILE_MB))


@router.message(F.document)
async def handle_document(
    message: Message, bot: Bot, semaphore: asyncio.Semaphore
) -> None:
    document = message.document
    name = document.file_name or "fayl"
    extension = Path(name).suffix.lower()

    if extension == ".docx":
        await message.answer("📄 Bu fayl allaqachon Word (.docx) formatida.")
        return

    if extension not in SUPPORTED_EXTENSIONS:
        await message.answer(
            "❌ Bu format qo'llab-quvvatlanmaydi.\n\n"
            "Yuborish mumkin: <b>PDF, XLSX, XLSM, XLS, CSV, TSV, TXT, MD</b>."
        )
        return

    if document.file_size and document.file_size > MAX_FILE_BYTES:
        await message.answer(f"❌ Fayl juda katta. Maksimal hajm: {MAX_FILE_MB} MB.")
        return

    user_id = message.from_user.id
    now = time.monotonic()
    remaining = COOLDOWN_SECONDS - (now - _last_request.get(user_id, 0))
    if remaining > 0:
        await message.answer(f"⏳ Iltimos, {int(remaining) + 1} soniya kutib turing.")
        return
    _last_request[user_id] = now

    status = await message.answer("⏳ Fayl qabul qilindi. Word ga o'girilmoqda...")
    workdir = Path(tempfile.mkdtemp(prefix="wordbot-"))
    source = workdir / f"input{extension}"
    target = workdir / "output.docx"

    try:
        await bot.download(document, destination=source)

        if semaphore.locked():
            await _edit(status, "⏳ Navbatda. Oldingi fayl tugagach boshlanadi...")

        async with semaphore:
            await _edit(status, "⏳ Konvertatsiya qilinmoqda...")
            # shortcut: wait_for threadni to'xtatmaydi — juda katta fayl
            # timeoutdan keyin ham fonda davom etishi mumkin.
            await asyncio.wait_for(
                convert_file(source, target), timeout=CONVERT_TIMEOUT
            )

        await message.answer_document(
            FSInputFile(target, filename=f"{_safe_stem(name)}.docx"),
            caption="✅ Tayyor! Word faylingiz.",
        )
        try:
            await status.delete()
        except Exception:
            pass

    except asyncio.TimeoutError:
        await _edit(
            status,
            "❌ Konvertatsiya juda uzoq davom etdi. "
            "Faylni kichikroqlab yoki bo'lib yuboring.",
        )
    except ConversionError as exc:
        await _edit(status, f"❌ {html.escape(str(exc))}")
    except Exception:
        logger.exception("Foydalanuvchi %s faylini konvertatsiya qilishda xato", user_id)
        await _edit(
            status,
            "❌ Kutilmagan xatolik. Iltimos, keyinroq qayta urinib ko'ring.",
        )
    finally:
        shutil.rmtree(workdir, ignore_errors=True)


async def _health(_request) -> web.Response:
    return web.Response(text="OK")


async def _set_webhook(bot: Bot) -> None:
    if not WEBHOOK_BASE:
        return
    url = f"{WEBHOOK_BASE}{WEBHOOK_PATH}"
    await bot.set_webhook(
        url,
        secret_token=WEBHOOK_SECRET,
        drop_pending_updates=False,
        allowed_updates=["message"],
    )
    logger.info("Webhook o'rnatildi: %s", url)


def _build_dispatcher() -> Dispatcher:
    dispatcher = Dispatcher()
    dispatcher.include_router(router)
    dispatcher["semaphore"] = asyncio.Semaphore(1)
    return dispatcher


async def _run_polling(bot: Bot, dispatcher: Dispatcher) -> None:
    await bot.delete_webhook(drop_pending_updates=False)
    logger.info("Long-polling rejimida ishga tushdi")
    await dispatcher.start_polling(bot, allowed_updates=["message"])


def main() -> None:
    if not BOT_TOKEN:
        raise SystemExit(
            "BOT_TOKEN topilmadi. .env fayliga yoki Render Environment ga qo'shing."
        )

    bot = Bot(BOT_TOKEN, default=DefaultBotProperties(parse_mode=ParseMode.HTML))
    dispatcher = _build_dispatcher()

    if WEBHOOK_BASE:
        from aiogram.webhook.aiohttp_server import SimpleRequestHandler, setup_application

        app = web.Application()
        app.router.add_get("/", _health)
        app.router.add_get("/health", _health)
        SimpleRequestHandler(
            dispatcher=dispatcher,
            bot=bot,
            secret_token=WEBHOOK_SECRET,
        ).register(app, path=WEBHOOK_PATH)
        setup_application(app, dispatcher, bot=bot)
        dispatcher.startup.register(_set_webhook)

        logger.info("Webhook rejimida, port %s", PORT)
        web.run_app(app, host="0.0.0.0", port=PORT)
    else:
        asyncio.run(_run_polling(bot, dispatcher))


if __name__ == "__main__":
    main()
