# Word Converter Bot — UI_UX_28_10_2026

Telegram bot: **PDF, Excel, CSV, TXT va Markdown** fayllarini **Word (.docx)** ga
aylantiradi.

## Imkoniyatlar

| Format | Natija |
|---|---|
| `.pdf` | `.docx` — `pdf2docx` bilan original joylashuv saqlanadi |
| `.xlsx`, `.xlsm`, `.xls` | Har bir varaq alohida jadval sifatida |
| `.csv`, `.tsv` | Ajratuvchi avtomatik aniqlanadi, jadvalga o'tkaziladi |
| `.txt`, `.md` | Matn; Markdown sarlavha, ro'yxat va kod bloklari bilan |
| `.docx` | Allaqachon Word — qayta ishlanmaydi |
| `.doc` | Qo'llab-quvvatlanmaydi (LibreOffice talab qiladi) |

- Fayl hajmi: **20 MB** gacha (Telegram Bot API chegarasi).
- Bir vaqtda bitta konvertatsiya — Render free (512 MB) xotirasini himoya qiladi.
- 100 sahifadan katta PDF matn rejimida o'giriladi (xotira tejash uchun).
- Skanerlangan PDF uchun **OCR yo'q**; parol bilan himoyalangan PDF ochilmaydi.

## Lokal ishga tushirish

```bash
cd UI_UX_28_10_2026
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt

cp .env.example .env
# .env ichiga @BotFather dan olingan BOT_TOKEN ni yozing

python bot.py
```

Token bo'lmasa yoki noto'g'ri bo'lsa bot aniq xato bilan to'xtaydi.
`RENDER_EXTERNAL_URL` yo'q bo'lsa **long-polling**, bor bo'lsa **webhook**
rejimi avtomatik ishlaydi.

## Test

```bash
python -m unittest discover -s tests -v
```

Testlar PDF, XLSX, CSV, TXT va Markdown konvertatsiyasini tekshiradi.

## GitHub

Repo: <https://github.com/nematjanoovv-max/UI_UX_28_10_2026>

```bash
git add .
git commit -m "Word converter telegram bot"
git push
```

## Render.com ga joylash

1. Yuqoridagi repo Render akkauntiga ulangan bo'lsin.
2. Bir marta bosish bilan Blueprint:

   <https://render.com/deploy?repo=https://github.com/nematjanoovv-max/UI_UX_28_10_2026>

3. `BOT_TOKEN` so'ralganda @BotFather dan olingan tokenni kiriting.
4. **Apply** — `render.yaml` bo'yicha web service yaratiladi.

Render o'zi `RENDER_EXTERNAL_URL` va `PORT` ni beradi; bot webhookni
avtomatik o'rnatadi. `GET /health` — health-check uchun.

> **Free plan:** servis 15 daqiqa harakatsizlikdan keyin uxlaydi. Telegram
> webhook kelganda uyg'onadi, lekin birinchi javob ~30–60 soniya kechikishi
> mumkin. Keyingi xabarlar tez ishlaydi.

## Sozlamalar

| Env | Standart | Ma'nosi |
|---|---|---|
| `BOT_TOKEN` | — | Majburiy, @BotFather tokeni |
| `MAX_FILE_MB` | `20` | Maksimal fayl hajmi |
| `CONVERT_TIMEOUT` | `300` | Konvertatsiya uchun soniyalarda limit |
| `WEBHOOK_BASE_URL` | `RENDER_EXTERNAL_URL` | Webhook manzili (lokal test uchun) |
| `WEBHOOK_SECRET` | token hash | Telegram webhook maxfiy kaliti |

## Struktura

```
UI_UX_28_10_2026/
├── bot.py            # Telegram handlerlar, webhook/polling
├── converters.py     # PDF/Excel/CSV/TXT/MD -> DOCX
├── requirements.txt
├── render.yaml       # Render Blueprint
├── .env.example
└── tests/
    └── test_converters.py
```

## Xavfsizlik

- `.env` va tokenlar `.gitignore` orqali repoga tushmaydi.
- Webhook `WEBHOOK_SECRET` bilan tekshiriladi.
- Har bir foydalanuvchi uchun 5 soniyalik cooldown.
- Vaqtinchalik fayllar har bir so'rovdan keyin o'chiriladi.
