"""
چت‌بات فارسی با Flask و OpenRouter
قابلیت‌ها: چند مکالمه، ذخیره دائمی، جدا بودن کاربران، آپلود عکس،
محدودیت پیام روزانه، پاسخ استریم (تایپ‌شونده)، جستجو در تاریخچه
"""

from flask import Flask, render_template, render_template_string, request, jsonify, g, Response, session, redirect, url_for
import requests
import os
import json
import uuid
import time
import datetime
import io
import base64
from urllib.parse import quote as url_quote
from PIL import Image
from huggingface_hub import InferenceClient
import re
from functools import wraps
from pypdf import PdfReader
from docx import Document
import openpyxl

app = Flask(__name__)
app.secret_key = os.environ.get("SECRET_KEY", "یه-رشته-تصادفی-برای-امنیت-عوضش-کن")

API_KEY = os.environ.get("OPENROUTER_API_KEY")

API_URL = "https://openrouter.ai/api/v1/chat/completions"
MODEL = "openrouter/free"

# تعریف سرویس‌دهنده‌های هوش مصنوعی قابل انتخاب.
# هر کدوم کلیدش رو از متغیر محیطی روی Railway می‌خونه - همه‌شون متعلق به خود توئه،
# کاربر فقط از بینشون یکی رو انتخاب می‌کنه، نیازی به وارد کردن کلید شخصی نیست.
PROVIDERS = {
    "openrouter": {
        "label": "🆓 رایگان (OpenRouter)",
        "url": API_URL,
        "model": MODEL,
        "supports_vision": True,
        "api_key": API_KEY,
    },
    "atria": {
        "label": "🚀 Atria Dawn",
        "url": "https://api.atria-asi.ai/v1/chat/completions",
        "model": "Atria-Dawn-Preview",
        "supports_vision": False,  # این مدل فقط متنی هست
        "api_key": os.environ.get("ATRIA_API_KEY"),
    },
    "gemini": {
        "label": "✨ Gemini",
        "url": "https://generativelanguage.googleapis.com/v1beta/openai/chat/completions",
        "model": "gemini-3.8-flash",
        "supports_vision": True,
        "api_key": os.environ.get("GEMINI_API_KEY"),
    },
    "groq": {
        "label": "⚡ Groq",
        "url": "https://api.groq.com/openai/v1/chat/completions",
        "model": "openai/gpt-oss-20b",
        "supports_vision": False,
        "api_key": os.environ.get("GROQ_API_KEY"),
    },
}
DEFAULT_PROVIDER = "openrouter"


# ==================== ساخت و ویرایش تصویر ====================

import logging

logger = logging.getLogger(__name__)

# کلیدها فقط از متغیرهای محیطی Railway خوانده می‌شوند.
CF_ACCOUNT_ID = os.environ.get("CF_ACCOUNT_ID")
CF_API_TOKEN = os.environ.get("CF_API_TOKEN")
HF_API_KEY = os.environ.get("HF_API_KEY")
SENSENOVA_API_KEY = os.environ.get("SENSENOVA_API_KEY")
POLLINATIONS_API_KEY = os.environ.get("POLLINATIONS_API_KEY")

CF_IMAGE_MODEL = "@cf/black-forest-labs/flux-1-schnell"

POLLINATIONS_EDIT_URL = (
    "https://gen.pollinations.ai/v1/images/edits"
)

MAX_IMAGE_BYTES = 8 * 1024 * 1024
IMAGE_TIMEOUT = (15, 120)
MAX_IMAGE_SIDE = 1024

IMAGE_GEN_TRIGGERS = [
    "عکس بساز", "تصویر بساز", "یه عکس از", "یک عکس از",
    "عکسی از", "نقاشی بکش", "نقاشی کن", "برام بکش",
    "یه نقاشی از", "یک نقاشی از",
]

IMAGE_EDIT_TRIGGERS = [
    "تغییر", "تبدیل", "ویرایش", "عوض", "رنگی", "رنگیش",
]

IMAGE_EDIT_COLOR_PATTERN = re.compile(
    r"(سفید|سیاه|قرمز|آبی|سبز|زرد|نارنجی|بنفش|صورتی|"
    r"طلایی|خاکستری)\s*(ش)?\s*کن"
)


def detect_image_generation_request(text):
    text = (text or "").strip()
    return any(trigger in text for trigger in IMAGE_GEN_TRIGGERS)


def detect_image_edit_request(text):
    text = (text or "").strip()
    return (
        any(trigger in text for trigger in IMAGE_EDIT_TRIGGERS)
        or bool(IMAGE_EDIT_COLOR_PATTERN.search(text))
    )


def translate_to_image_prompt(persian_text):
    """تبدیل دستور فارسی به دستور مناسب مدل تصویر."""
    if not API_KEY:
        return persian_text

    try:
        response = requests.post(
            API_URL,
            headers={
                "Authorization": f"Bearer {API_KEY}",
                "Content-Type": "application/json",
            },
            json={
                "model": MODEL,
                "messages": [
                    {
                        "role": "system",
                        "content": (
                            "Translate the user's image editing or "
                            "generation request into a concise English "
                            "prompt. Preserve the intended changes and "
                            "the identity and composition of the source "
                            "image when editing. Return only the prompt."
                        ),
                    },
                    {"role": "user", "content": persian_text},
                ],
            },
            timeout=30,
        )
        response.raise_for_status()
        result = response.json()
        prompt = result["choices"][0]["message"]["content"].strip()
        return prompt or persian_text
    except Exception:
        logger.exception("Image prompt translation failed")
        return persian_text


def _decode_input_image(image_data_url):
    """اعتبارسنجی، کوچک‌سازی و آماده‌سازی تصویر ورودی."""
    if not isinstance(image_data_url, str):
        raise ValueError("تصویر ورودی معتبر نیست.")

    if not image_data_url.startswith("data:image/"):
        raise ValueError("تصویر باید به شکل Data URL ارسال شود.")

    try:
        header, encoded = image_data_url.split(",", 1)
        raw = base64.b64decode(encoded, validate=True)
    except Exception as exc:
        raise ValueError("رمزگذاری تصویر معتبر نیست.") from exc

    if len(raw) > MAX_IMAGE_BYTES:
        raise ValueError("حجم تصویر بیشتر از ۸ مگابایت است.")

    try:
        with Image.open(io.BytesIO(raw)) as source:
            if source.format not in ("PNG", "JPEG", "WEBP"):
                raise ValueError("فرمت تصویر پشتیبانی نمی‌شود.")

            image = source.convert("RGB")
            image.thumbnail((MAX_IMAGE_SIDE, MAX_IMAGE_SIDE))

            output = io.BytesIO()
            image.save(output, format="PNG")
            result = output.getvalue()
    except ValueError:
        raise
    except Exception as exc:
        raise ValueError("خواندن تصویر ورودی ناموفق بود.") from exc

    return result


def _image_data_url(raw_bytes, mime="image/png"):
    encoded = base64.b64encode(raw_bytes).decode("ascii")
    return f"data:{mime};base64,{encoded}"


def _extract_api_error(response):
    """استخراج خطای قابل‌فهم بدون ثبت کلیدهای محرمانه."""
    try:
        data = response.json()
        error = data.get("error", data.get("errors", data))
        if isinstance(error, list) and error:
            error = error[0]
        if isinstance(error, dict):
            message = error.get("message") or str(error)
        else:
            message = str(error)
    except Exception:
        message = response.text[:500]

    return f"HTTP {response.status_code}: {message[:500]}"


def _try_pollinations_edit(image_bytes, prompt):
    """
    ویرایش از طریق API رسمی Pollinations.
    به کلید POLLINATIONS_API_KEY نیاز دارد.
    رایگان بودن درخواست به شرایط اعتبار حساب بستگی دارد.
    """
    if not POLLINATIONS_API_KEY:
        raise RuntimeError(
            "کلید POLLINATIONS_API_KEY تنظیم نشده است."
        )

    response = requests.post(
        POLLINATIONS_EDIT_URL,
        headers={
            "Authorization": f"Bearer {POLLINATIONS_API_KEY}",
            "Accept": "application/json",
        },
        files={
            "image": ("source.png", image_bytes, "image/png"),
        },
        data={
            "prompt": prompt,
            "model": "kontext",
            "size": "1024x1024",
        },
        timeout=IMAGE_TIMEOUT,
    )

    if not response.ok:
        raise RuntimeError(_extract_api_error(response))

    content_type = response.headers.get("Content-Type", "").lower()

    # بعضی APIها مستقیماً فایل تصویر برمی‌گردانند.
    if content_type.startswith("image/"):
        if len(response.content) > 15 * 1024 * 1024:
            raise RuntimeError("حجم تصویر خروجی بیش از حد مجاز است.")
        try:
            with Image.open(io.BytesIO(response.content)) as img:
                img.verify()
        except Exception as exc:
            raise RuntimeError("تصویر خروجی معتبر نیست.") from exc

        return _image_data_url(response.content)

    # پاسخ JSON سازگار با API تصاویر OpenAI.
    try:
        data = response.json()
        items = data.get("data", [])
        if not items:
            raise RuntimeError(
                "پاسخ سرویس فاقد تصویر خروجی است."
            )

        item = items[0]

        if item.get("b64_json"):
            raw = base64.b64decode(item["b64_json"], validate=True)
            with Image.open(io.BytesIO(raw)) as img:
                img.verify()
            return _image_data_url(raw)

        if item.get("url"):
            # برای جلوگیری از SSRF، URL خروجی را خودکار دریافت نمی‌کنیم.
            # اگر سرویس URL برگرداند، باید دریافت آن جداگانه امن‌سازی شود.
            raise RuntimeError(
                "سرویس URL تصویر برگرداند، نه داده تصویر. "
                "پاسخ API باید با فرمت b64_json تنظیم شود."
            )

    except RuntimeError:
        raise
    except Exception as exc:
        raise RuntimeError(
            "قالب پاسخ ویرایش تصویر قابل‌شناسایی نیست."
        ) from exc

    raise RuntimeError("تصویر خروجی در پاسخ سرویس یافت نشد.")


def _try_huggingface_edit(image_bytes, prompt):
    """
    گزینه جایگزین Hugging Face.
    ممکن است اعتبار مصرفی بخواهد و رایگان بودن دائمی تضمین نیست.
    """
    if not HF_API_KEY:
        raise RuntimeError("کلید HF_API_KEY تنظیم نشده است.")

    client = InferenceClient(
        provider="fal-ai",
        api_key=HF_API_KEY,
        timeout=120,
    )

    result = client.image_to_image(
        image_bytes,
        prompt=prompt,
        model="black-forest-labs/FLUX.1-Kontext-dev",
    )

    output = io.BytesIO()
    result.save(output, format="PNG")
    return _image_data_url(output.getvalue())


def _try_sensenova_edit(image_bytes, prompt):
    """گزینه اختیاری SenseNova؛ نیازمند کلید و مجوز معتبر."""
    if not SENSENOVA_API_KEY:
        raise RuntimeError("کلید SenseNova تنظیم نشده است.")

    response = requests.post(
        "https://token.sensenova.cn/v1/images/edits",
        headers={
            "Authorization": f"Bearer {SENSENOVA_API_KEY}",
        },
        files={
            "image": ("source.png", image_bytes, "image/png"),
        },
        data={
            "prompt": prompt,
            "model": "sensenova-u1.5-lite",
        },
        timeout=IMAGE_TIMEOUT,
    )

    if not response.ok:
        raise RuntimeError(_extract_api_error(response))

    data = response.json()
    items = data.get("data") or []

    if not items:
        raise RuntimeError(
            str(data.get("error", "تصویری برگردانده نشد"))[:500]
        )

    item = items[0]

    if item.get("b64_json"):
        raw = base64.b64decode(item["b64_json"], validate=True)
        with Image.open(io.BytesIO(raw)) as img:
            img.verify()
        return _image_data_url(raw)

    if item.get("url"):
        # URL خارجی بدون بررسی دانلود نمی‌شود.
        raise RuntimeError(
            "سرویس URL تصویر برگرداند؛ دریافت امن آن نیازمند بررسی جداگانه است."
        )

    raise RuntimeError("تصویر خروجی معتبر نیست.")


def edit_image(image_data_url, prompt):
    """
    ویرایش تصویر با ترتیب مشخص و خطاهای قابل‌ردیابی.
    این تابع مستقل از تنظیم بودن Cloudflare کار می‌کند.
    """
    try:
        image_bytes = _decode_input_image(image_data_url)
    except Exception as exc:
        return None, f"خطا در تصویر ورودی: {exc}"

    # فقط ارائه‌دهنده‌هایی که کلیدشان تنظیم شده امتحان می‌شوند.
    candidates = []

    if POLLINATIONS_API_KEY:
        candidates.append(("Pollinations", _try_pollinations_edit))

    if HF_API_KEY:
        candidates.append(("Hugging Face", _try_huggingface_edit))

    if SENSENOVA_API_KEY:
        candidates.append(("SenseNova", _try_sensenova_edit))

    if not candidates:
        return None, (
            "هیچ سرویس ویرایش تصویر فعالی تنظیم نشده است. "
            "در Railway حداقل یکی از این متغیرها را تنظیم کن: "
            "POLLINATIONS_API_KEY، HF_API_KEY یا SENSENOVA_API_KEY."
        )

    errors = []

    for name, handler in candidates:
        try:
            result = handler(image_bytes, prompt)
            if result:
                logger.info("Image edit succeeded: %s", name)
                return result, None

            errors.append(f"{name}: خروجی خالی")
        except Exception as exc:
            # جزئیات فقط در گزارش سرور ثبت می‌شوند.
            logger.warning(
                "Image edit failed for %s: %s",
                name,
                str(exc)[:500],
            )
            errors.append(f"{name}: {str(exc)[:200]}")

    return None, (
        "ویرایش تصویر با هیچ‌یک از سرویس‌های فعال موفق نشد. "
        + " | ".join(errors)
    )


def generate_image(prompt):
    """
    ساخت تصویر از متن با Cloudflare Workers AI.
    این تابع برای ساخت تصویر است، نه ویرایش تصویر ورودی.
    """
    if not CF_ACCOUNT_ID or not CF_API_TOKEN:
        return None, (
            "ساخت تصویر فعال نیست؛ متغیرهای CF_ACCOUNT_ID "
            "و CF_API_TOKEN را در Railway تنظیم کن."
        )

    url = (
        "https://api.cloudflare.com/client/v4/accounts/"
        f"{CF_ACCOUNT_ID}/ai/run/{CF_IMAGE_MODEL}"
    )

    try:
        response = requests.post(
            url,
            headers={
                "Authorization": f"Bearer {CF_API_TOKEN}",
                "Content-Type": "application/json",
            },
            json={"prompt": prompt},
            timeout=IMAGE_TIMEOUT,
        )

        if not response.ok:
            return None, (
                "خطا در ساخت تصویر: "
                + _extract_api_error(response)
            )

        data = response.json()
        result = data.get("result") or {}
        b64 = result.get("image")

        if not data.get("success") or not b64:
            return None, "سرویس ساخت تصویر خروجی معتبری نداد."

        raw = base64.b64decode(b64, validate=True)

        with Image.open(io.BytesIO(raw)) as img:
            img.verify()

        return _image_data_url(raw), None

    except Exception as exc:
        logger.exception("Image generation failed")
        return None, f"خطا در ساخت تصویر: {str(exc)[:300]}"


def prepare_chat(chat_id, user_chats, user_message):
    if not chat_id or chat_id not in user_chats:
        chat_id = str(uuid.uuid4())
        user_chats[chat_id] = {
            "title": "مکالمه جدید",
            "messages": [],
            "created": time.time(),
        }
    chat_obj = user_chats[chat_id]
    if len(chat_obj["messages"]) == 0:
        chat_obj["title"] = make_title(user_message or "تصویر")
    return chat_id, chat_obj


@app.before_request
def ensure_user_id():
    user_id = request.cookies.get(COOKIE_NAME)
    if not user_id:
        user_id = str(uuid.uuid4())
        g.new_user_id = user_id
    g.user_id = user_id


@app.after_request
def set_user_cookie(response):
    if hasattr(g, "new_user_id"):
        response.set_cookie(
            COOKIE_NAME,
            g.new_user_id,
            max_age=COOKIE_MAX_AGE,
            httponly=True,
            samesite="Lax",
        )
    return response


@app.route("/")
def index():
    return render_template("index.html")


@app.route("/usage", methods=["GET"])
def usage_status():
    user_entry = get_user_entry()
    today = datetime.date.today().isoformat()
    usage = user_entry.get("usage", {"date": today, "count": 0})
    if usage["date"] != today:
        usage = {"date": today, "count": 0}
    provider_id = user_entry.get("provider", DEFAULT_PROVIDER)
    if provider_id not in PROVIDERS or not PROVIDERS[provider_id]["api_key"]:
        provider_id = DEFAULT_PROVIDER
    return jsonify({
        "used": usage["count"],
        "limit": MAX_DAILY_MESSAGES,
        "provider": provider_id,
        "provider_label": PROVIDERS[provider_id]["label"],
    })


@app.route("/settings", methods=["GET"])
def get_settings():
    user_entry = get_user_entry()
    current = user_entry.get("provider", DEFAULT_PROVIDER)
    if current not in PROVIDERS or not PROVIDERS[current]["api_key"]:
        current = DEFAULT_PROVIDER
    return jsonify({
        "provider": current,
        "providers": [
            {
                "id": pid,
                "label": p["label"],
                "supports_vision": p["supports_vision"],
                "available": bool(p["api_key"]),
            }
            for pid, p in PROVIDERS.items()
        ],
    })


@app.route("/settings", methods=["POST"])
def update_settings():
    user_entry = get_user_entry()
    body = request.json or {}

    provider_id = body.get("provider")
    if provider_id:
        if provider_id not in PROVIDERS:
            return jsonify({"error": "سرویس‌دهنده‌ی نامعتبر"}), 400
        if not PROVIDERS[provider_id]["api_key"]:
            return jsonify({"error": "کلید این مدل هنوز روی سرور تنظیم نشده"}), 400
        user_entry["provider"] = provider_id

    save_data(data_store)
    return jsonify({"status": "ok"})


@app.route("/upload-document", methods=["POST"])
def upload_document():
    if "file" not in request.files:
        return jsonify({"error": "فایلی ارسال نشد"}), 400

    file = request.files["file"]
    if file.filename == "":
        return jsonify({"error": "فایلی انتخاب نشد"}), 400

    file_bytes = file.read()
    if len(file_bytes) > MAX_UPLOAD_SIZE:
        return jsonify({"error": "حجم فایل بیشتر از ۸ مگابایته"}), 400

    text, note = extract_text_from_file(file.filename, file_bytes)
    if text is None:
        return jsonify({"error": note}), 400

    return jsonify({
        "filename": file.filename,
        "text": text,
        "truncated": note == "truncated",
    })


@app.route("/memory", methods=["GET"])
def get_memory():
    return jsonify({"notes": get_user_memory()})


@app.route("/memory", methods=["POST"])
def add_memory_manual():
    note = (request.json or {}).get("text", "").strip()
    if not note:
        return jsonify({"error": "متن نمی‌تونه خالی باشه"}), 400
    add_memory_note(note)
    return jsonify({"status": "ok", "notes": get_user_memory()})


@app.route("/memory/<int:index>", methods=["DELETE"])
def delete_memory(index):
    notes = get_user_memory()
    if 0 <= index < len(notes):
        del notes[index]
        save_data(data_store)
    return jsonify({"notes": get_user_memory()})


@app.route("/chats", methods=["GET"])
def list_chats():
    user_chats = get_user_chats()
    chats = [
        {"id": cid, "title": c["title"], "created": c["created"]}
        for cid, c in user_chats.items()
    ]
    chats.sort(key=lambda c: c["created"], reverse=True)
    return jsonify({"chats": chats})


@app.route("/chats/search", methods=["GET"])
def search_chats():
    query = request.args.get("q", "").strip()
    user_chats = get_user_chats()
    if not query:
        return jsonify({"chats": []})

    results = []
    for cid, c in user_chats.items():
        found = query in c["title"]
        if not found:
            for m in c["messages"]:
                content = m["content"]
                if isinstance(content, str):
                    text = content
                else:
                    text = " ".join(p.get("text", "") for p in content if p.get("type") == "text")
                if query in text:
                    found = True
                    break
        if found:
            results.append({"id": cid, "title": c["title"], "created": c["created"]})

    results.sort(key=lambda c: c["created"], reverse=True)
    return jsonify({"chats": results})


@app.route("/chats/new", methods=["POST"])
def new_chat():
    user_chats = get_user_chats()
    chat_id = str(uuid.uuid4())
    user_chats[chat_id] = {
        "title": "مکالمه جدید",
        "messages": [],
        "created": time.time(),
    }
    save_data(data_store)
    return jsonify({"chat_id": chat_id})


@app.route("/chats/<chat_id>", methods=["GET"])
def get_chat(chat_id):
    user_chats = get_user_chats()
    chat = user_chats.get(chat_id)
    if not chat:
        return jsonify({"error": "مکالمه پیدا نشد"}), 404
    return jsonify({"messages": chat["messages"]})


@app.route("/chats/<chat_id>", methods=["DELETE"])
def delete_chat(chat_id):
    user_chats = get_user_chats()
    if chat_id in user_chats:
        del user_chats[chat_id]
        save_data(data_store)
    return jsonify({"status": "حذف شد"})


@app.route("/chats/<chat_id>/rename", methods=["POST"])
def rename_chat(chat_id):
    user_chats = get_user_chats()
    new_title = (request.json or {}).get("title", "").strip()
    if not new_title:
        return jsonify({"error": "اسم نمی‌تونه خالی باشه"}), 400
    if chat_id not in user_chats:
        return jsonify({"error": "مکالمه پیدا نشد"}), 404
    user_chats[chat_id]["title"] = new_title[:50]
    save_data(data_store)
    return jsonify({"status": "ok", "title": user_chats[chat_id]["title"]})


@app.route("/chat", methods=["POST"])
def chat():
    """نسخه غیر-استریم (پشتیبان/جایگزین)"""
    user_entry = get_user_entry()
    user_chats = user_entry["chats"]
    chat_id = request.json.get("chat_id")
    user_message = request.json.get("message", "").strip()
    image_data = request.json.get("image")

    if not user_message and not image_data:
        return jsonify({"error": "پیام خالی است"}), 400

    provider_key, provider_url, provider_model, supports_vision, provider_err = get_provider_config(user_entry)
    if provider_err:
        return jsonify({"error": provider_err}), 400
    if image_data and not supports_vision:
        return jsonify({"error": "مدل انتخابی‌ات فقط متنی هست و عکس رو نمی‌فهمه. از تنظیمات مدل رو عوض کن یا عکس رو حذف کن."}), 400

    allowed, err_msg = check_and_increment_usage()
    if not allowed:
        return jsonify({"error": err_msg}), 429

    memory_note = check_memory_trigger(user_message)
    if memory_note:
        add_memory_note(memory_note)

    chat_id, chat_obj = prepare_chat(chat_id, user_chats, user_message)
    content = build_content(user_message, image_data)
    chat_obj["messages"].append({"role": "user", "content": content})
    save_data(data_store)

    try:
        headers = {"Authorization": f"Bearer {provider_key}", "content-type": "application/json"}
        safe_messages = sanitize_messages_for_provider(chat_obj["messages"], supports_vision)
        payload = {
            "model": provider_model,
            "messages": [{"role": "system", "content": SYSTEM_PROMPT + build_memory_context()}] + safe_messages,
        }
        res = requests.post(provider_url, headers=headers, json=payload, timeout=60)
        response_data = res.json()

        if res.status_code != 200:
            error_msg = response_data.get("error", {}).get("message", str(response_data))
            return jsonify({"error": f"خطا در ارتباط با API: {error_msg}"}), 500

        assistant_reply = response_data["choices"][0]["message"]["content"]
        chat_obj["messages"].append({"role": "assistant", "content": assistant_reply})
        save_data(data_store)

        return jsonify({"reply": assistant_reply, "chat_id": chat_id})

    except Exception as e:
        return jsonify({"error": f"خطا در ارتباط با API: {str(e)}"}), 500


@app.route("/chat/stream", methods=["POST"])
def chat_stream():
    """نسخه استریم - جواب رو به‌صورت تکه‌تکه (تایپ‌شونده) برمی‌گردونه"""
    user_entry = get_user_entry()
    user_chats = user_entry["chats"]
    chat_id = request.json.get("chat_id")
    user_message = request.json.get("message", "").strip()
    image_data = request.json.get("image")

    if not user_message and not image_data:
        return jsonify({"error": "پیام خالی است"}), 400

    # اگه درخواست ساخت عکس بود (و خودش عکسی آپلود نکرده)، مسیر جدا از چت متنی می‌ره
    image_gen_requested = (not image_data) and detect_image_generation_request(user_message)
    # اگه یه عکس آپلود کرده و خواسته تغییرش بدیم (نه فقط توضیحش بدیم)
    image_edit_requested = bool(image_data) and detect_image_edit_request(user_message)

    if not image_gen_requested and not image_edit_requested:
        provider_key, provider_url, provider_model, supports_vision, provider_err = get_provider_config(user_entry)
        if provider_err:
            return jsonify({"error": provider_err}), 400
        if image_data and not supports_vision:
            return jsonify({"error": "مدل انتخابی‌ات فقط متنی هست و عکس رو نمی‌فهمه. از تنظیمات مدل رو عوض کن یا عکس رو حذف کن."}), 400

    allowed, err_msg = check_and_increment_usage()
    if not allowed:
        return jsonify({"error": err_msg}), 429

    memory_note = check_memory_trigger(user_message)
    if memory_note:
        add_memory_note(memory_note)

    chat_id, chat_obj = prepare_chat(chat_id, user_chats, user_message)
    content = build_content(user_message, image_data)
    chat_obj["messages"].append({"role": "user", "content": content})
    save_data(data_store)

    if image_gen_requested:
        def generate_image_stream():
            english_prompt = translate_to_image_prompt(user_message)
            img_url, err = generate_image(english_prompt)
            if err:
                yield f"data: {json.dumps({'error': err})}\n\n"
            else:
                assistant_content = [{"type": "image_url", "image_url": {"url": img_url}}]
                chat_obj["messages"].append({"role": "assistant", "content": assistant_content})
                save_data(data_store)
                yield f"data: {json.dumps({'image': img_url})}\n\n"
            yield f"data: {json.dumps({'done': True, 'chat_id': chat_id})}\n\n"

        return Response(generate_image_stream(), mimetype="text/event-stream")

    if image_edit_requested:
        def edit_image_stream():
            english_prompt = translate_to_image_prompt(user_message)
            img_url, err = edit_image(image_data, english_prompt)
            if err:
                yield f"data: {json.dumps({'error': err})}\n\n"
            else:
                assistant_content = [{"type": "image_url", "image_url": {"url": img_url}}]
                chat_obj["messages"].append({"role": "assistant", "content": assistant_content})
                save_data(data_store)
                yield f"data: {json.dumps({'image': img_url})}\n\n"
            yield f"data: {json.dumps({'done': True, 'chat_id': chat_id})}\n\n"

        return Response(edit_image_stream(), mimetype="text/event-stream")

    system_content = SYSTEM_PROMPT + build_memory_context()

    def generate():
        full_reply = ""
        try:
            headers = {"Authorization": f"Bearer {provider_key}", "content-type": "application/json"}
            safe_messages = sanitize_messages_for_provider(chat_obj["messages"], supports_vision)
            payload = {
                "model": provider_model,
                "messages": [{"role": "system", "content": system_content}] + safe_messages,
                "stream": True,
            }
            with requests.post(provider_url, headers=headers, json=payload, timeout=120, stream=True) as res:
                if res.status_code != 200:
                    try:
                        err_body = res.json()
                        err_detail = err_body.get("error", {})
                        if isinstance(err_detail, dict):
                            err_detail = err_detail.get("message", str(err_body))
                    except Exception:
                        err_detail = res.text[:300]
                    yield f"data: {json.dumps({'error': f'کد {res.status_code}: {err_detail}'})}\n\n"
                    return
                for line in res.iter_lines():
                    if not line:
                        continue
                    decoded = line.decode("utf-8", errors="ignore")
                    if not decoded.startswith("data: "):
                        continue
                    data_str = decoded[6:]
                    if data_str.strip() == "[DONE]":
                        break
                    try:
                        chunk = json.loads(data_str)
                        delta = chunk["choices"][0]["delta"].get("content", "")
                        if delta:
                            full_reply += delta
                            yield f"data: {json.dumps({'delta': delta})}\n\n"
                    except Exception:
                        continue
        except Exception as e:
            yield f"data: {json.dumps({'error': str(e)})}\n\n"
        finally:
            if full_reply:
                chat_obj["messages"].append({"role": "assistant", "content": full_reply})
                save_data(data_store)
            yield f"data: {json.dumps({'done': True, 'chat_id': chat_id})}\n\n"

    return Response(generate(), mimetype="text/event-stream")


# ==================== پنل مدیریت ====================

ADMIN_LOGIN_HTML = """
<!DOCTYPE html>
<html lang="fa" dir="rtl">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>ورود مدیر</title>
<style>
    body { font-family: Tahoma, sans-serif; background: #14141a; color: #eee;
           display: flex; align-items: center; justify-content: center; height: 100vh; margin: 0; }
    .box { background: #1e1e26; padding: 30px; border-radius: 14px; width: 90%; max-width: 340px; }
    h2 { margin-top: 0; color: #8a6fff; text-align: center; }
    input { width: 100%; padding: 12px; margin: 10px 0; border-radius: 8px; border: 1px solid #3a3a46;
            background: #26262f; color: #eee; font-size: 15px; box-sizing: border-box; }
    button { width: 100%; padding: 12px; border-radius: 8px; border: none; background: #6c47ff;
             color: white; font-size: 15px; cursor: pointer; }
    .error { color: #ff6b6b; text-align: center; font-size: 13px; }
</style>
</head>
<body>
<div class="box">
    <h2>🔐 ورود به پنل مدیریت</h2>
    {% if error %}<p class="error">{{ error }}</p>{% endif %}
    <form method="POST">
        <input type="password" name="password" placeholder="رمز عبور" autofocus>
        <button type="submit">ورود</button>
    </form>
</div>
</body>
</html>
"""

ADMIN_PANEL_HTML = """
<!DOCTYPE html>
<html lang="fa" dir="rtl">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>پنل مدیریت</title>
<style>
    * { box-sizing: border-box; }
    body { font-family: Tahoma, sans-serif; background: #14141a; color: #eee; margin: 0; padding: 16px; }
    h2 { color: #8a6fff; }
    .top-bar { display: flex; justify-content: space-between; align-items: center; margin-bottom: 16px; }
    .logout-btn { background: #3a3a46; color: #eee; border: none; padding: 8px 14px; border-radius: 8px;
                  text-decoration: none; font-size: 13px; }
    .stats-grid { display: grid; grid-template-columns: repeat(auto-fit, minmax(140px, 1fr)); gap: 10px; margin-bottom: 20px; }
    .stat-card { background: #1e1e26; padding: 16px; border-radius: 12px; text-align: center; }
    .stat-card .num { font-size: 26px; font-weight: bold; color: #8a6fff; }
    .stat-card .label { font-size: 12px; color: #999; margin-top: 4px; }
    table { width: 100%; border-collapse: collapse; background: #1e1e26; border-radius: 12px; overflow: hidden; font-size: 13px; }
    th, td { padding: 10px; text-align: right; border-bottom: 1px solid #2a2a34; }
    th { background: #26262f; color: #8a6fff; }
    .del-btn { background: #c0392b; color: white; border: none; padding: 5px 10px; border-radius: 6px;
               font-size: 12px; cursor: pointer; }
    .scroll-x { overflow-x: auto; }
</style>
</head>
<body>
<div class="top-bar">
    <h2>📊 پنل مدیریت چت‌بات</h2>
    <a class="logout-btn" href="{{ url_for('admin_logout') }}">خروج</a>
</div>

<div class="stats-grid">
    <div class="stat-card"><div class="num">{{ total_users }}</div><div class="label">کاربران</div></div>
    <div class="stat-card"><div class="num">{{ total_chats }}</div><div class="label">مکالمات</div></div>
    <div class="stat-card"><div class="num">{{ total_messages }}</div><div class="label">کل پیام‌ها</div></div>
    <div class="stat-card"><div class="num">{{ messages_today }}</div><div class="label">پیام‌های امروز</div></div>
</div>

<div class="scroll-x">
<table>
    <tr>
        <th>شناسه کاربر</th>
        <th>تعداد مکالمات</th>
        <th>پیام امروز</th>
        <th>یادداشت حافظه</th>
        <th></th>
    </tr>
    {% for u in users %}
    <tr>
        <td>{{ u.id }}</td>
        <td>{{ u.chat_count }}</td>
        <td>{{ u.today_count }} / {{ limit }}</td>
        <td>{{ u.memory_count }}</td>
        <td>
            <form method="POST" action="{{ url_for('admin_delete_user', user_id=u.full_id) }}"
                  onsubmit="return confirm('کل داده‌های این کاربر پاک بشه؟');">
                <button class="del-btn" type="submit">حذف</button>
            </form>
        </td>
    </tr>
    {% endfor %}
</table>
</div>
</body>
</html>
"""


def admin_required(view_func):
    @wraps(view_func)
    def wrapper(*args, **kwargs):
        if not session.get("is_admin"):
            return redirect(url_for("admin_login"))
        return view_func(*args, **kwargs)
    return wrapper


@app.route("/admin/login", methods=["GET", "POST"])
def admin_login():
    error = None
    if request.method == "POST":
        if request.form.get("password") == ADMIN_PASSWORD:
            session["is_admin"] = True
            return redirect(url_for("admin_panel"))
        error = "رمز عبور اشتباهه"
    return render_template_string(ADMIN_LOGIN_HTML, error=error)


@app.route("/admin/logout")
def admin_logout():
    session.pop("is_admin", None)
    return redirect(url_for("admin_login"))


@app.route("/admin")
@admin_required
def admin_panel():
    today = datetime.date.today().isoformat()
    users_data = data_store.get("users", {})

    total_users = len(users_data)
    total_chats = 0
    total_messages = 0
    messages_today = 0
    users_list = []

    for uid, entry in users_data.items():
        chats = entry.get("chats", {})
        chat_count = len(chats)
        msg_count = sum(len(c.get("messages", [])) for c in chats.values())
        usage = entry.get("usage", {"date": today, "count": 0})
        today_count = usage["count"] if usage.get("date") == today else 0
        memory_count = len(entry.get("memory", []))

        total_chats += chat_count
        total_messages += msg_count
        messages_today += today_count

        users_list.append({
            "id": uid[:8] + "...",
            "full_id": uid,
            "chat_count": chat_count,
            "today_count": today_count,
            "memory_count": memory_count,
        })

    users_list.sort(key=lambda u: u["today_count"], reverse=True)

    return render_template_string(
        ADMIN_PANEL_HTML,
        total_users=total_users,
        total_chats=total_chats,
        total_messages=total_messages,
        messages_today=messages_today,
        users=users_list,
        limit=MAX_DAILY_MESSAGES,
    )


@app.route("/admin/delete-user/<user_id>", methods=["POST"])
@admin_required
def admin_delete_user(user_id):
    if user_id in data_store.get("users", {}):
        del data_store["users"][user_id]
        save_data(data_store)
    return redirect(url_for("admin_panel"))


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(debug=True, host="0.0.0.0", port=port)
