"""
یک چت‌بات ساده با استفاده از Flask و OpenRouter API
با قابلیت چند مکالمه جداگانه، ذخیره دائمی تاریخچه و ارسال تصویر
"""

from flask import Flask, render_template, request, jsonify
import requests
import os
import json
import uuid
import time

app = Flask(__name__)

# فقط از متغیر محیطی Railway استفاده می‌شود
API_KEY = os.environ.get("OPENROUTER_API_KEY")

API_URL = "https://openrouter.ai/api/v1/chat/completions"
MODEL = "openrouter/free"

HISTORY_FILE = os.path.join(
    os.path.dirname(os.path.abspath(__file__)),
    "chat_history.json"
)


def load_data():
    """
    ساختار داده:
    {
        "chats": {
            "chat_id": {
                "title": "...",
                "messages": [...],
                "created": ...
            }
        }
    }
    """

    if os.path.exists(HISTORY_FILE):
        try:
            with open(HISTORY_FILE, "r", encoding="utf-8") as f:
                data = json.load(f)

                if "chats" in data:
                    return data

        except (json.JSONDecodeError, IOError):
            pass

    return {"chats": {}}


def save_data(data):
    with open(HISTORY_FILE, "w", encoding="utf-8") as f:
        json.dump(
            data,
            f,
            ensure_ascii=False,
            indent=2
        )


data_store = load_data()


def make_title(first_message):
    title = first_message.strip()

    if len(title) > 30:
        title = title[:30] + "..."

    return title or "مکالمه جدید"


@app.route("/")
def index():
    return render_template("index.html")


@app.route("/chats", methods=["GET"])
def list_chats():
    """لیست همه مکالمات، جدیدترین اول"""

    chats = [
        {
            "id": chat_id,
            "title": chat["title"],
            "created": chat["created"]
        }
        for chat_id, chat in data_store["chats"].items()
    ]

    chats.sort(
        key=lambda chat: chat["created"],
        reverse=True
    )

    return jsonify({"chats": chats})


@app.route("/chats/new", methods=["POST"])
def new_chat():
    chat_id = str(uuid.uuid4())

    data_store["chats"][chat_id] = {
        "title": "مکالمه جدید",
        "messages": [],
        "created": time.time(),
    }

    save_data(data_store)

    return jsonify({
        "chat_id": chat_id
    })


@app.route("/chats/<chat_id>", methods=["GET"])
def get_chat(chat_id):
    chat = data_store["chats"].get(chat_id)

    if not chat:
        return jsonify({
            "error": "مکالمه پیدا نشد"
        }), 404

    return jsonify({
        "messages": chat["messages"]
    })


@app.route("/chats/<chat_id>", methods=["DELETE"])
def delete_chat(chat_id):

    if chat_id in data_store["chats"]:
        del data_store["chats"][chat_id]
        save_data(data_store)

    return jsonify({
        "status": "حذف شد"
    })


@app.route("/chat", methods=["POST"])
def chat():

    # بررسی وجود API Key
    if not API_KEY:
        return jsonify({
            "error": "OPENROUTER_API_KEY در تنظیمات محیطی تنظیم نشده است."
        }), 500

    data = request.get_json(silent=True) or {}

    chat_id = data.get("chat_id")
    user_message = data.get("message", "").strip()

    # تصویر به صورت Data URL
    # مثال:
    # data:image/jpeg;base64,...
    image_data = data.get("image")

    # حداقل باید متن یا تصویر وجود داشته باشد
    if not user_message and not image_data:
        return jsonify({
            "error": "پیام خالی است"
        }), 400

    # اگر چت وجود نداشت، چت جدید بساز
    if not chat_id or chat_id not in data_store["chats"]:

        chat_id = str(uuid.uuid4())

        data_store["chats"][chat_id] = {
            "title": "مکالمه جدید",
            "messages": [],
            "created": time.time(),
        }

    chat_obj = data_store["chats"][chat_id]

    # اگر اولین پیام مکالمه است، عنوان بساز
    if len(chat_obj["messages"]) == 0:

        chat_obj["title"] = make_title(
            user_message or "تصویر"
        )

    # ساخت محتوای پیام
    if image_data:

        text_part = (
            user_message
            if user_message
            else "این تصویر را توضیح بده"
        )

        content = [
            {
                "type": "text",
                "text": text_part
            },
            {
                "type": "image_url",
                "image_url": {
                    "url": image_data
                }
            }
        ]

    else:

        content = user_message

    # ذخیره پیام کاربر
    chat_obj["messages"].append({
        "role": "user",
        "content": content
    })

    save_data(data_store)

    try:

        headers = {
            "Authorization": f"Bearer {API_KEY}",
            "Content-Type": "application/json"
        }

        payload = {
            "model": MODEL,
            "messages": chat_obj["messages"]
        }

        response = requests.post(
            API_URL,
            headers=headers,
            json=payload,
            timeout=60
        )

        # تلاش برای دریافت JSON
        try:
            response_data = response.json()

        except ValueError:
            return jsonify({
                "error": "پاسخ نامعتبر از OpenRouter دریافت شد."
            }), 500

        # بررسی خطای API
        if response.status_code != 200:

            error_msg = (
                response_data
                .get("error", {})
                .get("message", str(response_data))
            )

            return jsonify({
                "error": f"خطا در ارتباط با API: {error_msg}"
            }), 500

        # بررسی وجود پاسخ
        if (
            "choices" not in response_data
            or not response_data["choices"]
        ):
            return jsonify({
                "error": "پاسخ معتبری از مدل دریافت نشد."
            }), 500

        assistant_reply = (
            response_data["choices"][0]
            ["message"]
            ["content"]
        )

        # ذخیره پاسخ دستیار
        chat_obj["messages"].append({
            "role": "assistant",
            "content": assistant_reply
        })

        save_data(data_store)

        return jsonify({
            "reply": assistant_reply,
            "chat_id": chat_id
        })

    except requests.exceptions.Timeout:

        return jsonify({
            "error": "زمان اتصال به OpenRouter تمام شد."
        }), 504

    except requests.exceptions.RequestException as e:

        return jsonify({
            "error": f"خطا در ارتباط با OpenRouter: {str(e)}"
        }), 500

    except Exception as e:

        return jsonify({
            "error": f"خطای غیرمنتظره: {str(e)}"
        }), 500


if __name__ == "__main__":

    port = int(
        os.environ.get("PORT", 5000)
    )

    app.run(
        debug=True,
        host="0.0.0.0",
        port=port
    )
