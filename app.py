"""
یک چت‌بات ساده با استفاده از Flask و OpenRouter API
با قابلیت چند مکالمه جداگانه و ذخیره دائمی تاریخچه
"""

from flask import Flask, render_template, request, jsonify
import requests
import os
import json
import uuid
import time

app = Flask(__name__)

API_KEY = os.environ.get("OPENROUTER_API_KEY")

API_URL = "https://openrouter.ai/api/v1/chat/completions"
MODEL = "openrouter/free"

HISTORY_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "chat_history.json")


def load_data():
    """ساختار داده: {"chats": {chat_id: {"title": ..., "messages": [...], "created": ...}}}"""
    if os.path.exists(HISTORY_FILE):
        try:
            with open(HISTORY_FILE, "r", encoding="utf-8") as f:
                data = json.load(f)
                if "chats" in data:
                    return data
        except (json.JSONDecodeError, IOError):
            pass

    # اگه فایل قدیمی (فقط لیست پیام) بود یا فایلی نبود، یه چت جدید بساز
    return {"chats": {}}


def save_data(data):
    with open(HISTORY_FILE, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)


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
        {"id": cid, "title": c["title"], "created": c["created"]}
        for cid, c in data_store["chats"].items()
    ]
    chats.sort(key=lambda c: c["created"], reverse=True)
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
    return jsonify({"chat_id": chat_id})


@app.route("/chats/<chat_id>", methods=["GET"])
def get_chat(chat_id):
    chat = data_store["chats"].get(chat_id)
    if not chat:
        return jsonify({"error": "مکالمه پیدا نشد"}), 404
    return jsonify({"messages": chat["messages"]})


@app.route("/chats/<chat_id>", methods=["DELETE"])
def delete_chat(chat_id):
    if chat_id in data_store["chats"]:
        del data_store["chats"][chat_id]
        save_data(data_store)
    return jsonify({"status": "حذف شد"})


@app.route("/chat", methods=["POST"])
def chat():
    chat_id = request.json.get("chat_id")
    user_message = request.json.get("message", "").strip()

    if not user_message:
        return jsonify({"error": "پیام خالی است"}), 400

    if not chat_id or chat_id not in data_store["chats"]:
        # اگه چتی مشخص نشده بود، یه چت جدید بساز
        chat_id = str(uuid.uuid4())
        data_store["chats"][chat_id] = {
            "title": "مکالمه جدید",
            "messages": [],
            "created": time.time(),
        }

    chat_obj = data_store["chats"][chat_id]

    # اگه اولین پیام این مکالمه‌ست، عنوانشو از روش بساز
    if len(chat_obj["messages"]) == 0:
        chat_obj["title"] = make_title(user_message)

    chat_obj["messages"].append({"role": "user", "content": user_message})
    save_data(data_store)

    try:
        headers = {
            "Authorization": f"Bearer {API_KEY}",
            "content-type": "application/json",
        }
        payload = {
            "model": MODEL,
            "messages": chat_obj["messages"],
        }

        res = requests.post(API_URL, headers=headers, json=payload, timeout=60)
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


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(debug=True, host="0.0.0.0", port=port)
