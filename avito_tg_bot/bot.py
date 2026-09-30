"""Пересылка входящих сообщений из мессенджера Авито в Telegram.

Бот периодически опрашивает Avito Messenger API, находит новые входящие
сообщения и отправляет их в указанный чат Telegram.
"""

import json
import logging
import os
import sys
import time
from html import escape
from pathlib import Path

import requests
from dotenv import load_dotenv

load_dotenv()

AVITO_API = "https://api.avito.ru"
TELEGRAM_API = "https://api.telegram.org"

log = logging.getLogger("avito_tg_bot")


def env(name, default=None, required=True):
    value = os.getenv(name, default)
    if required and not value:
        sys.exit(f"Не задана переменная окружения {name} (см. .env.example)")
    return value


class AvitoClient:
    def __init__(self, client_id, client_secret):
        self.client_id = client_id
        self.client_secret = client_secret
        self.session = requests.Session()
        self._token = None
        self._token_expires = 0
        self.user_id = None

    def _auth(self):
        resp = self.session.post(
            f"{AVITO_API}/token",
            data={
                "grant_type": "client_credentials",
                "client_id": self.client_id,
                "client_secret": self.client_secret,
            },
            timeout=30,
        )
        resp.raise_for_status()
        data = resp.json()
        if "access_token" not in data:
            raise RuntimeError(f"Авито не выдал токен: {data}")
        self._token = data["access_token"]
        # Обновляем токен заранее, за минуту до истечения.
        self._token_expires = time.time() + int(data.get("expires_in", 3600)) - 60

    def _request(self, method, path, _retry=True, **kwargs):
        if not self._token or time.time() >= self._token_expires:
            self._auth()
        resp = self.session.request(
            method,
            f"{AVITO_API}{path}",
            headers={"Authorization": f"Bearer {self._token}"},
            timeout=30,
            **kwargs,
        )
        if resp.status_code == 401 and _retry:
            self._auth()
            return self._request(method, path, _retry=False, **kwargs)
        resp.raise_for_status()
        return resp.json() if resp.content else {}

    def load_user_id(self):
        self.user_id = self._request("GET", "/core/v1/accounts/self")["id"]
        return self.user_id

    def unread_chats(self):
        data = self._request(
            "GET",
            f"/messenger/v2/accounts/{self.user_id}/chats",
            params={"unread_only": "true", "limit": 100},
        )
        return data.get("chats", [])

    def messages(self, chat_id, limit=20):
        data = self._request(
            "GET",
            f"/messenger/v3/accounts/{self.user_id}/chats/{chat_id}/messages/",
            params={"limit": limit},
        )
        # v3 возвращает список, v2 — объект {"messages": [...]}.
        return data if isinstance(data, list) else data.get("messages", [])

    def mark_read(self, chat_id):
        self._request("POST", f"/messenger/v1/accounts/{self.user_id}/chats/{chat_id}/read")


class Telegram:
    def __init__(self, token, chat_id):
        self.url = f"{TELEGRAM_API}/bot{token}"
        self.chat_id = chat_id

    def send(self, text):
        for _ in range(3):
            resp = requests.post(
                f"{self.url}/sendMessage",
                json={
                    "chat_id": self.chat_id,
                    "text": text[:4096],
                    "parse_mode": "HTML",
                    "disable_web_page_preview": True,
                },
                timeout=30,
            )
            if resp.status_code == 429:
                time.sleep(resp.json().get("parameters", {}).get("retry_after", 5))
                continue
            resp.raise_for_status()
            return
        raise RuntimeError("Telegram не принял сообщение после нескольких попыток")


class State:
    """Хранит id уже пересланных сообщений, чтобы не слать их повторно."""

    MAX_IDS = 5000

    def __init__(self, path):
        self.path = Path(path)
        try:
            self.ids = json.loads(self.path.read_text())
        except (FileNotFoundError, json.JSONDecodeError):
            self.ids = []
        self._set = set(self.ids)

    def __contains__(self, msg_id):
        return msg_id in self._set

    def add(self, msg_id):
        self.ids.append(msg_id)
        self._set.add(msg_id)

    def save(self):
        self.ids = self.ids[-self.MAX_IDS:]
        self._set = set(self.ids)
        tmp = self.path.with_suffix(".tmp")
        tmp.write_text(json.dumps(self.ids))
        tmp.replace(self.path)


def describe_content(msg):
    content = msg.get("content") or {}
    kind = msg.get("type")
    if kind == "text":
        return escape(content.get("text", ""))
    if kind == "image":
        sizes = (content.get("image") or {}).get("sizes") or {}
        url = sizes.get("1280x960") or next(iter(sizes.values()), None)
        return f'🖼 <a href="{escape(url)}">Изображение</a>' if url else "🖼 Изображение"
    if kind == "link":
        link = content.get("link") or {}
        return f"🔗 {escape(link.get('text') or link.get('url', ''))}"
    if kind == "location":
        loc = content.get("location") or {}
        return f"📍 {escape(loc.get('text') or loc.get('title') or 'Геопозиция')}"
    if kind == "voice":
        return "🎤 Голосовое сообщение (откройте в приложении Авито)"
    if kind == "item":
        item = content.get("item") or {}
        return f"📦 Объявление: {escape(item.get('title', ''))}"
    return f"[{escape(str(kind))}] {escape(json.dumps(content, ensure_ascii=False))}"


def format_message(chat, msg, my_id):
    sender = next(
        (u.get("name") for u in chat.get("users", []) if u.get("id") != my_id),
        "Покупатель",
    )
    item = (chat.get("context") or {}).get("value") or {}
    lines = [f"💬 <b>{escape(sender or 'Без имени')}</b>"]
    if item.get("title"):
        title = escape(item["title"])
        if item.get("price_string"):
            title += f" — {escape(item['price_string'])}"
        if item.get("url"):
            title = f'<a href="{escape(item["url"])}">{title}</a>'
        lines.append(f"📦 {title}")
    lines += ["", describe_content(msg)]
    lines.append(f'\n<a href="https://www.avito.ru/profile/messenger/channel/{chat["id"]}">Ответить на Авито</a>')
    return "\n".join(lines)


def poll_once(avito, tg, state, mark_read):
    sent = 0
    for chat in avito.unread_chats():
        new = [
            m for m in avito.messages(chat["id"])
            if m.get("direction") == "in" and m.get("id") not in state
        ]
        for msg in sorted(new, key=lambda m: m.get("created", 0)):
            tg.send(format_message(chat, msg, avito.user_id))
            state.add(msg["id"])
            state.save()
            sent += 1
        if new and mark_read:
            avito.mark_read(chat["id"])
    return sent


def main():
    logging.basicConfig(
        level=os.getenv("LOG_LEVEL", "INFO"),
        format="%(asctime)s %(levelname)s %(message)s",
    )
    avito = AvitoClient(env("AVITO_CLIENT_ID"), env("AVITO_CLIENT_SECRET"))
    tg = Telegram(env("TELEGRAM_BOT_TOKEN"), env("TELEGRAM_CHAT_ID"))
    state = State(env("STATE_FILE", "state.json"))
    interval = int(env("POLL_INTERVAL", "30"))
    mark_read = env("MARK_AS_READ", "false").lower() in ("1", "true", "yes")

    log.info("Авторизация в Авито...")
    log.info("Аккаунт Авито: %s", avito.load_user_id())
    tg.send("✅ Бот запущен: новые сообщения из Авито будут приходить сюда.")

    while True:
        try:
            sent = poll_once(avito, tg, state, mark_read)
            if sent:
                log.info("Переслано сообщений: %d", sent)
        except requests.RequestException as e:
            log.warning("Ошибка сети/API: %s", e)
        except Exception:
            log.exception("Непредвиденная ошибка")
        time.sleep(interval)


if __name__ == "__main__":
    main()
