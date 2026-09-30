"""Пересылка уведомлений Авито о новых сообщениях из почты в Telegram.

Подходит для обычного (частного) профиля Авито, где нет доступа к API:
Авито присылает письмо о каждом новом сообщении, бот читает эти письма
по IMAP и пересылает их в Telegram. Письма не помечаются прочитанными.

Режимы:
  python mail_bot.py          — один проход (для GitHub Actions / cron)
  python mail_bot.py --loop   — работать постоянно, проверяя почту каждые POLL_INTERVAL секунд
"""

import email
import imaplib
import logging
import os
import re
import sys
import time
from datetime import datetime, timedelta
from email.header import decode_header, make_header
from html import escape, unescape

import requests

from bot import TELEGRAM_API, State, Telegram, env

log = logging.getLogger("avito_mail_bot")

# Письма какого отправителя и с какой темой пересылать.
FROM_FILTER = os.getenv("MAIL_FROM_FILTER", "avito.ru")
SUBJECT_FILTER = [
    s.strip().lower()
    for s in os.getenv("MAIL_SUBJECT_FILTER", "сообщен,message").split(",")
    if s.strip()
]
LOOKBACK_DAYS = 3
MAX_TEXT = 1500


def decode(value):
    return str(make_header(decode_header(value or ""))).strip()


def html_to_text(html):
    html = re.sub(r"(?is)<(script|style|head).*?</\1>", " ", html)
    html = re.sub(r"(?i)<br\s*/?>|</p>|</div>|</tr>|</h\d>", "\n", html)
    return unescape(re.sub(r"<[^>]+>", " ", html))


def body_and_links(msg):
    plain, html = None, None
    for part in msg.walk():
        ctype = part.get_content_type()
        if part.get_content_maintype() == "multipart" or part.get_filename():
            continue
        payload = part.get_payload(decode=True)
        if payload is None:
            continue
        text = payload.decode(part.get_content_charset() or "utf-8", "replace")
        if ctype == "text/plain" and plain is None:
            plain = text
        elif ctype == "text/html" and html is None:
            html = text
    text = plain if plain and plain.strip() else html_to_text(html or "")
    links = re.findall(r"https?://[^\s\"'<>)]+", (html or "") + " " + (plain or ""))
    return text, [unescape(link) for link in links]


def clean_text(text):
    lines = [re.sub(r"[ \t ]+", " ", line).strip() for line in text.splitlines()]
    text = "\n".join(line for line in lines if line)
    if len(text) > MAX_TEXT:
        text = text[:MAX_TEXT].rsplit(" ", 1)[0] + "…"
    return text


def chat_link(links):
    for key in ("messenger", "channel", "/im/"):
        for link in links:
            if "avito" in link and key in link:
                return link
    return None


def is_message_notification(subject):
    if not SUBJECT_FILTER:
        return True
    subject = subject.lower()
    return any(word in subject for word in SUBJECT_FILTER)


def format_mail(subject, text, link):
    parts = [f"💬 <b>{escape(subject or 'Новое сообщение на Авито')}</b>"]
    if text:
        parts += ["", escape(text)]
    if link:
        parts.append(f'\n<a href="{escape(link)}">Открыть чат на Авито</a>')
    return "\n".join(parts)


def fetch_avito_mails(imap):
    since = (datetime.now() - timedelta(days=LOOKBACK_DAYS)).strftime("%d-%b-%Y")
    status, data = imap.uid("SEARCH", None, "FROM", f'"{FROM_FILTER}"', "SINCE", since)
    if status != "OK":
        raise RuntimeError(f"IMAP SEARCH: {status} {data}")
    for uid in data[0].split():
        # BODY.PEEK не снимает с письма отметку «непрочитано».
        status, parts = imap.uid("FETCH", uid, "(BODY.PEEK[])")
        if status != "OK" or not parts or not isinstance(parts[0], tuple):
            continue
        msg = email.message_from_bytes(parts[0][1])
        msg_id = msg.get("Message-ID") or f"uid:{uid.decode()}"
        yield msg_id.strip(), msg


def resolve_chat_id(token, state_path):
    """TELEGRAM_CHAT_ID из настроек, иначе — чат того, кто последним написал боту."""
    chat_id = os.getenv("TELEGRAM_CHAT_ID")
    if chat_id:
        return chat_id
    saved = state_path + ".chat"
    if os.path.exists(saved):
        return open(saved).read().strip()
    updates = requests.get(f"{TELEGRAM_API}/bot{token}/getUpdates", timeout=30).json()
    chats = [
        u["message"]["chat"]["id"]
        for u in updates.get("result", [])
        if u.get("message", {}).get("chat", {}).get("type") == "private"
    ]
    if not chats:
        sys.exit("Не задан TELEGRAM_CHAT_ID и никто не писал боту. Напишите боту любое сообщение.")
    with open(saved, "w") as f:
        f.write(str(chats[-1]))
    log.info("Chat id определён автоматически: %s", chats[-1])
    return str(chats[-1])


def check_mail(tg, state):
    first_run = not state.ids
    imap = imaplib.IMAP4_SSL(env("MAIL_IMAP_HOST", "imap.yandex.ru"), 993)
    try:
        imap.login(env("MAIL_USER"), env("MAIL_PASSWORD"))
        imap.select("INBOX", readonly=True)
        new = [(mid, m) for mid, m in fetch_avito_mails(imap) if mid not in state]
    finally:
        try:
            imap.logout()
        except Exception:
            pass

    if first_run:
        # Не пересылаем старые письма при первом запуске, только показываем, что нашли.
        subjects = [decode(m["Subject"]) for _, m in new]
        for mid, _ in new:
            state.add(mid)
        state.ids.append("__initialized__")
        state.save()
        lines = [f"✅ Бот подключился к почте. Писем от Авито за {LOOKBACK_DAYS} дн.: {len(subjects)}."]
        if subjects:
            lines.append("Последние темы писем:")
            lines += [f"• {escape(s)}" for s in subjects[-5:]]
        lines.append("\nНовые сообщения с Авито будут приходить сюда.")
        tg.send("\n".join(lines))
        return 0

    sent = 0
    for mid, msg in new:
        subject = decode(msg["Subject"])
        if is_message_notification(subject):
            text, links = body_and_links(msg)
            tg.send(format_mail(subject, clean_text(text), chat_link(links)))
            sent += 1
        state.add(mid)
        state.save()
    return sent


def main():
    logging.basicConfig(
        level=os.getenv("LOG_LEVEL", "INFO"),
        format="%(asctime)s %(levelname)s %(message)s",
    )
    token = env("TELEGRAM_BOT_TOKEN")
    state_path = env("STATE_FILE", "state.json")
    tg = Telegram(token, resolve_chat_id(token, state_path))
    state = State(state_path)

    if "--loop" not in sys.argv:
        log.info("Переслано писем: %d", check_mail(tg, state))
        return

    interval = int(env("POLL_INTERVAL", "60"))
    while True:
        try:
            sent = check_mail(tg, state)
            if sent:
                log.info("Переслано писем: %d", sent)
        except (imaplib.IMAP4.error, OSError, requests.RequestException) as e:
            log.warning("Ошибка почты/сети: %s", e)
        except Exception:
            log.exception("Непредвиденная ошибка")
        time.sleep(interval)


if __name__ == "__main__":
    main()
