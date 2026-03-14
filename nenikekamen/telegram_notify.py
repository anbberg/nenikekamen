"""Send messages via Telegram Bot API. Used for run success/failure notifications."""

import requests

MAX_MESSAGE_LENGTH = 4096
TRUNCATE_AT = 3500


def send_telegram_message(token: str, chat_id: str, text: str) -> bool:
    """
    Send a text message to a Telegram chat via the Bot API.

    Truncates text to fit Telegram's limit. Swallows errors (network, API)
    so that a failed notification does not affect the main script.
    Returns True if the message was sent successfully, False otherwise.
    """
    if len(text) > MAX_MESSAGE_LENGTH:
        text = text[:TRUNCATE_AT].rstrip() + "\n\n[... trunkerat]"

    url = f"https://api.telegram.org/bot{token}/sendMessage"
    payload = {"chat_id": chat_id, "text": text, "parse_mode": "Markdown"}

    try:
        resp = requests.post(url, json=payload, timeout=10)
        resp.raise_for_status()
        return True
    except requests.HTTPError as e:
        detail = ""
        if e.response is not None and e.response.text:
            detail = f" — {e.response.text.strip()}"
        print(f"Telegram-notis kunde inte skickas: {e}{detail}")
        return False
    except Exception as e:
        print(f"Telegram-notis kunde inte skickas: {e}")
        return False
