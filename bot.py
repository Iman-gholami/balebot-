import json
import logging
import os
import re
import time
from io import BytesIO
from pathlib import Path

import openpyxl
import requests


BOT_TOKEN = os.environ.get("BALE_BOT_TOKEN", "").strip()
SOURCE_CHANNEL_ID = os.environ.get("BALE_SOURCE_CHANNEL_ID", "").strip()
DESTINATION_CHAT_ID = os.environ.get("BALE_DESTINATION_CHAT_ID", "").strip()

API_BASE = os.environ.get("BALE_API_BASE", "https://tapi.bale.ai").rstrip("/")
STATE_FILE = Path(os.environ.get("STATE_FILE", "processed_files.json"))
POLL_TIMEOUT = int(os.environ.get("POLL_TIMEOUT", "25"))

FILE_PATTERN = re.compile(
    r"^NEW_FIRE_(\d{8})_IPS(?:\.xlsx?)?$",
    re.IGNORECASE,
)

logging.basicConfig(
    level=os.environ.get("LOG_LEVEL", "INFO"),
    format="%(asctime)s | %(levelname)s | %(message)s",
)
log = logging.getLogger("bale-fire-bot")


def api_url(method: str) -> str:
    return f"{API_BASE}/bot{BOT_TOKEN}/{method}"


def bale_request(method: str, data=None, timeout=40):
    response = requests.post(api_url(method), data=data or {}, timeout=timeout)
    response.raise_for_status()
    payload = response.json()
    if not payload.get("ok"):
        raise RuntimeError(payload.get("description") or f"Bale API error in {method}")
    return payload.get("result")


def get_updates(offset: int):
    return bale_request(
        "getUpdates",
        {"offset": offset, "timeout": POLL_TIMEOUT},
        timeout=POLL_TIMEOUT + 15,
    )


def send_message(chat_id: str, text: str):
    return bale_request("sendMessage", {"chat_id": chat_id, "text": text})


def get_file_path(file_id: str) -> str:
    result = bale_request("getFile", {"file_id": file_id})
    file_path = result.get("file_path") if isinstance(result, dict) else None
    if not file_path:
        raise RuntimeError("getFile returned no file_path")
    return file_path


def download_file(file_path: str) -> bytes:
    candidates = [
        f"{API_BASE}/file/bot{BOT_TOKEN}/{file_path}",
        f"{API_BASE}/file/{BOT_TOKEN}/{file_path}",
    ]
    last_error = None
    for url in candidates:
        try:
            response = requests.get(url, timeout=60)
            if response.ok and response.content:
                return response.content
            last_error = RuntimeError(f"download failed: HTTP {response.status_code}")
        except Exception as exc:
            last_error = exc
    raise RuntimeError(f"Could not download file: {last_error}")


def load_processed() -> set[str]:
    if not STATE_FILE.exists():
        return set()
    try:
        data = json.loads(STATE_FILE.read_text(encoding="utf-8"))
        return set(data if isinstance(data, list) else [])
    except Exception:
        log.exception("Could not read state file; starting with empty state.")
        return set()


def save_processed(processed: set[str]):
    STATE_FILE.write_text(
        json.dumps(sorted(processed), ensure_ascii=False, indent=2),
        encoding="utf-8",
    )


def normalize_header(value) -> str:
    if value is None:
        return ""
    return re.sub(r"[^a-z0-9]+", "", str(value).strip().lower())


IP_HEADER_NAMES = {
    "ip", "ipaddress", "ipaddr", "address", "destinationip", "srcip", "sourceip"
}
COUNTRY_HEADER_NAMES = {
    "country", "countryname", "location", "nation", "geo", "countrycode"
}


def find_columns(ws):
    rows = ws.iter_rows(values_only=True)
    try:
        first = next(rows)
    except StopIteration:
        raise ValueError("Excel file is empty")

    headers = [normalize_header(v) for v in first]
    ip_col = next((i for i, h in enumerate(headers) if h in IP_HEADER_NAMES), None)
    country_col = next((i for i, h in enumerate(headers) if h in COUNTRY_HEADER_NAMES), None)

    if ip_col is not None:
        return rows, ip_col, country_col

    # If there is no recognizable header, treat the first row as data.
    def with_first():
        yield first
        yield from rows

    return with_first(), 0, 1 if len(first) > 1 else None


def is_valid_ip(value) -> bool:
    if value is None:
        return False
    text = str(value).strip()
    parts = text.split(".")
    if len(parts) != 4:
        return False
    try:
        nums = [int(p) for p in parts]
    except ValueError:
        return False
    return all(0 <= n <= 255 for n in nums)


def extract_first_ten(excel_bytes: bytes):
    wb = openpyxl.load_workbook(BytesIO(excel_bytes), read_only=True, data_only=True)
    ws = wb.active
    rows, ip_col, country_col = find_columns(ws)

    result = []
    for row in rows:
        if ip_col >= len(row):
            continue
        ip = row[ip_col]
        if not is_valid_ip(ip):
            continue

        country = ""
        if country_col is not None and country_col < len(row) and row[country_col] is not None:
            country = str(row[country_col]).strip()

        result.append((str(ip).strip(), country))
        if len(result) == 10:
            break

    wb.close()
    return result


def build_message(base_name: str, rows) -> str:
    lines = [base_name, ""]
    for idx, (ip, country) in enumerate(rows, 1):
        suffix = f" - {country}" if country else ""
        lines.append(f"{idx}. {ip}{suffix}")
    return "\n".join(lines)


def extract_document(update):
    # Channel posts normally arrive under channel_post. message is supported too
    # so the same bot can be tested in a private/group chat.
    message = update.get("channel_post") or update.get("message")
    if not isinstance(message, dict):
        return None, None

    chat = message.get("chat") or {}
    chat_id = str(chat.get("id", ""))
    if SOURCE_CHANNEL_ID and chat_id != SOURCE_CHANNEL_ID:
        return None, None

    document = message.get("document")
    if not isinstance(document, dict):
        return None, None

    return message, document


def handle_update(update, processed: set[str]):
    message, document = extract_document(update)
    if not document:
        return

    file_name = (document.get("file_name") or "").strip()
    match = FILE_PATTERN.match(file_name)
    if not match:
        return

    # Some clients may omit the extension in the displayed name.
    logical_name = re.sub(r"\.xlsx?$", "", file_name, flags=re.IGNORECASE)
    file_id = str(document.get("file_id") or "")
    unique_key = file_id or logical_name

    if unique_key in processed:
        log.info("Already processed: %s", file_name)
        return

    log.info("Processing %s", file_name)

    file_path = get_file_path(file_id)
    content = download_file(file_path)
    rows = extract_first_ten(content)

    if len(rows) < 10:
        raise ValueError(f"Only {len(rows)} valid IP rows found in {file_name}; expected at least 10")

    text = build_message(logical_name, rows)
    send_message(DESTINATION_CHAT_ID, text)

    processed.add(unique_key)
    save_processed(processed)
    log.info("Sent first 10 IPs from %s", file_name)


def validate_config():
    missing = []
    if not BOT_TOKEN:
        missing.append("BALE_BOT_TOKEN")
    if not DESTINATION_CHAT_ID:
        missing.append("BALE_DESTINATION_CHAT_ID")
    if missing:
        raise SystemExit("Missing required environment variables: " + ", ".join(missing))


def main():
    validate_config()
    processed = load_processed()
    offset = 0

    log.info("Bot started. Waiting for NEW_FIRE_YYYYMMDD_IPS Excel files...")

    while True:
        try:
            updates = get_updates(offset)
            for update in updates or []:
                update_id = int(update.get("update_id", 0))
                offset = max(offset, update_id + 1)
                try:
                    handle_update(update, processed)
                except Exception:
                    log.exception("Failed to process update_id=%s", update_id)
        except KeyboardInterrupt:
            log.info("Stopped.")
            return
        except Exception:
            log.exception("Polling error; retrying shortly.")
            time.sleep(5)


if __name__ == "__main__":
    main()
