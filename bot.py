import asyncio
import json
import logging
import os
import re
from datetime import datetime
from io import BytesIO
from pathlib import Path
from zoneinfo import ZoneInfo

import jdatetime
import openpyxl
import requests
from dotenv import load_dotenv

from bale import BaleClient, pb
from bale.peer import Peer


load_dotenv()

BALE_TOKEN = os.environ.get("BALE_TOKEN", "").strip()
SOURCE = os.environ.get("BALE_SOURCE", "").strip()
DESTINATION = os.environ.get("BALE_DESTINATION", "").strip()
TIMEZONE = os.environ.get("BALE_TIMEZONE", "Asia/Tehran").strip()
CHECK_INTERVAL = int(os.environ.get("CHECK_INTERVAL", "60"))
HISTORY_LIMIT = int(os.environ.get("HISTORY_LIMIT", "100"))
STATE_FILE = Path(os.environ.get("STATE_FILE", "processed_files.json"))

logging.basicConfig(
    level=os.environ.get("LOG_LEVEL", "INFO"),
    format="%(asctime)s | %(levelname)s | %(message)s",
)
log = logging.getLogger("bale-fire-userbot")


def today_jalali_code() -> str:
    local_date = datetime.now(ZoneInfo(TIMEZONE)).date()
    jalali = jdatetime.date.fromgregorian(date=local_date)
    return f"{jalali.year:04d}{jalali.month:02d}{jalali.day:02d}"


def expected_filename() -> str:
    return f"NEW_FIRE_{today_jalali_code()}_IPS"


def file_stem(name: str) -> str:
    return re.sub(r"\.[^.]+$", "", name.strip())


def ref_from_config(value: str):
    value = value.strip()
    if value.startswith("user:"):
        return Peer.user(int(value.split(":", 1)[1]))
    if value.startswith("channel:") or value.startswith("group:"):
        return Peer.channel(int(value.split(":", 1)[1]))
    return value


def load_processed() -> set[str]:
    if not STATE_FILE.exists():
        return set()
    try:
        data = json.loads(STATE_FILE.read_text(encoding="utf-8"))
        return set(data if isinstance(data, list) else [])
    except Exception:
        log.exception("Could not read state file; starting empty.")
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
    country_col = next(
        (i for i, h in enumerate(headers) if h in COUNTRY_HEADER_NAMES), None
    )

    if ip_col is not None:
        return rows, ip_col, country_col

    def with_first():
        yield first
        yield from rows

    return with_first(), 0, 1 if len(first) > 1 else None


def is_valid_ipv4(value) -> bool:
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
        if not is_valid_ipv4(ip):
            continue

        country = ""
        if (
            country_col is not None
            and country_col < len(row)
            and row[country_col] is not None
        ):
            country = str(row[country_col]).strip()

        result.append((str(ip).strip(), country))
        if len(result) == 10:
            break

    wb.close()
    return result


def build_message(file_name: str, rows) -> str:
    base = re.sub(r"\.xlsx?$", "", file_name, flags=re.IGNORECASE)
    lines = [base, ""]
    for index, (ip, country) in enumerate(rows, 1):
        suffix = f" - {country}" if country else ""
        lines.append(f"{index}. {ip}{suffix}")
    return "\n".join(lines)


async def resolve_source(client: BaleClient):
    configured = ref_from_config(SOURCE)
    if SOURCE.startswith("@") or SOURCE.startswith("channel:") or SOURCE.startswith("group:"):
        return configured

    # Convenience: allow exact channel title from dialog list.
    dialogs = await client.get_dialogs(limit=300)
    matches = [d for d in dialogs if (d.title or "").strip() == SOURCE]
    if len(matches) == 1:
        return matches[0].peer
    if not matches:
        raise RuntimeError(
            f"Source {SOURCE!r} not found. Run: python list_dialogs.py"
        )
    raise RuntimeError(
        f"More than one dialog is named {SOURCE!r}; use channel:<ID> from list_dialogs.py"
    )


async def get_download_url(client: BaleClient, media) -> tuple[str, dict]:
    req = pb.GetNasimFileUrlRequest()
    req.file.fileId = int(media.file_id)
    req.file.accessHash = int(media.access_hash)

    resp = await client.files.GetNasimFileUrl(req)
    file_url = resp.fileUrl

    url = getattr(file_url, "url", "") or ""
    headers = {}

    # Some Bale file responses use an unsigned URL plus required headers.
    try:
        if not url and file_url.HasField("unsignedUrl"):
            url = file_url.unsignedUrl.value
    except Exception:
        pass

    for item in getattr(file_url, "unsignedUrlHeaders", []) or []:
        key = getattr(item, "key", "") or ""
        value = getattr(item, "value", "") or ""
        if key:
            headers[key] = value

    if not url:
        raise RuntimeError("Bale returned an empty download URL")
    return url, headers


async def download_media(client: BaleClient, media) -> bytes:
    url, headers = await get_download_url(client, media)

    def _download():
        response = requests.get(url, headers=headers, timeout=90)
        response.raise_for_status()
        return response.content

    return await asyncio.to_thread(_download)


async def find_today_file(client: BaleClient, source_ref):
    target = expected_filename().lower()
    async for message in client.iter_messages(source_ref, limit=HISTORY_LIMIT):
        content = message.content
        media = getattr(content, "media", None)
        if not media:
            continue
        name = (media.name or "").strip()
        if file_stem(name).lower() == target:
            return message, media
    return None, None


async def process_once(client: BaleClient, source_ref, processed: set[str]) -> bool:
    message, media = await find_today_file(client, source_ref)
    if not media:
        log.info("Today's file not found yet: %s", expected_filename())
        return False

    key = f"{today_jalali_code()}:{media.file_id}:{message.rid}"
    if key in processed:
        log.info("Already processed today: %s", media.name)
        return True

    log.info("Found today's file: %s", media.name)
    excel_bytes = await download_media(client, media)
    rows = extract_first_ten(excel_bytes)

    if len(rows) < 10:
        raise RuntimeError(
            f"Only {len(rows)} valid IPv4 rows found in {media.name}; expected at least 10"
        )

    await client.send_message(ref_from_config(DESTINATION), build_message(media.name, rows))

    processed.add(key)
    save_processed(processed)
    log.info("Sent 10 IP rows to destination.")
    return True


async def main():
    if not BALE_TOKEN:
        raise SystemExit("BALE_TOKEN is missing. Run: python login.py")
    if not SOURCE:
        raise SystemExit("BALE_SOURCE is missing. Run: python list_dialogs.py")
    if not DESTINATION:
        raise SystemExit("BALE_DESTINATION is missing. Run: python list_dialogs.py")

    processed = load_processed()

    async with BaleClient(BALE_TOKEN) as client:
        me = await client.get_me()
        log.info("Logged in as %s (user id=%s)", me.title, me.peer.id)

        source_ref = await resolve_source(client)
        source_info = await client.resolve(source_ref)
        log.info(
            "Source: %s (id=%s). Today's target: %s",
            source_info.title,
            source_info.peer.id,
            expected_filename(),
        )

        while True:
            try:
                await process_once(client, source_ref, processed)
            except Exception:
                log.exception("Daily processing failed.")
            await asyncio.sleep(CHECK_INTERVAL)


if __name__ == "__main__":
    asyncio.run(main())
