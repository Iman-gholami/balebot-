import asyncio
import csv
import fcntl
import json
import logging
import os
import re
import zipfile
from datetime import datetime, timedelta
from io import BytesIO, StringIO
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
MAX_WAIT_MINUTES = int(os.environ.get("MAX_WAIT_MINUTES", "360"))
HISTORY_LIMIT = int(os.environ.get("HISTORY_LIMIT", "100"))
STATE_FILE = Path(os.environ.get("STATE_FILE", "processed_files.json"))
LOCK_FILE = Path(os.environ.get("LOCK_FILE", ".balebot.lock"))

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


def load_state() -> dict:
    default = {"days": {}}
    if not STATE_FILE.exists():
        return default

    try:
        data = json.loads(STATE_FILE.read_text(encoding="utf-8"))

        # Backward compatibility with the old list-based state format.
        if isinstance(data, list):
            return default

        if not isinstance(data, dict):
            return default

        days = data.get("days")
        if not isinstance(days, dict):
            data["days"] = {}

        return data
    except Exception:
        log.exception("Could not read state file; starting empty.")
        return default


def save_state(state: dict):
    temp = STATE_FILE.with_suffix(STATE_FILE.suffix + ".tmp")
    temp.write_text(
        json.dumps(state, ensure_ascii=False, indent=2, sort_keys=True),
        encoding="utf-8",
    )
    temp.replace(STATE_FILE)


def previous_jalali_code() -> str:
    local_date = datetime.now(ZoneInfo(TIMEZONE)).date() - timedelta(days=1)
    jalali = jdatetime.date.fromgregorian(date=local_date)
    return f"{jalali.year:04d}{jalali.month:02d}{jalali.day:02d}"


def same_ten_ips(today_rows, yesterday_ips) -> bool:
    if len(today_rows) != 10 or not isinstance(yesterday_ips, list) or len(yesterday_ips) != 10:
        return False

    today_ips = [ip for ip, _country in today_rows]
    # Order does not matter, but duplicates do.
    return sorted(today_ips) == sorted(str(ip) for ip in yesterday_ips)


def acquire_process_lock():
    LOCK_FILE.parent.mkdir(parents=True, exist_ok=True)
    lock_handle = LOCK_FILE.open("w")
    try:
        fcntl.flock(lock_handle.fileno(), fcntl.LOCK_EX | fcntl.LOCK_NB)
    except BlockingIOError:
        lock_handle.close()
        return None
    lock_handle.write(str(os.getpid()))
    lock_handle.flush()
    return lock_handle


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


def _extract_rows_from_xlsx(excel_bytes: bytes):
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


def _extract_rows_from_csv(csv_bytes: bytes):
    text = None
    for encoding in ("utf-8-sig", "utf-8", "cp1256", "latin-1"):
        try:
            text = csv_bytes.decode(encoding)
            break
        except UnicodeDecodeError:
            continue

    if text is None:
        raise RuntimeError("Could not decode CSV file.")

    sample = text[:4096]
    try:
        dialect = csv.Sniffer().sniff(sample, delimiters=",;\t|")
    except csv.Error:
        dialect = csv.excel

    reader = csv.reader(StringIO(text), dialect)
    rows = list(reader)
    if not rows:
        raise RuntimeError("CSV file is empty.")

    headers = [normalize_header(v) for v in rows[0]]
    ip_col = next((i for i, h in enumerate(headers) if h in IP_HEADER_NAMES), None)
    country_col = next(
        (i for i, h in enumerate(headers) if h in COUNTRY_HEADER_NAMES), None
    )

    data_rows = rows[1:]
    if ip_col is None:
        ip_col = 0
        country_col = 1 if len(rows[0]) > 1 else None
        data_rows = rows

    result = []
    for row in data_rows:
        if ip_col >= len(row):
            continue
        ip = row[ip_col]
        if not is_valid_ipv4(ip):
            continue

        country = ""
        if country_col is not None and country_col < len(row):
            country = str(row[country_col]).strip()

        result.append((str(ip).strip(), country))
        if len(result) == 10:
            break

    return result


def extract_first_ten(file_bytes: bytes):
    """
    Accept either a direct .xlsx file or a ZIP/archive containing an .xlsx file.
    Bale news channels may publish a container named NEW_FIRE_..._IPS with the
    actual Excel workbook inside it.
    """
    stream = BytesIO(file_bytes)

    if not zipfile.is_zipfile(stream):
        raise RuntimeError(
            "Downloaded file is neither an XLSX/ZIP file nor a supported archive."
        )

    stream.seek(0)
    with zipfile.ZipFile(stream) as archive:
        names = archive.namelist()

        # A real XLSX is itself a ZIP and contains [Content_Types].xml.
        if "[Content_Types].xml" in names and "xl/workbook.xml" in names:
            return _extract_rows_from_xlsx(file_bytes)

        # Otherwise treat it as an outer archive and find the workbook inside.
        candidates = [
            name for name in names
            if not name.endswith("/")
            and name.lower().endswith((".xlsx", ".csv"))
            and not name.startswith("__MACOSX/")
        ]

        if not candidates:
            visible = ", ".join(names[:20])
            raise RuntimeError(
                "Archive opened successfully, but no .xlsx or .csv file was found inside. "
                f"First archive entries: {visible}"
            )

        candidates.sort(
            key=lambda name: (
                0 if "new_fire" in name.lower() else 1,
                0 if name.lower().endswith(".csv") else 1,
                len(name),
            )
        )
        data_name = candidates[0]
        log.info("Data file found inside archive: %s", data_name)

        info = archive.getinfo(data_name)
        if info.file_size > 100 * 1024 * 1024:
            raise RuntimeError("Data file inside archive is unexpectedly large (>100 MB).")

        data_bytes = archive.read(data_name)

    if data_name.lower().endswith(".csv"):
        return _extract_rows_from_csv(data_bytes)
    return _extract_rows_from_xlsx(data_bytes)


def build_message(file_name: str, rows) -> str:
    lines = [
        "🔔 معرفی جهت مسدودسازی",
        "📋 فهرست آدرس‌های پویشگر خارجی:",
        "",
    ]
    for index, (ip, country) in enumerate(rows, 1):
        suffix = f" - {country}" if country else ""
        lines.append(f"{index}. {ip}{suffix}")
    return "\n".join(lines)


def build_unchanged_message() -> str:
    return (
        "🔔 معرفی جهت مسدودسازی\n"
        "📋 هر ۱۰ آدرس IP امروز دقیقاً با روز قبل یکسان است."
    )


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


async def process_once(client: BaleClient, source_ref, destination_ref, state: dict) -> bool:
    today_code = today_jalali_code()
    days = state.setdefault("days", {})

    # Idempotency by calendar day: never send twice for the same date.
    if today_code in days and days[today_code].get("sent"):
        log.info("Already processed today: %s", today_code)
        return True

    message, media = await find_today_file(client, source_ref)
    if not media:
        log.info("Today's file not found yet: %s", expected_filename())
        return False

    log.info("Found today's file: %s", media.name)
    file_bytes = await download_media(client, media)
    rows = extract_first_ten(file_bytes)

    if len(rows) < 10:
        raise RuntimeError(
            f"Only {len(rows)} valid IPv4 rows found in {media.name}; expected at least 10"
        )

    yesterday_code = previous_jalali_code()
    yesterday = days.get(yesterday_code, {})
    yesterday_ips = yesterday.get("ips") if isinstance(yesterday, dict) else None
    unchanged = same_ten_ips(rows, yesterday_ips)

    if unchanged:
        sent = await client.send_message(
            destination_ref,
            build_unchanged_message(),
        )
        log.info("Message send returned: %r", sent)
        log.info("All 10 IPs are identical to previous day (%s).", yesterday_code)
    else:
        sent = await client.send_message(
            destination_ref,
            build_message(media.name, rows),
        )
        log.info("Message send returned: %r", sent)
        if yesterday_ips is None:
            log.info("No processed file for previous calendar day (%s); sent normal list.", yesterday_code)
        else:
            log.info("Today's IPs differ from previous day; sent normal list.")

    days[today_code] = {
        "sent": True,
        "file_name": media.name,
        "message_rid": int(message.rid),
        "file_id": int(media.file_id),
        "ips": [ip for ip, _country in rows],
        "unchanged_from_previous_day": unchanged,
        "compared_with": yesterday_code if yesterday_ips is not None else None,
    }
    save_state(state)

    log.info("Daily state saved for %s.", today_code)
    return True


async def main():
    if not BALE_TOKEN:
        raise SystemExit("BALE_TOKEN is missing. Run: python login.py")
    if not SOURCE:
        raise SystemExit("BALE_SOURCE is missing. Run: python list_dialogs.py")
    if not DESTINATION:
        raise SystemExit("BALE_DESTINATION is missing. Run: python list_dialogs.py")

    lock_handle = acquire_process_lock()
    if lock_handle is None:
        log.warning("Another balebot process is already running. Exiting.")
        return

    state = load_state()

    async with BaleClient(BALE_TOKEN) as client:
        me = await client.get_me()
        log.info("Logged in as %s (user id=%s)", me.title, me.peer.id)

        source_ref = await resolve_source(client)
        source_info = await client.resolve(source_ref)

        destination_ref = ref_from_config(DESTINATION)
        destination_info = await client.resolve(destination_ref)

        log.info(
            "Source: %s (id=%s). Today's target: %s",
            source_info.title,
            source_info.peer.id,
            expected_filename(),
        )
        log.info(
            "Destination: %s (id=%s, configured=%s)",
            destination_info.title,
            destination_info.peer.id,
            DESTINATION,
        )

        deadline = asyncio.get_running_loop().time() + (MAX_WAIT_MINUTES * 60)

        while True:
            try:
                done = await process_once(client, source_ref, destination_ref, state)
                if done:
                    log.info("Daily job completed. Exiting.")
                    return
            except Exception:
                log.exception("Daily processing failed.")

            if asyncio.get_running_loop().time() >= deadline:
                log.warning(
                    "Today's file was not successfully processed within %s minutes. Exiting.",
                    MAX_WAIT_MINUTES,
                )
                return

            await asyncio.sleep(CHECK_INTERVAL)


if __name__ == "__main__":
    asyncio.run(main())
