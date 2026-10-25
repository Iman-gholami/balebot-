import asyncio
import csv
import fcntl
import json
import logging
import os
import re
import zipfile
from datetime import datetime, timedelta
from html.parser import HTMLParser
from io import BytesIO, StringIO
from pathlib import Path
from urllib.parse import unquote, urljoin, urlparse
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
AFTA_INDEX_URL = os.environ.get("AFTA_INDEX_URL", "https://afta.gov.ir/XYp").strip()
AFTA_MAX_DOWNLOAD_MB = int(os.environ.get("AFTA_MAX_DOWNLOAD_MB", "25"))
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


def expected_filename_for(jalali_code: str) -> str:
    return f"NEW_FIRE_{jalali_code}_IPS"


def expected_filename() -> str:
    return expected_filename_for(today_jalali_code())


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



class _LinkCollector(HTMLParser):
    def __init__(self):
        super().__init__()
        self.links = []
        self._href = None
        self._text = []

    def handle_starttag(self, tag, attrs):
        if tag.lower() == "a":
            self._href = dict(attrs).get("href")
            self._text = []

    def handle_data(self, data):
        if self._href is not None:
            self._text.append(data)

    def handle_endtag(self, tag):
        if tag.lower() == "a" and self._href is not None:
            self.links.append((self._href, " ".join(self._text).strip()))
            self._href = None
            self._text = []


def _today_jalali_variants() -> tuple[str, ...]:
    code = today_jalali_code()
    year, month, day = code[:4], code[4:6], code[6:8]
    return (
        code,
        f"{year}.{month}.{day}",
        f"{year}_{month}_{day}",
        f"{year}-{month}-{day}",
    )


def _archive_looks_like_today(data: bytes, filename: str = "") -> bool:
    variants = tuple(v.lower() for v in _today_jalali_variants())
    lower_name = (filename or "").lower()

    if any(v in lower_name for v in variants):
        return True

    try:
        with zipfile.ZipFile(BytesIO(data)) as archive:
            names = [n.lower() for n in archive.namelist()]
            return any(any(v in n for v in variants) for n in names)
    except Exception:
        return False


def _response_filename(response) -> str:
    disposition = response.headers.get("Content-Disposition", "")
    match = re.search(r"filename\*?=(?:UTF-8''|\")?([^\";]+)", disposition, re.I)
    if match:
        return unquote(match.group(1).strip().strip('"'))
    return Path(urlparse(response.url).path).name


def _candidate_score(url: str, text: str) -> tuple:
    combined = f"{url} {text}".lower()
    return (
        0 if "new_fire" in combined else 1,
        0 if ".zip" in combined else 1,
        0 if "شاخص" in combined and "آلودگی" in combined else 1,
        0 if "دانلود" in combined else 1,
        len(url),
    )


def _download_afta_archive_sync() -> tuple[bytes, str]:
    session = requests.Session()
    session.headers.update({
        "User-Agent": (
            "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 "
            "(KHTML, like Gecko) Chrome/130.0 Safari/537.36"
        ),
        "Accept-Language": "fa-IR,fa;q=0.9,en;q=0.7",
    })

    max_bytes = AFTA_MAX_DOWNLOAD_MB * 1024 * 1024
    start_host = urlparse(AFTA_INDEX_URL).netloc.lower()
    queue = [(AFTA_INDEX_URL, 0)]
    visited = set()
    last_error = None

    while queue:
        url, depth = queue.pop(0)
        if url in visited or depth > 3:
            continue
        visited.add(url)

        try:
            response = session.get(url, timeout=45, allow_redirects=True, stream=True)
            response.raise_for_status()

            final_host = urlparse(response.url).netloc.lower()
            if final_host and start_host and final_host != start_host:
                raise RuntimeError(
                    f"AFTA link redirected outside expected host: {response.url}"
                )

            content_type = response.headers.get("Content-Type", "").lower()
            filename = _response_filename(response)

            # Download file responses in a bounded way.
            if (
                "zip" in content_type
                or filename.lower().endswith(".zip")
                or "octet-stream" in content_type
            ):
                chunks = []
                total = 0
                for chunk in response.iter_content(chunk_size=65536):
                    if not chunk:
                        continue
                    total += len(chunk)
                    if total > max_bytes:
                        raise RuntimeError(
                            f"AFTA download exceeded {AFTA_MAX_DOWNLOAD_MB} MB limit."
                        )
                    chunks.append(chunk)
                data = b"".join(chunks)

                if zipfile.is_zipfile(BytesIO(data)):
                    if not _archive_looks_like_today(data, filename):
                        log.info("Skipping archive that does not match today's Jalali date: %s", filename)
                        continue
                    return data, filename or "AFTA_NEW_FIRE.zip"

            # HTML page: collect links and continue crawling.
            body = response.content
            if len(body) > 5 * 1024 * 1024:
                raise RuntimeError("AFTA HTML page is unexpectedly large.")

            text = body.decode(response.encoding or "utf-8", errors="replace")
            parser = _LinkCollector()
            parser.feed(text)

            links = []
            for href, anchor_text in parser.links:
                if not href:
                    continue
                absolute = urljoin(response.url, href.strip())
                parsed = urlparse(absolute)
                if parsed.scheme not in ("http", "https"):
                    continue
                if parsed.netloc.lower() != start_host:
                    continue
                links.append((absolute, anchor_text))

            links.sort(key=lambda item: _candidate_score(item[0], item[1]))
            for absolute, anchor_text in links:
                if absolute not in visited:
                    queue.append((absolute, depth + 1))

        except Exception as exc:
            last_error = exc
            log.warning("AFTA fetch failed for %s: %s", url, exc)

    if last_error:
        raise RuntimeError(f"Could not locate today's AFTA ZIP: {last_error}")
    raise RuntimeError("Could not locate today's AFTA ZIP from the configured AFTA page.")


async def download_today_from_afta() -> tuple[bytes, str]:
    log.info("Checking AFTA download page: %s", AFTA_INDEX_URL)
    data, filename = await asyncio.to_thread(_download_afta_archive_sync)
    log.info("Downloaded today's AFTA archive: %s (%s bytes)", filename, len(data))
    return data, filename


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


async def find_file_by_target(client: BaleClient, source_ref, target_name: str):
    target = target_name.lower()
    async for message in client.iter_messages(source_ref, limit=HISTORY_LIMIT):
        content = message.content
        media = getattr(content, "media", None)
        if not media:
            continue
        name = (media.name or "").strip()
        if file_stem(name).lower() == target:
            return message, media
    return None, None


async def find_today_file(client: BaleClient, source_ref):
    return await find_file_by_target(client, source_ref, expected_filename())


async def load_previous_day_ips(client: BaleClient, source_ref):
    yesterday_code = previous_jalali_code()
    target = expected_filename_for(yesterday_code)

    _message, media = await find_file_by_target(client, source_ref, target)
    if not media:
        log.info("Previous day's file not found in source: %s", target)
        return yesterday_code, None

    log.info("Found previous day's file for comparison: %s", media.name)
    file_bytes = await download_media(client, media)
    rows = extract_first_ten(file_bytes)

    if len(rows) < 10:
        log.warning(
            "Previous day's file contains only %s valid IPv4 rows; skipping comparison.",
            len(rows),
        )
        return yesterday_code, None

    return yesterday_code, [ip for ip, _country in rows]


async def process_once(client: BaleClient, destination_ref, state: dict) -> bool:
    today_code = today_jalali_code()
    days = state.setdefault("days", {})

    # Idempotency by calendar day: never send twice for the same date.
    if today_code in days and days[today_code].get("sent"):
        log.info("Already processed today: %s", today_code)
        return True

    try:
        file_bytes, source_file_name = await download_today_from_afta()
    except Exception as exc:
        log.info("Today's AFTA file is not available yet: %s", exc)
        return False

    rows = extract_first_ten(file_bytes)

    if len(rows) < 10:
        raise RuntimeError(
            f"Only {len(rows)} valid IPv4 rows found in {source_file_name}; expected at least 10"
        )

    yesterday_code = previous_jalali_code()

    # Prefer locally saved state, but if it is missing (fresh deploy/state loss),
    # fetch yesterday's actual source file and compare against that.
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
            build_message(source_file_name, rows),
        )
        log.info("Message send returned: %r", sent)
        if yesterday_ips is None:
            log.info("No usable file for previous calendar day (%s); sent normal list.", yesterday_code)
        else:
            log.info("Today's IPs differ from previous day; sent normal list.")

    days[today_code] = {
        "sent": True,
        "file_name": source_file_name,
        "source": AFTA_INDEX_URL,
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
    if not DESTINATION:
        raise SystemExit("BALE_DESTINATION is missing. Run: python list_dialogs.py")

    lock_handle = acquire_process_lock()
    if lock_handle is None:
        log.warning("Another balebot process is already running. Exiting.")
        return

    state = load_state()
    deadline = asyncio.get_running_loop().time() + (MAX_WAIT_MINUTES * 60)

    while True:
        if asyncio.get_running_loop().time() >= deadline:
            log.warning(
                "Today's file was not successfully processed within %s minutes. Exiting.",
                MAX_WAIT_MINUTES,
            )
            return

        try:
            async with BaleClient(BALE_TOKEN) as client:
                me = await client.get_me()
                log.info("Logged in as %s (user id=%s)", me.title, me.peer.id)

                destination_ref = ref_from_config(DESTINATION)
                destination_info = await client.resolve(destination_ref)

                log.info("Today's target: %s", expected_filename())
                log.info(
                    "Destination: %s (id=%s, configured=%s)",
                    destination_info.title,
                    destination_info.peer.id,
                    DESTINATION,
                )

                while True:
                    try:
                        done = await process_once(client, destination_ref, state)
                        if done:
                            log.info("Daily job completed. Exiting.")
                            return
                    except RuntimeError as exc:
                        if "not connected" in str(exc).lower():
                            log.warning("Bale connection lost; reconnecting with a fresh client.")
                            break
                        log.exception("Daily processing failed.")
                    except Exception:
                        log.exception("Daily processing failed.")

                    if asyncio.get_running_loop().time() >= deadline:
                        log.warning(
                            "Today's file was not successfully processed within %s minutes. Exiting.",
                            MAX_WAIT_MINUTES,
                        )
                        return

                    await asyncio.sleep(CHECK_INTERVAL)

        except Exception:
            log.exception("Bale client session failed; retrying with a fresh connection.")

        if asyncio.get_running_loop().time() >= deadline:
            log.warning(
                "Today's file was not successfully processed within %s minutes. Exiting.",
                MAX_WAIT_MINUTES,
            )
            return

        await asyncio.sleep(min(CHECK_INTERVAL, 60))


if __name__ == "__main__":
    asyncio.run(main())
