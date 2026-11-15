"""Daily seller-report service with a neutral import layer for authorized sources."""

from __future__ import annotations

import csv
import hashlib
import json
import logging
import os
import shutil
import smtplib
import sqlite3
import ssl
import time
from datetime import datetime
from email.message import EmailMessage
from pathlib import Path
from typing import Any
from zoneinfo import ZoneInfo

from dotenv import load_dotenv
from openpyxl import load_workbook

ROOT = Path(__file__).resolve().parent.parent
DB_PATH = Path(os.getenv("DATABASE_PATH", ROOT / "data" / "seller_tracker.db"))
EXPORT_DIR = Path(os.getenv("EXPORT_DIR", ROOT / "data" / "seller_exports"))
INBOX_DIR = Path(os.getenv("INBOX_DIR", ROOT / "data" / "seller_inbox"))
PROCESSED_DIR = Path(os.getenv("PROCESSED_DIR", ROOT / "data" / "seller_processed"))
CSV_COLUMNS = [
    "update_type", "platform", "seller_id", "seller_name", "store_name",
    "business_email", "business_phone", "registered_at", "latest_upload_name",
    "latest_upload_at", "latest_upload_url", "source_url", "first_seen",
]

logging.basicConfig(
    level=os.getenv("LOG_LEVEL", "INFO").upper(),
    format="%(asctime)s %(levelname)s %(message)s",
)
log = logging.getLogger("seller_tracker")

FIELD_ALIASES = {
    "platform": "platform",
    "source": "platform",
    "sellerid": "seller_id",
    "accountid": "seller_id",
    "vendorid": "seller_id",
    "id": "seller_id",
    "sellername": "seller_name",
    "name": "seller_name",
    "businessname": "store_name",
    "storename": "store_name",
    "shopname": "store_name",
    "businessemail": "business_email",
    "selleremail": "business_email",
    "publicemail": "business_email",
    "businessphone": "business_phone",
    "sellerphone": "business_phone",
    "publicphone": "business_phone",
    "registeredat": "registered_at",
    "registrationdate": "registered_at",
    "createdat": "registered_at",
    "latestuploadname": "latest_upload_name",
    "latestuploadfilename": "latest_upload_name",
    "latestuploadtitle": "latest_upload_name",
    "latestuploadid": "latest_upload_id",
    "latestuploadat": "latest_upload_at",
    "latestuploadedat": "latest_upload_at",
    "latestuploaddate": "latest_upload_at",
    "latestuploadurl": "latest_upload_url",
    "sourceurl": "source_url",
    "sellerurl": "source_url",
}


def connect_db() -> sqlite3.Connection:
    DB_PATH.parent.mkdir(parents=True, exist_ok=True)
    db = sqlite3.connect(DB_PATH)
    db.row_factory = sqlite3.Row
    db.execute(
        """CREATE TABLE IF NOT EXISTS seller_events (
            id INTEGER PRIMARY KEY,
            event_key TEXT NOT NULL UNIQUE,
            platform TEXT NOT NULL,
            seller_id TEXT NOT NULL,
            update_type TEXT NOT NULL,
            seller_name TEXT NOT NULL DEFAULT '',
            store_name TEXT NOT NULL DEFAULT '',
            business_email TEXT NOT NULL DEFAULT '',
            business_phone TEXT NOT NULL DEFAULT '',
            registered_at TEXT NOT NULL DEFAULT '',
            latest_upload_id TEXT NOT NULL DEFAULT '',
            latest_upload_name TEXT NOT NULL DEFAULT '',
            latest_upload_at TEXT NOT NULL DEFAULT '',
            latest_upload_url TEXT NOT NULL DEFAULT '',
            source_url TEXT NOT NULL DEFAULT '',
            first_seen TEXT NOT NULL,
            emailed_at TEXT
        )"""
    )
    db.execute("CREATE TABLE IF NOT EXISTS state (key TEXT PRIMARY KEY, value TEXT NOT NULL)")
    db.commit()
    return db


def clean(value: Any) -> str:
    if value is None:
        return ""
    if isinstance(value, float) and value.is_integer():
        value = int(value)
    return " ".join(str(value).split()).strip()


def normalize_field(value: str) -> str:
    return "".join(character.lower() for character in value if character.isalnum())


def normalize_record(raw: dict[str, Any], default_platform: str) -> dict[str, str]:
    record: dict[str, str] = {field: "" for field in set(FIELD_ALIASES.values())}
    for key, value in raw.items():
        target = FIELD_ALIASES.get(normalize_field(str(key)))
        if target:
            record[target] = clean(value)

    # Many APIs return the latest upload as a nested object; flatten known fields.
    upload = raw.get("latest_upload") or raw.get("latestUpload") or raw.get("last_upload")
    if isinstance(upload, dict):
        upload_aliases = {
            "id": "latest_upload_id", "name": "latest_upload_name",
            "filename": "latest_upload_name", "title": "latest_upload_name",
            "created_at": "latest_upload_at", "uploaded_at": "latest_upload_at",
            "date": "latest_upload_at", "url": "latest_upload_url",
        }
        for key, value in upload.items():
            target = upload_aliases.get(str(key).lower())
            if target and not record.get(target):
                record[target] = clean(value)

    record["platform"] = record.get("platform") or default_platform
    return record


def read_csv(path: Path, default_platform: str) -> list[dict[str, str]]:
    with path.open("r", encoding="utf-8-sig", newline="") as file:
        return [normalize_record(row, default_platform) for row in csv.DictReader(file)]


def read_xlsx(path: Path, default_platform: str) -> list[dict[str, str]]:
    workbook = load_workbook(path, read_only=True, data_only=True)
    sheet = workbook.active
    rows = list(sheet.iter_rows(values_only=True))
    workbook.close()
    header_index = next(
        (index for index, row in enumerate(rows[:20])
         if any(normalize_field(clean(cell)) in FIELD_ALIASES for cell in row if cell is not None)),
        None,
    )
    if header_index is None:
        raise ValueError("Could not find seller field headers in the first 20 spreadsheet rows")
    headers = [clean(value) for value in rows[header_index]]
    records = []
    for row in rows[header_index + 1:]:
        if any(value is not None and clean(value) for value in row):
            records.append(normalize_record(
                {headers[i]: row[i] for i in range(min(len(headers), len(row)))}, default_platform
            ))
    return records


def read_json(path: Path, default_platform: str) -> list[dict[str, str]]:
    payload = json.loads(path.read_text(encoding="utf-8-sig"))
    if isinstance(payload, dict):
        payload = payload.get("sellers", payload.get("data", [payload]))
    if not isinstance(payload, list):
        raise ValueError("JSON input must be a seller list or an object containing 'sellers'/'data'")
    return [normalize_record(row, default_platform) for row in payload if isinstance(row, dict)]


def read_import(path: Path) -> list[dict[str, str]]:
    relative_parts = path.relative_to(INBOX_DIR).parts
    platform = relative_parts[0] if len(relative_parts) > 1 else path.stem
    extension = path.suffix.lower()
    if extension == ".csv":
        return read_csv(path, platform)
    if extension == ".xlsx":
        return read_xlsx(path, platform)
    if extension == ".json":
        return read_json(path, platform)
    raise ValueError("Supported input files are .csv, .xlsx, and .json")


def make_event_key(record: dict[str, str]) -> str:
    identity = record.get("seller_id") or record.get("business_email")
    if not identity:
        raise ValueError("A record needs seller_id (preferred) or business_email to deduplicate")
    upload_version = (
        record.get("latest_upload_id") or record.get("latest_upload_at")
        or record.get("latest_upload_name") or "registered"
    )
    key = "|".join((record["platform"].casefold(), identity.casefold(), upload_version.casefold()))
    return hashlib.sha256(key.encode("utf-8")).hexdigest()


def archive_import(path: Path) -> None:
    PROCESSED_DIR.mkdir(parents=True, exist_ok=True)
    target = PROCESSED_DIR / path.name
    if target.exists():
        target = PROCESSED_DIR / f"{path.stem}_{datetime.now():%Y%m%d%H%M%S}{path.suffix}"
    shutil.move(str(path), str(target))


def collect_imports(db: sqlite3.Connection) -> int:
    INBOX_DIR.mkdir(parents=True, exist_ok=True)
    now = datetime.now(ZoneInfo(os.getenv("TIMEZONE", "Asia/Kolkata"))).isoformat(timespec="seconds")
    first_run = db.execute("SELECT value FROM state WHERE key='baseline_done'").fetchone() is None
    inserted = 0
    files = sorted(path for path in INBOX_DIR.rglob("*") if path.is_file() and not path.name.startswith("~$"))
    for path in files:
        try:
            records = read_import(path)
            valid = [record for record in records if record.get("seller_id") or record.get("business_email")]
            if not valid:
                raise ValueError("No records with seller IDs were found; file left for review")
            for record in valid:
                event_key = make_event_key(record)
                exists = db.execute(
                    "SELECT 1 FROM seller_events WHERE platform=? AND seller_id=? LIMIT 1",
                    (record["platform"], record.get("seller_id") or record.get("business_email")),
                ).fetchone()
                update_type = "New upload" if exists else "New seller"
                cursor = db.execute(
                    """INSERT OR IGNORE INTO seller_events
                       (event_key, platform, seller_id, update_type, seller_name, store_name,
                        business_email, business_phone, registered_at, latest_upload_id,
                        latest_upload_name, latest_upload_at, latest_upload_url, source_url,
                        first_seen, emailed_at)
                       VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)""",
                    (
                        event_key, record["platform"], record.get("seller_id") or record.get("business_email"),
                        update_type, record.get("seller_name", ""), record.get("store_name", ""),
                        record.get("business_email", ""), record.get("business_phone", ""),
                        record.get("registered_at", ""), record.get("latest_upload_id", ""),
                        record.get("latest_upload_name", ""), record.get("latest_upload_at", ""),
                        record.get("latest_upload_url", ""), record.get("source_url", ""),
                        now, now if first_run else None,
                    ),
                )
                inserted += cursor.rowcount
            db.commit()
            archive_import(path)
            log.info("Read %d seller row(s) from %s; added %d update(s)", len(valid), path.name, inserted)
        except Exception:
            db.rollback()
            log.exception("Could not import %s; file remains available for review", path)

    if first_run and inserted:
        db.execute("INSERT OR REPLACE INTO state(key, value) VALUES('baseline_done', ?)", (now,))
        db.execute("UPDATE seller_events SET emailed_at=? WHERE emailed_at IS NULL", (now,))
        db.commit()
        log.info("Initial seller baseline recorded; existing rows will not be emailed as new")
    return inserted


def collect_ebay_books(db: sqlite3.Connection) -> int:
    """Fetch eBay Singapore Books listings and save new events."""
    from app.ebay import fetch_latest_books_as_records

    now = datetime.now(
        ZoneInfo(os.getenv("TIMEZONE", "Asia/Kolkata"))
    ).isoformat(timespec="seconds")

    baseline = db.execute(
        "SELECT value FROM state WHERE key='ebay_baseline_done'"
    ).fetchone()

    records = fetch_latest_books_as_records(max_results=600)
    inserted = 0

    for record in records:
        event_key = make_event_key(record)

        exists = db.execute(
            """SELECT 1 FROM seller_events
               WHERE platform=? AND seller_id=? LIMIT 1""",
            (record["platform"], record["seller_id"]),
        ).fetchone()

        update_type = "New upload" if exists else "New seller"

        cursor = db.execute(
            """INSERT OR IGNORE INTO seller_events
            (event_key, platform, seller_id, update_type,
             seller_name, store_name, business_email, business_phone,
             registered_at, latest_upload_id, latest_upload_name,
             latest_upload_at, latest_upload_url, source_url,
             first_seen, emailed_at)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)""",
            (
                event_key, record["platform"], record["seller_id"],
                update_type, record["seller_name"], record["store_name"],
                record["business_email"], record["business_phone"],
                record["registered_at"], record["latest_upload_id"],
                record["latest_upload_name"], record["latest_upload_at"],
                record["latest_upload_url"], record["source_url"],
                now, now if baseline else None,
            ),
        )
        inserted += cursor.rowcount

    db.commit()

    if not baseline and records:
        db.execute(
            "INSERT OR REPLACE INTO state(key, value) VALUES(?, ?)",
            ("ebay_baseline_done", now),
        )
        db.execute(
            """UPDATE seller_events SET emailed_at=?
               WHERE platform='ebay_sg' AND emailed_at IS NULL""",
            (now,),
        )
        db.commit()
        log.info("Initial eBay Books baseline recorded")

    log.info("eBay Books: fetched %d listings; added %d events",
             len(records), inserted)
    return inserted

def export_pending_csv(db: sqlite3.Connection, run_time: datetime) -> tuple[Path, list[sqlite3.Row]]:
    rows = db.execute(
        "SELECT * FROM seller_events WHERE emailed_at IS NULL ORDER BY first_seen, id"
    ).fetchall()
    EXPORT_DIR.mkdir(parents=True, exist_ok=True)
    path = EXPORT_DIR / f"seller_updates_{run_time:%Y-%m-%d}.csv"
    with path.open("w", newline="", encoding="utf-8-sig") as file:
        writer = csv.DictWriter(file, fieldnames=CSV_COLUMNS)
        writer.writeheader()
        for row in rows:
            writer.writerow({column: row[column] for column in CSV_COLUMNS})
    return path, rows


def send_report(db: sqlite3.Connection, path: Path, rows: list[sqlite3.Row]) -> None:
    recipients = [item.strip() for item in os.getenv("TEAM_EMAILS", "").split(",") if item.strip()]
    required = ("SMTP_HOST", "SMTP_USERNAME", "SMTP_PASSWORD", "EMAIL_FROM")
    missing = [name for name in required if not os.getenv(name)]
    if not recipients or missing:
        raise RuntimeError("Email settings missing: " + ", ".join(missing or ["TEAM_EMAILS"]))

    message = EmailMessage()
    message["Subject"] = f"Seller tracker daily CSV — {len(rows)} update(s)"
    message["From"] = os.environ["EMAIL_FROM"]
    message["To"] = ", ".join(recipients)
    message.set_content(f"The daily seller tracker found {len(rows)} new seller/upload update(s).")
    message.add_attachment(path.read_bytes(), maintype="text", subtype="csv", filename=path.name)

    host = os.environ["SMTP_HOST"]
    port = int(os.getenv("SMTP_PORT", "587"))
    if port == 465:
        with smtplib.SMTP_SSL(host, port, context=ssl.create_default_context(), timeout=30) as smtp:
            smtp.login(os.environ["SMTP_USERNAME"], os.environ["SMTP_PASSWORD"])
            smtp.send_message(message)
    else:
        with smtplib.SMTP(host, port, timeout=30) as smtp:
            smtp.starttls(context=ssl.create_default_context())
            smtp.login(os.environ["SMTP_USERNAME"], os.environ["SMTP_PASSWORD"])
            smtp.send_message(message)

    sent_at = datetime.now(ZoneInfo(os.getenv("TIMEZONE", "Asia/Kolkata"))).isoformat(timespec="seconds")
    db.executemany("UPDATE seller_events SET emailed_at=? WHERE id=?", [(sent_at, row["id"]) for row in rows])
    db.execute("INSERT OR REPLACE INTO state(key, value) VALUES('last_successful_run', ?)", (sent_at,))
    db.commit()
    log.info("Sent seller report to %d recipient(s)", len(recipients))



def main() -> None:
    load_dotenv(ROOT / ".env")

    timezone = ZoneInfo(os.getenv("TIMEZONE", "Asia/Kolkata"))

    try:
        hour, minute = (
            int(part) for part in os.getenv("RUN_AT", "02:30").split(":")
        )
        if not (0 <= hour <= 23 and 0 <= minute <= 59):
            raise ValueError
    except ValueError as error:
        raise ValueError(
            "RUN_AT must use 24-hour HH:MM format, such as 02:30"
        ) from error

    db = connect_db()

    log.info(
        "Seller tracker ready; daily run at %02d:%02d %s",
        hour, minute, timezone.key
    )

    run_now = os.getenv(
        "POLL_NOW_ON_START", "false"
    ).strip().lower() in {"true", "1", "yes"}

    while True:
        now = datetime.now(timezone)
        today = now.date().isoformat()

        done = db.execute(
            "SELECT value FROM state WHERE key='run_date'"
        ).fetchone()

        due = (now.hour, now.minute) >= (hour, minute)

        if run_now or (due and (not done or done[0] != today)):
            run_now = False

            try:
                collect_imports(db)

                from app.isbn_import import collect_isbn_exports
                collect_isbn_exports(db)

                path, rows = export_pending_csv(db, now)

                required = (
                    "SMTP_HOST", "SMTP_USERNAME",
                    "SMTP_PASSWORD", "EMAIL_FROM"
                )

                missing = [
                    key for key in required if not os.getenv(key)
                ]

                recipients = os.getenv("TEAM_EMAILS", "").strip()

                if rows and not missing and recipients:
                    send_report(db, path, rows)
                elif rows:
                    log.warning(
                        "CSV exported to %s, but email settings are pending",
                        path
                    )
                else:
                    log.info("No new seller updates to report")

                db.execute(
                    "INSERT OR REPLACE INTO state(key, value) VALUES('run_date', ?)",
                    (today,)
                )
                db.commit()

            except Exception:
                log.exception(
                    "Daily job could not complete; will retry in 5 minutes"
                )
                time.sleep(300)
                continue

        time.sleep(20)


if __name__ == "__main__":
    main()
