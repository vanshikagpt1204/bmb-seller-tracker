import csv
import logging
import shutil
import sqlite3
from pathlib import Path
from openpyxl import load_workbook

ROOT = Path(__file__).resolve().parent.parent
INBOX = ROOT / "data" / "isbn_inbox"
PROCESSED = ROOT / "data" / "isbn_processed"
EXPORTS = ROOT / "data" / "isbn_exports"
SOURCE_URL = "https://isbn.gov.in/Home/IsbnSearch"

log = logging.getLogger("seller_tracker")


def collect_isbn_exports(db: sqlite3.Connection) -> int:
    db.execute("""
        CREATE TABLE IF NOT EXISTS isbn_books (
            isbn TEXT PRIMARY KEY,
            title TEXT,
            edition TEXT,
            product_form TEXT,
            language TEXT,
            applicant_type TEXT,
            publisher TEXT,
            imprint TEXT,
            author TEXT,
            publication_date TEXT,
            cover_uploaded TEXT,
            source_url TEXT
        )
    """)
    db.execute("""
        CREATE TABLE IF NOT EXISTS isbn_publishers (
            publisher_key TEXT PRIMARY KEY,
            publisher TEXT,
            imprint TEXT,
            applicant_type TEXT,
            book_count INTEGER DEFAULT 0
        )
    """)

    INBOX.mkdir(parents=True, exist_ok=True)
    PROCESSED.mkdir(parents=True, exist_ok=True)
    EXPORTS.mkdir(parents=True, exist_ok=True)
    inserted = 0

    for path in sorted(INBOX.glob("*.xlsx")):
        if path.name.startswith("~$"):
            continue
        workbook = load_workbook(path, read_only=True, data_only=True)
        sheet = workbook.active
        rows = sheet.iter_rows(values_only=True)
        headers = [str(v or "").strip().lower() for v in next(rows)]
        indexes = {name: i for i, name in enumerate(headers)}

        required = {"isbn_number", "title", "publisher"}
        if not required.issubset(indexes):
            workbook.close()
            raise ValueError(f"Missing ISBN columns in {path.name}: {required - set(indexes)}")

        def value(row, name):
            i = indexes.get(name)
            return str(row[i]).strip() if i is not None and i < len(row) and row[i] is not None else ""

        for row in rows:
            isbn = value(row, "isbn_number")
            publisher = value(row, "publisher")
            title = value(row, "title")
            if not isbn or not publisher or not title:
                continue

            imprint = value(row, "imprint")
            db.execute("""
                INSERT OR IGNORE INTO isbn_books
                (isbn, title, edition, product_form, language, applicant_type,
                 publisher, imprint, author, publication_date, cover_uploaded, source_url)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """, (
                isbn, title, value(row, "edition"), value(row, "productform"),
                value(row, "language"), value(row, "applicant_type"),
                publisher, imprint, value(row, "author"),
                value(row, "publicationdate"),
                value(row, "finalcoverpageuploaded"), SOURCE_URL
            ))
            inserted += db.execute("SELECT changes()").fetchone()[0]

            key = (publisher + "|" + imprint).casefold()
            db.execute("""
                INSERT INTO isbn_publishers
                (publisher_key, publisher, imprint, applicant_type, book_count)
                VALUES (?, ?, ?, ?, 1)
                ON CONFLICT(publisher_key) DO UPDATE SET
                book_count = (SELECT COUNT(*) FROM isbn_books
                              WHERE lower(publisher)=lower(?) AND lower(imprint)=lower(?))
            """, (key, publisher, imprint, value(row, "applicant_type"), publisher, imprint))

        workbook.close()
        db.commit()
        target = PROCESSED / path.name
        if target.exists():
            target = PROCESSED / f"{path.stem}_{len(list(PROCESSED.iterdir()))}{path.suffix}"
        shutil.move(str(path), str(target))
        log.info("Imported ISBN spreadsheet %s", path.name)

    for table, filename in (
        ("isbn_books", "isbn_books.csv"),
        ("isbn_publishers", "isbn_publishers.csv"),
    ):
        cursor = db.execute(f"SELECT * FROM {table}")
        with (EXPORTS / filename).open("w", newline="", encoding="utf-8-sig") as f:
            writer = csv.writer(f)
            writer.writerow([col[0] for col in cursor.description])
            writer.writerows(cursor.fetchall())

    log.info("ISBN import complete: %d new book records", inserted)
    return inserted
