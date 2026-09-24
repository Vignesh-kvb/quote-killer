"""
db.py — the SQLite database shared by every part of the app.

Tables
  skus               the line items being sourced (from the sent RFx, or skus.csv via extract.py)
  responses          one row per vendor file: document-level terms (freight, discount, GST, ...)
  quote_lines        one row per (vendor file, SKU): the price exactly as the vendor wrote it
  normalised_prices  one row per (vendor file, SKU): comparable INR price, status, steps, flags
  rfx, outbox        the sent RFx and its simulated emails
View
  comparison         flat join of the above, used by the Analyst

Values in `responses` and `quote_lines` are RAW: as written by the vendor, not yet
converted to INR per piece. Normalisation is a separate, later step.
"""

import csv
import sqlite3
from pathlib import Path

ROOT = Path(__file__).parent
DB_PATH = ROOT / "quotes.db"

SCHEMA = """
CREATE TABLE IF NOT EXISTS skus (
    sku_id          TEXT PRIMARY KEY,
    description     TEXT,
    item_type       TEXT,
    ply             INTEGER,
    flute           TEXT,
    paper_spec      TEXT,
    bursting_factor INTEGER,
    total_gsm       INTEGER,
    length_mm       INTEGER,
    width_mm        INTEGER,
    height_mm       INTEGER,
    printing        TEXT,
    uom             TEXT,
    annual_qty      INTEGER
);

CREATE TABLE IF NOT EXISTS responses (
    response_id     INTEGER PRIMARY KEY AUTOINCREMENT,
    file_name       TEXT UNIQUE NOT NULL,
    file_type       TEXT,
    vendor_name     TEXT,
    quote_reference TEXT,
    quote_date      TEXT,
    currency        TEXT,
    freight         TEXT,
    discount_text   TEXT,
    discount_pct    REAL,
    gst_text        TEXT,
    payment_terms   TEXT,
    validity        TEXT,
    delivery        TEXT,
    terms_source    TEXT,
    unmatched_lines TEXT,   -- JSON list of vendor lines that matched no SKU
    warnings        TEXT,   -- JSON list: model warnings + our own validation checks
    raw_json        TEXT,   -- the full model output, for audit
    model           TEXT,
    input_tokens    INTEGER,
    output_tokens   INTEGER,
    extracted_at    TEXT
);

CREATE TABLE IF NOT EXISTS quote_lines (
    response_id     INTEGER NOT NULL REFERENCES responses(response_id) ON DELETE CASCADE,
    sku_id          TEXT NOT NULL REFERENCES skus(sku_id),
    quoted          INTEGER NOT NULL,  -- 1 = vendor explicitly priced this SKU
    price           REAL,              -- number as written; NULL if not quoted / unreadable
    currency        TEXT,
    unit            TEXT,              -- basis as written, e.g. "per box", "per 100 pcs"
    pack_size       INTEGER,           -- pieces the price covers (1, 100, ...); NULL if unclear
    tax_included    INTEGER,           -- 1 incl. GST, 0 excl. GST, NULL not stated
    conditions      TEXT,              -- JSON list of line-specific conditions
    confidence      REAL,              -- 0..1, from the model
    matched_by      TEXT,              -- how the vendor line was matched to the SKU
    source_note     TEXT,              -- page / cell / paragraph + short quote
    PRIMARY KEY (response_id, sku_id)
);

CREATE TABLE IF NOT EXISTS normalised_prices (
    response_id     INTEGER NOT NULL REFERENCES responses(response_id) ON DELETE CASCADE,
    sku_id          TEXT NOT NULL REFERENCES skus(sku_id),
    status          TEXT NOT NULL,     -- ok | flagged | needs_review | not_quoted
    inr_per_unit    REAL,              -- comparable price: INR per piece/sheet/set, ex-GST, after discount
                                       -- (NULL unless status is ok or flagged)
    indicative_inr  REAL,              -- rough estimate shown for needs_review cells only; never compared
    fx_rate         REAL,              -- INR per unit of the quoted currency, as used
    steps           TEXT,              -- JSON list: every conversion applied, in order
    flags           TEXT,              -- JSON list: why the cell is flagged / needs review
    normalised_at   TEXT,
    PRIMARY KEY (response_id, sku_id)
);

CREATE TABLE IF NOT EXISTS rfx (
    rfx_id          INTEGER PRIMARY KEY AUTOINCREMENT,
    rfx_no          TEXT,
    title           TEXT,
    scope           TEXT,
    questionnaire   TEXT,   -- JSON list of questions
    terms           TEXT,   -- JSON list of {term, value}
    recipients      TEXT,   -- JSON list of {vendor, email}
    status          TEXT,   -- sent
    sent_at         TEXT
);

CREATE TABLE IF NOT EXISTS outbox (   -- simulated email: nothing is actually sent
    rfx_id          INTEGER REFERENCES rfx(rfx_id) ON DELETE CASCADE,
    vendor          TEXT,
    email           TEXT,
    subject         TEXT,
    body            TEXT,
    sent_at         TEXT
);

-- One flat row per (vendor, SKU): what the Analyst queries. Recreated on every connect.
DROP VIEW IF EXISTS comparison;
CREATE VIEW comparison AS
SELECT n.response_id, r.vendor_name, r.file_name,
       n.sku_id, s.description, s.item_type, s.ply, s.flute, s.printing, s.uom, s.annual_qty,
       s.length_mm, s.width_mm, s.height_mm,
       n.status, n.inr_per_unit, n.flags,
       l.quoted, l.price AS raw_price, l.currency AS raw_currency, l.unit AS raw_unit,
       l.confidence, l.source_note,
       r.freight, r.gst_text, r.discount_pct, r.payment_terms, r.validity, r.delivery
FROM normalised_prices n
JOIN quote_lines l USING (response_id, sku_id)
JOIN responses r USING (response_id)
JOIN skus s USING (sku_id);
"""


def connect():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA foreign_keys = ON")
    conn.executescript(SCHEMA)
    return conn


def load_skus(conn, csv_path=ROOT / "skus.csv"):
    """(Re)load the SKU master from skus.csv. Returns the SKUs as a list of dicts."""
    with open(csv_path, newline="", encoding="utf-8") as f:
        rows = list(csv.DictReader(f))
    cols = list(rows[0].keys())
    conn.executemany(
        f"INSERT OR REPLACE INTO skus ({', '.join(cols)}) VALUES ({', '.join('?' * len(cols))})",
        [[r[c] if r[c] != "" else None for c in cols] for r in rows],
    )
    conn.commit()
    return rows


def save_skus(conn, items):
    """Replace the SKU master with the RFx line items (list of dicts with skus columns)."""
    cols = ["sku_id", "description", "item_type", "ply", "flute", "paper_spec", "bursting_factor",
            "total_gsm", "length_mm", "width_mm", "height_mm", "printing", "uom", "annual_qty"]
    conn.execute("DELETE FROM skus")
    conn.executemany(f"INSERT INTO skus ({', '.join(cols)}) VALUES ({', '.join('?' * len(cols))})",
                     [[item.get(c) for c in cols] for item in items])
    conn.commit()


def clear_responses(conn):
    """Remove every vendor reply and everything derived from it."""
    for table in ("normalised_prices", "quote_lines", "responses"):
        conn.execute(f"DELETE FROM {table}")
    conn.commit()


def reset_all(conn):
    """Back to a blank project: no RFx, no SKUs, no replies."""
    clear_responses(conn)
    for table in ("outbox", "rfx", "skus"):
        conn.execute(f"DELETE FROM {table}")
    conn.commit()
