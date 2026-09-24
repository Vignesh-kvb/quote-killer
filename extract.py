"""
extract.py — read every vendor reply with Claude and save what it says to SQLite.

For each file in vendor_replies/:
  - PDF and images are sent to Claude directly (Claude reads them visually)
  - Excel and Word are converted to text first (with cell / paragraph references)
  - Plain text (email) is sent as text with line numbers
Claude returns one JSON record per SKU (structured output, validated with Pydantic).

Extraction records values EXACTLY AS WRITTEN. It does not convert currency, remove GST,
or apply discounts — that is the normalisation step, done later in plain Python so the
buyer can see every conversion. Missing values stay null. Nothing is inferred.

Cache: every successful extraction is saved in cache/extractions/ (committed to the repo), keyed
by a hash of the file's bytes, the model, the prompt, the output schema and the SKU list. If all of
those are unchanged, the saved result is reused instead of calling the API (so a live demo is fast
and free). Change any of them and it is a cache miss: Claude is called again.

Run:
  python extract.py                      extract every file not yet in the database
  python extract.py --force              re-process everything (uses the cache where valid)
  python extract.py --force --no-cache   re-extract everything with fresh API calls
  python extract.py 04_om_sai_rate_card.jpg   extract just one file
"""

import argparse
import base64
import hashlib
import json
import sys
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import datetime, timezone
from pathlib import Path
from types import SimpleNamespace
from typing import List, Optional

import anthropic
from docx import Document
from dotenv import load_dotenv
from openpyxl import load_workbook
from pydantic import BaseModel

import db

MODEL = "claude-sonnet-5"
REPLIES_DIR = Path(__file__).parent / "vendor_replies"
CACHE_DIR = Path(__file__).parent / "cache" / "extractions"
RFQ_NO = "NHA/PKG/RFQ/2026/031"


# ---------------------------------------------------------------------------
# The shape Claude must return (enforced by structured outputs)
# ---------------------------------------------------------------------------
class LineQuote(BaseModel):
    sku_id: str
    quoted: bool
    price: Optional[float]
    currency: Optional[str]
    unit: Optional[str]
    pack_size: Optional[int]
    tax_included: Optional[bool]
    conditions: List[str]
    confidence: float
    matched_by: str
    source_note: str


class DocumentTerms(BaseModel):
    vendor_name: Optional[str]
    quote_reference: Optional[str]
    quote_date: Optional[str]
    currency: Optional[str]
    freight: Optional[str]
    discount_text: Optional[str]
    discount_pct: Optional[float]
    gst_text: Optional[str]
    payment_terms: Optional[str]
    validity: Optional[str]
    delivery: Optional[str]
    source_note: str


class Extraction(BaseModel):
    terms: DocumentTerms
    lines: List[LineQuote]
    unmatched_lines: List[str]
    warnings: List[str]


SYSTEM_PROMPT = f"""You are a procurement data-extraction engine. You read ONE vendor's reply to \
the buyer's RFQ {RFQ_NO} for corrugated packaging and record what it says about each SKU in the \
buyer's SKU list. A buyer will commit real money based on your output, so accuracy and honesty \
about uncertainty matter more than completeness.

Rules
1. Record values EXACTLY AS WRITTEN. Do not convert currency, do not convert units, do not remove \
or add GST, do not apply discounts. A later deterministic step does all of that.
2. NEVER infer or guess. If a value is not explicitly stated for a SKU, return null. Phrases like \
"same as last year", "as discussed", "on request", a range, or an illegible number give price = null. \
Never copy a price from a similar SKU, never compute one from a per-kg rate.
3. quoted = true only if the document explicitly gives a price for that SKU. Otherwise quoted = false \
and price = null; explain in source_note (e.g. "not listed on rate card").
4. Matching: vendors may not use the buyer's SKU codes. Match on dimensions, ply and description. \
Sizes may be in inches (1 inch = 25.4 mm; allow about 3 mm rounding). Say how you matched in \
matched_by (e.g. "SKU code", "size + ply", "description"). If a vendor line could match more than one \
SKU, or none, do not force it: add it to unmatched_lines and leave the SKUs unquoted.
5. currency: ISO 4217 code ("INR" for Rs / Rs. / ₹ / rupees, "USD" for $ / US dollars); null if \
not stated. unit: the price basis as written ("per box", "per pc", "per 100 pcs", "per kg", ...). pack_size: the \
number of pieces the price covers (1 for per piece/box/sheet/set, 100 for per 100 pcs); null if the \
basis is not stated or not a piece count (e.g. per kg).
6. tax_included: true only if the document says the rates include GST; false if it says GST is extra \
or the price column is explicitly ex-GST; null if not stated. If both ex-GST and GST-inclusive prices \
are shown, record the ex-GST price with tax_included = false.
7. If a printed value is struck out and corrected by hand, record the corrected value, mention both \
in source_note and add a warning.
8. Document-level terms (freight, discounts, GST, payment, validity, delivery) go in `terms`. Read \
EVERYTHING, including footnotes, small print, asterisks, later pages, notes sheets and the end of \
letters. discount_pct only if an explicit percentage applies to this quotation; describe any \
conditions in discount_text. Line-specific conditions go in that line's `conditions`.
9. confidence (0 to 1): how sure you are that the price, unit AND SKU match are all correct. Use \
below 0.7 when legibility, matching or price basis is uncertain, and say why in source_note.
10. source_note: where you found it (page, sheet + cell, paragraph, line or photo row) plus a short \
verbatim quote.
11. Return exactly one entry in `lines` for EVERY SKU in the buyer's list, in the same order.
12. Put anything a buyer should double-check in `warnings`."""


# ---------------------------------------------------------------------------
# Turning each file into content Claude can read
# ---------------------------------------------------------------------------
def excel_to_text(path):
    """Every non-empty cell with its sheet and cell reference, row by row."""
    wb = load_workbook(path, data_only=True)
    out = []
    for ws in wb.worksheets:
        out.append(f"=== Sheet: {ws.title} ===")
        for row in ws.iter_rows():
            cells = [f"{c.coordinate}={c.value!r}" for c in row if c.value is not None]
            if cells:
                out.append(" | ".join(cells))
    return "\n".join(out)


def docx_to_text(path):
    """Paragraphs and tables, each labelled so Claude can cite them."""
    doc = Document(path)
    out = [f"[para {i}] {p.text}" for i, p in enumerate(doc.paragraphs, start=1) if p.text.strip()]
    for t_i, table in enumerate(doc.tables, start=1):
        for r_i, row in enumerate(table.rows, start=1):
            out.append(f"[table {t_i} row {r_i}] " + " | ".join(c.text for c in row.cells))
    return "\n".join(out)


def text_with_line_numbers(path):
    lines = path.read_text(encoding="utf-8").splitlines()
    return "\n".join(f"L{i:03d}: {line}" for i, line in enumerate(lines, start=1))


def b64(path):
    return base64.standard_b64encode(path.read_bytes()).decode("utf-8")


IMAGE_TYPES = {".jpg": "image/jpeg", ".jpeg": "image/jpeg", ".png": "image/png", ".webp": "image/webp"}


def file_to_content(path):
    """Returns (file_type, list of content blocks) for the user message."""
    ext = path.suffix.lower()
    if ext == ".pdf":
        return "pdf", [{"type": "document",
                        "source": {"type": "base64", "media_type": "application/pdf", "data": b64(path)}}]
    if ext in IMAGE_TYPES:
        return "image", [{"type": "image",
                          "source": {"type": "base64", "media_type": IMAGE_TYPES[ext], "data": b64(path)}}]
    if ext in (".xlsx", ".xlsm"):
        text, kind = excel_to_text(path), "excel"
    elif ext == ".docx":
        text, kind = docx_to_text(path), "word"
    elif ext in (".txt", ".eml"):
        text, kind = text_with_line_numbers(path), "email"
    else:
        raise ValueError(f"Unsupported file type: {path.name}")
    return kind, [{"type": "text", "text": f"<vendor_file name=\"{path.name}\">\n{text}\n</vendor_file>"}]


def sku_table(skus):
    # Inch equivalents are included because many Indian vendors quote sizes in inches.
    lines = ["sku_id | description | type | ply | L x W x H mm | same in inches | uom"]
    for s in skus:
        vals = [int(v) for v in (s["length_mm"], s["width_mm"], s["height_mm"]) if v]
        mm_txt = "x".join(str(v) for v in vals)
        in_txt = "x".join(f"{v / 25.4:.1f}" for v in vals)
        lines.append(f"{s['sku_id']} | {s['description']} | {s['item_type']} | {s['ply']} | "
                     f"{mm_txt} | {in_txt} | {s['uom']}")
    return "\n".join(lines)


# ---------------------------------------------------------------------------
# Calling Claude
# ---------------------------------------------------------------------------
def cache_key(path, skus):
    """Hash of everything that can change the model's answer for this file."""
    h = hashlib.sha256()
    for part in (path.read_bytes(), MODEL.encode(), SYSTEM_PROMPT.encode(), sku_table(skus).encode(),
                 json.dumps(Extraction.model_json_schema(), sort_keys=True).encode()):
        h.update(hashlib.sha256(part).digest())
    return h.hexdigest()[:16]


def cache_path(path, skus):
    return CACHE_DIR / f"{path.name}.{cache_key(path, skus)}.json"


def read_cache(path, skus):
    """(result, usage, cached_at) from the cache, or None on a miss."""
    cp = cache_path(path, skus)
    if not cp.exists():
        return None
    data = json.loads(cp.read_text(encoding="utf-8"))
    usage = SimpleNamespace(**data["usage"])
    return Extraction.model_validate(data["extraction"]), usage, data["created_at"]


def write_cache(path, skus, result, usage, created_at=None):
    CACHE_DIR.mkdir(parents=True, exist_ok=True)
    data = {"file": path.name, "model": MODEL,
            "created_at": created_at or datetime.now(timezone.utc).isoformat(timespec="seconds"),
            "usage": {"input_tokens": usage.input_tokens, "output_tokens": usage.output_tokens},
            "extraction": result.model_dump()}
    cache_path(path, skus).write_text(json.dumps(data, indent=1, ensure_ascii=False), encoding="utf-8")


def extract_file(client, path, skus):
    """One API call. Returns (file_type, Extraction, usage)."""
    file_type, blocks = file_to_content(path)
    prompt = (f"Buyer's SKU list ({len(skus)} SKUs):\n{sku_table(skus)}\n\n"
              f"Extract this vendor reply ({path.name}, {file_type}) following the rules.")
    # Streaming lets us allow a large output budget (30 lines + thinking) without HTTP timeouts.
    with client.messages.stream(
        model=MODEL,
        max_tokens=64000,
        system=SYSTEM_PROMPT,
        messages=[{"role": "user", "content": blocks + [{"type": "text", "text": prompt}]}],
        output_format=Extraction,
    ) as stream:
        response = stream.get_final_message()
    if response.stop_reason == "refusal":
        raise RuntimeError(f"Claude declined to process {path.name}")
    if response.stop_reason == "max_tokens":
        raise RuntimeError(f"Output was cut off for {path.name} (max_tokens reached)")
    return file_type, response.parsed_output, response.usage


def validate(result, skus):
    """Our own checks on top of the model's output. Never changes a price; only flags."""
    warnings = list(result.warnings)
    valid_ids = [s["sku_id"] for s in skus]
    by_sku = {}
    for line in result.lines:
        if line.sku_id not in valid_ids:
            warnings.append(f"Model returned unknown SKU '{line.sku_id}' - ignored.")
            continue
        if line.sku_id in by_sku:
            warnings.append(f"Model returned {line.sku_id} twice - kept the first, please review.")
            continue
        if not line.quoted and line.price is not None:
            warnings.append(f"{line.sku_id}: marked not quoted but had a price - price discarded.")
            line.price = None
        if line.quoted and line.price is None:
            warnings.append(f"{line.sku_id}: marked quoted but no price given - treated as not quoted.")
            line.quoted = False
        line.confidence = min(max(line.confidence, 0.0), 1.0)
        by_sku[line.sku_id] = line
    for sid in valid_ids:  # every SKU must have a row; a missing one is "not quoted", never guessed
        if sid not in by_sku:
            warnings.append(f"{sid}: missing from model output - recorded as not quoted.")
            by_sku[sid] = LineQuote(sku_id=sid, quoted=False, price=None, currency=None, unit=None,
                                    pack_size=None, tax_included=None, conditions=[], confidence=0.0,
                                    matched_by="none", source_note="Not returned by extraction.")
    unquoted = [sid for sid in valid_ids if not by_sku[sid].quoted]
    if result.unmatched_lines and unquoted:
        warnings.append(f"POSSIBLE MISSED MATCH: {len(result.unmatched_lines)} vendor line(s) matched no SKU "
                        f"while {', '.join(unquoted)} show as not quoted. Review before treating them as missing.")
    return [by_sku[sid] for sid in valid_ids], warnings


def save(conn, path, file_type, result, lines, warnings, usage):
    t = result.terms
    conn.execute("DELETE FROM responses WHERE file_name = ?", (path.name,))  # cascades to quote_lines
    cur = conn.execute(
        """INSERT INTO responses (file_name, file_type, vendor_name, quote_reference, quote_date, currency,
               freight, discount_text, discount_pct, gst_text, payment_terms, validity, delivery,
               terms_source, unmatched_lines, warnings, raw_json, model, input_tokens, output_tokens,
               extracted_at)
           VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)""",
        (path.name, file_type, t.vendor_name, t.quote_reference, t.quote_date, t.currency, t.freight,
         t.discount_text, t.discount_pct, t.gst_text, t.payment_terms, t.validity, t.delivery,
         t.source_note, json.dumps(result.unmatched_lines), json.dumps(warnings),
         result.model_dump_json(), MODEL, usage.input_tokens, usage.output_tokens,
         datetime.now(timezone.utc).isoformat(timespec="seconds")))
    conn.executemany(
        """INSERT INTO quote_lines (response_id, sku_id, quoted, price, currency, unit, pack_size,
               tax_included, conditions, confidence, matched_by, source_note)
           VALUES (?,?,?,?,?,?,?,?,?,?,?,?)""",
        [(cur.lastrowid, l.sku_id, int(l.quoted), l.price, l.currency, l.unit, l.pack_size,
          None if l.tax_included is None else int(l.tax_included), json.dumps(l.conditions),
          l.confidence, l.matched_by, l.source_note) for l in lines])
    conn.commit()


def process_files(paths, skus, on_progress=lambda path, msg: None, max_workers=5, use_cache=True):
    """
    Extract several files and save them to SQLite. Cached results are used when valid;
    the rest are sent to Claude in parallel and then cached.
    Returns a list of (path, summary or None, error or None).
    """
    outcomes, jobs = [], []
    conn = db.connect()

    def store(path, file_type, result, usage, source):
        lines, warnings = validate(result, skus)
        save(conn, path, file_type, result, lines, warnings, usage)
        quoted = sum(l.quoted for l in lines)
        summary = (f"{result.terms.vendor_name}: {quoted}/{len(lines)} SKUs quoted, "
                   f"{len(warnings)} warnings ({source})")
        on_progress(path, summary)
        outcomes.append((path, summary, None))

    for path in paths:
        hit = read_cache(path, skus) if use_cache else None
        if hit:
            result, usage, cached_at = hit
            store(path, file_to_content(path)[0], result, usage, f"cached result from {cached_at[:10]}")
        else:
            jobs.append(path)

    if jobs:
        client = anthropic.Anthropic()
        with ThreadPoolExecutor(max_workers=max_workers) as pool:
            futures = {pool.submit(extract_file, client, path, skus): path for path in jobs}
            for path in jobs:
                on_progress(path, "reading with Claude ...")
            for future in as_completed(futures):
                path = futures[future]
                try:
                    file_type, result, usage = future.result()
                except (anthropic.APIError, RuntimeError, ValueError) as e:
                    on_progress(path, f"FAILED: {e}")
                    outcomes.append((path, None, str(e)))
                    continue
                write_cache(path, skus, result, usage)
                store(path, file_type, result, usage, "live API call")
    conn.close()
    return outcomes


def main():
    parser = argparse.ArgumentParser(description="Extract vendor replies with Claude into quotes.db")
    parser.add_argument("files", nargs="*", help="specific file names in vendor_replies/ (default: all)")
    parser.add_argument("--force", action="store_true", help="re-process files already in the database")
    parser.add_argument("--no-cache", action="store_true", help="ignore cached results and call the API")
    args = parser.parse_args()

    load_dotenv()  # puts ANTHROPIC_API_KEY from .env into the environment
    conn = db.connect()
    skus = db.load_skus(conn)
    done = {r["file_name"] for r in conn.execute("SELECT file_name FROM responses")}
    conn.close()

    paths = [REPLIES_DIR / f for f in args.files] if args.files else sorted(
        p for p in REPLIES_DIR.iterdir() if p.is_file() and not p.name.startswith("."))
    todo = []
    for path in paths:
        if path.name in done and not args.force and not args.files:
            print(f"skip  {path.name} (already extracted; use --force to redo)")
        else:
            todo.append(path)
    outcomes = process_files(todo, skus, use_cache=not args.no_cache,
                             on_progress=lambda path, msg: print(f"{path.name}: {msg}", flush=True))
    sys.exit(1 if any(err for _, _, err in outcomes) else 0)


if __name__ == "__main__":
    main()
