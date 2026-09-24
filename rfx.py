"""
rfx.py — the RFx co-pilot and the (simulated) send step.

The co-pilot is a chat with Claude that edits a structured RFx draft through tools:
  set_header         title, RFx number, scope
  set_line_items     the SKUs being sourced (replaces the whole list)
  set_questionnaire  supplier questions, with knockout (disqualifying) rules
  set_terms          commercial terms
  load_item_master   pull the buyer's standard SKU list from skus.csv
The buyer can also edit every section directly in the app; the model always sees the latest
version, including those edits.

send_rfx() "sends" the RFx: it saves it, makes its line items the SKU master used by extraction,
and writes one email per vendor to the outbox table. No email leaves the machine.
"""

import csv
import json
from datetime import datetime, timezone
from pathlib import Path
from typing import List, Literal, Optional

import anthropic
from pydantic import BaseModel, ValidationError

import db

MODEL = "claude-sonnet-5"
MAX_TOOL_ROUNDS = 8
ROOT = Path(__file__).parent

DEFAULT_RFX_NO = "NHA/PKG/RFQ/2026/031"
BUYER = "Nilgiri Home Appliances Ltd., Plot 14, MIDC Chakan, Pune 410501"

# Fictional vendor directory for the demo (example.* addresses; nothing is really sent).
DEFAULT_RECIPIENTS = [
    {"vendor": "Shree Ganesh Corrugators Pvt. Ltd.", "email": "sales@shreeganesh.example"},
    {"vendor": "Deccan Kraft Packaging LLP", "email": "sales@deccankraft.example"},
    {"vendor": "Sabarmati Box Industries", "email": "offers@sabarmati.example"},
    {"vendor": "Om Sai Carton Works", "email": "omsai.cartons@example.com"},
    {"vendor": "Coromandel Packaging Exports (EOU)", "email": "karthik.r@coromandelpack.example"},
]

ITEM_COLUMNS = ["sku_id", "description", "item_type", "ply", "flute", "paper_spec", "bursting_factor",
                "total_gsm", "length_mm", "width_mm", "height_mm", "printing", "uom", "annual_qty"]
QUESTION_COLUMNS = ["qid", "question", "answer_type", "knockout", "pass_rule"]
TERM_COLUMNS = ["term", "value"]


# ---------------------------------------------------------------------------
# Draft structure (validated with Pydantic when the model writes it)
# ---------------------------------------------------------------------------
class LineItem(BaseModel):
    sku_id: str
    description: str
    item_type: Literal["RSC", "DIECUT", "PAD", "PARTITION"]
    ply: int
    flute: Optional[str] = None
    paper_spec: Optional[str] = None
    bursting_factor: Optional[int] = None
    total_gsm: Optional[int] = None
    length_mm: int
    width_mm: int
    height_mm: Optional[int] = None
    printing: Optional[str] = None
    uom: Literal["piece", "sheet", "set"]
    annual_qty: Optional[int] = None


class Question(BaseModel):
    qid: str
    question: str
    answer_type: Literal["yes_no", "number", "text", "choice"]
    knockout: bool
    pass_rule: str = ""


class Term(BaseModel):
    term: str
    value: str


def empty_draft():
    return {"rfx_no": DEFAULT_RFX_NO, "title": "", "scope": "", "line_items": [], "questionnaire": [], "terms": []}


def load_item_master():
    """skus.csv -> list of line-item dicts with proper types."""
    ints = {"ply", "bursting_factor", "total_gsm", "length_mm", "width_mm", "height_mm", "annual_qty"}
    with open(ROOT / "skus.csv", newline="", encoding="utf-8") as f:
        rows = list(csv.DictReader(f))
    return [{c: (int(r[c]) if c in ints and r[c] != "" else (r[c] or None)) for c in ITEM_COLUMNS} for r in rows]


def load_sent_draft():
    """The last sent RFx from the database (line items from the SKU master), or None."""
    conn = db.connect()
    row = conn.execute("SELECT * FROM rfx ORDER BY rfx_id DESC LIMIT 1").fetchone()
    items = [dict(r) for r in conn.execute("SELECT * FROM skus ORDER BY sku_id")]
    conn.close()
    if row is None:
        return None, None
    draft = {"rfx_no": row["rfx_no"], "title": row["title"], "scope": row["scope"],
             "line_items": items, "questionnaire": json.loads(row["questionnaire"]),
             "terms": json.loads(row["terms"])}
    return draft, dict(row)


# ---------------------------------------------------------------------------
# Tools
# ---------------------------------------------------------------------------
_ITEM_SCHEMA = {
    "type": "object",
    "properties": {
        "sku_id": {"type": "string", "description": "buyer code, e.g. CB-001"},
        "description": {"type": "string"},
        "item_type": {"type": "string", "enum": ["RSC", "DIECUT", "PAD", "PARTITION"]},
        "ply": {"type": "integer"},
        "flute": {"type": ["string", "null"]},
        "paper_spec": {"type": ["string", "null"], "description": "GSM per layer, e.g. 150/120/150"},
        "bursting_factor": {"type": ["integer", "null"]},
        "total_gsm": {"type": ["integer", "null"]},
        "length_mm": {"type": "integer"},
        "width_mm": {"type": "integer"},
        "height_mm": {"type": ["integer", "null"], "description": "null for pads"},
        "printing": {"type": ["string", "null"]},
        "uom": {"type": "string", "enum": ["piece", "sheet", "set"]},
        "annual_qty": {"type": ["integer", "null"], "description": "null if the buyer has not given it"},
    },
    "required": ["sku_id", "description", "item_type", "ply", "length_mm", "width_mm", "uom"],
}

TOOLS = [
    {
        "name": "set_header",
        "description": "Set the RFx title, number and scope of work (replaces them).",
        "input_schema": {
            "type": "object",
            "properties": {
                "title": {"type": "string"},
                "rfx_no": {"type": "string"},
                "scope": {"type": "string", "description": "Scope of work, a few short paragraphs or bullets."},
            },
            "required": ["title", "rfx_no", "scope"],
        },
    },
    {
        "name": "set_line_items",
        "description": "Replace the full list of line items. Include every item that should remain.",
        "input_schema": {"type": "object", "properties": {"items": {"type": "array", "items": _ITEM_SCHEMA}},
                         "required": ["items"]},
    },
    {
        "name": "set_questionnaire",
        "description": ("Replace the supplier questionnaire. knockout=true means failing the pass_rule "
                        "disqualifies the vendor."),
        "input_schema": {
            "type": "object",
            "properties": {"questions": {"type": "array", "items": {
                "type": "object",
                "properties": {
                    "qid": {"type": "string", "description": "e.g. Q1"},
                    "question": {"type": "string"},
                    "answer_type": {"type": "string", "enum": ["yes_no", "number", "text", "choice"]},
                    "knockout": {"type": "boolean"},
                    "pass_rule": {"type": "string", "description": "e.g. 'Yes', '>= 10000 boxes/day'; empty if none"},
                },
                "required": ["qid", "question", "answer_type", "knockout", "pass_rule"],
            }}},
            "required": ["questions"],
        },
    },
    {
        "name": "set_terms",
        "description": "Replace the commercial terms (payment, freight, price basis, validity, due date, ...).",
        "input_schema": {
            "type": "object",
            "properties": {"terms": {"type": "array", "items": {
                "type": "object",
                "properties": {"term": {"type": "string"}, "value": {"type": "string"}},
                "required": ["term", "value"],
            }}},
            "required": ["terms"],
        },
    },
    {
        "name": "load_item_master",
        "description": ("Load the buyer's standard corrugated SKU master (item codes, specs, sizes, annual "
                        "volumes) into the line items, replacing the current list. Use when the buyer refers "
                        "to their usual / standard / existing items."),
        "input_schema": {"type": "object", "properties": {}},
    },
]

SYSTEM_PROMPT = f"""You are an RFx co-pilot for a category buyer at {BUYER}, sourcing corrugated \
packaging in India. You help the buyer build a request for quotation with four parts: header \
(title, RFx number, scope), line items, a supplier questionnaire, and commercial terms.

How to work
- Change the draft ONLY through the tools. The buyer sees the draft as editable tables next to this \
chat and may edit it; each message includes the current draft in <current_draft>. Always build on it \
(keep the buyer's edits) and pass complete lists, since each tool replaces its section.
- Never invent the buyer's facts: quantities, sizes, specs, volumes, dates or targets they have not \
given. Leave such fields null/empty and ask. You may PROPOSE standard content (common questions, \
standard terms); say clearly that it is a proposal for the buyer to confirm.
- Use Indian corrugated conventions: ply, flute (B, C, BC, E), paper GSM per layer, bursting factor \
(BF), internal dimensions in mm, uom piece/sheet/set.
- Questionnaire: include knockout questions that decide whether a vendor is eligible (e.g. quality \
certification, test certificates per lot, capacity, lead time), each with a clear pass_rule, plus \
non-knockout questions for comparison.
- Terms should cover: price basis (INR per unit, ex-GST, GST rate stated separately), freight basis \
(ask the buyer: delivered to plant vs ex-works), payment, delivery lead time, quote validity, quote due \
date, and that vendors may reply in any format.
- Keep chat replies short: say what you changed in a line or two, then ask the single most useful \
next question."""


def _validate_list(model, rows, key):
    items = [model(**r).model_dump() for r in rows]
    ids = [i[key] for i in items]
    dupes = sorted({x for x in ids if ids.count(x) > 1})
    if dupes:
        raise ValueError(f"duplicate {key}: {dupes}")
    return items


class RfxCopilot:
    def __init__(self):
        self.client = anthropic.Anthropic()
        self.messages = []

    def chat(self, text, draft, on_change=lambda section, summary: None):
        """
        Send one buyer message. `draft` is modified in place by tool calls.
        Returns (reply_text, list of change summaries).
        """
        content = (f"<current_draft>\n{json.dumps(draft, ensure_ascii=False)}\n</current_draft>\n\n"
                   f"Buyer: {text}")
        self.messages.append({"role": "user", "content": content})
        changes, reply = [], ""

        for _ in range(MAX_TOOL_ROUNDS):
            response = self.client.messages.create(
                model=MODEL, max_tokens=16000, system=SYSTEM_PROMPT, tools=TOOLS,
                messages=self.messages, cache_control={"type": "ephemeral"})
            self.messages.append({"role": "assistant", "content": response.content})
            reply = "\n".join(b.text for b in response.content if b.type == "text").strip()
            if response.stop_reason == "refusal":
                return "The model declined this request.", changes
            tool_uses = [b for b in response.content if b.type == "tool_use"]
            if not tool_uses:
                break
            results = []
            for tu in tool_uses:
                try:
                    summary = self._apply(tu.name, tu.input, draft)
                    changes.append(summary)
                    on_change(tu.name, summary)
                    results.append({"type": "tool_result", "tool_use_id": tu.id, "content": f"Done: {summary}"})
                except (ValidationError, ValueError, KeyError, TypeError) as e:
                    results.append({"type": "tool_result", "tool_use_id": tu.id, "is_error": True,
                                    "content": f"Rejected, draft unchanged: {e}"})
            self.messages.append({"role": "user", "content": results})
        return reply, changes

    @staticmethod
    def _apply(name, args, draft):
        if name == "set_header":
            draft.update(title=args["title"], rfx_no=args["rfx_no"], scope=args["scope"])
            return "Updated title and scope"
        if name == "set_line_items":
            draft["line_items"] = _validate_list(LineItem, args["items"], "sku_id")
            return f"Line items: {len(draft['line_items'])}"
        if name == "set_questionnaire":
            draft["questionnaire"] = _validate_list(Question, args["questions"], "qid")
            knockouts = sum(q["knockout"] for q in draft["questionnaire"])
            return f"Questionnaire: {len(draft['questionnaire'])} questions ({knockouts} knockout)"
        if name == "set_terms":
            draft["terms"] = _validate_list(Term, args["terms"], "term")
            return f"Terms: {len(draft['terms'])}"
        if name == "load_item_master":
            draft["line_items"] = load_item_master()
            return f"Loaded {len(draft['line_items'])} items from the item master"
        raise KeyError(f"unknown tool {name}")


# ---------------------------------------------------------------------------
# Send (simulated)
# ---------------------------------------------------------------------------
def check_draft(draft, recipients):
    """Problems that block sending (empty list = OK to send)."""
    problems = []
    if not draft["line_items"]:
        problems.append("Add at least one line item.")
    else:
        try:
            _validate_list(LineItem, draft["line_items"], "sku_id")
        except (ValidationError, ValueError) as e:
            problems.append(f"Line items are incomplete or invalid: {str(e).splitlines()[0]}")
    if not draft["title"].strip():
        problems.append("Add a title.")
    if not [r for r in recipients if r.get("email")]:
        problems.append("Add at least one recipient.")
    return problems


def render_email(draft, vendor):
    lines = [f"Dear {vendor} team,", "",
             f"{BUYER.split(',')[0]} invites your quotation for: {draft['title']} (ref. {draft['rfx_no']}).", "",
             "SCOPE", draft["scope"], "", "LINE ITEMS"]
    for i in draft["line_items"]:
        size = "x".join(str(v) for v in (i["length_mm"], i["width_mm"], i.get("height_mm")) if v)
        lines.append(f"- {i['sku_id']}: {i['description']} | {i['ply']}-ply {i.get('flute') or ''} "
                     f"{i.get('paper_spec') or ''} BF{i.get('bursting_factor') or '-'} | {size} mm | "
                     f"{i.get('printing') or ''} | per {i['uom']} | annual qty {i.get('annual_qty') or 'TBC'}")
    lines += ["", "SUPPLIER QUESTIONNAIRE"]
    for q in draft["questionnaire"]:
        lines.append(f"- {q['qid']}. {q['question']}" + (" [mandatory]" if q["knockout"] else ""))
    lines += ["", "COMMERCIAL TERMS"] + [f"- {t['term']}: {t['value']}" for t in draft["terms"]]
    lines += ["", "You may reply in any format you like (Excel, PDF, Word, email, or a photo of your rate card).",
              "", "Regards,", "Purchase Department", BUYER]
    return f"RFQ {draft['rfx_no']}: {draft['title']}", "\n".join(lines)


def send_rfx(draft, recipients):
    """Save the RFx, make its line items the SKU master, write the outbox. Returns rfx_id."""
    items = _validate_list(LineItem, draft["line_items"], "sku_id")
    now = datetime.now(timezone.utc).isoformat(timespec="seconds")
    conn = db.connect()
    db.clear_responses(conn)   # a new RFx starts a new comparison
    db.save_skus(conn, items)
    cur = conn.execute(
        "INSERT INTO rfx (rfx_no, title, scope, questionnaire, terms, recipients, status, sent_at) "
        "VALUES (?,?,?,?,?,?,?,?)",
        (draft["rfx_no"], draft["title"], draft["scope"], json.dumps(draft["questionnaire"]),
         json.dumps(draft["terms"]), json.dumps(recipients), "sent", now))
    rfx_id = cur.lastrowid
    for r in recipients:
        if r.get("email"):
            subject, body = render_email(draft, r["vendor"])
            conn.execute("INSERT INTO outbox (rfx_id, vendor, email, subject, body, sent_at) VALUES (?,?,?,?,?,?)",
                         (rfx_id, r["vendor"], r["email"], subject, body, now))
    conn.commit()
    conn.close()
    return rfx_id
