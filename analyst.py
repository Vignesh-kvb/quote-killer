"""
analyst.py — the Analyst: Claude answers buyer questions by querying quotes.db with tools.

Tools the model can call
  run_sql         read-only SELECT over the database; every result gets an id (Q1, Q2, ...)
  make_chart      a chart built ONLY from a previous query result (the model cannot type in numbers)
  submit_answer   the final answer: text, cited cells, and whether there was enough data

Guardrails in code (not just in the prompt)
  - SQL runs on a read-only connection with an authorizer that allows only reads,
    and a step limit that stops runaway queries.
  - Charts must reference an existing query id and existing column names.
  - Cited cells are checked against the database; unknown ones are dropped and reported.
  - Every number in the answer (prices, totals, percentages, counts) must match a value in a
    query result (number audit); otherwise the answer is sent back to the model to fix.
  - An answer with no queries behind it is marked "not grounded".

Try it from the terminal:  python analyst.py "Which vendor is cheapest overall?"
"""

import json
import re
import sqlite3
import sys

import anthropic
import pandas as pd
from dotenv import load_dotenv

import db
from normalise import load_settings

MODEL = "claude-sonnet-5"
MAX_TOOL_ROUNDS = 15
MAX_ROWS_TO_MODEL = 200
MAX_ANSWER_RETRIES = 2

SCHEMA_DOC = """
Main view (use this for almost everything):
comparison — one row per (vendor reply, SKU). 150 rows = 5 vendors x 30 SKUs.
  response_id      INTEGER  identifies the vendor reply (use with sku_id to cite a cell)
  vendor_name      TEXT     vendor as written in their document
  file_name        TEXT     the vendor's original file
  sku_id           TEXT     buyer's SKU code, CB-001 .. CB-030
  description      TEXT     buyer's item description
  item_type        TEXT     RSC (regular slotted carton) | DIECUT (mailer) | PAD (layer pad) | PARTITION
  ply              INTEGER  3, 5 or 7
  flute, printing  TEXT
  uom              TEXT     piece | sheet | set — the unit inr_per_unit is expressed in
  annual_qty       INTEGER  buyer's expected annual volume for the SKU
  length_mm, width_mm, height_mm  INTEGER (height NULL for pads)
  status           TEXT     ok           -> comparable price, no assumptions
                            flagged      -> comparable price, but relies on an assumption (see flags)
                            needs_review -> NO comparable price (ambiguous unit etc.); inr_per_unit is NULL
                            not_quoted   -> vendor gave no price; inr_per_unit is NULL
  inr_per_unit     REAL     comparable price: INR per uom, EXCLUDING GST, AFTER stated discount,
                            freight AS QUOTED (see freight). NULL unless status is ok or flagged.
  flags            TEXT     JSON list of reasons for flagged / needs_review
  quoted           INTEGER  1 if the vendor explicitly priced this SKU
  raw_price, raw_currency, raw_unit  what the vendor literally wrote
  confidence       REAL     extraction confidence 0..1
  source_note      TEXT     where in the vendor's document the price came from
  freight          TEXT     vendor's freight terms (they DIFFER between vendors)
  gst_text, payment_terms, validity, delivery  TEXT  vendor's terms as written
  discount_pct     REAL     discount already applied in inr_per_unit (NULL if none)

Underlying tables (rarely needed): skus, responses (one row per vendor reply, incl. warnings and
unmatched_lines as JSON text), quote_lines (raw extraction), normalised_prices (steps = JSON list of
conversion steps).
"""

TOOLS = [
    {
        "name": "run_sql",
        "description": (
            "Run ONE read-only SQLite SELECT (or WITH ... SELECT) query against the procurement "
            "database and get the rows back. Each successful call returns a query id like Q3 that "
            "you must reference when you use its numbers. Results over "
            f"{MAX_ROWS_TO_MODEL} rows are truncated, so aggregate in SQL where possible."),
        "strict": True,
        "input_schema": {
            "type": "object",
            "properties": {
                "sql": {"type": "string", "description": "A single SQLite SELECT statement."},
                "purpose": {"type": "string", "description": "One short line: what this query finds out."},
            },
            "required": ["sql", "purpose"],
            "additionalProperties": False,
        },
    },
    {
        "name": "make_chart",
        "description": (
            "Draw a chart for the buyer from the rows of a previous run_sql result. You cannot pass "
            "numbers directly; the chart uses the query's rows. Use it when a visual makes the "
            "comparison clearer (e.g. price by vendor, spend by option)."),
        "strict": True,
        "input_schema": {
            "type": "object",
            "properties": {
                "query_id": {"type": "string", "description": "e.g. Q2"},
                "chart_type": {"type": "string", "enum": ["bar", "grouped_bar", "line", "scatter"]},
                "x": {"type": "string", "description": "column for the x axis"},
                "y": {"type": "string", "description": "numeric column for the y axis"},
                "color": {"type": "string", "description": "optional: column to colour/group by (omit for none)"},
                "title": {"type": "string"},
            },
            "required": ["query_id", "chart_type", "x", "y", "title"],
            "additionalProperties": False,
        },
    },
    {
        "name": "submit_answer",
        "description": "Submit the final answer to the buyer. Call exactly once, at the end.",
        "strict": True,
        "input_schema": {
            "type": "object",
            "properties": {
                "answer": {
                    "type": "string",
                    "description": "Markdown answer. Reference query ids like [Q2] next to figures."},
                "enough_data": {
                    "type": "boolean",
                    "description": "false if the question needs data the database does not contain"},
                "missing_data": {
                    "type": "string",
                    "description": "What data is missing and would be needed (empty string if none)."},
                "cited_cells": {
                    "type": "array",
                    "description": "The (response_id, sku_id) cells whose values drive the answer.",
                    "items": {
                        "type": "object",
                        "properties": {
                            "response_id": {"type": "integer"},
                            "sku_id": {"type": "string"},
                            "why": {"type": "string", "description": "few words: how this cell was used"},
                        },
                        "required": ["response_id", "sku_id", "why"],
                        "additionalProperties": False,
                    },
                },
            },
            "required": ["answer", "enough_data", "missing_data", "cited_cells"],
            "additionalProperties": False,
        },
    },
]


def vendor_list():
    conn = db.connect()
    rows = conn.execute("SELECT response_id, vendor_name, file_name FROM responses ORDER BY file_name").fetchall()
    conn.close()
    return "\n".join(f"  response_id {r['response_id']}: {r['vendor_name']} ({r['file_name']})" for r in rows)


def system_prompt():
    settings = load_settings()
    fx = ", ".join(f"1 {k} = INR {v:g}" for k, v in settings["fx_to_inr"].items() if k != "INR")
    return f"""You are a procurement analyst helping a category buyer at Nilgiri Home Appliances \
evaluate vendor quotes for corrugated packaging (RFQ NHA/PKG/RFQ/2026/031). The buyer may commit \
crores of rupees based on your answer, so every figure must be traceable and every gap must be stated.

Database (SQLite):
{SCHEMA_DOC}
FX rate used in normalisation: {fx}.
Vendor replies:
{vendor_list()}

How to work
1. Answer ONLY from data you retrieved with run_sql in this conversation. Every number you state \
must appear in a query result; put its query id next to it, like "₹7.36 [Q2]". Do ALL arithmetic \
(sums, subtotals per vendor, differences, percentages, savings) in SQL, never yourself: if you need a \
subtotal, run a query that returns it. Answers are automatically checked and figures not found in \
any query result are rejected.
2. Never estimate, impute or fill in a missing price, and never use outside knowledge of market \
prices. A not_quoted or needs_review cell has no comparable price: exclude it and say how many cells \
were excluded and why.
3. Comparable prices are inr_per_unit where status IN ('ok','flagged'). If flagged cells affect the \
answer, say which assumption they rely on (from flags).
4. Prices exclude GST and are freight AS QUOTED: vendors' freight terms differ (one may include \
delivery). Mention this whenever you compare vendors on price. Annual spend = inr_per_unit x annual_qty.
5. If the question needs information the database does not have (for example a quality \
questionnaire, audit results, delivery performance, freight costs, past prices, capacity), set \
enough_data = false, say plainly "not enough data" and what is missing, and answer only the part the \
data supports. Do not invent a proxy without clearly labelling it as a proxy the buyer did not ask for.
6. Include response_id in queries whose rows you will cite. Cite the cells (response_id, sku_id) that drive your answer in submit_answer: e.g. the winning \
price per SKU, the cells behind a total, or the excluded cells you mention. For answers built on \
many cells, cite the most decision-relevant ones (up to 60).
7. Use make_chart when a chart genuinely helps; charts can only use a query's rows.
8. Refer to vendors by name; response_id is internal, never show it to the buyer.
9. Be concise: lead with the direct answer, then a short table or bullets. End by calling \
submit_answer exactly once."""


# ---------------------------------------------------------------------------
# Tool implementations
# ---------------------------------------------------------------------------
_ALLOWED_SQL_ACTIONS = {sqlite3.SQLITE_SELECT, sqlite3.SQLITE_READ, sqlite3.SQLITE_FUNCTION,
                        sqlite3.SQLITE_RECURSIVE}


def _read_only_authorizer(action, *_):
    return sqlite3.SQLITE_OK if action in _ALLOWED_SQL_ACTIONS else sqlite3.SQLITE_DENY


def run_readonly_sql(sql):
    """Run one SELECT on a read-only connection. Returns a DataFrame or raises."""
    conn = sqlite3.connect(f"file:{db.DB_PATH}?mode=ro", uri=True)
    try:
        conn.set_authorizer(_read_only_authorizer)
        steps = {"n": 0}

        def stop_runaway():
            steps["n"] += 1
            return 1 if steps["n"] > 20000 else 0  # ~20M VM steps, then abort

        conn.set_progress_handler(stop_runaway, 1000)
        cur = conn.execute(sql)
        cols = [d[0] for d in cur.description] if cur.description else []
        return pd.DataFrame(cur.fetchall(), columns=cols)
    finally:
        conn.close()


_UNITS = {"crore": 1e7, "crores": 1e7, "cr": 1e7, "lakh": 1e5, "lakhs": 1e5, "lac": 1e5,
          "million": 1e6, "mn": 1e6}
_NUMBER = re.compile(r"(?<![\w.])(\d{1,3}(?:,\d{2,3})+|\d+)(\.\d+)?(?:\s*(crores?|cr|lakhs?|lac|million|mn)\b)?",
                     re.IGNORECASE)


def audit_numbers(answer, frames):
    """
    Return the figures in `answer` that don't match any value in the query results.
    Every number is checked (prices, totals, percentages AND counts), except markdown list numbering,
    SKU codes, sizes, years and query/chart ids. A figure matches if it equals a value (numeric cell,
    number inside a text cell, or a query's row count) rounded to the precision shown.
    """
    text = re.sub(r"(?m)^\s*\d+\.\s", " ", answer)  # "1. " list numbering
    text = re.sub(r"\[[QC\d,\s]+\]|\b[QC]\d+\b|CB-\d+|\d+\s*[x×]\s*\d+(?:\s*[x×]\s*\d+)?|RFQ\S*"
                  r"|\b(?:19|20)\d{2}\b", " ", text)
    values = []
    for df in frames:
        values.append(float(len(df)))
        for col in df.columns:
            nums = pd.to_numeric(df[col], errors="coerce")
            values.extend(float(v) for v in nums.dropna())
            for cell in df[col][nums.isna()].dropna().astype(str):
                values.extend(float(n.replace(",", "")) for n in re.findall(r"\d[\d,]*\.?\d*", cell)
                              if n.replace(",", "").replace(".", "").isdigit())
    unverified = []
    for m in _NUMBER.finditer(text):
        whole, frac, unit = m.group(1), m.group(2) or "", (m.group(3) or "").lower()
        scale = _UNITS.get(unit, 1)
        decimals = len(frac) - 1 if frac else 0
        target = float(whole.replace(",", "") + frac) * scale
        tolerance = 0.5 * 10 ** -decimals * scale + 1e-9
        if not any(abs(abs(v) - target) <= tolerance for v in values):
            unverified.append(m.group(0).strip())
    return sorted(set(unverified))


_MARKUP = re.compile(r"</?(?:answer|parameter|invoke|function_calls)\b[^>]*>")


def valid_cells():
    conn = db.connect()
    cells = {(r["response_id"], r["sku_id"]) for r in conn.execute("SELECT response_id, sku_id FROM normalised_prices")}
    conn.close()
    return cells


class AnalystSession:
    """Holds the conversation and every query/chart produced, across questions."""

    def __init__(self):
        self.client = anthropic.Anthropic()
        self.messages = []
        self.queries = {}   # "Q1" -> {"sql", "purpose", "df"}
        self.charts = {}    # "C1" -> {"query_id", "chart_type", "x", "y", "color", "title"}
        self.turns = []     # finished questions, for display

    # --- tools -----------------------------------------------------------------
    def _tool_run_sql(self, args):
        sql = args["sql"].strip().rstrip(";")
        try:
            df = run_readonly_sql(sql)
        except sqlite3.Error as e:
            return f"SQL error: {e}. Only single read-only SELECT statements are allowed.", True, None
        qid = f"Q{len(self.queries) + 1}"
        self.queries[qid] = {"sql": sql, "purpose": args["purpose"], "df": df}
        shown = df.head(MAX_ROWS_TO_MODEL)
        payload = {"query_id": qid, "columns": list(df.columns), "row_count": len(df),
                   "truncated": len(df) > MAX_ROWS_TO_MODEL,
                   "rows": json.loads(shown.to_json(orient="values"))}
        return json.dumps(payload), False, qid

    def _tool_make_chart(self, args):
        args = dict(args)
        color = (args.get("color") or "").strip()
        args["color"] = "" if color.lower() in ("", "none", "null") else color
        q = self.queries.get(args["query_id"])
        if q is None:
            return f"Unknown query id {args['query_id']}.", True, None
        cols = list(q["df"].columns)
        needed = [args["x"], args["y"]] + ([args["color"]] if args["color"] else [])
        missing = [c for c in needed if c not in cols]
        if missing:
            return f"Columns {missing} not in {args['query_id']} (has {cols}).", True, None
        if not pd.api.types.is_numeric_dtype(q["df"][args["y"]]):
            return f"y column '{args['y']}' is not numeric.", True, None
        cid = f"C{len(self.charts) + 1}"
        self.charts[cid] = dict(args)
        return f"Chart {cid} created from {args['query_id']} ({len(q['df'])} rows).", False, cid

    # --- the loop ----------------------------------------------------------------
    def ask(self, question, on_step=lambda kind, text: None):
        """Run one question to completion. Returns the finished turn dict."""
        turn = {"question": question, "answer": None, "enough_data": True, "missing_data": "",
                "cited_cells": [], "query_ids": [], "chart_ids": [], "warnings": [],
                "input_tokens": 0, "output_tokens": 0, "retries": 0, "chart_failures": 0,
                "submitted": False}
        self.messages.append({"role": "user", "content": question})
        reminded = False

        for round_no in range(1, MAX_TOOL_ROUNDS + 1):
            response = self.client.messages.create(
                model=MODEL,
                max_tokens=16000,
                system=system_prompt(),
                tools=TOOLS,
                messages=self.messages,
                cache_control={"type": "ephemeral"},
            )
            u = response.usage
            turn["input_tokens"] += (u.input_tokens + (u.cache_creation_input_tokens or 0)
                                     + (u.cache_read_input_tokens or 0))
            turn["output_tokens"] += response.usage.output_tokens
            self.messages.append({"role": "assistant", "content": response.content})

            if response.stop_reason == "refusal":
                turn["answer"] = "The model declined to answer this question."
                break
            if response.stop_reason == "max_tokens":
                turn["warnings"].append("The response was cut off (max_tokens).")

            tool_uses = [b for b in response.content if b.type == "tool_use"]
            if not tool_uses:
                if not reminded:  # ended without submit_answer: ask once for the structured answer
                    reminded = True
                    self.messages.append({"role": "user", "content":
                                          "Please submit your answer with the submit_answer tool."})
                    continue
                turn["answer"] = "\n".join(b.text for b in response.content if b.type == "text")
                turn["warnings"].append("Answer was not submitted with citations.")
                break

            results, submitted = [], None
            for tu in tool_uses:
                if tu.name == "run_sql":
                    on_step("sql", tu.input["purpose"])
                    content, is_error, qid = self._tool_run_sql(tu.input)
                    if qid:
                        turn["query_ids"].append(qid)
                elif tu.name == "make_chart":
                    if turn["chart_failures"] >= 2:
                        content, is_error, cid = ("Charting is unavailable for this question. Continue "
                                                  "without a chart and submit your answer."), True, None
                    else:
                        on_step("chart", tu.input.get("title", "chart"))
                        content, is_error, cid = self._tool_make_chart(tu.input)
                    if cid:
                        turn["chart_ids"].append(cid)
                    elif is_error:
                        turn["chart_failures"] += 1
                elif tu.name == "submit_answer":
                    problems = self._check_answer(tu.input)
                    if problems and turn["retries"] < MAX_ANSWER_RETRIES:
                        turn["retries"] += 1
                        on_step("check", "Answer rejected by checks; model is correcting it")
                        content, is_error = "Answer rejected:\n- " + "\n- ".join(problems), True
                    else:
                        submitted = tu.input
                        turn["warnings"].extend(problems)
                        content, is_error = "Answer received.", False
                else:
                    content, is_error = f"Unknown tool {tu.name}.", True
                if is_error:
                    on_step("error", f"{tu.name}: {content[:300]}")
                results.append({"type": "tool_result", "tool_use_id": tu.id,
                                "content": content, "is_error": is_error})
            if submitted is None and round_no == MAX_TOOL_ROUNDS - 2:
                results.append({"type": "text", "text": "Tool budget nearly used up. Stop exploring and "
                                "call submit_answer now with what the queries so far support."})
            self.messages.append({"role": "user", "content": results})

            if submitted is not None:
                self._finish(turn, submitted)
                break
        else:
            turn["answer"] = turn["answer"] or "Stopped: too many tool calls without a final answer."
            turn["warnings"].append(f"Hit the limit of {MAX_TOOL_ROUNDS} tool rounds.")

        if not turn["query_ids"]:
            turn["warnings"].append("Not grounded: no database query was run for this answer.")
        self.turns.append(turn)
        return turn

    def _check_answer(self, submitted):
        """Deterministic checks on a proposed answer. Returns a list of problems (empty = OK)."""
        problems = []
        if _MARKUP.search(submitted["answer"]):
            problems.append("The answer text contains tool/XML markup (e.g. </answer> or <parameter>). "
                            "Put only the buyer-facing markdown in `answer`.")
        unverified = audit_numbers(submitted["answer"], [q["df"] for q in self.queries.values()])
        if unverified:
            problems.append("These figures do not appear in any query result: " + ", ".join(unverified)
                            + ". Compute them with run_sql (do the arithmetic in SQL) and resubmit.")
        known = valid_cells()
        bad = [(c["response_id"], c["sku_id"]) for c in submitted["cited_cells"]
               if (c["response_id"], c["sku_id"]) not in known]
        if bad:
            problems.append(f"Unknown cited cells {bad[:10]}. Use response_id values from the vendor list.")
        if submitted["enough_data"] and not submitted["cited_cells"] and self.queries:
            problems.append("No cells cited. Cite the (response_id, sku_id) cells behind the answer.")
        return problems

    def _finish(self, turn, submitted):
        turn["submitted"] = True
        turn["answer"] = _MARKUP.split(submitted["answer"])[0].rstrip()  # last resort if retries ran out
        turn["enough_data"] = submitted["enough_data"]
        turn["missing_data"] = submitted["missing_data"]
        known = valid_cells()
        seen = set()
        for c in submitted["cited_cells"]:
            key = (c["response_id"], c["sku_id"])
            if key in seen:
                continue
            seen.add(key)
            if key in known:
                turn["cited_cells"].append(c)
            else:
                turn["warnings"].append(f"Dropped citation to unknown cell {key}.")

    # --- helpers for display / export ------------------------------------------
    def cited_cells_frame(self, turn):
        if not turn["cited_cells"]:
            return pd.DataFrame()
        cited = pd.DataFrame(turn["cited_cells"])
        conn = db.connect()
        data = pd.read_sql("""SELECT response_id, sku_id, vendor_name, description, status, inr_per_unit,
                                     raw_price, raw_currency, raw_unit, source_note, file_name
                              FROM comparison""", conn)
        conn.close()
        return cited.merge(data, on=["response_id", "sku_id"], how="left")


def main():
    load_dotenv()
    question = " ".join(sys.argv[1:]) or "Which vendor is cheapest overall?"
    session = AnalystSession()
    turn = session.ask(question, on_step=lambda kind, text: print(f"  [{kind}] {text}"))
    print("\n" + (turn["answer"] or ""))
    print(f"\nenough_data={turn['enough_data']} missing={turn['missing_data']!r}")
    print(f"queries={turn['query_ids']} charts={turn['chart_ids']} cited={len(turn['cited_cells'])}")
    print(f"warnings={turn['warnings']}")
    print(f"tokens in={turn['input_tokens']} out={turn['output_tokens']}")


if __name__ == "__main__":
    main()
