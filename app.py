"""
app.py — Quote Killer buyer screen (Streamlit).

Tabs
  RFx           co-pilot chat + editable draft, simulated send, vendor replies in (folder or upload)
  Comparison    30 x N grid of normalised prices; click a cell for its source and conversion
  Review Queue  only the uncertain cells
  Analyst       questions in plain English, answered by Claude via read-only SQL
The Comparison and Review Queue tabs make no AI calls.

Run:  streamlit run app.py
"""

import io
import json
import re
from pathlib import Path

import altair as alt
import pandas as pd
import streamlit as st
from dotenv import load_dotenv

import db
import extract
import normalise
import rfx
from analyst import AnalystSession
from normalise import load_settings

load_dotenv()  # ANTHROPIC_API_KEY for the Analyst tab

REPLIES_DIR = Path(__file__).parent / "vendor_replies"
UPLOAD_DIR = Path(__file__).parent / "uploads"
UPLOAD_TYPES = ["pdf", "xlsx", "docx", "jpg", "jpeg", "png", "webp", "txt", "eml"]

# Cell colours: semi-transparent so they read in both light and dark themes.
AMBER = "background-color: rgba(255, 176, 0, 0.35)"
RED = "background-color: rgba(230, 60, 60, 0.30)"
GREY = "color: rgba(128, 128, 128, 0.9)"

st.set_page_config(page_title="Quote Killer", page_icon="📦", layout="wide")


# ---------------------------------------------------------------------------
# Data
# ---------------------------------------------------------------------------
def load_cells():
    conn = db.connect()
    df = pd.read_sql("""
        SELECT n.response_id, n.sku_id, n.status, n.inr_per_unit, n.indicative_inr, n.fx_rate,
               n.steps, n.flags, n.normalised_at,
               l.quoted, l.price, l.currency, l.unit, l.pack_size, l.tax_included, l.conditions,
               l.confidence, l.matched_by, l.source_note,
               r.file_name, r.file_type, r.vendor_name, r.freight, r.gst_text, r.discount_text,
               r.discount_pct, r.payment_terms, r.validity, r.delivery, r.warnings,
               s.description, s.uom, s.ply
        FROM normalised_prices n
        JOIN quote_lines l USING (response_id, sku_id)
        JOIN responses r USING (response_id)
        JOIN skus s USING (sku_id)
        ORDER BY r.file_name, n.sku_id""", conn)
    conn.close()
    return df


def short_name(vendor_name):
    """'SHREE GANESH CORRUGATORS PVT. LTD.' -> 'Shree Ganesh'"""
    words = (vendor_name or "Unknown").replace("(", " ").split()
    return " ".join(words[:2]).title()


def is_low_confidence(row, threshold):
    return row.status != "not_quoted" and row.confidence is not None and row.confidence < threshold


def is_uncertain(row, threshold):
    return row.status in ("flagged", "needs_review") or is_low_confidence(row, threshold)


def cell_text(row):
    if row.status == "not_quoted":
        return "—"
    if row.status == "needs_review":
        return "Review"
    text = f"₹{row.inr_per_unit:,.2f}"
    return text + " ⚠" if row.status == "flagged" else text


def cell_style(row, threshold):
    if row.status == "needs_review":
        return RED
    if row.status == "flagged" or is_low_confidence(row, threshold):
        return AMBER
    if row.status == "not_quoted":
        return GREY
    return ""


# ---------------------------------------------------------------------------
# Cell detail panel
# ---------------------------------------------------------------------------
def yes_no_unknown(v):
    return {1: "yes", 0: "no"}.get(v, "not stated")


def show_cell_detail(cell, threshold):
    st.subheader(f"{cell.sku_id} · {short_name(cell.vendor_name)}")
    st.caption(f"{cell.description} · {cell.ply}-ply · priced per {cell.uom}")

    status_colour = {"ok": "green", "flagged": "orange", "needs_review": "red", "not_quoted": "gray"}
    st.badge(cell.status.replace("_", " "), color=status_colour[cell.status])

    if cell.status in ("ok", "flagged"):
        st.metric("Comparable price (ex-GST, after discount)", f"₹{cell.inr_per_unit:,.2f} per {cell.uom}")
    elif cell.indicative_inr:
        st.metric("Indicative only — not used for comparison", f"≈ ₹{cell.indicative_inr:,.2f}")

    for flag in json.loads(cell.flags or "[]"):
        st.warning(flag)
    if is_low_confidence(cell, threshold):
        st.warning(f"Extraction confidence {cell.confidence:.2f} is below {threshold}.")

    st.markdown("**What the vendor wrote**")
    if cell.quoted:
        st.markdown(
            f"- Price: **{cell.price:g} {cell.currency or ''}** {cell.unit or ''}\n"
            f"- Pack size: {int(cell.pack_size) if pd.notna(cell.pack_size) else 'not stated'} · "
            f"GST included: {yes_no_unknown(cell.tax_included)}\n"
            f"- Matched to SKU by: {cell.matched_by} · confidence {cell.confidence:.2f}")
    conditions = json.loads(cell.conditions or "[]")
    if conditions:
        st.markdown("- Conditions: " + "; ".join(conditions))
    st.markdown("**Source**")
    st.info(f"{cell.file_name}: {cell.source_note}")

    steps = json.loads(cell.steps or "[]")
    if steps:
        st.markdown("**Conversion steps**")
        st.markdown("\n".join(f"{i}. {s}" for i, s in enumerate(steps, start=1)))

    with st.expander("Vendor terms (apply to all lines)"):
        st.markdown(
            f"- Freight: {cell.freight or 'not stated'}\n"
            f"- GST: {cell.gst_text or 'not stated'}\n"
            f"- Discount: {cell.discount_text or 'none stated'}\n"
            f"- Payment: {cell.payment_terms or 'not stated'}\n"
            f"- Validity: {cell.validity or 'not stated'} · Delivery: {cell.delivery or 'not stated'}")

    path = REPLIES_DIR / cell.file_name
    if not path.exists():
        path = UPLOAD_DIR / cell.file_name
    if path.exists():
        with st.expander("Original document"):
            if cell.file_type == "image":
                st.image(str(path))
            elif cell.file_type == "email":
                st.code(path.read_text(encoding="utf-8"), language=None)
            st.download_button("Download original", path.read_bytes(), file_name=cell.file_name,
                               key=f"dl-{cell.response_id}-{cell.sku_id}")


# ---------------------------------------------------------------------------
# Analyst rendering
# ---------------------------------------------------------------------------
def render_chart(spec, df):
    nominal_x = spec["chart_type"] in ("bar", "grouped_bar") or not pd.api.types.is_numeric_dtype(df[spec["x"]])
    enc = {"x": alt.X(f"{spec['x']}:{'N' if nominal_x else 'Q'}", sort=None, title=spec["x"]),
           "y": alt.Y(f"{spec['y']}:Q", title=spec["y"]),
           "tooltip": list(df.columns)}
    if spec["color"]:
        enc["color"] = alt.Color(f"{spec['color']}:N", title=spec["color"])
        if spec["chart_type"] == "grouped_bar":
            enc["xOffset"] = f"{spec['color']}:N"
    mark = {"bar": "mark_bar", "grouped_bar": "mark_bar", "line": "mark_line", "scatter": "mark_circle"}
    chart = getattr(alt.Chart(df), mark[spec["chart_type"]])().encode(**enc).properties(
        title=spec["title"], height=320)
    st.altair_chart(chart, width="stretch")


def excel_bytes(sheets):
    """sheets: dict of sheet name -> DataFrame. Returns an .xlsx file as bytes."""
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        for name, frame in sheets.items():
            frame.to_excel(writer, sheet_name=name[:31], index=False)
    return buf.getvalue()


def render_turn(session, turn, idx):
    with st.chat_message("user"):
        st.markdown(turn["question"])
    with st.chat_message("assistant"):
        if not turn["enough_data"]:
            st.warning(f"**Not enough data.** {turn['missing_data']}")
        st.markdown(turn["answer"] or "_No answer._")
        for cid in turn["chart_ids"]:
            spec = session.charts[cid]
            render_chart(spec, session.queries[spec["query_id"]]["df"])
        for w in turn["warnings"]:
            st.error(f"Check: {w}")

        unverified = any("do not appear in any query result" in w for w in turn["warnings"])
        if turn["query_ids"]:
            check = ("✗ some figures unverified" if unverified else "✓ every figure matched a query result") \
                if turn.get("submitted") else "no answer submitted"
            st.caption(f"Grounded in {len(turn['query_ids'])} queries ({', '.join(turn['query_ids'])}) · "
                       f"{len(turn['cited_cells'])} cited cells · {check} · "
                       f"{turn['input_tokens']:,} in / {turn['output_tokens']:,} out tokens")

        cited = session.cited_cells_frame(turn)
        if not cited.empty:
            with st.expander(f"Cited cells ({len(cited)})"):
                st.dataframe(cited, hide_index=True,
                             column_config={"inr_per_unit": st.column_config.NumberColumn("₹ comparable", format="₹%.2f")})
        with st.expander(f"Queries used ({len(turn['query_ids'])})"):
            for qid in turn["query_ids"]:
                q = session.queries[qid]
                st.markdown(f"**{qid}** — {q['purpose']}")
                st.code(q["sql"], language="sql")
                st.dataframe(q["df"], hide_index=True)
                st.download_button(f"{qid} as CSV", q["df"].to_csv(index=False).encode("utf-8"),
                                   file_name=f"{qid}.csv", mime="text/csv", key=f"csv-{idx}-{qid}")

        summary = pd.DataFrame([
            ("Question", turn["question"]), ("Answer", turn["answer"]),
            ("Enough data", "yes" if turn["enough_data"] else "no"), ("Missing data", turn["missing_data"]),
            ("Warnings", "; ".join(turn["warnings"])), ("Queries", ", ".join(turn["query_ids"]))],
            columns=["Field", "Value"])
        sheets = {"Answer": summary, "Cited cells": cited}
        sheets.update({qid: session.queries[qid]["df"] for qid in turn["query_ids"]})
        c1, c2 = st.columns(2)
        c1.download_button("Export answer + data (Excel)", excel_bytes(sheets), file_name=f"analysis_{idx + 1}.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                           key=f"xlsx-{idx}")
        if not cited.empty:
            c2.download_button("Cited cells (CSV)", cited.to_csv(index=False).encode("utf-8"),
                               file_name=f"cited_cells_{idx + 1}.csv", mime="text/csv", key=f"cited-{idx}")


# ---------------------------------------------------------------------------
# RFx tab helpers
# ---------------------------------------------------------------------------
INT_COLS = ["ply", "bursting_factor", "total_gsm", "length_mm", "width_mm", "height_mm", "annual_qty"]


def frame(rows, columns, int_cols=()):
    df = pd.DataFrame(rows, columns=columns)
    for c in int_cols:
        df[c] = pd.to_numeric(df[c], errors="coerce").astype("Int64")
    return df


def records(df):
    """DataFrame from st.data_editor -> list of plain dicts (NaN -> None, numpy ints -> int)."""
    df = df.dropna(how="all")
    return json.loads(df.to_json(orient="records"))


def set_draft(draft):
    """Replace the draft shown in the editors (after a co-pilot change or a load)."""
    st.session_state.rfx_base = draft
    st.session_state.rfx_ver = st.session_state.get("rfx_ver", 0) + 1


def render_draft_editor(base, ver):
    """Editable draft. Returns the current draft including the buyer's edits."""
    c1, c2 = st.columns([3, 1])
    title = c1.text_input("Title", base["title"], key=f"title-{ver}")
    rfx_no = c2.text_input("RFx no.", base["rfx_no"], key=f"rfxno-{ver}")
    scope = st.text_area("Scope", base["scope"], height=120, key=f"scope-{ver}")

    st.markdown(f"**Line items** ({len(base['line_items'])})")
    items = st.data_editor(
        frame(base["line_items"], rfx.ITEM_COLUMNS, INT_COLS), num_rows="dynamic", key=f"items-{ver}",
        hide_index=True, height=min(len(base["line_items"]) + 2, 12) * 35 + 3,
        column_config={"item_type": st.column_config.SelectboxColumn(options=["RSC", "DIECUT", "PAD", "PARTITION"]),
                       "uom": st.column_config.SelectboxColumn(options=["piece", "sheet", "set"]),
                       "description": st.column_config.TextColumn(width="medium")})
    if st.button("Load item master (skus.csv)", help="Fill line items with the standard 30 corrugated SKUs"):
        draft = {**base, "title": title, "rfx_no": rfx_no, "scope": scope, "line_items": rfx.load_item_master()}
        set_draft(draft)
        st.rerun()

    st.markdown("**Supplier questionnaire**")
    questions = st.data_editor(
        frame(base["questionnaire"], rfx.QUESTION_COLUMNS), num_rows="dynamic", key=f"q-{ver}", hide_index=True,
        column_config={"answer_type": st.column_config.SelectboxColumn(options=["yes_no", "number", "text", "choice"]),
                       "knockout": st.column_config.CheckboxColumn(help="Failing disqualifies the vendor"),
                       "question": st.column_config.TextColumn(width="large")})
    st.markdown("**Commercial terms**")
    terms = st.data_editor(frame(base["terms"], rfx.TERM_COLUMNS), num_rows="dynamic", key=f"t-{ver}",
                           hide_index=True, column_config={"value": st.column_config.TextColumn(width="large")})

    q_records = records(questions)
    for q in q_records:
        q["knockout"] = bool(q.get("knockout"))
    return {"rfx_no": rfx_no, "title": title, "scope": scope, "line_items": records(items),
            "questionnaire": q_records, "terms": records(terms)}


def run_pipeline(paths, use_cache=True):
    """Extract (cache or Claude, in parallel) + normalise (Python) the given files; remember a summary."""
    conn = db.connect()
    skus = [dict(r) for r in conn.execute("SELECT * FROM skus ORDER BY sku_id")]
    conn.close()
    if not skus:
        st.error("Send the RFx first: its line items are what vendor replies are matched against.")
        return
    lines = []
    with st.status(f"Reading {len(paths)} vendor repl{'y' if len(paths) == 1 else 'ies'}…",
                   expanded=True) as status:
        def progress(path, msg):
            status.write(f"**{path.name}** — {msg}")
            if not msg.startswith("reading"):
                lines.append(f"{path.name} — {msg}")
        outcomes = extract.process_files(paths, skus, on_progress=progress, use_cache=use_cache)
        status.write("Normalising prices (plain Python)…")
        normalise.run()
        failed = sum(1 for _, _, err in outcomes if err)
        status.update(label=f"Processed {len(outcomes) - failed} of {len(outcomes)} files",
                      state="error" if failed else "complete")
    st.session_state.last_pipeline = lines
    st.session_state.pop("analyst", None)  # new data: start a fresh analyst conversation


def render_rfx_tab(sent_row, responses):
    if "rfx_base" not in st.session_state:
        set_draft(rfx.load_sent_draft()[0] or rfx.empty_draft())
    if "rfx_chat" not in st.session_state:
        st.session_state.rfx_chat = []
    base, ver = st.session_state.rfx_base, st.session_state.rfx_ver

    chat_col, draft_col = st.columns([2, 3], gap="large")
    with draft_col:  # rendered first so the chat sees the buyer's latest edits
        st.subheader("1 · Draft")
        draft = render_draft_editor(base, ver)

        st.subheader("2 · Send")
        recipients = records(st.data_editor(pd.DataFrame(st.session_state.get("recipients", rfx.DEFAULT_RECIPIENTS)),
                                            num_rows="dynamic", hide_index=True, key="recipients-editor"))
        if sent_row:
            st.badge(f"Sent {sent_row['sent_at']} UTC to {len(json.loads(sent_row['recipients']))} vendors "
                     "(simulated)", color="green")
        problems = rfx.check_draft(draft, recipients)
        for pr in problems:
            st.caption(f"⚠ {pr}")
        if st.button("Send RFx", type="primary", disabled=bool(problems),
                     help="Simulated: saves the RFx and writes emails to the outbox. Replaces any current replies."):
            rfx.send_rfx(draft, recipients)
            st.session_state.recipients = recipients
            set_draft(draft)
            st.session_state.pop("analyst", None)
            st.rerun()
        if sent_row:
            conn = db.connect()
            outbox = conn.execute("SELECT vendor, email, subject, body FROM outbox WHERE rfx_id = ?",
                                  (sent_row["rfx_id"],)).fetchall()
            conn.close()
            with st.expander(f"Outbox ({len(outbox)} emails, not actually sent)"):
                for m in outbox:
                    st.markdown(f"**To:** {m['vendor']} <{m['email']}>  \n**Subject:** {m['subject']}")
                with st.container(height=300):
                    st.code(outbox[0]["body"] if outbox else "", language=None)

        st.subheader("3 · Replies")
        done = set(responses.file_name) if not responses.empty else set()
        waiting = sorted(p for p in REPLIES_DIR.iterdir() if p.is_file() and not p.name.startswith(".")
                         and p.name not in done)
        st.caption(f"{len(done)} replies processed · {len(waiting)} waiting in vendor_replies/")
        use_cache = st.toggle("Use cached extractions when available", value=True,
                              help="Reuses a saved Claude extraction when the file, prompt, model and line items "
                                   "are unchanged (cache/extractions/). Turn off to call the API live.")
        if st.button(f"Vendor replies arrive ({len(waiting)})", disabled=not sent_row or not waiting,
                     help="Runs every new file in vendor_replies/ through extraction (Claude) and normalisation."):
            run_pipeline(waiting, use_cache)
            st.rerun()
        uploads = st.file_uploader("Or upload replies (PDF, Excel, Word, image, email text)", type=UPLOAD_TYPES,
                                   accept_multiple_files=True, key=f"up-{st.session_state.get('up_ver', 0)}")
        if uploads and st.button(f"Process {len(uploads)} uploaded file(s)", disabled=not sent_row):
            UPLOAD_DIR.mkdir(exist_ok=True)
            paths = []
            for f in uploads:
                path = UPLOAD_DIR / Path(f.name).name
                path.write_bytes(f.getvalue())
                paths.append(path)
            run_pipeline(paths, use_cache)
            st.session_state.up_ver = st.session_state.get("up_ver", 0) + 1
            st.rerun()
        if st.session_state.get("last_pipeline"):
            with st.expander("Last run", expanded=True):
                st.markdown("\n".join(f"- {line}" for line in st.session_state.last_pipeline))
        if not responses.empty:
            st.dataframe(responses, hide_index=True)

    with chat_col:
        st.subheader("RFx co-pilot")
        st.caption("Describe what you need; the co-pilot fills the draft. You can edit any table directly.")
        for role, text in st.session_state.rfx_chat:
            with st.chat_message(role):
                st.markdown(text)
        prompt = st.chat_input("e.g. Quote for our standard 30 corrugated SKUs, delivered to Chakan…", key="rfx_input")
        if prompt:
            if "copilot" not in st.session_state:
                st.session_state.copilot = rfx.RfxCopilot()
            st.session_state.rfx_chat.append(("user", prompt))
            with st.chat_message("user"):
                st.markdown(prompt)
            with st.chat_message("assistant"):
                with st.status("Drafting…", expanded=True) as status:
                    working = json.loads(json.dumps(draft))
                    try:
                        reply, changes = st.session_state.copilot.chat(
                            prompt, working, on_change=lambda section, summary: status.write(f"✏️ {summary}"))
                        status.update(label="Draft updated" if changes else "Done", state="complete", expanded=False)
                    except Exception as e:
                        status.update(label="Failed", state="error")
                        reply, changes = f"The co-pilot hit an error: {e}", []
            note = ("\n\n_" + " · ".join(changes) + "_") if changes else ""
            st.session_state.rfx_chat.append(("assistant", (reply or "Done.") + note))
            if changes:
                set_draft(working)
            st.rerun()


# ---------------------------------------------------------------------------
# Page
# ---------------------------------------------------------------------------
settings = load_settings()
threshold = settings["low_confidence_below"]
cells = load_cells()
_, sent_row = rfx.load_sent_draft()
_conn = db.connect()
responses = pd.read_sql("SELECT file_name, vendor_name, file_type, extracted_at FROM responses ORDER BY file_name",
                        _conn)
_conn.close()

with st.sidebar:
    st.header("Pipeline")
    st.markdown(f"- RFx: {'sent ' + sent_row['sent_at'][:16].replace('T', ' ') if sent_row else 'draft'}\n"
                f"- Replies processed: {len(responses)}\n"
                f"- Priced cells: {int(cells.status.isin(['ok', 'flagged']).sum()) if not cells.empty else 0}")
    with st.popover("Reset demo", width="stretch"):
        st.write("Deletes the RFx, outbox, all extracted replies, uploads and chats. "
                 "Files in vendor_replies/ are kept.")
        if st.button("Yes, reset everything", type="primary"):
            _conn = db.connect()
            db.reset_all(_conn)
            _conn.close()
            normalise.run()
            if UPLOAD_DIR.exists():
                for f in UPLOAD_DIR.iterdir():
                    if f.is_file():
                        f.unlink()
            for key in list(st.session_state.keys()):
                del st.session_state[key]
            st.rerun()

st.title("Quote Killer")
fx = ", ".join(f"{k} ₹{v:g}" for k, v in settings["fx_to_inr"].items() if k != "INR")
rfx_label = f"RFQ {sent_row['rfx_no']}" if sent_row else "No RFx sent yet"
st.caption(f"{rfx_label} · prices compared as INR per piece/sheet/set, ex-GST, after stated discounts · "
           f"FX {fx} (settings.json)")

n_uncertain = sum(is_uncertain(c, threshold) for c in cells.itertuples())
tab_rfx, tab_grid, tab_queue, tab_analyst = st.tabs(
    ["RFx", "Comparison", f"Review Queue ({n_uncertain})", "Analyst"])

with tab_rfx:
    render_rfx_tab(sent_row, responses)

NO_DATA = "No vendor replies processed yet. In the RFx tab: send the RFx, then click **Vendor replies arrive**."

if not cells.empty:
    # Vendor columns in a fixed order, with unique short names
    vendors = cells[["response_id", "vendor_name", "file_name"]].drop_duplicates().sort_values("file_name")
    labels = {}
    for r in vendors.itertuples():
        name = short_name(r.vendor_name)
        labels[r.response_id] = name if name not in labels.values() else f"{name} ({r.file_name[:2]})"

# --- Comparison grid ----------------------------------------------------------
with tab_grid:
    if cells.empty:
        st.info(NO_DATA)
    else:
        badge_cols = st.columns(len(vendors))
        for col, r in zip(badge_cols, vendors.itertuples()):
            v = cells[cells.response_id == r.response_id]
            total = len(v)
            quoted = int(v.quoted.sum())
            comparable = int(v.status.isin(["ok", "flagged"]).sum())
            with col:
                st.markdown(f"**{labels[r.response_id]}**")
                st.badge(f"{quoted}/{total} quoted", color="green" if quoted == total else "orange")
                if comparable < quoted:
                    st.badge(f"{comparable}/{total} comparable", color="red")
                freight = re.sub(r"^\s*freight\s*[:\-]?\s*", "", v.freight.iloc[0] or "", flags=re.I) or "not stated"
                st.caption(f"Freight: {freight[0].lower() + freight[1:]}")

        sku_ids = sorted(cells.sku_id.unique())
        by_key = {(c.response_id, c.sku_id): c for c in cells.itertuples()}
        descriptions = cells.drop_duplicates("sku_id").set_index("sku_id").description

        grid = pd.DataFrame({"SKU": sku_ids, "Item": [descriptions[s] for s in sku_ids]})
        styles = pd.DataFrame("", index=grid.index, columns=grid.columns)
        for rid, label in labels.items():
            grid[label] = [cell_text(by_key[(rid, s)]) for s in sku_ids]
            styles[label] = [cell_style(by_key[(rid, s)], threshold) for s in sku_ids]

        st.caption("Click a price to see where it came from. "
                   "Amber = converted but uncertain (⚠ assumption or low confidence) · "
                   "Red = needs review, no comparable price · — = not quoted")
        left, right = st.columns([3, 2], gap="large")
        with left:
            event = st.dataframe(
                grid.style.apply(lambda _: styles, axis=None),
                hide_index=True,
                height=(len(grid) + 1) * 35 + 3,
                on_select="rerun",
                selection_mode="single-cell",
                key="grid",
                column_config={"Item": st.column_config.TextColumn(width="medium")},
            )
        with right:
            selected = event.selection.cells if event else []
            label_to_rid = {v: k for k, v in labels.items()}
            if selected and selected[0][1] in label_to_rid:
                row_i, col_name = selected[0]
                show_cell_detail(by_key[(label_to_rid[col_name], sku_ids[row_i])], threshold)
            else:
                st.info("Select a price cell to see its source document, what the vendor wrote, "
                        "and every conversion step.")

# --- Review queue -------------------------------------------------------------
with tab_queue:
    if cells.empty:
        st.info(NO_DATA)
    else:
        queue = [c for c in cells.itertuples() if is_uncertain(c, threshold)]
        if not queue:
            st.success("Nothing to review.")
        else:
            order = {"needs_review": 0, "flagged": 1, "ok": 2}
            queue.sort(key=lambda c: (order.get(c.status, 3), c.file_name, c.sku_id))
            table = pd.DataFrame([{
                "Status": c.status.replace("_", " "),
                "Vendor": labels[c.response_id],
                "SKU": c.sku_id,
                "Why": " · ".join(json.loads(c.flags or "[]")
                                  + ([f"confidence {c.confidence:.2f}"] if is_low_confidence(c, threshold) else [])),
                "Vendor wrote": f"{c.price:g} {c.currency or ''} {c.unit or ''}".strip() if c.quoted else "—",
                "Comparable ₹": c.inr_per_unit if pd.notna(c.inr_per_unit) else float("nan"),
                "Item": c.description,
            } for c in queue])

            counts = table.Status.value_counts()
            st.caption(" · ".join(f"{n} {s}" for s, n in counts.items())
                       + " — needs review cells are excluded from comparisons until resolved.")
            q_left, q_right = st.columns([3, 2], gap="large")
            with q_left:
                q_event = st.dataframe(
                    table, hide_index=True, on_select="rerun", selection_mode="single-row", key="queue",
                    height=min(len(table) + 1, 20) * 35 + 3,
                    column_config={"Comparable ₹": st.column_config.NumberColumn(format="₹%.2f"),
                                   "Why": st.column_config.TextColumn(width="large")},
                )
            with q_right:
                rows = q_event.selection.rows if q_event else []
                if rows:
                    show_cell_detail(queue[rows[0]], threshold)
                else:
                    st.info("Select a row to see its source and conversion.")

# --- Analyst ------------------------------------------------------------------
STEP_ICONS = {"sql": "🔎", "chart": "📊", "check": "↩️"}

with tab_analyst:
    if cells.empty:
        st.info(NO_DATA)
    else:
        st.caption("Ask in plain English. Claude answers only by querying the database above; every figure "
                   "is checked against the query results, and it says when there is not enough data.")
        if "analyst" not in st.session_state:
            try:
                st.session_state.analyst = AnalystSession()
            except Exception as e:  # e.g. no API key
                st.error(f"Analyst unavailable: {e}. Check ANTHROPIC_API_KEY in .env.")
        session = st.session_state.get("analyst")

        if session and session.messages and st.button("New conversation"):
            st.session_state.analyst = AnalystSession()
            st.rerun()

        for i, turn in enumerate(session.turns if session else []):
            render_turn(session, turn, i)

        question = st.chat_input("e.g. What is the total annual spend if we award each line to the cheapest vendor?",
                                 key="analyst_input")
        if question and session:
            with st.chat_message("user"):
                st.markdown(question)
            with st.chat_message("assistant"):
                with st.status("Analysing…", expanded=True) as status:
                    try:
                        session.ask(question, on_step=lambda kind, text: status.write(f"{STEP_ICONS.get(kind, '•')} {text}"))
                        status.update(label="Done", state="complete", expanded=False)
                    except Exception as e:
                        status.update(label="Failed", state="error")
                        st.error(f"The Analyst hit an error: {e}")
                        st.stop()
            st.rerun()
