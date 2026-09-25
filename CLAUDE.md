# RFx Analyst — Project Memory

## What this is
A prototype for the Aerchain PM take-home assignment "Kill the Quote Spreadsheet".
A buyer drafts an RFx with an AI co-pilot, vendors reply in messy formats, the system
extracts every reply into one normalised side-by-side comparison, and the buyer asks
questions in plain language to reach a defensible award decision.

## Fixed decisions
- **Category:** corrugated packaging (boxes, 3-ply/5-ply/7-ply, sheets, partitions, etc.)
- **Scale:** 30 line items, 5 vendors, one quality questionnaire, attached documents
- **Stack:** Python + Streamlit (UI) + SQLite (storage)
- **AI:** real Anthropic API, model `claude-sonnet-5`, via the `anthropic` Python SDK
- **API key:** read `ANTHROPIC_API_KEY` from `.env` (python-dotenv). `.env` is gitignored. Never print or commit it.
- **Virtualenv:** `.venv/` (Python 3.14)
- **Currency of comparison:** INR (₹)

## Hard rules (from the assignment + the user)
1. Stub the plumbing (e.g. fake email/SMTP), but **AI loops must be real**:
   extraction and reasoning always go through the Claude API.
2. **Never hardcode answers** to demo questions. The analyst chat must compute from the database.
3. **Never guess missing prices.** A line a vendor didn't quote is stored as missing (NULL) and
   shown as "Not quoted" — never filled in, averaged, or invented.
4. When the system is unsure (unit ambiguity, low-confidence OCR, currency), it must **flag it
   to the buyer** with the source snippet, not silently decide.

## Dataset (built — `python generate_data.py`, deterministic, seed 42)
- Buyer: Nilgiri Home Appliances Ltd. (fictional), Chakan, Pune. RFQ no. NHA/PKG/RFQ/2026/031.
- `skus.csv`: 30 SKUs CB-001..CB-030 (3/5/7-ply RSC cartons, E-flute die-cut mailers, layer pads, partitions)
  with ply, flute, paper_spec (GSM per layer), BF, total_gsm, L/W/H mm, printing, uom, annual_qty.
- `vendor_replies/`:
  - V1 `01_shree_ganesh_quotation.xlsx`: own layout, no SKU codes, sizes in INCHES, sorted by ply, has ex-GST and incl-GST columns.
  - V2 `02_deccan_kraft_quotation.pdf`: LIST rates; 6% discount only in a footnote on page 2. Freight included.
  - V3 `03_sabarmati_offer_letter.docx`: prices in prose, INCLUDING 18% GST (stated late in the letter).
  - V4 `04_om_sai_rate_card.jpg`: phone photo, tilted, 27/30 lines (no CB-024/025/026), CB-014 price struck out & hand-corrected.
  - V5 `05_coromandel_email.txt`: USD per 100 pcs, mixed line formats.
- `ground_truth.csv`: 150 rows (vendor x SKU). Price = INR per piece, ex-GST, after discount, freight as quoted.
  USD converted at fixed **USD_INR = 88.00** (app must use and show the same rate). Only for scoring — the app must never read it.

## Extraction (built — `python extract.py`)
- `db.py`: SQLite schema in `quotes.db` (gitignored): `skus`, `responses` (per-file terms), `quote_lines` (per file x SKU).
- `extract.py`: one streamed `messages.stream(..., output_format=Extraction)` call per file with Pydantic
  structured output. PDF/image sent natively; Excel -> text with cell refs; Word -> text with [para N]; email -> L001 line numbers.
- Stores values AS WRITTEN (price, ISO currency, unit, pack_size, tax_included, conditions, confidence, matched_by,
  source_note). No conversion, no GST removal, no discount applied — normalisation is a separate deterministic step.
- `validate()` never changes prices; it only flags: unknown/duplicate SKUs, quoted-without-price, missing SKUs
  (recorded as not quoted), and POSSIBLE MISSED MATCH (unmatched vendor lines + unquoted SKUs).
- SKU list sent to Claude includes inch equivalents (a 599 vs 600 mm miss happened once without them).
- Result on 2026-09-25: 150/150 prices correct vs ground_truth.csv after simple normalisation.
- Lesson: model `confidence` stayed >= 0.7 even on the missed match — do not rely on it alone for trust; use deterministic checks.
- Uses max_tokens 64000 with streaming (the Excel file needed ~16k output tokens incl. thinking).

## Normalisation (built — `python normalise.py`, no AI)
- Reads quote_lines + responses, writes `normalised_prices` table and `normalise_log.csv` (gitignored).
- Comparable price = INR per piece/sheet/set, ex-GST, after stated discount. Order: ÷ pack count -> × FX -> ÷ (1+GST) -> × (1-discount).
- `settings.json` is user-editable: fx_to_inr (USD 88.0), default_gst_pct, low_confidence_below, rsc_glue_flap_mm.
- GST rate is read from the vendor's own gst_text ("@ 18%"); default only if absent (and flagged).
- Status per cell: ok | flagged (value usable, assumption listed) | needs_review (NO comparable value) | not_quoted.
  - needs_review: unrecognised unit, per kg (indicative_inr only, never compared), pack-size conflict,
    unit word doesn't fit SKU uom (e.g. "per box" on a pad), unknown currency, or unquoted SKU when vendor had unmatched lines.
  - flagged: GST not stated (Coromandel "GST as applicable"), low confidence, default GST rate used.
- Every step logged in `steps` JSON with the vendor's own words. Current result: 143/143 converted values correct;
  Shree Ganesh CB-027..030 needs_review ("per box" on pads/partitions); Coromandel all flagged (GST not stated).
- Later idea: buyer override (resolve a needs_review cell in the UI, recorded with who/why).

## Buyer screen (built — `streamlit run app.py`, no AI calls)
- Reads quotes.db via one join (normalised_prices + quote_lines + responses + skus). Settings from settings.json.
- Header: RFQ, basis of comparison, FX used. Per vendor: "N/30 quoted" badge (green/orange), "N/30 comparable"
  (red, only if fewer comparable than quoted), freight terms.
- Comparison tab: 30 x 5 grid (st.dataframe, single-cell selection). ok = ₹x.xx; flagged = amber + ⚠;
  low confidence = amber; needs_review = red "Review"; not_quoted = grey "—". Click -> detail panel on the right:
  comparable price, flags, what vendor wrote, source note, numbered conversion steps, vendor terms, original doc
  (image/email shown inline, all downloadable).
- Review Queue tab: only uncertain cells (needs_review first, then flagged/low confidence) with a "Why" column; row click -> same detail.
- `.claude/launch.json` config "rfx-analyst" runs it on port 8501 for the preview browser.
- Known gap: the hand-corrected price (Om Sai CB-014) is not amber — extraction confidence 0.85 is above threshold and the
  model's warning lives at document level. Fix idea: add a per-line `needs_attention` reason to the extraction schema.

## Analyst tab (built — `analyst.py`, also `python analyst.py "question"`)
- Manual tool-use loop on claude-sonnet-5: tools run_sql (read-only), make_chart (only from a query's rows), submit_answer
  (answer, enough_data, missing_data, cited_cells). Max 15 rounds; nudges to submit at round 13.
- SQL guard: `file:...?mode=ro` + sqlite authorizer (SELECT/READ/FUNCTION/RECURSIVE only) + progress-handler step cap.
- `comparison` VIEW (in db.py SCHEMA, recreated on connect) is the flat table the model queries.
- submit_answer is REJECTED (max 2 retries, then shown with warnings) if: any number in the answer isn't in a query result
  (audit_numbers: all numbers incl. counts; skips list ordinals, CB codes, dims, years, Q/C ids; accepts lakh/crore,
  numbers inside text cells, row counts), tool markup leaked into text, unknown cited cells, or no citations.
- Lessons: model summed per-vendor subtotals wrong (off by exactly ₹10 lakh) and miscounted wins → audit catches both.
  A required empty-string `color` param made the model emit garbage 11 times → optional params instead.
- UI: chat with live steps, "Not enough data" banner, charts (Altair), cited cells table, queries+SQL, CSV per query,
  Excel export (answer + cited cells + every query). Verified: VP questionnaire question -> enough_data=false.

## RFx tab (built — `rfx.py`)
- Co-pilot chat (claude-sonnet-5, tools set_header / set_line_items / set_questionnaire / set_terms / load_item_master,
  Pydantic-validated). Current draft incl. buyer edits is sent in <current_draft> each turn. Never invents buyer facts.
- Draft editors: text inputs + st.data_editor keyed by `rfx_ver` (bumped when the model/load changes the base draft).
- Send RFx (simulated): clears replies, writes skus from line items, rfx row, outbox emails (example.* addresses).
- "Vendor replies arrive": extract.process_files (5 parallel API calls) + normalise.run(); ~1.5 min for 5 files.
  Manual upload saves to uploads/ (gitignored) then same pipeline. Sidebar "Reset demo" (popover confirm) -> db.reset_all.
- normalise cross-checks tax_included against the vendor's GST wording (gst_wording_supports) so "GST as applicable"
  is always flagged regardless of how the model read it (it flip-flopped between runs).
- Streamlit: `.streamlit/config.toml` headless=true (skips email prompt). Preview config uses port 8502.

## Extraction cache (built)
- `cache/extractions/<file>.<hash16>.json` — COMMITTED to the repo so the live demo needs no extraction API calls.
- Key = sha256 of file bytes + MODEL + SYSTEM_PROMPT + sku_table(skus) + Extraction JSON schema. Any change -> miss -> live call -> cache written.
- extract.process_files(use_cache=True): hits are stored instantly (label "cached result from <date>"), misses go to Claude in parallel.
  No Anthropic client is created when everything hits, so extraction works without an API key.
- CLI: `--no-cache` forces live calls. App: "Use cached extractions when available" toggle (default on).
- Verified 2026-09-25: full pipeline from cache in <0.1 s with no API key, 143/143 correct on a scratch DB.
- If extract.py's prompt/schema changes, the cache silently misses: re-run extraction once and commit the new files.

## Known gaps / next ideas
- Questionnaire answers are not extracted from replies (fabricated replies contain none) -> knockout filtering is "not enough data".
- Hand-corrected price (Om Sai CB-014) not amber; needs a per-line attention reason in the extraction schema.
- Chat histories live in Streamlit session only (lost on page reload); the sent RFx persists in the DB.
- Buyer override for needs_review cells.

## Planned flow (build in this order)
1. **Setup** — requirements.txt, `.env` loading, SQLite schema.
2. **Dataset** — fabricate RFx (30 lines + questionnaire) and generate the 5 messy vendor files.
3. **RFx co-pilot** — chat with Claude to draft scope, line items, questionnaire, terms.
4. **Send (stubbed)** — fake "email out" that just logs to the DB.
5. **Extraction** — Claude reads each file (text/PDF/DOCX/XLSX/image) → structured JSON
   (price, unit, currency, confidence, source quote) → saved to SQLite.
6. **Normalisation** — convert units & currency with explicit, visible rules; flag anything uncertain.
7. **Comparison view** — 30 lines × 5 vendors, with questionnaire answers, flags, and source links.
8. **Analyst chat** — Claude uses tools (SQL queries over SQLite) to answer questions with text,
   tables, charts and CSV/Excel exports. E.g. "split cheapest per line, only among vendors who
   passed the quality questionnaire".
9. **Deliverables** — live demo, recorded walkthrough, one-page decision note.

## How to run
See README.md. Short version: `pip install -r requirements.txt` (pinned versions), `.env` with ANTHROPIC_API_KEY,
`.venv/bin/streamlit run app.py`. The user often runs commands in fresh terminal tabs without the venv active —
suggest `.venv/bin/...` paths.

## Working notes for Claude
- User is a beginner: explain each step briefly and say how to run it.
- Don't build ahead of what the user asked for.
