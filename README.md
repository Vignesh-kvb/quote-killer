# Quote Killer

A prototype for the Aerchain take-home *"Kill the Quote Spreadsheet"*: a buyer drafts an RFx with an
AI co-pilot, five vendors reply in whatever format they like, the system reads every reply into one
side-by-side comparison, and the buyer questions the result in plain English.

Category: **corrugated packaging** — 30 line items, 5 vendors, a supplier questionnaire.
Stack: Python, Streamlit, SQLite, Claude (`claude-sonnet-5`) via the Anthropic API.

## Quick start

```bash
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
echo "ANTHROPIC_API_KEY=sk-ant-..." > .env      # your own key; .env is gitignored
streamlit run app.py                             # then open http://localhost:8501
```

The demo runs without an API key for extraction (cached results are committed, see below), but the
RFx co-pilot and the Analyst always call Claude live.

## Demo walkthrough

1. **Reset demo** (sidebar) to start from a blank project.
2. **RFx tab** – tell the co-pilot what you need, e.g. *"Annual rate contract for our standard 30
   corrugated SKUs for the Chakan plant, delivered, ISO 9001 and a lab certificate per lot mandatory,
   payment 45 days"*. It fills title, scope, line items, questionnaire (with knockout questions) and
   terms. Every section is an editable table.
3. **Send RFx** – simulated: the RFx is saved and one email per vendor goes to an outbox you can read.
4. **Vendor replies arrive** – the five files in `vendor_replies/` are extracted and normalised
   (cached results by default; switch the toggle off to call Claude live, ~1.5 min). You can also
   upload your own reply files.
5. **Comparison** – 30 × 5 grid of comparable prices. Amber = converted with an assumption, red =
   needs review (no comparable price), — = not quoted. Click any cell for the vendor's original words,
   the source location, every conversion step and the original document.
6. **Review Queue** – only the uncertain cells, with the reason.
7. **Analyst** – ask e.g. *"What if we split it, cheapest per line, but only among vendors who cleared
   the quality questionnaire?"* (answer: not enough data – replies contain no questionnaire answers)
   or *"Compare single-vendor awards with splitting to the cheapest per line"*. Answers cite the cells
   and queries they use and export to CSV / Excel.

## The vendor replies (the ugly edges)

| File | Format | Trap |
|---|---|---|
| `01_shree_ganesh_quotation.xlsx` | Excel, vendor's own layout | no item codes, sizes in inches, ex- and incl-GST columns, pads priced "per box" |
| `02_deccan_kraft_quotation.pdf` | PDF on letterhead | list prices; 6% discount only in a footnote on page 2 |
| `03_sabarmati_offer_letter.docx` | Word letter | prices in prose, **including** GST (stated late in the letter) |
| `04_om_sai_rate_card.jpg` | phone photo, tilted | 27 of 30 lines; one price struck out and hand-corrected |
| `05_coromandel_email.txt` | plain email | USD, per 100 pieces, "GST as applicable" |

`ground_truth.csv` holds the correct INR price for every cell. It is used only to measure accuracy;
the app never reads it. Latest measured result: **143 / 143** converted prices correct, the 3 missing
lines left empty, 4 ambiguous cells sent to review instead of guessed.

## How it earns trust

- **Extraction records, never converts.** Claude returns each price exactly as written (price,
  currency, unit, pack size, GST treatment, confidence, source location + quote) as validated JSON.
  Missing values stay null.
- **Normalisation is plain Python** (`normalise.py`): pack size → currency (editable rate in
  `settings.json`) → remove GST (using the rate the vendor wrote) → stated discount. Every step is
  logged and shown on screen. Ambiguous units, per-kg prices, unknown currencies and conflicting pack
  sizes become **needs review** with no number, never a guess. The model's GST reading is cross-checked
  against the vendor's own wording.
- **The Analyst can only query.** Read-only SQL (authorizer + read-only connection + step limit),
  charts only from query results, and every number in an answer must appear in a query result or the
  answer is sent back to be recomputed in SQL. It must say "not enough data" when that is true.

## Project layout

| File | What it does |
|---|---|
| `app.py` | Streamlit app: RFx, Comparison, Review Queue, Analyst tabs |
| `rfx.py` | RFx co-pilot (Claude tool use) and the simulated send |
| `extract.py` | reads vendor files with Claude into SQLite; extraction cache |
| `normalise.py` | converts every price to INR per unit, ex-GST, with a logged audit trail |
| `analyst.py` | Analyst tool loop: `run_sql`, `make_chart`, `submit_answer`, number audit |
| `db.py` | SQLite schema (`quotes.db`) and the `comparison` view |
| `generate_data.py` | fabricates `skus.csv`, `vendor_replies/` and `ground_truth.csv` (seeded, repeatable) |
| `settings.json` | FX rates, default GST, confidence threshold |
| `cache/extractions/` | saved Claude extractions (committed) |

## Command line

```bash
python generate_data.py                      # rebuild the dataset (same output every time)
python extract.py                            # extract new files in vendor_replies/ (uses cache)
python extract.py --force --no-cache         # re-extract everything with live API calls
python normalise.py                          # recompute comparable prices + normalise_log.csv
python analyst.py "Which vendor is cheapest on 5-ply cartons?"
```

## Extraction cache

Each extraction is saved as `cache/extractions/<file>.<hash>.json`. The hash covers the file's bytes,
the model, the extraction prompt, the output schema and the RFx line items, so a cached result is
reused only when nothing that could change the answer has changed; anything else is a cache miss and
Claude is called. Cached results are real model outputs from earlier runs and the app labels them
"cached result from <date>". Delete the folder, or untick the toggle / pass `--no-cache`, to force
live extraction.

## Known limitations

- Questionnaire answers are not extracted (the sample replies contain none), so knockout filtering
  answers "not enough data".
- Freight is compared "as quoted": one vendor includes delivery, the others don't; no freight costs
  are invented.
- The hand-corrected price on the photo is noted in its source but not highlighted amber.
- Buyers cannot yet resolve a needs-review cell in the UI.
- Chat histories live in the browser session; sending is simulated (no real email).
