"""
normalise.py — turn every extracted price into a comparable number, in plain Python (no AI).

Comparable price = INR per piece (per sheet for pads, per set for partitions),
                   EXCLUDING GST, AFTER any stated discount.

Every cell gets a status:
  ok            converted with no assumptions
  flagged       converted, but relies on something the buyer should know (e.g. GST not stated)
  needs_review  NOT converted: the unit is ambiguous, per-kg, conflicting, or the currency is unknown.
                No comparable price is produced; the buyer must resolve it.
  not_quoted    the vendor gave no price

Every conversion step is logged: in the `normalised_prices` table (column `steps`)
and in normalise_log.csv.

Settings (edit settings.json, then re-run):
  fx_to_inr            exchange rates, e.g. "USD": 88.0
  default_gst_pct      used only if a GST-inclusive document doesn't state the rate
  low_confidence_below extraction confidence below this is flagged
  rsc_glue_flap_mm     used only for the indicative weight of per-kg quotes

Run:  python normalise.py
"""

import csv
import json
import re
from datetime import datetime, timezone
from pathlib import Path

import db

ROOT = Path(__file__).parent
SETTINGS_PATH = ROOT / "settings.json"
LOG_PATH = ROOT / "normalise_log.csv"

PIECE_WORDS = {"pc", "pcs", "piece", "pieces", "no", "nos", "number", "numbers", "each", "ea", "unit", "units"}
BOX_WORDS = {"box", "boxes", "carton", "cartons", "ctn", "ctns"}
SHEET_WORDS = {"sheet", "sheets", "pad", "pads"}
SET_WORDS = {"set", "sets"}
KG_WORDS = {"kg", "kgs", "kilo", "kilos", "kilogram", "kilograms"}
CURRENCY_TOKENS = r"(rs\.?|inr|usd|us\$|\$|₹|rupees?|dollars?)"


def load_settings():
    return json.loads(SETTINGS_PATH.read_text(encoding="utf-8"))


# ---------------------------------------------------------------------------
# Reading the unit text
# ---------------------------------------------------------------------------
def parse_unit(unit_text):
    """
    Understand a price basis like "per 100 pcs", "Rs./pc", "per kg", "per box".
    Returns dict(kind, count, word) where kind is 'count' or 'kg', or None if not understood.
    """
    if not unit_text:
        return None
    t = unit_text.lower().strip()
    t = t.replace("/", " per ")
    t = re.sub(CURRENCY_TOKENS, " ", t)
    t = re.sub(r"\s+", " ", t).strip()
    if t in ("each", "ea", "apiece"):
        return {"kind": "count", "count": 1, "word": None}
    m = re.search(r"\bper\s+(\d[\d,]*)?\s*([a-z]+)?\b", t)
    if not m:
        return None
    count = int(m.group(1).replace(",", "")) if m.group(1) else None
    word = m.group(2)
    if word in KG_WORDS:
        return {"kind": "kg", "count": count or 1, "word": word}
    known = PIECE_WORDS | BOX_WORDS | SHEET_WORDS | SET_WORDS
    if word is not None and word not in known:
        return None
    if count is None and word is None:
        return None
    return {"kind": "count", "count": count or 1, "word": word}


def unit_fits_sku(word, sku_uom):
    """Does the vendor's unit word clearly mean one of the buyer's units? None = no noun given."""
    if word is None or word in PIECE_WORDS:
        return True
    if sku_uom == "piece":
        return word in BOX_WORDS          # a carton SKU priced "per box" is per piece
    if sku_uom == "sheet":
        return word in SHEET_WORDS
    if sku_uom == "set":
        return word in SET_WORDS
    return False


def gst_pct_from_text(gst_text):
    """Read the GST rate the vendor wrote, e.g. 'inclusive of GST @ 18%'. None if not exactly one rate."""
    if not gst_text:
        return None
    rates = set(re.findall(r"(\d+(?:\.\d+)?)\s*%", gst_text))
    return float(rates.pop()) if len(rates) == 1 else None


GST_EXTRA_WORDS = re.compile(r"\b(extra|excl\.?|exclusive|excluding|plus|additional)\b|\+\s*gst", re.I)
GST_INCL_WORDS = re.compile(r"\b(incl\.?|inclusive|including|included)\b", re.I)


def gst_wording_supports(tax_included, gst_text):
    """
    Does the vendor's own GST wording back up the extracted tax_included value?
    Used so a vague phrase ("GST as applicable") is always flagged, whatever the model concluded.
    No document-level GST text (e.g. an Excel column labelled ex-GST) is accepted as-is.
    """
    if not gst_text:
        return True
    pattern = GST_INCL_WORDS if tax_included == 1 else GST_EXTRA_WORDS
    return bool(pattern.search(gst_text))


def rsc_weight_kg(sku, glue_flap_mm):
    """Theoretical weight of a regular slotted carton from its spec. Used only for indicative per-kg values."""
    if sku["item_type"] != "RSC" or not sku["height_mm"] or not sku.get("total_gsm"):
        return None
    L, W, H = int(sku["length_mm"]), int(sku["width_mm"]), int(sku["height_mm"])
    area_m2 = (2 * (L + W) + glue_flap_mm) * (W + H) / 1e6
    return area_m2 * int(sku["total_gsm"]) / 1000


# ---------------------------------------------------------------------------
# One cell
# ---------------------------------------------------------------------------
def normalise_cell(line, resp, sku, settings):
    """Returns dict(status, inr_per_unit, indicative_inr, fx_rate, steps, flags)."""
    steps, flags = [], []
    out = dict(status="ok", inr_per_unit=None, indicative_inr=None, fx_rate=None, steps=steps, flags=flags)

    if not line["quoted"] or line["price"] is None:
        out["status"] = "not_quoted"
        steps.append("No price given by vendor.")
        unmatched = json.loads(resp["unmatched_lines"] or "[]")
        if unmatched:
            out["status"] = "needs_review"
            flags.append(f"Vendor had {len(unmatched)} line(s) that matched no SKU; "
                         "this SKU may be one of them.")
        return out

    price = line["price"]
    steps.append(f"Start: {price:g} {line['currency'] or '?'} {line['unit'] or '(no unit)'} "
                 f"[{line['source_note'][:80]}]")

    # 1. Currency ------------------------------------------------------------
    currency = line["currency"] or resp["currency"]
    if not line["currency"] and resp["currency"]:
        steps.append(f"Currency taken from document header: {currency}.")
    rate = settings["fx_to_inr"].get(currency) if currency else None
    if rate is None:
        out["status"] = "needs_review"
        flags.append(f"Currency '{currency}' unknown or not in settings.json fx_to_inr.")
        return out
    out["fx_rate"] = rate

    # 2. Unit ----------------------------------------------------------------
    unit = parse_unit(line["unit"])
    if unit is None:
        out["status"] = "needs_review"
        flags.append(f"Unit '{line['unit']}' is ambiguous or not recognised.")
        return out
    if line["pack_size"] is not None and line["pack_size"] != unit["count"]:
        out["status"] = "needs_review"
        flags.append(f"Unit text says {unit['count']} but extracted pack size is {line['pack_size']}.")
        return out

    if unit["kind"] == "kg":
        out["status"] = "needs_review"
        flags.append("Priced per kg: needs the agreed box weight from the vendor.")
        weight = rsc_weight_kg(sku, settings["rsc_glue_flap_mm"])
        if weight:
            est = price / unit["count"] * weight * rate
            out["indicative_inr"] = round(est, 4)
            steps.append(f"Indicative only: {price:g}/kg x theoretical weight {weight:.3f} kg "
                         f"x {rate:g} = {est:.2f} INR (not used for comparison).")
        return out

    if not unit_fits_sku(unit["word"], sku["uom"]):
        out["status"] = "needs_review"
        flags.append(f"Vendor unit '{line['unit']}' does not clearly mean one {sku['uom']} "
                     f"for this {sku['item_type'].lower()} item.")
        return out

    value = price
    if unit["count"] != 1:
        value = value / unit["count"]
        steps.append(f"Per {unit['count']} -> per 1 {sku['uom']}: {price:g} / {unit['count']} = {value:.4f}")

    if rate != 1:
        before = value
        value = value * rate
        steps.append(f"{currency} -> INR at {rate:g} (settings.json): {before:.4f} x {rate:g} = {value:.4f}")

    # 3. GST -----------------------------------------------------------------
    tax_included = line["tax_included"]
    if tax_included is not None and not gst_wording_supports(tax_included, resp["gst_text"]):
        steps.append(f"Extraction read GST as {'included' if tax_included else 'excluded'}, but the vendor's "
                     f"wording (\"{resp['gst_text']}\") does not say so explicitly.")
        tax_included = None

    if tax_included == 1:
        pct = gst_pct_from_text(resp["gst_text"])
        if pct is None:
            pct = settings["default_gst_pct"]
            flags.append(f"Price includes GST but rate not stated; default {pct:g}% used.")
        before = value
        value = value / (1 + pct / 100)
        steps.append(f"Remove GST {pct:g}% (vendor: \"{resp['gst_text']}\"): "
                     f"{before:.4f} / {1 + pct / 100:g} = {value:.4f}")
    elif tax_included is None:
        flags.append(f"GST treatment not stated (vendor: \"{resp['gst_text'] or 'nothing'}\"); "
                     "price used as written, assumed ex-GST.")
    else:
        steps.append("Already ex-GST.")

    # 4. Discount ------------------------------------------------------------
    if resp["discount_pct"]:
        before = value
        value = value * (1 - resp["discount_pct"] / 100)
        steps.append(f"Apply {resp['discount_pct']:g}% discount (vendor: \"{resp['discount_text']}\"): "
                     f"{before:.4f} x {1 - resp['discount_pct'] / 100:g} = {value:.4f}")

    # 5. Confidence --------------------------------------------------------------
    if line["confidence"] is not None and line["confidence"] < settings["low_confidence_below"]:
        flags.append(f"Low extraction confidence ({line['confidence']:.2f}).")

    out["inr_per_unit"] = round(value, 4)
    steps.append(f"Result: INR {value:.2f} per {sku['uom']}, ex-GST.")
    if flags:
        out["status"] = "flagged"
    return out


# ---------------------------------------------------------------------------
def run():
    """Normalise every extracted cell. Returns {vendor: {status: count}} (empty if nothing to do)."""
    settings = load_settings()
    conn = db.connect()
    skus = {s["sku_id"]: s for s in (dict(r) for r in conn.execute("SELECT * FROM skus"))}
    rows = conn.execute("""
        SELECT l.*, r.file_name, r.vendor_name
        FROM quote_lines l JOIN responses r USING (response_id)
        ORDER BY r.file_name, l.sku_id""").fetchall()
    responses = {r["response_id"]: r for r in conn.execute("SELECT * FROM responses")}
    if not rows:
        conn.execute("DELETE FROM normalised_prices")
        conn.commit()
        conn.close()
        return {}

    now = datetime.now(timezone.utc).isoformat(timespec="seconds")
    conn.execute("DELETE FROM normalised_prices")
    log_rows, counts = [], {}
    for line in rows:
        res = normalise_cell(line, responses[line["response_id"]], skus[line["sku_id"]], settings)
        conn.execute("""INSERT INTO normalised_prices (response_id, sku_id, status, inr_per_unit,
                            indicative_inr, fx_rate, steps, flags, normalised_at)
                        VALUES (?,?,?,?,?,?,?,?,?)""",
                     (line["response_id"], line["sku_id"], res["status"], res["inr_per_unit"],
                      res["indicative_inr"], res["fx_rate"], json.dumps(res["steps"]),
                      json.dumps(res["flags"]), now))
        counts.setdefault(line["vendor_name"], {}).setdefault(res["status"], 0)
        counts[line["vendor_name"]][res["status"]] += 1
        log_rows.append({"file": line["file_name"], "sku_id": line["sku_id"], "raw_price": line["price"],
                         "raw_currency": line["currency"], "raw_unit": line["unit"], "status": res["status"],
                         "inr_per_unit": res["inr_per_unit"], "indicative_inr": res["indicative_inr"],
                         "steps": " | ".join(res["steps"]), "flags": " | ".join(res["flags"])})
    conn.commit()
    conn.close()

    with open(LOG_PATH, "w", newline="", encoding="utf-8") as f:
        w = csv.DictWriter(f, fieldnames=list(log_rows[0].keys()))
        w.writeheader()
        w.writerows(log_rows)

    return counts


def main():
    counts = run()
    if not counts:
        print("Nothing to normalise. Run extract.py first.")
        return
    print(f"Settings: fx_to_inr={load_settings()['fx_to_inr']}")
    for vendor, c in counts.items():
        print(f"  {vendor}: " + ", ".join(f"{k} {v}" for k, v in sorted(c.items())))
    print(f"Wrote {sum(sum(c.values()) for c in counts.values())} cells to quotes.db "
          f"(normalised_prices) and {LOG_PATH.name}")


if __name__ == "__main__":
    main()
