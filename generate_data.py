"""
generate_data.py — fabricate the demo dataset for RFx Analyst.

Creates:
  skus.csv                     30 corrugated SKUs the buyer is sourcing (the RFx line items)
  vendor_replies/              5 vendor replies, each in a different messy format
  ground_truth.csv             the correct INR per-piece price for every vendor x SKU

This data is ONLY used as input for the app and to score extraction accuracy.
The app itself must never read ground_truth.csv to answer questions.

Price definition used in ground_truth.csv (the "comparable unit price"):
  INR per piece (per sheet for pads, per set for partitions),
  EXCLUDING GST, AFTER any discount, AS QUOTED for freight
  (freight terms differ by vendor and are recorded in a separate column).
  USD is converted at the fixed reference rate USD_INR below.

Run:  python generate_data.py
"""

import csv
import random
from pathlib import Path

import numpy as np
from docx import Document
from docx.shared import Pt
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from PIL import Image, ImageDraw, ImageFilter, ImageFont
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import mm
from reportlab.platypus import Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle

ROOT = Path(__file__).parent
OUT_DIR = ROOT / "vendor_replies"
SEED = 42
USD_INR = 88.00  # fixed reference FX rate; the app must use (and show) the same rate
GST = 0.18

BUYER = "Nilgiri Home Appliances Ltd."
BUYER_PLANT = "Plot 14, MIDC Chakan, Pune 410501"
RFQ_NO = "NHA/PKG/RFQ/2026/031"

# ---------------------------------------------------------------------------
# 1. SKUs
# ---------------------------------------------------------------------------
# (sku_id, description, item_type, ply, L, W, H, printing, annual_qty, extra)
# item_type: RSC (regular slotted carton), DIECUT (mailer), PAD (layer pad), PARTITION
# For PARTITION, extra = (rows, cols) of cells; L/W/H are the box inner dimensions.
SKU_ROWS = [
    ("CB-001", "Unit carton - Dry iron",             "RSC", 3, 280, 130, 120, "1 colour", 180000, None),
    ("CB-002", "Unit carton - Steam iron",           "RSC", 3, 300, 150, 140, "1 colour", 150000, None),
    ("CB-003", "Unit carton - Electric kettle 1.5L", "RSC", 3, 230, 200, 250, "2 colour", 220000, None),
    ("CB-004", "Unit carton - Hand blender",         "RSC", 3, 380, 110, 100, "1 colour",  90000, None),
    ("CB-005", "Unit carton - Sandwich maker",       "RSC", 3, 270, 250, 110, "2 colour",  80000, None),
    ("CB-006", "Unit carton - Hair dryer",           "RSC", 3, 250, 200,  90, "2 colour",  60000, None),
    ("CB-007", "Spare parts carton - Small",         "RSC", 3, 200, 150, 100, "Plain",    250000, None),
    ("CB-008", "Spare parts carton - Medium",        "RSC", 3, 300, 200, 150, "Plain",    160000, None),
    ("CB-009", "Accessory carton - Mixer jar set",   "RSC", 3, 320, 240, 200, "1 colour",  70000, None),
    ("CB-010", "E-commerce shipper - Small",         "RSC", 3, 350, 250, 150, "Plain",    300000, None),
    ("CB-011", "E-commerce shipper - Medium",        "RSC", 3, 450, 350, 200, "Plain",    200000, None),
    ("CB-012", "Master carton - Dry iron (6 nos)",   "RSC", 5, 420, 290, 260, "1 colour",  30000, None),
    ("CB-013", "Master carton - Kettle (4 nos)",     "RSC", 5, 470, 410, 260, "1 colour",  55000, None),
    ("CB-014", "Unit carton - Mixer grinder 750W",   "RSC", 5, 420, 260, 300, "2 colour", 140000, None),
    ("CB-015", "Unit carton - Mixer grinder 1000W",  "RSC", 5, 450, 280, 320, "2 colour",  90000, None),
    ("CB-016", "Unit carton - Induction cooktop",    "RSC", 5, 400, 330, 120, "2 colour", 110000, None),
    ("CB-017", "Air cooler spares carton",           "RSC", 5, 520, 380, 300, "1 colour",  25000, None),
    ("CB-018", "Unit carton - Table fan 400mm",      "RSC", 5, 460, 180, 460, "1 colour",  75000, None),
    ("CB-019", "Unit carton - Pedestal fan",         "RSC", 5, 700, 200, 350, "1 colour",  60000, None),
    ("CB-020", "E-commerce shipper - Large",         "RSC", 5, 600, 400, 400, "Plain",     90000, None),
    ("CB-021", "Master carton - Ceiling fan (4 nos)", "RSC", 7, 560, 480, 400, "1 colour",  20000, None),
    ("CB-022", "Unit carton - OTG 28L",              "RSC", 7, 600, 450, 420, "2 colour",  30000, None),
    ("CB-023", "Unit carton - Water heater 15L",     "RSC", 7, 480, 460, 560, "1 colour",  25000, None),
    ("CB-024", "Export master carton - Mixed",       "RSC", 7, 800, 600, 500, "Plain",      8000, None),
    ("CB-025", "Mailer box - Beard trimmer",         "DIECUT", 3, 250, 180, 70, "4 colour", 120000, None),
    ("CB-026", "Mailer box - Accessory kit",         "DIECUT", 3, 300, 220, 80, "4 colour",  80000, None),
    ("CB-027", "Layer pad - Pallet 1000x800",        "PAD", 3, 1000, 800, None, "Plain",   100000, None),
    ("CB-028", "Layer pad - Pallet 1200x1000",       "PAD", 5, 1200, 1000, None, "Plain",   40000, None),
    ("CB-029", "Partition 4x3 cells - Jar set",      "PARTITION", 3, 400, 300, 200, "Plain", 50000, (3, 4)),
    ("CB-030", "Partition 2x2 cells - Kettle master", "PARTITION", 3, 470, 410, 260, "Plain", 55000, (2, 2)),
]

# Paper specification per ply: (liners gsm, flutes gsm, bursting factor, flute profile)
PAPER = {
    3: ([150, 150], [120], 18, "B"),
    5: ([180, 150, 180], [120, 120], 20, "BC"),
    7: ([200, 180, 180, 200], [140, 140, 140], 22, "BCB"),
}
DIECUT_PAPER = ([180, 150], [100], 18, "E")  # white-top E-flute for printed mailers
TAKE_UP = 1.45  # fluting paper uses ~45% more paper than the flat area


def paper_for(sku):
    return DIECUT_PAPER if sku["item_type"] == "DIECUT" else PAPER[sku["ply"]]


def build_skus():
    skus = []
    for sid, desc, typ, ply, L, W, H, printing, qty, extra in SKU_ROWS:
        sku = dict(sku_id=sid, description=desc, item_type=typ, ply=ply,
                   length_mm=L, width_mm=W, height_mm=H, printing=printing,
                   annual_qty=qty, cells=extra)
        liners, flutes, bf, flute = paper_for(sku)
        sku["flute"] = flute
        sku["paper_spec"] = "/".join(str(g) for g in _interleave(liners, flutes))
        sku["bursting_factor"] = bf
        sku["total_gsm"] = round(sum(liners) + sum(flutes) * TAKE_UP)
        sku["uom"] = {"PAD": "sheet", "PARTITION": "set"}.get(typ, "piece")
        skus.append(sku)
    return skus


def _interleave(liners, flutes):
    out = []
    for i, liner in enumerate(liners):
        out.append(liner)
        if i < len(flutes):
            out.append(flutes[i])
    return out


def board_area_m2(sku):
    """Approximate area of corrugated board needed for one piece."""
    L, W, H = sku["length_mm"], sku["width_mm"], sku["height_mm"]
    t = sku["item_type"]
    if t == "RSC":        # blank = (2L + 2W + glue flap) x (W + H)
        area = (2 * (L + W) + 35) * (W + H)
    elif t == "DIECUT":   # FEFCO 0427 style mailer, rough blank size
        area = (L + 4 * H) * (2 * W + 3 * H)
    elif t == "PAD":
        area = L * W
    else:                 # PARTITION: (rows-1) strips of length L + (cols-1) strips of length W
        rows, cols = sku["cells"]
        area = ((rows - 1) * L + (cols - 1) * W) * H
    return area / 1e6


def dims_text(sku, sep=" x ", unit=" mm"):
    L, W, H = sku["length_mm"], sku["width_mm"], sku["height_mm"]
    return f"{L}{sep}{W}{unit}" if H is None else f"{L}{sep}{W}{sep}{H}{unit}"


def write_skus_csv(skus):
    cols = ["sku_id", "description", "item_type", "ply", "flute", "paper_spec",
            "bursting_factor", "total_gsm", "length_mm", "width_mm", "height_mm",
            "printing", "uom", "annual_qty"]
    with open(ROOT / "skus.csv", "w", newline="") as f:
        w = csv.DictWriter(f, fieldnames=cols, extrasaction="ignore")
        w.writeheader()
        for s in skus:
            w.writerow({**s, "height_mm": "" if s["height_mm"] is None else s["height_mm"]})


# ---------------------------------------------------------------------------
# 2. Vendors and their "true" net prices
# ---------------------------------------------------------------------------
# rate_kg: converted-board rate in INR/kg by ply (typical Indian kraft box pricing, 2026)
VENDORS = [
    dict(id="V1", name="Shree Ganesh Corrugators Pvt. Ltd.", city="Bhiwandi, Maharashtra",
         file="01_shree_ganesh_quotation.xlsx", rate_kg={3: 56.0, 5: 53.5, 7: 51.0},
         freight="Extra at actuals"),
    dict(id="V2", name="Deccan Kraft Packaging LLP", city="Bhosari, Pune",
         file="02_deccan_kraft_quotation.pdf", rate_kg={3: 55.0, 5: 51.5, 7: 50.0},
         freight="Included (delivered to Chakan plant)"),
    dict(id="V3", name="Sabarmati Box Industries", city="Vatva GIDC, Ahmedabad",
         file="03_sabarmati_offer_letter.docx", rate_kg={3: 54.0, 5: 54.0, 7: 53.0},
         freight="Extra, charged on actuals"),
    dict(id="V4", name="Om Sai Carton Works", city="Vapi, Gujarat",
         file="04_om_sai_rate_card.jpg", rate_kg={3: 51.5, 5: 52.0, 7: 54.5},
         freight="Ex-works Vapi"),
    dict(id="V5", name="Coromandel Packaging Exports (EOU)", city="Sriperumbudur, Chennai",
         file="05_coromandel_email.txt", rate_kg={3: 57.5, 5: 53.0, 7: 49.0},
         freight="Extra"),
]
PRINT_ADDER = {"Plain": 0.0, "1 colour": 0.80, "2 colour": 1.60, "4 colour": 3.50}
TYPE_ADDER = {"RSC": 0.0, "DIECUT": 1.50, "PAD": 0.0, "PARTITION": 1.00}
V4_NOT_QUOTED = {"CB-025", "CB-026", "CB-024"}  # no die-cut tooling, no 800mm 7-ply


def true_prices(skus, rng):
    """Net INR per piece, ex-GST, for every vendor x SKU (before format quirks)."""
    prices = {}
    for v in VENDORS:
        for s in skus:
            weight_kg = board_area_m2(s) * s["total_gsm"] / 1000
            base = weight_kg * v["rate_kg"][s["ply"]] * rng.uniform(0.95, 1.05)
            price = base + PRINT_ADDER[s["printing"]] + TYPE_ADDER[s["item_type"]] + 1.0
            prices[(v["id"], s["sku_id"])] = round(price, 2)
    return prices


# ---------------------------------------------------------------------------
# 3. Vendor reply generators. Each returns a list of ground-truth rows.
# ---------------------------------------------------------------------------
def gt_row(v, sku_id, price, as_written, conversion):
    return dict(vendor_id=v["id"], vendor_name=v["name"], source_file=v["file"],
                sku_id=sku_id, quoted="Y" if price is not None else "N",
                unit_price_inr="" if price is None else f"{price:.4f}".rstrip("0").rstrip("."),
                as_written_in_source=as_written, conversion_applied=conversion,
                freight_terms=v["freight"])


def inches(mm_val):
    return f"{mm_val / 25.4:.1f}"


# --- V1: Excel that ignores the template ------------------------------------
def make_v1_excel(v, skus, prices):
    """Own layout, own item codes, sizes in INCHES, sorted by ply, shows GST-inclusive column."""
    wb = Workbook()
    ws = wb.active
    ws.title = "Quotation"
    bold = Font(bold=True)
    ws["A1"] = v["name"].upper()
    ws["A1"].font = Font(bold=True, size=14, color="8B0000")
    ws["A2"] = f"Mfrs. of Corrugated Boxes & Sheets | {v['city']} | GSTIN 27AAXCS0000X1Z0"
    ws["A4"] = f"To: {BUYER}, {BUYER_PLANT}"
    ws["A5"] = "Kind Attn: Purchase Dept."
    ws["A6"] = "Our Ref: SGC/Q/26-27/0412    Date: 16-09-2026"
    ws["A7"] = "Sub: Quotation for corrugated boxes as per your enquiry"
    for r in range(1, 8):
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=9)

    headers = ["Sr.", "Item", "Size L x W x H (inch)", "Ply", "Paper (GSM) / BF",
               "Min. Order Qty", "Rate/Box (Rs.)", "GST 18% (Rs.)", "Rate incl. GST (Rs.)"]
    hdr_row = 9
    fill = PatternFill("solid", fgColor="FFE699")
    thin = Side(style="thin", color="999999")
    for c, h in enumerate(headers, start=1):
        cell = ws.cell(row=hdr_row, column=c, value=h)
        cell.font = bold
        cell.fill = fill
        cell.alignment = Alignment(wrap_text=True, horizontal="center", vertical="center")
        cell.border = Border(top=thin, bottom=thin, left=thin, right=thin)

    # Vendor sorts by ply (big boxes first) then by size, and uses its own codes.
    ordered = sorted(skus, key=lambda s: (-s["ply"], -(s["length_mm"] * s["width_mm"])))
    gt = []
    for i, s in enumerate(ordered, start=1):
        price = prices[(v["id"], s["sku_id"])]
        size = " x ".join(inches(x) for x in (s["length_mm"], s["width_mm"], s["height_mm"]) if x is not None)
        kind = {"RSC": "Box", "DIECUT": "Die-cut Mailer", "PAD": "Sheet/Pad", "PARTITION": "Partition"}[s["item_type"]]
        item = f"{s['ply']} Ply {kind}"
        if s["item_type"] == "PARTITION":
            r, c = s["cells"]
            item += f" ({r}x{c} cells)"
        if s["printing"] != "Plain":
            item += f", {s['printing']} print"
        row = [i, item, size, f"{s['ply']} Ply", f"{s['paper_spec']} / {s['bursting_factor']} BF",
               max(1000, s["annual_qty"] // 12 // 1000 * 1000), price,
               round(price * GST, 2), round(price * (1 + GST), 2)]
        for c, val in enumerate(row, start=1):
            cell = ws.cell(row=hdr_row + i, column=c, value=val)
            cell.border = Border(top=thin, bottom=thin, left=thin, right=thin)
            if c >= 7:
                cell.number_format = "#,##0.00"
        gt.append(gt_row(v, s["sku_id"], price, f"Rs. {price:.2f}/box ex-GST (Excel row {i}, size in inch)", "none"))

    r = hdr_row + len(ordered) + 2
    for line in ["Terms & Conditions:",
                 "1. Rates are per box, ex-works Bhiwandi.",
                 "2. GST @ 18% extra as shown.",
                 "3. Freight: extra at actuals.",
                 "4. Payment: 30 days from date of invoice.",
                 "5. Delivery: 7-10 days from receipt of PO & approved artwork.",
                 "6. Quantity tolerance +/- 10%.",
                 "7. Validity: 30 days."]:
        ws.cell(row=r, column=1, value=line).font = bold if line.endswith(":") else Font()
        r += 1
    for col, width in zip("ABCDEFGHI", [5, 34, 22, 7, 22, 14, 14, 14, 16]):
        ws.column_dimensions[col].width = width
    ws.freeze_panes = ws.cell(row=hdr_row + 1, column=1)

    notes = wb.create_sheet("Notes")
    notes["A1"] = "Board specs are our standard. Sizes are internal dimensions converted to inch."
    notes["A2"] = "Printing plates / cylinders charged extra one-time (Rs. 3,500 per colour)."
    wb.save(OUT_DIR / v["file"])
    return gt


# --- V2: PDF with the discount hidden in a footnote -------------------------
def make_v2_pdf(v, skus, prices):
    """Letterhead PDF. Rates shown are LIST rates; a 6% discount appears only in a footnote."""
    discount = 0.06
    path = OUT_DIR / v["file"]

    def letterhead(canvas, doc):
        canvas.saveState()
        canvas.setFillColor(colors.HexColor("#1F4E79"))
        canvas.rect(0, A4[1] - 32 * mm, A4[0], 32 * mm, fill=1, stroke=0)
        canvas.setFillColor(colors.white)
        canvas.setFont("Helvetica-Bold", 18)
        canvas.drawString(18 * mm, A4[1] - 16 * mm, v["name"].upper())
        canvas.setFont("Helvetica", 9)
        canvas.drawString(18 * mm, A4[1] - 23 * mm,
                          f"Gat No. 112, {v['city']} 411026  |  GSTIN 27AAXFD0000X1Z0  |  sales@deccankraft.example")
        canvas.setFillColor(colors.grey)
        canvas.setFont("Helvetica", 7)
        canvas.drawString(18 * mm, 10 * mm, f"Page {doc.page}  |  ISO 9001:2015 certified corrugated packaging plant")
        canvas.restoreState()

    styles = getSampleStyleSheet()
    small = ParagraphStyle("small", parent=styles["Normal"], fontSize=8, leading=10)
    tiny = ParagraphStyle("tiny", parent=styles["Normal"], fontSize=6.5, leading=8, textColor=colors.HexColor("#555555"))
    story = [
        Paragraph(f"<b>QUOTATION</b> &nbsp;&nbsp; No. DKP/QTN/2026/0877 &nbsp;&nbsp; Date: 18 Sep 2026", styles["Normal"]),
        Spacer(1, 3 * mm),
        Paragraph(f"To,<br/>The Purchase Manager<br/>{BUYER}<br/>{BUYER_PLANT}", small),
        Spacer(1, 3 * mm),
        Paragraph(f"Ref: Your RFQ {RFQ_NO}. We thank you for your enquiry and are pleased to quote as under:", small),
        Spacer(1, 4 * mm),
    ]
    data = [["S.No", "Your Code", "Description", "Size (mm)", "Board", "Rate*\n(Rs./pc)"]]
    gt = []
    for i, s in enumerate(skus, start=1):
        true = prices[(v["id"], s["sku_id"])]
        listed = round(true / (1 - discount), 2)
        net = round(listed * (1 - discount), 2)
        data.append([str(i), s["sku_id"], s["description"], dims_text(s, " x ", ""),
                     f"{s['ply']}P {s['flute']} {s['bursting_factor']}BF", f"{listed:.2f}"])
        gt.append(gt_row(v, s["sku_id"], net, f"Rs. {listed:.2f}/pc list", "list x (1 - 6% footnote discount)"))
    table = Table(data, colWidths=[10 * mm, 18 * mm, 62 * mm, 32 * mm, 28 * mm, 20 * mm], repeatRows=1)
    table.setStyle(TableStyle([
        ("FONT", (0, 0), (-1, -1), "Helvetica", 8),
        ("FONT", (0, 0), (-1, 0), "Helvetica-Bold", 8),
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#DDEBF7")),
        ("GRID", (0, 0), (-1, -1), 0.3, colors.grey),
        ("ALIGN", (-1, 1), (-1, -1), "RIGHT"),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
    ]))
    story += [table, Spacer(1, 5 * mm),
              Paragraph("<b>Terms:</b> GST @ 18% extra. Freight included up to your Chakan plant. "
                        "Payment 45 days from invoice. Delivery within 12 days of PO. Validity 30 days.", small),
              Spacer(1, 12 * mm),
              Paragraph("For DECCAN KRAFT PACKAGING LLP<br/><br/>Authorised Signatory", small),
              Spacer(1, 10 * mm),
              Paragraph("* Rates shown are our standard list rates. A special trade discount of 6% on list rates "
                        f"is applicable on all items against this RFQ ({RFQ_NO}) and will be passed on in the invoice. "
                        "Plate/die charges, if any, extra at actuals.", tiny)]
    SimpleDocTemplate(str(path), pagesize=A4, topMargin=38 * mm, bottomMargin=18 * mm,
                      leftMargin=15 * mm, rightMargin=15 * mm).build(story, onFirstPage=letterhead, onLaterPages=letterhead)
    return gt


# --- V3: Word doc with commercials in prose, prices INCLUDING GST ------------
def make_v3_docx(v, skus, prices, rng):
    doc = Document()
    style = doc.styles["Normal"]
    style.font.name = "Calibri"
    style.font.size = Pt(11)
    doc.add_heading(v["name"], level=1)
    doc.add_paragraph(f"Plot 7/B, {v['city']} 382445  |  Ph: +91 79 0000 0000")
    doc.add_paragraph("Date: 19th September 2026")
    doc.add_paragraph(f"To\nThe Purchase Head\n{BUYER}\n{BUYER_PLANT}")
    doc.add_paragraph(f"Subject: Our best offer against your RFQ {RFQ_NO} for corrugated packaging")
    doc.add_paragraph("Dear Sir / Madam,")
    doc.add_paragraph(
        "Greetings from Sabarmati Box Industries. We have been supplying corrugated boxes to leading "
        "appliance brands in Gujarat and Maharashtra for over two decades, and we thank you for the "
        "opportunity. Having studied your specifications carefully, we are happy to submit our offer below.")

    groups = [
        ("For the 3-ply unit and spare-parts cartons, we can offer", lambda s: s["item_type"] == "RSC" and s["ply"] == 3),
        ("Coming to the 5-ply range, our prices are", lambda s: s["item_type"] == "RSC" and s["ply"] == 5),
        ("For the heavy-duty 7-ply cartons, we are pleased to quote", lambda s: s["item_type"] == "RSC" and s["ply"] == 7),
        ("The printed E-flute mailer boxes would come to", lambda s: s["item_type"] == "DIECUT"),
        ("Finally, for the layer pads and partitions,", lambda s: s["item_type"] in ("PAD", "PARTITION")),
    ]
    gt = []
    for opener, pick in groups:
        parts = []
        for s in [s for s in skus if pick(s)]:
            true = prices[(v["id"], s["sku_id"])]
            incl = round(true * (1 + GST), 2)
            net = round(incl / (1 + GST), 2)
            name = s["description"].split(" - ")[-1].lower()
            code = f"{s['sku_id']}, " if rng.random() < 0.5 else ""
            per = {"PAD": "per sheet", "PARTITION": "per set"}.get(s["item_type"], "per box")
            parts.append(f"₹{incl:.2f} {per} for the {name} ({code}{s['ply']}-ply, {dims_text(s, '×')})")
            gt.append(gt_row(v, s["sku_id"], net, f"₹{incl:.2f} incl. GST (prose)", "÷ 1.18 (GST-inclusive)"))
        text = opener + " " + "; ".join(parts[:-1]) + (", and " if len(parts) > 1 else "") + parts[-1] + "."
        doc.add_paragraph(text)

    doc.add_paragraph(
        "All the above boxes will be manufactured from virgin kraft paper of the GSM and BF specified in your "
        "RFQ, and our in-house lab will provide a test certificate (BS, ECT, moisture) with every lot. "
        "Delivery can start within 10 days of receipt of your purchase order and approved artwork.")
    doc.add_paragraph(
        "As regards commercials, please note that all rates mentioned in this letter are inclusive of GST "
        "@ 18%, so no further tax will be charged. Freight will be extra and charged on actuals. Payment "
        "terms requested are 60 days from the date of invoice. This offer is valid for 45 days.")
    doc.add_paragraph("We look forward to a long association.\n\nWarm regards,\n\nNitin Patel\nDirector - Sales")
    doc.save(OUT_DIR / v["file"])
    return gt


# --- V4: phone photo of a printed rate card, rotated, 27 of 30 lines --------
FONT_DIRS = [Path("/System/Library/Fonts/Supplemental"), Path("/System/Library/Fonts"),
             Path("/usr/share/fonts/truetype/dejavu"), Path("C:/Windows/Fonts")]


def load_font(names, size):
    for d in FONT_DIRS:
        for n in names:
            p = d / n
            if p.exists():
                try:
                    return ImageFont.truetype(str(p), size)
                except OSError:
                    pass
    return ImageFont.load_default(size=size)


def perspective_coeffs(dst, src):
    """Coefficients for Image.transform(PERSPECTIVE): maps output points dst -> input points src."""
    m = []
    for (x, y), (u, v) in zip(dst, src):
        m.append([x, y, 1, 0, 0, 0, -u * x, -u * y])
        m.append([0, 0, 0, x, y, 1, -v * x, -v * y])
    return np.linalg.solve(np.array(m, float), np.array(src, float).reshape(8))


def make_v4_photo(v, skus, prices, rng):
    W, H = 1400, 2050
    paper = Image.new("RGB", (W, H), (246, 243, 232))
    d = ImageDraw.Draw(paper)
    f_title = load_font(["Arial Bold.ttf", "DejaVuSans-Bold.ttf", "arialbd.ttf"], 46)
    f_sub = load_font(["Arial.ttf", "DejaVuSans.ttf", "arial.ttf"], 26)
    f_row = load_font(["Courier New.ttf", "DejaVuSansMono.ttf", "cour.ttf"], 27)
    f_hand = load_font(["Bradley Hand Bold.ttf", "Noteworthy.ttc", "DejaVuSans-Oblique.ttf"], 34)

    d.text((W // 2, 70), "OM SAI CARTON WORKS", font=f_title, fill=(20, 20, 20), anchor="mm")
    d.text((W // 2, 120), "Survey No. 88, GIDC Phase II, Vapi (Gujarat)  |  Mob. 98XXX XXXXX",
           font=f_sub, fill=(40, 40, 40), anchor="mm")
    d.text((W // 2, 175), f"RATE CARD  -  {BUYER}  -  Ref {RFQ_NO}", font=f_sub, fill=(20, 20, 20), anchor="mm")
    d.line((60, 210, W - 60, 210), fill=(0, 0, 0), width=3)

    cols = [70, 150, 560, 800, 900, 1130]  # x positions
    y = 235
    for x, h in zip(cols, ["No", "Item", "Size mm", "Ply", "Paper", "Rs/pc"]):
        d.text((x, y), h, font=f_row, fill=(0, 0, 0))
    y += 45
    d.line((60, y - 8, W - 60, y - 8), fill=(0, 0, 0), width=2)

    gt = []
    quoted = [s for s in skus if s["sku_id"] not in V4_NOT_QUOTED]
    corrected_sku = "CB-014"  # one printed rate is struck out and hand-corrected
    for i, s in enumerate(quoted, start=1):
        price = prices[(v["id"], s["sku_id"])]
        short = s["description"].split(" - ")[-1][:22]
        d.text((cols[0], y), f"{i:>2}", font=f_row, fill=(15, 15, 15))
        d.text((cols[1], y), short, font=f_row, fill=(15, 15, 15))
        d.text((cols[2], y), dims_text(s, "x", ""), font=f_row, fill=(15, 15, 15))
        d.text((cols[3], y), f"{s['ply']}", font=f_row, fill=(15, 15, 15))
        d.text((cols[4], y), f"{s['bursting_factor']}BF", font=f_row, fill=(15, 15, 15))
        if s["sku_id"] == corrected_sku:
            old = round(price * 1.08, 2)
            d.text((cols[5], y), f"{old:>7.2f}", font=f_row, fill=(15, 15, 15))
            d.line((cols[5] - 5, y + 16, cols[5] + 130, y + 12), fill=(20, 40, 160), width=4)
            ink = Image.new("RGBA", (200, 60), (0, 0, 0, 0))
            ImageDraw.Draw(ink).text((5, 5), f"{price:.2f}", font=f_hand, fill=(20, 40, 160, 255))
            ink = ink.rotate(8, expand=True, resample=Image.BICUBIC)
            paper.paste(ink, (cols[5] + 120, y - 28), ink)
            written = f"Rs. {old:.2f} printed, struck out, hand-corrected to {price:.2f}"
        else:
            d.text((cols[5], y), f"{price:>7.2f}", font=f_row, fill=(15, 15, 15))
            written = f"Rs. {price:.2f}/pc (photo row {i})"
        gt.append(gt_row(v, s["sku_id"], price, written, "none"))
        y += 56
        d.line((60, y - 10, W - 60, y - 10), fill=(185, 185, 180), width=1)

    y += 20
    for line in ["* Rates per piece, ex-works Vapi. GST 18% extra. Freight by party.",
                 "* Payment: 15 days.  Validity: till 31-Oct-2026.",
                 "* Die-cut mailers & 800mm export carton: not in our range."]:
        d.text((70, y), line, font=f_sub, fill=(30, 30, 30))
        y += 38
    for s in skus:
        if s["sku_id"] in V4_NOT_QUOTED:
            gt.append(gt_row(v, s["sku_id"], None, "not on rate card", "n/a - not quoted"))

    # --- turn the clean page into a phone photo ---
    bg_w, bg_h = W + 360, H + 360
    photo = Image.new("RGB", (bg_w, bg_h), (92, 70, 52))  # wooden desk
    desk = np.array(photo, dtype=np.float32)
    desk += np.random.default_rng(SEED).normal(0, 9, desk.shape)  # wood grain-ish texture
    photo = Image.fromarray(desk.clip(0, 255).astype(np.uint8))
    # the page is skewed, as if the phone was held at an angle
    # (right edge raised a little so that, with the rotation below, the net tilt is ~2 degrees)
    dst = [(210, 200), (W + 150, 160), (W + 215, H + 170), (150, H + 150)]
    src = [(0, 0), (W, 0), (W, H), (0, H)]
    warped = paper.convert("RGBA").transform((bg_w, bg_h), Image.PERSPECTIVE,
                                             perspective_coeffs(dst, src), Image.BICUBIC)
    photo.paste(warped, (0, 0), warped)
    photo = photo.rotate(-2.0, resample=Image.BICUBIC, fillcolor=(92, 70, 52))

    arr = np.array(photo, dtype=np.float32)
    yy, xx = np.mgrid[0:bg_h, 0:bg_w]
    light = 1.05 - 0.30 * (xx / bg_w) * 0.6 - 0.25 * (yy / bg_h)  # uneven light, darker bottom-right
    arr *= light[..., None]
    shadow = (np.abs(xx - bg_w * 0.55) < 25) * 0.08  # faint fold line down the page
    arr *= (1 - shadow)[..., None]
    arr += np.random.default_rng(SEED + 1).normal(0, 7, arr.shape)  # sensor noise
    photo = Image.fromarray(arr.clip(0, 255).astype(np.uint8)).filter(ImageFilter.GaussianBlur(1.1))
    photo = photo.resize((int(bg_w * 0.8), int(bg_h * 0.8)), Image.LANCZOS)
    photo.save(OUT_DIR / v["file"], "JPEG", quality=68)
    return gt


# --- V5: plain-text email, USD per 100 pieces --------------------------------
def make_v5_email(v, skus, prices, rng):
    lines = [
        f"From: Karthik Raman <karthik.r@coromandelpack.example>",
        f"To: purchase@nilgiri-appliances.example",
        f"Date: Mon, 21 Sep 2026 18:42:10 +0530",
        f"Subject: RE: {RFQ_NO} - Corrugated boxes - our rates",
        "",
        "Dear Sir,",
        "",
        "Thanks for the RFQ. As we are an EOU our price list is maintained in USD, pls find our best",
        "rates below. All prices are USD per 100 pcs (pads per 100 sheets, partitions per 100 sets).",
        "",
    ]
    gt = []
    styles = [
        lambda s, p: f"{s['sku_id']}  {dims_text(s, 'x', '')} {s['ply']}ply  - USD {p:.2f} / 100 pcs",
        lambda s, p: f"{s['sku_id']} ({s['description'].split(' - ')[-1].lower()}) ..... ${p:.2f} per 100",
        lambda s, p: f"{s['sku_id']}: {p:.2f} USD/100nos",
    ]
    for s in skus:
        true = prices[(v["id"], s["sku_id"])]
        usd100 = round(true * 100 / USD_INR, 2)
        inr = round(usd100 * USD_INR / 100, 4)
        lines.append(rng.choice(styles)(s, usd100))
        gt.append(gt_row(v, s["sku_id"], inr, f"USD {usd100:.2f} per 100 pcs",
                         f"USD/100 x {USD_INR:.2f} INR/USD / 100"))
    lines += [
        "",
        "Freight extra, ex-works Sriperumbudur. GST as applicable. Payment 30 days.",
        "Board: virgin kraft as per your spec, 7-ply is our strength - we run a 2.2m BHS line.",
        "Rates valid 30 days. Let me know if you need samples.",
        "",
        "Rgds,",
        "Karthik Raman",
        "Sr. Manager - Sales, Coromandel Packaging Exports (EOU)",
        "+91 44 0000 0000",
        "",
        "Sent from my iPhone",
    ]
    (OUT_DIR / v["file"]).write_text("\n".join(lines) + "\n", encoding="utf-8")
    return gt


# ---------------------------------------------------------------------------
def main():
    rng = random.Random(SEED)
    OUT_DIR.mkdir(exist_ok=True)
    skus = build_skus()
    write_skus_csv(skus)
    prices = true_prices(skus, rng)

    v1, v2, v3, v4, v5 = VENDORS
    ground_truth = []
    ground_truth += make_v1_excel(v1, skus, prices)
    ground_truth += make_v2_pdf(v2, skus, prices)
    ground_truth += make_v3_docx(v3, skus, prices, rng)
    ground_truth += make_v4_photo(v4, skus, prices, rng)
    ground_truth += make_v5_email(v5, skus, prices, rng)

    order = {s["sku_id"]: i for i, s in enumerate(skus)}
    ground_truth.sort(key=lambda r: (r["vendor_id"], order[r["sku_id"]]))
    with open(ROOT / "ground_truth.csv", "w", newline="", encoding="utf-8") as f:
        w = csv.DictWriter(f, fieldnames=list(ground_truth[0].keys()))
        w.writeheader()
        w.writerows(ground_truth)

    print(f"Wrote skus.csv ({len(skus)} SKUs)")
    for v in VENDORS:
        print(f"Wrote vendor_replies/{v['file']}")
    print(f"Wrote ground_truth.csv ({len(ground_truth)} rows, USD_INR = {USD_INR})")


if __name__ == "__main__":
    main()
