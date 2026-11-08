#!/usr/bin/env python3
"""
Weekly invoice intake: folder in -> spreadsheet out (one row per file).

Columns: file, vendor, invoice_number, due_date, amount, flag, reason

Steps
  1. READ     PDFs (text layer), scans/photos (OCR), emails (pulls the attachment)
  2. EXTRACT  vendor / invoice # / dates / total / line items.
              AI extractor (Claude) when ANTHROPIC_API_KEY is set,
              otherwise a local regex parser so it still runs offline.
  3. FLAG     plain, visible Python rules (not the AI) decide the Y/N, so every
              flag can be explained to a reviewer.
  4. WRITE    invoice_review.xlsx (Invoices, Review Queue, Summary, Rules tabs),
              a CSV, and monday_digest.txt (the email a reviewer would get).

Usage
  python3 invoice_pipeline.py /path/to/invoice/folder [--out outdir] [--as-of 2026-11-08]
            [--threshold 5000] [--history history.json]
  (--out may contain {date}, e.g. "Weekly Review/{date}")
Needs: pip install pdfplumber pypdfium2 pillow openpyxl, plus Tesseract (the OCR program).
(poppler is used instead of pdfplumber/pypdfium2 if it happens to be installed.)
Optional: pip install anthropic   and   export ANTHROPIC_API_KEY=...
"""
import argparse, base64, csv, email, hashlib, json, os, re, shutil, subprocess, sys, tempfile
from datetime import date, datetime, timedelta
from email import policy

DEFAULT_RULES = {
    "threshold":         {"enabled": True,  "amount": 5000},
    "duplicate":         {"enabled": True},
    "line_item_math":    {"enabled": True,  "tolerance": 0.01},
    "foreign_currency":  {"enabled": True,  "eur_to_usd": 1.10},
    "unreadable_field":  {"enabled": True},
    "portfolio_company": {"enabled": True,  "names": ["Grasp", "PlanIt", "WorthPoint"]},
    "overdue":           {"enabled": False, "days": 30},
    "new_vendor":        {"enabled": False, "known_vendors": []},
    "custom": [],        # your own "if the invoice mentions X, flag it" rules
}

# Plain-English names, used by the Rules tab and the settings window.
RULE_INFO = {
    "threshold":         ("Over the dollar limit", "Larger payments need sign-off."),
    "duplicate":         ("Duplicate invoice", "Stops paying the same bill twice. Also checks invoices from past weeks."),
    "line_item_math":    ("Totals do not add up", "Catches vendor math errors and overbilling."),
    "foreign_currency":  ("Foreign currency", "Exchange rate, fees and VAT need a human decision."),
    "unreadable_field":  ("Unreadable field", "Vendor, invoice number, due date or amount could not be read. Never guess on a payment."),
    "portfolio_company": ("Names a portfolio company", "Invoice mentions one of these companies but is billed to the holding company, so the cost may belong elsewhere."),
    "overdue":           ("Past due", "Invoice is this many days or more past its due date on the run date."),
    "new_vendor":        ("Vendor not on the approved list", "Guards against fake-vendor fraud. Does nothing until you add vendors to the list."),
}


def load_rules(path=None):
    """Built-in defaults, overridden by whatever is in rules.json (missing settings fall back to defaults)."""
    import copy
    rules = copy.deepcopy(DEFAULT_RULES)
    if path and os.path.exists(path):
        with open(path, encoding="utf-8") as f:
            user = json.load(f)
        for k, v in user.items():
            if k == "custom":
                rules["custom"] = v if isinstance(v, list) else []
            elif k in rules and isinstance(v, dict):
                rules[k].update(v)
    return rules
MONEY = re.compile(r"[$€]\s?[\d,]+\.\d{2}")
MODEL = os.environ.get("INVOICE_MODEL", "claude-haiku-5-5")   # small + cheap is enough here


# ---------------------------------------------------------------- 1. READ
def find_tool(name):
    """Find a command-line tool, including the usual Windows install folders."""
    if os.environ.get("INVOICE_PYTHON_ONLY") and name != "tesseract":
        return None                                  # test switch: pretend poppler is not installed
    p = shutil.which(name)
    if p:
        return p
    if os.name == "nt" and name == "tesseract":
        for base in (os.environ.get("ProgramFiles", r"C:\Program Files"),
                     os.environ.get("ProgramFiles(x86)", r"C:\Program Files (x86)"),
                     os.path.join(os.environ.get("LOCALAPPDATA", ""), "Programs")):
            c = os.path.join(base, "Tesseract-OCR", "tesseract.exe")
            if os.path.exists(c):
                return c
    return None


def run(cmd):
    return subprocess.run(cmd, capture_output=True, text=True).stdout


def pdf_text(path):
    """Text layer of a PDF. Uses poppler if installed, otherwise pure-Python pdfplumber."""
    exe = find_tool("pdftotext")
    if exe:
        return run([exe, "-layout", path, "-"])
    try:
        import pdfplumber
    except ImportError:
        sys.exit("Cannot read PDFs. Run: pip install pdfplumber pypdfium2 pillow openpyxl")
    with pdfplumber.open(path) as pdf:
        return "\n".join((pg.extract_text(layout=True) or "") for pg in pdf.pages)


def pdf_first_page_image(path):
    """Render page 1 of an image-only PDF to a PNG so it can be OCR'd."""
    td = tempfile.mkdtemp()
    exe = find_tool("pdftoppm")
    if exe:
        run([exe, "-r", "200", "-png", path, os.path.join(td, "p")])
        return os.path.join(td, sorted(os.listdir(td))[0])
    try:
        import pypdfium2 as pdfium
    except ImportError:
        sys.exit("Cannot render scanned PDFs. Run: pip install pdfplumber pypdfium2 pillow openpyxl")
    out = os.path.join(td, "p-1.png")
    pdfium.PdfDocument(path)[0].render(scale=200 / 72).to_pil().save(out)
    return out


def ocr_variants(path):
    """OCR an image a few ways; return list of texts (upscaled psm 6 usually best)."""
    from PIL import Image
    outs = []
    with tempfile.TemporaryDirectory() as td:
        im = Image.open(path).convert("L")
        big = os.path.join(td, "big.png")
        im.resize((im.width * 2, im.height * 2)).save(big)
        for src, psm in ((big, "6"), (big, "4"), (path, "6")):
            exe = find_tool("tesseract")
            if not exe:
                sys.exit("Tesseract (the OCR program) is not installed. Windows: winget install UB-Mannheim.TesseractOCR "
                         "then reopen the terminal. See README.")
            outs.append(run([exe, src, "-", "--psm", psm]))
    return outs


def read_file(path):
    """Return (text_variants, image_path_or_None, note)."""
    ext = os.path.splitext(path)[1].lower()
    if ext == ".pdf":
        txt = pdf_text(path)
        if len(txt.strip()) > 40:
            return [txt], None, ""
        img = pdf_first_page_image(path)
        return ocr_variants(img), img, "scanned PDF, read by OCR"
    if ext in (".jpg", ".jpeg", ".png"):
        return ocr_variants(path), path, "image, read by OCR"
    if ext == ".eml":
        msg = email.message_from_binary_file(open(path, "rb"), policy=policy.default)
        for part in msg.walk():
            fn = part.get_filename()
            if fn:
                td = tempfile.mkdtemp()
                p = os.path.join(td, fn)
                open(p, "wb").write(part.get_payload(decode=True))
                texts, img, note = read_file(p)
                return texts, img, f"email attachment {fn}" + (f" ({note})" if note else "")
        return [msg.get_body(("plain",)).get_content()], None, "email body only, no attachment"
    return [""], None, "unsupported file type"


# ---------------------------------------------------------------- 2. EXTRACT
def parse_date(s):
    s = s.strip().replace(",", "")
    for fmt in ("%Y-%m-%d", "%m/%d/%Y", "%d %B %Y", "%B %d %Y", "%b %d %Y", "%d %b %Y"):
        try:
            return datetime.strptime(s, fmt).date()
        except ValueError:
            pass
    return None


DATE_PAT = (r"(\d{4}-\d{2}-\d{2}|\d{1,2}/\d{1,2}/\d{4}|\d{1,2} [A-Za-z]+ \d{4}|[A-Za-z]{3,9} \d{1,2},? \d{4})")


def local_extract(texts, ocr=False):
    """Regex fallback. Tries each OCR variant, keeps the one that yields the most fields."""
    best, best_score = {}, -1
    for t in texts:
        r = _extract_one(t, ocr)
        score = sum(1 for k in ("vendor", "invoice_number", "invoice_date", "total") if r.get(k))
        score += 1 if r.get("invoice_date") else 0
        if score > best_score:
            best, best_score = r, score
    return best


def _extract_one(text, ocr=False):
    lines = [l for l in text.splitlines() if l.strip()]
    v = lines[0].strip() if lines else None
    if v:
        v = " ".join(v.split())                                 # tidy repeated spaces
        v = re.sub(r"^invoice\s+", "", v, flags=re.I)             # "Invoice  Acme Corp" when both share a row
        v = re.sub(r"^\W*\w{1,2}\s+(?=[A-Z])", "", v)          # OCR junk like "s Lumen ..."
        v = v.title() if v.isupper() else v                     # IRONWOOD SECURITY -> Ironwood Security
    r = {"vendor": v}
    # invoice number
    m = (re.search(r"Invoice\s*(?:Number|No\.?|#)\s*:?\s*([A-Z0-9][A-Z0-9-]*\d)", text, re.I)
         or re.search(r"^\s*INVOICE\s+([A-Z]{2,}[-A-Z0-9]*\d)\b", text, re.M)
         or re.search(r"^\s*([A-Z]{1,6}-[A-Z0-9-]*\d)\s+[·•\-]\s", text, re.M))
    inv = m.group(1).strip() if m else None
    if inv:
        inv = re.sub(r"(?<=\d)O(?=\d)", "0", inv)        # OCR reads 1041 as 1O41
    r["invoice_number"] = inv
    # invoice date
    m = (re.search(r"(?:Invoice Date|ISSUED|DATE|Date)\s*:?\s*" + DATE_PAT, text)
         or re.search(r"[·•\-]\s+" + DATE_PAT, text))
    r["invoice_date"] = parse_date(m.group(1)) if m else None
    # explicit due date
    m = re.search(r"(?:Payment Due|DUE|Due)\s*:?\s*" + DATE_PAT, text)
    r["due_date"] = parse_date(m.group(1)) if m else None
    # terms
    m = re.search(r"Net\s*(\d+)", text, re.I)
    r["terms_days"] = int(m.group(1)) if m else None
    r["due_on_receipt"] = bool(re.search(r"due upon receipt", text, re.I))
    # currency / total
    # A lone euro sign is only trusted in a real text layer. OCR often misreads "$" as the euro sign,
    # so in a scan we also need a word like EUR, IBAN, VAT or GmbH before calling it a euro invoice.
    euro_words = re.search(r"\bEUR\b|amounts in EUR|\bIBAN\b|\bVAT\b|\bGmbH\b|\bEuros?\b", text, re.I)
    r["currency"] = "EUR" if euro_words or ("€" in text and not ocr) else "USD"
    r["euro_misread"] = bool(ocr and "€" in text and not euro_words)
    monies = MONEY.findall(text)
    if monies:
        r["total"] = float(re.sub(r"[^\d.]", "", monies[-1]))
    # line items: lines ending in an amount, excluding the total line
    items = []
    for l in lines:
        if re.search(r"TOTAL DUE|AMOUNT DUE", l, re.I):
            continue
        m = re.search(r"(\d[\d,]*\.\d{2})\s*$", l)
        if m and not re.search(r"Suite|Blvd|Dr,|St,|Ave", l):
            items.append(float(m.group(1).replace(",", "")))
    r["line_items"] = items
    return r


AI_PROMPT = """Extract this invoice. Reply with JSON only:
{"vendor": str, "invoice_number": str, "invoice_date": "YYYY-MM-DD", "due_date": "YYYY-MM-DD or null",
 "terms_days": int or null, "due_on_receipt": bool, "currency": "USD|EUR|...",
 "total": number, "line_items": [number, ...]}
Use the invoice's own printed total, not a recomputed one. Dates with slashes are US MM/DD/YYYY.
If a field is unreadable use null. Do not guess."""


def ai_extract(texts, img):
    import anthropic
    client = anthropic.Anthropic()
    content = []
    if img:
        data = base64.b64encode(open(img, "rb").read()).decode()
        mt = "image/jpeg" if img.lower().endswith((".jpg", ".jpeg")) else "image/png"
        content.append({"type": "image", "source": {"type": "base64", "media_type": mt, "data": data}})
    else:
        content.append({"type": "text", "text": texts[0]})
    content.append({"type": "text", "text": AI_PROMPT})
    resp = client.messages.create(model=MODEL, max_tokens=800,
                                  messages=[{"role": "user", "content": content}])
    raw = resp.content[0].text
    d = json.loads(raw[raw.index("{"): raw.rindex("}") + 1])
    for k in ("invoice_date", "due_date"):
        d[k] = parse_date(d[k]) if d.get(k) else None
    return d


def extract(texts, img):
    rec, how = None, "local"
    if os.environ.get("ANTHROPIC_API_KEY"):
        try:
            rec, how = ai_extract(texts, img), "ai"
        except Exception as e:                       # never let one bad file stop the run
            print(f"  AI extract failed ({e}); using local parser", file=sys.stderr)
    if rec is None:
        rec = local_extract(texts, ocr=img is not None)
    blob = " ".join(texts).lower()
    rec["text"] = blob                                # custom and portfolio rules search this
    return rec, how


# ---------------------------------------------------------------- 3. FLAG
def norm_vendor(v):
    v = re.sub(r"[^a-z0-9 ]", "", (v or "").lower())
    v = re.sub(r"\b(llc|llp|inc|gmbh|co|corp|ltd)\b", "", v)
    return v.split()[0] if v.split() else ""


def norm_inv(n):
    return re.sub(r"[^A-Z0-9]", "", (n or "").upper()).replace("O", "0")


def finish(rec):
    """Work out due_date from invoice date + terms when it isn't printed."""
    d = rec.get("due_date")
    if not d and rec.get("invoice_date"):
        if rec.get("terms_days") is not None:
            d = rec["invoice_date"] + timedelta(days=rec["terms_days"])
        elif rec.get("due_on_receipt"):
            d = rec["invoice_date"]
    rec["due_final"] = d
    return rec


def apply_flags(rows, as_of, prior=(), rules=None):
    """rows: list of dicts with file + extracted fields. Adds flag + reason.
    prior: records from earlier runs, so a resend next week is still caught.
    rules: settings (see DEFAULT_RULES). Each rule can be switched off."""
    rules = rules or DEFAULT_RULES
    on = lambda name: rules.get(name, {}).get("enabled", False)
    limit = float(rules["threshold"].get("amount", 5000))
    rate = float(rules["foreign_currency"].get("eur_to_usd", 1.10))
    seen_num, seen_amt = {}, {}
    if on("duplicate"):
        for p_ in prior:
            pv = norm_vendor(p_.get("vendor"))
            seen_num.setdefault((pv, norm_inv(p_.get("invoice_number"))), p_)
            if p_.get("total") and p_.get("invoice_date"):
                seen_amt.setdefault((pv, p_["total"], p_["invoice_date"]), p_)
    known = {norm_vendor(v) for v in rules["new_vendor"].get("known_vendors", []) if v.strip()}

    for r in rows:
        reasons = []
        r["custom_hits"], r["custom_actions"] = [], []
        amt = r.get("total")
        usd = amt * rate if (amt and r.get("currency") == "EUR") else amt
        foreign = r.get("currency") not in (None, "USD")

        # missing / unreadable fields
        if on("unreadable_field"):
            missing = [lab for k, lab in (("vendor", "vendor"), ("invoice_number", "invoice number"),
                                          ("due_final", "due date"), ("total", "amount")) if not r.get(k)]
            if missing:
                reasons.append("could not read " + ", ".join(missing) + " - check the original")

        # over the dollar limit
        if on("threshold") and usd and usd > limit:
            if not foreign:
                reasons.append(f"over ${limit:,.0f} approval threshold")
            elif not on("foreign_currency"):
                reasons.append(f"over ${limit:,.0f} approval threshold (about ${usd:,.0f} USD)")

        # foreign currency
        if on("foreign_currency") and foreign:
            note = f"billed in {r['currency']}, not USD - confirm the exchange rate and payment method"
            if on("threshold") and usd and usd > limit:
                note += f" (about ${usd:,.0f} USD, likely over ${limit:,.0f})"
            reasons.append(note)

        # line items do not add up to the printed total
        items = r.get("line_items") or []
        if on("line_item_math") and items and amt is not None and \
                abs(sum(items) - amt) > float(rules["line_item_math"].get("tolerance", 0.01)):
            reasons.append(f"line items add to ${sum(items):,.2f} but invoice total says ${amt:,.2f} "
                           f"(total is ${abs(amt - sum(items)):,.2f} {'higher' if amt > sum(items) else 'lower'} than the lines)")

        # intercompany: names an operating company but is billed to the holding company
        if on("portfolio_company"):
            hits = [n for n in rules["portfolio_company"].get("names", [])
                    if n.strip() and n.lower() in (r.get("text") or "")]
            r["portfolio"] = hits
            if hits:
                reasons.append(f"mentions {', '.join(hits)} but is billed to Waverock Holdings - confirm which entity should be charged")

        # past due
        if on("overdue") and r.get("due_final"):
            late = (as_of - r["due_final"]).days
            if late >= int(rules["overdue"].get("days", 30)):
                reasons.append(f"{late} days past due (due {r['due_final'].isoformat()}) - late fees may apply")

        # vendor not on the approved list
        if on("new_vendor") and known and r.get("vendor") and norm_vendor(r["vendor"]) not in known:
            reasons.append("vendor is not on the approved vendor list - verify before paying")

        # your own rules
        for c in rules.get("custom", []):
            terms = [t.strip().lower() for t in c.get("contains", []) if t.strip()]
            if c.get("enabled", True) and terms and any(t in (r.get("text") or "") for t in terms):
                reasons.append(c.get("reason") or c.get("name") or "custom rule matched")
                r["custom_hits"].append(c.get("name", ""))
                r["custom_actions"].append((int(c.get("priority", 2)), c.get("next_step") or "Review the invoice."))

        # duplicates: same vendor + invoice number, or same vendor + amount + invoice date
        if on("duplicate"):
            v = norm_vendor(r.get("vendor"))
            k1 = (v, norm_inv(r.get("invoice_number")))
            k2 = (v, amt, r.get("invoice_date"))
            dup_of = None
            if r.get("invoice_number") and k1 in seen_num:
                dup_of = seen_num[k1]
            elif amt and r.get("invoice_date") and k2 in seen_amt:
                dup_of = seen_amt[k2]
            if dup_of:
                earlier = " (processed in an earlier run)" if dup_of.get("prior") else ""
                reasons.append(f"duplicate of {dup_of['file']}{earlier} (same vendor and {'invoice number' if k1 in seen_num else 'amount/date'}) - pay only once")
                dup_of.setdefault("dup_children", []).append(r["file"])
            else:
                seen_num[k1] = r
                if amt and r.get("invoice_date"):
                    seen_amt[k2] = r
        r["reasons"] = reasons

    # the first copy of a duplicate pair also needs a look
    for r in rows:
        if r.get("dup_children"):
            r["reasons"].append(f"also received as {', '.join(r['dup_children'])} - pay only once")
            due = r.get("due_final")
            if due and due < as_of:
                r["reasons"].append(f"due {due.isoformat()}, now {(as_of - due).days} days overdue - confirm it was not already paid")
        r["flag"] = "Y" if r["reasons"] else "N"
    return rows


# ---------------------------------------------------------------- 4. WRITE
# reason keyword -> (priority, suggested next step). Priority 1 = money may go out wrongly.
ACTIONS = [
    ("duplicate of",       1, "Hold this copy. Pay the original once."),
    ("also received",      1, "Confirm in QuickBooks/Bill.com it was not already paid, then pay once."),
    ("line items add",     1, "Ask the vendor for a corrected invoice before paying."),
    ("could not read",     1, "Open the original file and fill in the missing field by hand."),
    ("approval threshold", 2, "Needs approver sign-off (COO / fractional CFO)."),
    ("billed in",          2, "Confirm exchange rate and how to pay (wire/IBAN). Check VAT treatment."),
    ("confirm which entity", 2, "Confirm which company should be charged; record the intercompany allocation."),
    ("days past due",      2, "Pay or contact the vendor now to avoid late fees."),
    ("approved vendor list", 1, "Verify the vendor (call a known number, check W-9 and bank details) before paying."),
]


def triage(r):
    """(priority, next steps) for a flagged row."""
    text = "; ".join(r["reasons"])
    hits = [(pr, act) for key, pr, act in ACTIONS if key in text] + list(r.get("custom_actions", []))
    if not hits:
        return 2, "Review the invoice."
    return min(h[0] for h in hits), " ".join(dict.fromkeys(h[1] for h in hits))


def row_summary(rows):
    usd = [r for r in rows if r.get("total") is not None and r.get("currency", "USD") == "USD"]
    dup_copies = [r for r in rows if any(x.startswith("duplicate of") for x in r["reasons"])]
    over = 0.0
    for r in rows:
        items = r.get("line_items") or []
        if items and r.get("total") is not None and r["total"] > sum(items) + 0.01:
            over += r["total"] - sum(items)
    return {
        "n": len(rows), "y": sum(r["flag"] == "Y" for r in rows),
        "usd_total": sum(r["total"] for r in usd),
        "usd_flagged": sum(r["total"] for r in usd if r["flag"] == "Y"),
        "usd_clear": sum(r["total"] for r in usd if r["flag"] == "N"),
        "dup_exposure": sum(r["total"] for r in dup_copies if r.get("total")),
        "dup_n": len(dup_copies), "overbilled": over,
    }


def write_outputs(rows, outdir, as_of, rules=None):
    rules = rules or DEFAULT_RULES
    os.makedirs(outdir, exist_ok=True)
    cols = ["file", "vendor", "invoice_number", "due_date", "amount", "flag", "reason"]
    out = []
    for r in rows:
        out.append({
            "file": r["file"], "vendor": r.get("vendor") or "",
            "invoice_number": r.get("invoice_number") or "",
            "due_date": r["due_final"].isoformat() if r.get("due_final") else "",
            "amount": r.get("total") if r.get("total") is not None else "",
            "flag": r["flag"], "reason": "; ".join(r["reasons"]),
        })
    with open(os.path.join(outdir, "invoice_review.csv"), "w", newline="") as f:
        w = csv.DictWriter(f, fieldnames=cols); w.writeheader(); w.writerows(out)

    from openpyxl import Workbook
    from openpyxl.chart import BarChart, Reference
    from openpyxl.chart.label import DataLabelList
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
    from openpyxl.worksheet.datavalidation import DataValidation

    INK, SLATE = "1F2937", "334155"
    RED_BG, RED_TX, GRN_BG, GRN_TX = "FDE8E8", "9B1C1C", "E6F4EA", "166534"
    thin = Side(style="thin", color="D1D5DB")
    head_font, head_fill = Font(bold=True, color="FFFFFF"), PatternFill("solid", fgColor=SLATE)

    def style_header(ws_, row=1):
        for c in ws_[row]:
            c.font, c.fill = head_font, head_fill
            c.alignment = Alignment(vertical="center", wrap_text=True)

    wb = Workbook()

    # --- Invoices: exactly the seven requested columns
    ws = wb.active; ws.title = "Invoices"
    ws.append(cols); style_header(ws)
    for o, r in zip(out, rows):
        ws.append([o[c] for c in cols])
        row = ws[ws.max_row]
        sym = "[$\u20ac-x-euro2] #,##0.00" if r.get("currency") == "EUR" else "$#,##0.00"
        row[4].number_format = sym
        row[6].alignment = Alignment(wrap_text=True, vertical="top")
        for c in row:
            c.border = Border(bottom=thin)
        if o["flag"] == "Y":
            for c in row:
                c.fill = PatternFill("solid", fgColor=RED_BG)
            row[5].font = Font(bold=True, color=RED_TX)
        else:
            row[5].font = Font(bold=True, color=GRN_TX)
            row[5].fill = PatternFill("solid", fgColor=GRN_BG)
        row[5].alignment = Alignment(horizontal="center", vertical="top")
    for col, w in zip("ABCDEFG", (30, 32, 20, 12, 14, 7, 95)):
        ws.column_dimensions[col].width = w
    ws.freeze_panes = "A2"; ws.auto_filter.ref = ws.dimensions

    # --- Review Queue: only the Y rows, worst first, with a next step and a place to record the decision
    rq = wb.create_sheet("Review Queue")
    rq.append(["priority", "file", "vendor", "invoice_number", "amount", "due_date", "why flagged",
               "suggested next step", "reviewer decision", "notes"])
    style_header(rq)
    flagged = sorted((r for r in rows if r["flag"] == "Y"), key=lambda r: (triage(r)[0], -(r.get("total") or 0)))
    for r in flagged:
        pr, act = triage(r)
        rq.append(["1 - High" if pr == 1 else "2 - Medium", r["file"], r.get("vendor") or "",
                   r.get("invoice_number") or "", r.get("total"),
                   r["due_final"].isoformat() if r.get("due_final") else "",
                   "; ".join(r["reasons"]), act, "", ""])
        row = rq[rq.max_row]
        row[4].number_format = "[$\u20ac-x-euro2] #,##0.00" if r.get("currency") == "EUR" else "$#,##0.00"
        for c in row:
            c.alignment = Alignment(wrap_text=True, vertical="top"); c.border = Border(bottom=thin)
        row[0].font = Font(bold=True, color=RED_TX if pr == 1 else "92400E")
        row[8].fill = PatternFill("solid", fgColor="FEF9C3")
    dv = DataValidation(type="list", formula1='"Approve,Hold,Ask vendor,Reject"', allow_blank=True)
    rq.add_data_validation(dv); dv.add(f"I2:I{max(rq.max_row, 2)}")
    for col, w in zip("ABCDEFGHIJ", (11, 28, 28, 18, 13, 12, 60, 50, 16, 30)):
        rq.column_dimensions[col].width = w
    rq.freeze_panes = "A2"

    # --- Summary
    sm = wb.create_sheet("Summary")
    t = row_summary(rows)
    sm["A1"] = "Weekly invoice review"; sm["A1"].font = Font(size=16, bold=True, color=INK)
    sm["A2"] = f"Run date {as_of.isoformat()}  |  {t['n']} files"; sm["A2"].font = Font(color="6B7280")
    stats = [("Flagged for review", t["y"], "0"), ("Clear to pay", t["n"] - t["y"], "0"),
             ("USD on flagged invoices", t["usd_flagged"], "$#,##0.00"),
             ("USD clear to pay", t["usd_clear"], "$#,##0.00"),
             (f"Double payments avoided ({t['dup_n']} duplicate copies)", t["dup_exposure"], "$#,##0.00"),
             ("Overbilling caught (total vs line items)", t["overbilled"], "$#,##0.00")]
    r0 = 4
    for k, (lab, val, fmt) in enumerate(stats):
        sm.cell(r0 + k, 1, lab); c = sm.cell(r0 + k, 2, val); c.number_format = fmt
        c.font = Font(bold=True, size=12)
        sm.cell(r0 + k, 1).border = sm.cell(r0 + k, 2).border = Border(bottom=thin)
    hdr = r0 + len(stats) + 2
    sm.cell(hdr, 1, "Why invoices were flagged"); sm.cell(hdr, 2, "Files")
    style_header(sm, hdr)
    cats = [("Over $5,000", "approval threshold"), ("Duplicate or resend", "duplicate of"),
            ("Original of a duplicate", "also received"), ("Totals do not add up", "line items add"),
            ("Foreign currency", "billed in"), ("Wrong entity / intercompany", "confirm which entity"),
            ("Unreadable field", "could not read"), ("Past due", "days past due"),
            ("Vendor not on approved list", "approved vendor list")]
    for lab, key in cats:
        sm.append([lab, sum(1 for o in out if o["flag"] == "Y" and key in o["reason"])])
    for c in rules.get("custom", []):
        if c.get("enabled", True):
            sm.append([f"Custom: {c.get('name', '')}", sum(1 for r in rows if c.get("name", "") in r.get("custom_hits", []))])
    last = sm.max_row
    ch = BarChart(); ch.type = "bar"; ch.style = 10
    ch.title = "Flags by reason (a file can have more than one)"
    ch.add_data(Reference(sm, min_col=2, min_row=hdr, max_row=last), titles_from_data=True)
    ch.set_categories(Reference(sm, min_col=1, min_row=hdr + 1, max_row=last))
    ch.legend = None; ch.height, ch.width = 7.5, 16
    ch.dataLabels = DataLabelList(); ch.dataLabels.showVal = True
    ch.dataLabels.showSerName = ch.dataLabels.showCatName = ch.dataLabels.showLegendKey = ch.dataLabels.showPercent = False
    ch.y_axis.delete = False; ch.x_axis.delete = False
    sm.add_chart(ch, f"D{r0}")
    sm.column_dimensions["A"].width = 52; sm.column_dimensions["B"].width = 18

    # --- Rules: so a reviewer can see exactly how the Y/N was decided (read from the same settings the run used)
    ru = wb.create_sheet("Rules")
    ru.append(["rule", "on?", "setting", "why it matters"]); style_header(ru)
    settings = {
        "threshold": f"flag over ${float(rules['threshold'].get('amount', 5000)):,.0f}",
        "duplicate": "same vendor + invoice number, or same vendor + amount + date",
        "line_item_math": f"lines vs total, tolerance ${rules['line_item_math'].get('tolerance', 0.01)}",
        "foreign_currency": f"not USD (EUR counted at {rules['foreign_currency'].get('eur_to_usd', 1.10)} for the dollar limit)",
        "unreadable_field": "vendor, invoice number, due date, amount",
        "portfolio_company": ", ".join(rules["portfolio_company"].get("names", [])) or "(none)",
        "overdue": f"{rules['overdue'].get('days', 30)}+ days past due",
        "new_vendor": f"{len([v for v in rules['new_vendor'].get('known_vendors', []) if v.strip()])} approved vendors listed",
    }
    for key, (label, why) in RULE_INFO.items():
        ru.append([label, "ON" if rules[key].get("enabled") else "off", settings[key], why])
    for c in rules.get("custom", []):
        ru.append([f"Custom: {c.get('name', '')}", "ON" if c.get("enabled", True) else "off",
                   "invoice mentions: " + ", ".join(c.get("contains", [])), c.get("reason", "")])
    for row in ru.iter_rows(min_row=2):
        for c in row:
            c.alignment = Alignment(wrap_text=True, vertical="top")
        row[1].font = Font(bold=True, color=GRN_TX if row[1].value == "ON" else "6B7280")
    for col, w in zip("ABCD", (34, 8, 62, 70)):
        ru.column_dimensions[col].width = w
    ru.append([]); ru.append(["Due dates not printed on an invoice are worked out from invoice date + terms. "
                              "Slash dates are read as US month/day. Overdue is measured from the run date."])
    try:
        wb.save(os.path.join(outdir, "invoice_review.xlsx"))
    except PermissionError:
        sys.exit("Could not save invoice_review.xlsx because it is open in Excel (or locked by Drive). "
                 "Close the spreadsheet and click Run again.")
    with open(os.path.join(outdir, "rules_used.json"), "w", encoding="utf-8") as f:   # audit trail of this run
        json.dump(rules, f, indent=2)

    write_digest(rows, outdir, as_of, t, flagged)
    with open(os.path.join(outdir, "run_summary.json"), "w", encoding="utf-8") as f:   # read by the app window
        json.dump({
            "run_date": as_of.isoformat(), "files": t["n"], "flagged": t["y"], "clear": t["n"] - t["y"],
            "usd_clear": round(t["usd_clear"], 2), "dup_exposure": round(t["dup_exposure"], 2),
            "overbilled": round(t["overbilled"], 2),
            "queue": [{"file": r["file"], "vendor": r.get("vendor"), "invoice_number": r.get("invoice_number"),
                       "amount": r.get("total"), "currency": r.get("currency", "USD"),
                       "priority": triage(r)[0], "why": "; ".join(r["reasons"]), "next_step": triage(r)[1]}
                      for r in flagged]}, f, indent=1)


def write_digest(rows, outdir, as_of, t, flagged):
    """Plain-text Monday email: what needs a person, then what is clear to pay."""
    L = [f"Subject: Invoice review for the week of {as_of.isoformat()}: {t['y']} need a look, "
         f"{t['n'] - t['y']} clear", "",
         f"{t['n']} invoices processed. {t['y']} flagged. Nothing has been paid.", ""]
    if t["dup_exposure"] or t["overbilled"]:
        L.append(f"Caught this week: ${t['dup_exposure']:,.2f} in duplicate invoices and "
                 f"${t['overbilled']:,.2f} of overbilling.\n")
    for label, want in (("NEEDS A CLOSER LOOK FIRST", 1), ("NEEDS APPROVAL", 2)):
        grp = [r for r in flagged if triage(r)[0] == want]
        if not grp:
            continue
        L.append(label)
        for r in grp:
            amt = f"{r['total']:,.2f} {r.get('currency', 'USD')}" if r.get("total") is not None else "?"
            L.append(f"  - {r.get('vendor')} {r.get('invoice_number')}  {amt}  ({r['file']})")
            L.append(f"      Why: {'; '.join(r['reasons'])}")
            L.append(f"      Do: {triage(r)[1]}")
        L.append("")
    clear = sorted((r for r in rows if r["flag"] == "N"), key=lambda r: r.get("due_final") or date.max)
    L.append(f"CLEAR TO PAY ({len(clear)}, ${t['usd_clear']:,.2f}), earliest due first")
    for r in clear:
        d = r["due_final"].isoformat() if r.get("due_final") else "no date"
        L.append(f"  - {d}  {r.get('vendor')} {r.get('invoice_number')}  ${r['total']:,.2f}")
    open(os.path.join(outdir, "monday_digest.txt"), "w", encoding="utf-8").write("\n".join(L) + "\n")
    write_digest_html(rows, outdir, as_of, t, flagged, clear)


def write_digest_html(rows, outdir, as_of, t, flagged, clear):
    """Same content as the text digest, as tables with inline styles (email clients ignore <style> blocks)."""
    from html import escape as e
    th = "text-align:left;padding:6px 10px;background:#1f3a5f;color:#fff;font-size:13px"
    td = "padding:6px 10px;border-bottom:1px solid #ddd;font-size:13px;vertical-align:top"
    def table(head, body):
        return ('<table cellspacing="0" cellpadding="0" style="border-collapse:collapse;width:100%;margin:6px 0 18px">'
                "<tr>" + "".join(f'<th style="{th}">{h}</th>' for h in head) + "</tr>" + body + "</table>")
    def amt(r):
        return f"{r['total']:,.2f} {e(r.get('currency', 'USD'))}" if r.get("total") is not None else "?"
    h = ['<div style="font-family:Arial,Helvetica,sans-serif;color:#222;max-width:900px">',
         f'<h2 style="margin:0 0 4px">Invoice review, week of {as_of.isoformat()}</h2>',
         f'<p style="margin:0 0 10px"><b>{t["n"]}</b> invoices processed, <b>{t["y"]}</b> flagged, '
         f'<b>{t["n"] - t["y"]}</b> clear. Nothing has been paid.</p>']
    if t["dup_exposure"] or t["overbilled"]:
        h.append(f'<p style="margin:0 0 14px;padding:8px 10px;background:#fff4e5;border-left:4px solid #e08a00">'
                 f'Caught this week: <b>${t["dup_exposure"]:,.2f}</b> in duplicate invoices and '
                 f'<b>${t["overbilled"]:,.2f}</b> of overbilling.</p>')
    from urllib.parse import quote
    def ask_link(r):
        """mailto link that opens a ready-to-send approval request. The To address is filled in by the Apps Script."""
        subj = f"Approval needed: {r.get('vendor')} {r.get('invoice_number')} {amt(r)}"
        msg = (f"Please approve or decline this invoice. Nothing has been paid.\n\n"
               f"Vendor: {r.get('vendor')}\nInvoice: {r.get('invoice_number')}\nAmount: {amt(r)}\n"
               f"File: {r['file']}\nWhy flagged: {'; '.join(r['reasons'])}\n\nReply with: APPROVE or DECLINE (and a note).")
        return (f'<a href="mailto:?subject={quote(subj)}&body={quote(msg)}" '
                f'style="display:inline-block;padding:5px 10px;background:#1f3a5f;color:#fff;'
                f'text-decoration:none;border-radius:4px;font-size:12px;white-space:nowrap">Send for approval</a>')
    for label, want in (("Needs a closer look first", 1), ("Needs approval", 2)):
        grp = [r for r in flagged if triage(r)[0] == want]
        if not grp:
            continue
        body = "".join(
            f'<tr><td style="{td}"><b>{e(str(r.get("vendor")))}</b><br>{e(str(r.get("invoice_number")))}'
            f'<br><span style="color:#777">{e(r["file"])}</span></td>'
            f'<td style="{td};white-space:nowrap;text-align:right">{amt(r)}</td>'
            f'<td style="{td}">{e("; ".join(r["reasons"]))}</td>'
            f'<td style="{td}">{e(triage(r)[1])}</td>'
            + (f'<td style="{td}">{ask_link(r)}</td>' if want == 2 else "") + '</tr>' for r in grp)
        h.append(f'<h3 style="margin:14px 0 0">{label} ({len(grp)})</h3>')
        h.append(table(["Vendor / invoice", "Amount", "Why flagged", "What to do"] + (["Approval"] if want == 2 else []), body))
    body = "".join(
        f'<tr><td style="{td}">{r["due_final"].isoformat() if r.get("due_final") else "no date"}</td>'
        f'<td style="{td}">{e(str(r.get("vendor")))}</td><td style="{td}">{e(str(r.get("invoice_number")))}</td>'
        f'<td style="{td};text-align:right">${r["total"]:,.2f}</td></tr>' for r in clear)
    h.append(f'<h3 style="margin:14px 0 0">Clear to pay ({len(clear)}, ${t["usd_clear"]:,.2f}), earliest due first</h3>')
    h.append(table(["Due", "Vendor", "Invoice", "Amount"], body))
    h.append('<p style="color:#666;font-size:12px;margin-top:6px">Record each decision in the Review Queue tab of '
             'invoice_review.xlsx (Decision column). This tool never pays anything.</p>')
    h.append("</div>")
    open(os.path.join(outdir, "monday_digest.html"), "w", encoding="utf-8").write("\n".join(h))


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("folder"); ap.add_argument("--out", default="out")
    ap.add_argument("--as-of", default=date.today().isoformat(), help="date used for overdue checks")
    ap.add_argument("--rules", help="settings file (default: rules.json next to this script, if present)")
    ap.add_argument("--threshold", type=float, help="override the dollar limit from the settings file")
    ap.add_argument("--history", help="json file remembering past runs: skips files already done, catches cross-week duplicates")
    a = ap.parse_args()
    rules_path = a.rules or os.path.join(os.path.dirname(os.path.abspath(__file__)), "rules.json")
    rules = load_rules(rules_path)
    if a.threshold is not None:
        rules["threshold"]["amount"] = a.threshold
    as_of = date.fromisoformat(a.as_of)
    a.out = a.out.replace("{date}", as_of.isoformat())      # lets a weekly job write to a dated folder
    hist = json.load(open(a.history)) if a.history and os.path.exists(a.history) else {}
    prior = []
    for h in hist.values():
        h = dict(h, prior=True)
        h["invoice_date"] = date.fromisoformat(h["invoice_date"]) if h.get("invoice_date") else None
        prior.append(h)
    paths = [os.path.join(a.folder, f) for f in sorted(os.listdir(a.folder))
             if not f.startswith(".") and os.path.isfile(os.path.join(a.folder, f))]
    rows = []
    for path in paths:
        f = os.path.basename(path)
        digest = hashlib.sha256(open(path, "rb").read()).hexdigest()
        key = f"{digest}:{f}"            # same bytes under a NEW name is a resend, so it must not be skipped
        if key in hist:
            print(f"{f:36} skipped (already processed in an earlier run)")
            continue
        texts, img, note = read_file(path)
        rec, how = extract(texts, img)
        rec = finish(rec); rec["file"] = f; rec["sha"] = key
        print(f"{f:36} {how:5} {rec.get('vendor')!s:32.32} {rec.get('invoice_number')!s:18} {rec.get('total')}  {note}")
        if rec.get("euro_misread"):
            print("    note: the scan shows a euro sign but nothing else says euros, so it was read as USD "
                  "(OCR often misreads $).")
        rows.append(rec)
    rows = apply_flags(rows, as_of, prior, rules)
    write_outputs(rows, a.out, as_of, rules)
    if a.history:
        for r in rows:
            hist[r["sha"]] = {"file": r["file"], "vendor": r.get("vendor"), "invoice_number": r.get("invoice_number"),
                              "total": r.get("total"), "invoice_date": r["invoice_date"].isoformat() if r.get("invoice_date") else None}
        json.dump(hist, open(a.history, "w"), indent=1)
    print(f"\n{len(rows)} files, {sum(r['flag']=='Y' for r in rows)} flagged -> {a.out}/ (invoice_review.xlsx, invoice_review.csv, monday_digest.txt)")


if __name__ == "__main__":
    main()
