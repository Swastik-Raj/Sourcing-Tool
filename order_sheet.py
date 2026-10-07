"""Order sheet: turns the decision agent's approved-orders export into the Excel sheet a person uses to place orders by
hand. Fully deterministic: no network calls, no LLM, no cart automation, no purchasing. It only reads files.

    python order_sheet.py build --approved approved_orders.csv --results "Excel Output Sheets\\sourcing_results_<ts>.xlsx" [--duty 0.35] [--out "Order Sheets"]

Reads: approved_orders.csv (from `decision_agent.py export`), the results file, email_state.json (seller quotes, if any),
reviewer_exclusions.csv and lcom_prices.csv. Writes only Order Sheets\\order_sheet_<timestamp>.xlsx.
"""
import argparse
import csv
import glob
import hashlib
import io
import json
import os
import re
import shutil
import subprocess
import sys
import tempfile
from datetime import datetime, timedelta, timezone
from decimal import ROUND_HALF_UP, Decimal, InvalidOperation
from urllib.parse import urlsplit

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

from decision_agent import APPROVAL_TTL_DAYS, LARGE_TOTAL, NOT_AN_ORDER, parse_moq   # one source for the 7 days and the limit
from email_agent import excluded_reason, listing_key, load_exclusions

# --- Platform ordering behaviour: an UNVERIFIED ASSUMPTION, not a fact. Used only for the "How to order" column. ---
PLATFORM_ORDERING = (      # (text found in the URL host, label shown, how it is believed to work)
    ("alibaba.com", "Alibaba", "add-to-cart when logged in"),
    ("aliexpress.", "AliExpress", "add-to-cart when logged in"),
    ("made-in-china.com", "Made-in-China", "inquiry-based, no checkout seen"),
)
UNKNOWN_HOW_TO_ORDER = "unknown, check the listing"

HERE = os.path.dirname(os.path.abspath(__file__))
EMAIL_STATE_PATH = os.path.join(HERE, "email_state.json")                    # read only
REVIEWER_EXCLUSIONS_CSV = os.path.join(HERE, "reviewer_exclusions.csv")      # read only
LCOM_PRICES_CSV = os.path.join(HERE, "lcom_prices.csv")                      # read only; only a fallback for source/date
DEFAULT_OUT = os.path.join(HERE, "Order Sheets")
DEFAULT_DUTY = Decimal("0.35")
DUTY_SOURCE = ("Estimate only. Source: Srijan's 'China to Walmart: Top 3 Products' doc, Oct 5, 2026 "
               "(TariffTracker, March 2026); a customs broker should confirm.")
SCENARIOS = (200, 500)                       # the "Cost at N pcs" columns
CENT = Decimal("0.01")

REQUIRED_APPROVED = ["sku", "seller", "listing_url", "qty", "max_unit_price", "approval_id", "approved_by", "approved_at",
                     "expires_at", "source_results_file", "flags"]
REQUIRED_RESULTS = ["SKU", "Keyword", "Recommended Manufacturer", "Recommended URL", "Ordering Note", "L-Com Unit Price",
                    "L-Com Price Source", "Manufacturer 1 Name", "Manufacturer 1 Listed Price", "Manufacturer 1 Unit Price",
                    "Manufacturer 1 MOQ", "Manufacturer 1 URL", "Manufacturer 1 Match Tier", "Manufacturer 1 Comment"]
ORDER_HEADERS = ["#", "L-Com SKU", "Product", "Supplier", "Platform", "Listing link", "Unit price (USD)", "MOQ (pcs)",
                 "Qty to order", "Qty vs MOQ", "Line cost (USD)", "Est. with duty (USD)",
                 f"Cost at {SCENARIOS[0]} pcs (USD)", f"Cost at {SCENARIOS[1]} pcs (USD)", "How verified",
                 "Check before ordering", "How to order (unverified assumption)"]
OTHER_HEADERS = ["L-Com SKU", "Supplier", "Unit price (USD)", "MOQ (pcs)", "Min. order cost (USD)", "Platform",
                 "Why not in the order list", "Link / source"]
FIRST_ROW, DUTY_CELL = 8, "$C$4"


class OrderSheetError(Exception):
    """A refusal or failed check, shown as a plain message (nothing is kept)."""


# ---------- small helpers ----------

def now() -> datetime:
    return datetime.now(timezone.utc)


def money(value) -> Decimal:
    return Decimal(str(value)).quantize(CENT, ROUND_HALF_UP)


def fmt_money(d) -> str:
    return f"${money(d):,.2f}"


def to_decimal(value):
    """A cell or text -> finite Decimal, or None (also for 'UNDETERMINED', blanks and junk)."""
    if value is None or isinstance(value, bool):
        return None
    try:
        d = Decimal(str(value).replace(",", "").replace("$", "").strip())
    except InvalidOperation:
        return None
    return d if d.is_finite() else None


def parse_dt(text: str) -> datetime:
    return datetime.strptime(str(text).strip(), "%Y-%m-%dT%H:%M:%SZ").replace(tzinfo=timezone.utc)


def sha256(path: str) -> str:
    with open(path, "rb") as f:
        return hashlib.sha256(f.read()).hexdigest()


def run_id(results_path: str) -> str:
    m = re.search(r"(\d{8}_\d{6})", os.path.basename(results_path))
    return m[1] if m else os.path.basename(results_path)


def platform_of(url: str) -> tuple:
    """(label, how to order) from the URL host. The 'how' comes from PLATFORM_ORDERING, an unverified assumption."""
    host = (urlsplit(url or "").hostname or "").lower()
    for needle, label, how in PLATFORM_ORDERING:
        if needle in host:
            return label, how
    return (host or "unknown"), UNKNOWN_HOW_TO_ORDER


PRICE_RANGE = re.compile(r"\$\s*(\d+(?:\.\d+)?)\s*[-–]\s*\$?\s*(?![\d.]+\s*%)(\d+(?:\.\d+)?)")


def price_range_of(listed) -> tuple:
    """(low, high) when the listed price text is a range like '$2-16' or 'US$3.82-5.22', else None."""
    m = PRICE_RANGE.search(str(listed or ""))
    return (Decimal(m[1]), Decimal(m[2])) if m and Decimal(m[2]) > Decimal(m[1]) else None


# ---------- the approved-orders export ----------

def read_approved(path: str) -> tuple:
    """(valid rows, invalid rows as (line number, reason)). Refuses a file that isn't a decision-agent export."""
    try:
        raw = open(path, encoding="utf-8-sig", newline="").read()
    except FileNotFoundError:
        raise OrderSheetError(f"Approved-orders file not found: {path}")
    rows = list(csv.reader(io.StringIO(raw)))
    if not rows or [c.strip() for c in rows[0]] != [NOT_AN_ORDER]:
        raise OrderSheetError(f"{os.path.basename(path)} does not start with the line '{NOT_AN_ORDER}'. "
                              "This is not a decision-agent export (python decision_agent.py export). Nothing was built.")
    if len(rows) < 2:
        raise OrderSheetError(f"{os.path.basename(path)} has no header row. Nothing was built.")
    header = [h.strip() for h in rows[1]]
    missing = [h for h in REQUIRED_APPROVED if h not in header]
    if missing:
        raise OrderSheetError(f"{os.path.basename(path)} is missing required column(s): {', '.join(missing)}. Nothing was built.")
    valid, invalid = [], []
    for n, cells in enumerate(rows[2:], start=3):
        if not any(c.strip() for c in cells):
            continue
        r = dict(zip(header, cells))
        qty, cap = to_decimal(r.get("qty")), to_decimal(r.get("max_unit_price"))
        try:
            approved, expires = parse_dt(r["approved_at"]), parse_dt(r["expires_at"])
        except (ValueError, KeyError):
            invalid.append((n, "approved_at / expires_at is not a valid timestamp"))
            continue
        problem = ("no sku" if not r.get("sku", "").strip() else "no listing_url" if not r.get("listing_url", "").strip() else
                   "qty is not a positive whole number" if qty is None or qty <= 0 or qty != qty.to_integral_value() else
                   "max_unit_price is not a positive amount" if cap is None or cap <= 0 else
                   "no approval_id" if not r.get("approval_id", "").strip() else "")
        if problem:
            invalid.append((n, problem))
            continue
        valid.append({**r, "sku": r["sku"].strip(), "listing_url": r["listing_url"].strip(), "qty": int(qty), "max": cap,
                      "approved_dt": approved, "expires_dt": expires,
                      "flag_list": [f for f in (r.get("flags") or "").split(";") if f]})
    if not valid:
        detail = "; ".join(f"line {n}: {why}" for n, why in invalid)
        raise OrderSheetError(f"{os.path.basename(path)} has zero valid approval rows" + (f" ({detail})" if detail else "") + ". Nothing was built.")
    return valid, invalid


def approval_expiry(ap: dict) -> datetime:
    """The earlier of: approved_at + the decision agent's validity period (re-computed here) and the file's own expires_at."""
    return min(ap["approved_dt"] + timedelta(days=APPROVAL_TTL_DAYS), ap["expires_dt"])


# ---------- the results file ----------

def read_results(path: str) -> list:
    """Results sheet rows as {header: value}, read by header name. Refuses a file missing a required column."""
    try:
        wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    except FileNotFoundError:
        raise OrderSheetError(f"Results file not found: {path}")
    if "Results" not in wb.sheetnames:
        wb.close()
        raise OrderSheetError(f"{os.path.basename(path)} has no 'Results' sheet.")
    rows = list(wb["Results"].iter_rows(values_only=True))
    wb.close()
    header = [str(h).strip() if h is not None else "" for h in (rows[0] if rows else [])]
    missing = [h for h in REQUIRED_RESULTS if h not in header]
    if missing:
        raise OrderSheetError(f"{os.path.basename(path)}: the Results sheet is missing column(s): {', '.join(missing)}. Nothing was built.")
    return [dict(zip(header, r)) for r in rows[1:] if r and r[header.index("SKU")]]


def candidates_of(row: dict) -> list:
    out = []
    for n in range(1, 10):
        if f"Manufacturer {n} URL" not in row:
            break
        name, url = str(row.get(f"Manufacturer {n} Name") or "").strip(), str(row.get(f"Manufacturer {n} URL") or "").strip()
        if not (name or url):
            continue
        moq_text = str(row.get(f"Manufacturer {n} MOQ") or "").strip()
        out.append({"n": n, "name": name, "url": url, "listed": str(row.get(f"Manufacturer {n} Listed Price") or "").strip(),
                    "unit": to_decimal(row.get(f"Manufacturer {n} Unit Price")), "moq_text": moq_text, "moq": parse_moq(moq_text),
                    "tier": str(row.get(f"Manufacturer {n} Match Tier") or ""), "accuracy": row.get(f"Manufacturer {n} Accuracy"),
                    "comment": str(row.get(f"Manufacturer {n} Comment") or ""), "vs": str(row.get(f"Manufacturer {n} vs. L-Com Price") or ""),
                    "title": str(row.get(f"Manufacturer {n} Listing Title") or "").strip()})
    return out


def find_candidate(row: dict, sku: str, url: str):
    """The candidate with this exact listing (SKU + URL, never SKU alone), or None."""
    want = listing_key(sku, url, "")
    return next((c for c in candidates_of(row) if c["url"] and listing_key(sku, c["url"], "") == want), None)


def lcom_source(row: dict, sku: str, lcom_path: str) -> tuple:
    """(source, date): from the results file; only if that is blank, from lcom_prices.csv."""
    source, date = str(row.get("L-Com Price Source") or "").strip(), str(row.get("L-Com Price Date") or "").strip()
    if not source and os.path.exists(lcom_path):
        with open(lcom_path, newline="", encoding="utf-8") as f:
            for r in csv.DictReader(f):
                if r["sku"].strip().upper() == sku.upper():
                    return r.get("source", "").strip(), r.get("date_checked", "").strip()
    return source, date


# ---------- seller quotes (email_state.json, read only) ----------

QUOTE_PRICE_FIELDS = ("sample_unit_price", "sample_total_price", "branding_fee", "shipping_cost")


def load_quotes(path: str) -> dict:
    try:
        with open(path, encoding="utf-8") as f:
            data = json.load(f)
        return data if isinstance(data, dict) else {}
    except (OSError, ValueError):
        return {}


def quote_of(entry) -> dict:
    """The seller's recorded quote, or None when the entry carries no price or quantity."""
    if not isinstance(entry, dict):
        return None
    has = any(entry.get(k) not in (None, "", []) for k in QUOTE_PRICE_FIELDS + ("sample_quantity", "bulk_tiers"))
    return entry if has else None


def quote_priced(q: dict) -> bool:
    return any(q.get(k) not in (None, "") for k in QUOTE_PRICE_FIELDS) or bool(q.get("bulk_tiers"))


def field_verified(q: dict, field: str) -> bool:
    """Evidence-verified: the email agent stored a verbatim evidence quote for this field and warned nothing about it."""
    evidence = q.get("evidence") if isinstance(q.get("evidence"), dict) else {}
    warned = any(str(w).startswith(f"{field}:") for w in (q.get("quote_warnings") or []))
    return bool(str(evidence.get(field) or "").strip()) and not warned


def verified_tiers(q: dict) -> list:
    tiers = []
    for t in (q.get("bulk_tiers") or []):
        if isinstance(t, dict) and str(t.get("evidence") or "").strip():
            qty, price = to_decimal(t.get("min_qty")), to_decimal(t.get("unit_price"))
            if qty and qty > 0 and price is not None and price >= 0:
                tiers.append((int(qty), price))
    return sorted(tiers)


def quote_price_for(q: dict, qty: int):
    """(price, description) the seller's quote gives for THIS quantity, or (None, why not). Only a USD, evidence-verified
    quote is ever used: a bulk tier whose minimum is met, else the per-unit sample price when the order fits the sample quantity."""
    if (q.get("currency") or "") != "USD":
        return None, f"quote currency is {q.get('currency') or 'unknown'}, not USD"
    met = [t for t in verified_tiers(q) if t[0] <= qty]
    if met:
        n, price = met[-1]
        return price, f"seller quote, bulk tier >= {n:,} pcs"
    price, sample_qty = to_decimal(q.get("sample_unit_price")), to_decimal(q.get("sample_quantity"))
    if (price is not None and field_verified(q, "sample_unit_price") and q.get("price_basis") == "per_unit"
            and sample_qty and qty <= sample_qty):
        return price, "seller quote, sample price"
    return None, "no verified quoted price applies at this quantity"


def quote_summary(q: dict) -> str:
    cur = q.get("currency") or "currency unknown"
    parts = []
    if q.get("sample_quantity"):
        parts.append(f"sample qty {q['sample_quantity']}")
    if q.get("sample_unit_price") is not None:
        parts.append(f"sample ${q['sample_unit_price']} per unit" if cur == "USD" else f"sample {q['sample_unit_price']} {cur} per unit")
    if q.get("bulk_tiers"):
        parts.append("bulk " + ", ".join(f">= {t.get('min_qty')} pcs {t.get('unit_price')}" for t in q["bulk_tiers"] if isinstance(t, dict)))
    if q.get("moq"):
        parts.append(f"quoted MOQ {q['moq']}")
    return f"{', '.join(parts) or 'no prices'} ({cur})"


# ---------- building the rows ----------

FLAG_TEXT = {
    "lcom_price_unverified": "L-Com reference price is unverified.",
    "manual_review_tier": "Match tier was 'flagged for manual review' (accuracy 80-89%).",
    "needs_seller_quote_for_smaller_qty": "The MOQ total was over the limit when this was approved; the approval gave an explicit qty.",
    "large_total": "Large quantity total.",
    "moq_not_stated": "MOQ was not stated on the listing.",
    "quote_above_listed_price": "A seller quote was well above the listed price when this was approved.",
    "quote_no_longer_clears_margin_bar": "A seller quote did not clear the margin bar when this was approved.",
    "seller_needs_info": "The seller had asked for more information.",
    "waiting_on_us": "The seller was waiting on us when this was approved.",
    "no_seller_reply": "No seller reply was recorded when this was approved.",
    "listed_price_above_approved_max": "The listed price was above the approved max when exported.",
}


def build_rows(approved: list, results: list, results_path: str, quotes: dict, exclusions: list, lcom_path: str, at: datetime) -> dict:
    by_sku = {str(r["SKU"]).strip().upper(): r for r in results}
    out = {"rows": [], "expired": [], "could_not_build": [], "excluded": [], "non_usd": [], "warnings": [], "other_skus": []}
    rid = run_id(results_path)
    for ap in approved:
        sku, url = ap["sku"], ap["listing_url"]
        if at >= approval_expiry(ap):
            out["expired"].append((ap, f"expired {approval_expiry(ap).strftime('%Y-%m-%d %H:%M')} UTC"))
            continue
        row = by_sku.get(sku.upper())
        if row is None:
            out["could_not_build"].append((ap, "SKU not found in the results file"))
            continue
        cand = find_candidate(row, sku, url)
        if cand is None:
            out["could_not_build"].append((ap, "listing URL not found among that SKU's candidates in the results file (never matched by SKU alone)"))
            continue
        why = excluded_reason(sku, cand["name"], cand["url"], exclusions)
        if why:
            out["excluded"].append((ap, cand, why))
            continue
        if cand["unit"] is None or cand["unit"] <= 0:
            out["could_not_build"].append((ap, "the results file has no usable unit price for this listing"))
            continue
        is_pick = listing_key(sku, str(row.get("Recommended URL") or ""), "") == listing_key(sku, cand["url"], "")
        entry = quotes.get(listing_key(sku, url, ap.get("seller", "")))
        quote = quote_of(entry)
        csv_cur = (ap.get("quoted_currency") or "").strip()
        non_usd = bool((quote and quote_priced(quote) and (quote.get("currency") or "") != "USD") or (csv_cur and csv_cur != "USD"))
        listed_range = price_range_of(cand["listed"])
        unit, source = cand["unit"], "listing price, not a seller quote" + (" (range, high end)" if listed_range else "")
        moq, notes = cand["moq"], []
        if quote and not non_usd:
            q_price, q_how = quote_price_for(quote, ap["qty"])
            if q_price is not None:
                unit, source = q_price, q_how
            notes.append(f"Seller quote on file: {quote_summary(quote)}; "
                         + (f"price used: {q_how}." if q_price is not None else f"not used ({q_how}); listing price used."))
            q_moq = to_decimal(quote.get("moq"))
            if q_moq and field_verified(quote, "moq") and (moq is None or int(q_moq) > moq):
                moq = int(q_moq)
                notes.append(f"MOQ {moq:,} taken from the seller's verified quote (higher than the listing's).")
        if non_usd:
            cur = (quote.get("currency") if quote else None) or csv_cur or "unknown"
            notes.append(f"Seller quoted in {cur}; needs conversion. NOT in the totals.")
            out["non_usd"].append((ap, cur))
        if is_pick and str(row.get("Ordering Note") or "").strip():
            notes.insert(0, f"Search agent note: {str(row['Ordering Note']).strip()}")
        if not is_pick:
            notes.insert(0, "Not the search agent's own pick for this SKU.")
        if listed_range:
            notes.append(f"Listing shows a price range ({cand['listed']}); the high end is used. The price at your quantity is not stated.")
            if cand["unit"] != listed_range[1] and (cand["moq_text"] or True) and "quoted" not in source:
                notes.append("WARNING: the recorded unit price is not the range's high end - check the listing.")
        if unit > ap["max"]:
            notes.append(f"UNIT PRICE {fmt_money(unit)} IS ABOVE THE APPROVED MAX {fmt_money(ap['max'])}.")
            out["warnings"].append(f"PRICE OVER APPROVED MAX: {sku}: {fmt_money(unit)} > approved max {fmt_money(ap['max'])} ({ap['approval_id']})")
        if moq is None:
            notes.append("MOQ not stated on the listing: ask the seller.")
        else:
            if ap["qty"] < moq:
                out["warnings"].append(f"BELOW MOQ: {sku}: qty {ap['qty']:,} < MOQ {moq:,}")
            if unit * moq > LARGE_TOTAL:
                out["warnings"].append(f"MOQ COST OVER LIMIT: {sku}: the minimum order costs {fmt_money(unit * moq)} "
                                       f"({moq:,} x {fmt_money(unit)}), over the {fmt_money(LARGE_TOTAL)} limit")
                notes.append(f"The minimum order (MOQ x price) is over the {fmt_money(LARGE_TOTAL)} limit.")
        notes += [FLAG_TEXT[f] for f in ap["flag_list"] if f in FLAG_TEXT and FLAG_TEXT[f] not in notes]
        l_src, l_date = lcom_source(row, sku, lcom_path)
        platform, how = platform_of(cand["url"])
        acc = f" Match: {cand['accuracy']}% ({cand['tier']})." if cand["accuracy"] not in (None, "") else ""
        verified = (f"{'Pick from' if is_pick else f'Candidate #{cand['n']} in'} results file {rid}.{acc} Approved {ap['approval_id']} by "
                    f"{ap['approved_by']} on {ap['approved_at'][:10]} ({ap.get('channel') or 'channel not recorded'}): qty {ap['qty']:,}, max unit price "
                    f"{fmt_money(ap['max'])}. Price: {source}. L-Com reference price: {l_src or 'source not recorded'}"
                    f"{f', checked {l_date}' if l_date else ''}. Listing page not re-read by this script.")
        product = str(row.get("Keyword") or "").strip() + (f" - listing: {cand['title']}" if cand["title"] else "")
        out["rows"].append({"ap": ap, "sku": sku, "product": product, "supplier": cand["name"] or "(seller name not shown)",
                            "platform": platform, "how": how, "url": cand["url"], "unit": unit, "moq": moq, "qty": ap["qty"],
                            "in_total": not non_usd, "how_verified": verified, "check": " ".join(notes), "cand": cand, "row": row,
                            "price_source": source})
    return out


def other_options(built: dict, results: list, exclusions: list, results_path: str) -> list:
    """For every SKU on the sheet (or excluded), its other candidates from the results file with the file's own reason."""
    by_sku = {str(r["SKU"]).strip().upper(): r for r in results}
    used = {(r["sku"].upper(), listing_key(r["sku"], r["url"], "")) for r in built["rows"]}
    skus = list(dict.fromkeys([r["sku"] for r in built["rows"]] + [ap["sku"] for ap, _, _ in built["excluded"]]))
    out = []
    for sku in skus:
        row = by_sku[sku.upper()]
        rec = listing_key(sku, str(row.get("Recommended URL") or ""), "")
        approved_here = {ap["listing_url"] for ap in [r["ap"] for r in built["rows"]] + [x[0] for x in built["excluded"]] if ap["sku"].upper() == sku.upper()}
        for c in candidates_of(row):
            key = listing_key(sku, c["url"], "")
            if (sku.upper(), key) in used:
                continue
            reviewer = excluded_reason(sku, c["name"], c["url"], exclusions)
            if reviewer:
                why = f"Excluded by reviewer_exclusions.csv: {reviewer}"
            else:
                bits = []
                if c["url"] and key == rec:
                    bits.append("The search agent's own pick, but not an approved line.")
                bits.append(f"{c['tier'] or 'no tier recorded'}" + (f" ({c['accuracy']}% accuracy)" if c["accuracy"] not in (None, "") else "") + ".")
                if c["comment"]:
                    bits.append(c["comment"].rstrip(".") + ".")
                if c["vs"]:
                    bits.append(f"Price vs L-Com: {c['vs']}.")
                if c["comment"].startswith("All five rubric") or not c["comment"]:
                    bits.append("No rejection reason is recorded for this candidate; the search agent recommended a different seller (see the Recommendation column of the results file).")
                why = " ".join(bits)
            out.append({"sku": sku, "supplier": c["name"] or "(seller name not shown)", "unit": c["unit"], "moq": c["moq"],
                        "platform": platform_of(c["url"])[0] if c["url"] else "", "why": why, "link": c["url"] or "no URL in the results file"})
    return out


# ---------- the workbook ----------

NAVY = PatternFill("solid", fgColor="FF1F3864")
YELLOW = PatternFill("solid", fgColor="FFFFFF00")
THIN = Side(style="thin")
BOX = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
ARIAL = lambda **kw: Font(name="Arial", size=kw.pop("size", 10), **kw)
USD, INT = "\\$#,##0.00", "#,##0"


def fit_height(cells_and_widths, minimum=15.0) -> float:
    lines = 1
    for text, width in cells_and_widths:
        if text:
            lines = max(lines, sum(max(1, -(-len(part) // max(1, int(width * 1.05)))) for part in str(text).split("\n")))
    return max(minimum, 13.2 * lines + 3)


def style_cell(c, fmt=None, font=None, fill=None, wrap=True, border=True):
    c.font = font or ARIAL()
    c.alignment = Alignment(vertical="top", wrap_text=wrap)
    if border:
        c.border = BOX
    if fmt:
        c.number_format = fmt
    if fill:
        c.fill = fill


def write_header(ws, row, headers):
    for col, text in enumerate(headers, start=1):
        c = ws.cell(row, col, text)
        c.font, c.fill, c.border = ARIAL(bold=True, color="FFFFFFFF"), NAVY, BOX
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws.row_dimensions[row].height = 31.5


ORDER_WIDTHS = [4, 15, 30, 28, 14, 13, 11, 10, 10, 11, 13, 13, 13, 13, 52, 52, 28]


def write_workbook(path: str, rows: list, others: list, duty: Decimal, notes: list, prepared: datetime) -> None:
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Order List"
    ws["A1"], ws["A1"].font = "Sample / first order list - L-Com alternatives", ARIAL(size=14, bold=True)
    ws["A2"] = (f"Prepared {prepared.strftime('%Y-%m-%d')}. Listing prices in USD, before shipping, duties and taxes. "
                "Payment is made by Neeraj or Poonam.")
    ws["A2"].font = ARIAL()
    ws["A4"], ws["A4"].font = "Import duty estimate", ARIAL(bold=True)
    ws["C4"] = float(duty)
    ws["C4"].font, ws["C4"].fill = ARIAL(color="FF0000FF"), YELLOW
    ws["C4"].number_format = "0%" if (duty * 100) == (duty * 100).to_integral_value() else "0.0%"
    ws["D4"], ws["D4"].font = DUTY_SOURCE, ARIAL()
    ws["A5"], ws["A5"].font = "Yellow cells are inputs: change a Qty or the duty rate and every cost updates.", ARIAL()
    write_header(ws, 7, ORDER_HEADERS)
    for i, r in enumerate(rows):
        x = FIRST_ROW + i
        cells = [i + 1, r["sku"], r["product"], r["supplier"], r["platform"], "Open listing", float(r["unit"]),
                 r["moq"], r["qty"], f'=IF(H{x}="","MOQ NOT STATED",IF(I{x}<H{x},"BELOW MOQ","ok"))']
        if r["in_total"]:
            cells += [f"=G{x}*I{x}", f"=K{x}*(1+{DUTY_CELL})"] + [f"=G{x}*MAX({s},H{x})" for s in SCENARIOS]
        else:
            cells += ["not in total"] * 4
        cells += [r["how_verified"], r["check"], r["how"]]
        for col, value in enumerate(cells, start=1):
            c = ws.cell(x, col, value)
            style_cell(c, fmt={7: USD, 8: INT, 9: INT, 11: USD, 12: USD, 13: USD, 14: USD}.get(col))
        ws.cell(x, 6).hyperlink = r["url"]
        ws.cell(x, 6).font = ARIAL(color="FF0563C1", underline="single")
        ws.cell(x, 9).font, ws.cell(x, 9).fill = ARIAL(color="FF0000FF"), YELLOW
        ws.row_dimensions[x].height = fit_height([(r["product"], 30), (r["supplier"], 28), (r["how_verified"], 52), (r["check"], 52)], 31.5)
    last, total = FIRST_ROW + len(rows) - 1, FIRST_ROW + len(rows)
    for col in range(1, len(ORDER_HEADERS) + 1):
        style_cell(ws.cell(total, col), font=ARIAL(bold=True), fmt=USD if col in (11, 12, 13, 14) else None)
    ws.cell(total, 2, "Total")
    for col, letter in zip((11, 12, 13, 14), "KLMN"):
        ws.cell(total, col, f"=SUM({letter}{FIRST_ROW}:{letter}{last})")
    ws.cell(total + 2, 1, "Not included: shipping from China, US delivery, duties and taxes (duty shown as an estimate only). "
                          "Prices are listing prices, not confirmed seller quotes unless a row says otherwise.").font = ARIAL()
    for col, width in enumerate(ORDER_WIDTHS, start=1):
        ws.column_dimensions[openpyxl.utils.get_column_letter(col)].width = width
    ws.freeze_panes = f"C{FIRST_ROW}"

    wo = wb.create_sheet("Other options")
    write_header(wo, 1, OTHER_HEADERS)
    for i, o in enumerate(others):
        x = 2 + i
        cells = [o["sku"], o["supplier"], float(o["unit"]) if o["unit"] is not None else "n/a", o["moq"] if o["moq"] is not None else "n/a",
                 f"=C{x}*D{x}" if o["unit"] is not None and o["moq"] is not None else "n/a", o["platform"], o["why"], o["link"]]
        for col, value in enumerate(cells, start=1):
            style_cell(wo.cell(x, col, value), fmt={3: USD, 4: INT, 5: USD}.get(col))
        wo.row_dimensions[x].height = fit_height([(o["why"], 52), (o["link"], 60), (o["supplier"], 34)], 23.85)
    for col, width in enumerate([15, 34, 11, 10, 14, 14, 52, 60], start=1):
        wo.column_dimensions[openpyxl.utils.get_column_letter(col)].width = width
    wo.freeze_panes = "A2"

    wn = wb.create_sheet("Notes")
    wn["A1"], wn["A1"].font = "Notes", ARIAL(size=14, bold=True)
    for i, (label, text) in enumerate(notes, start=3):
        a, b = wn.cell(i, 1, label), wn.cell(i, 2, text)
        a.font, a.alignment = ARIAL(bold=True), Alignment(vertical="top", wrap_text=True)
        b.font, b.alignment = ARIAL(), Alignment(vertical="top", wrap_text=True)
        wn.row_dimensions[i].height = fit_height([(text, 110)], 15.0)
    wn.column_dimensions["A"].width, wn.column_dimensions["B"].width = 28, 110
    wb.calculation.fullCalcOnLoad = True
    wb.save(path)


def make_notes(built: dict, others: list, results_path: str, approved_path: str, approved: list, results: list,
               platforms: dict, extra_files: list) -> list:
    shared = "; ".join(f"{label} = {how}" for label, how in sorted(platforms.items())) or "no listings"
    left = [f"{ap['sku']} ({ap['approval_id']}): {why}" for ap, why in built["expired"]]
    left += [f"{ap['sku']} ({ap['approval_id']}): could not build - {why}" for ap, why in built["could_not_build"]]
    left += [f"{ap['sku']} ({ap['approval_id']}): excluded by reviewer_exclusions.csv ({why}); see 'Other options'" for ap, _, why in built["excluded"]]
    left += [f"{r['sku']} ({r['ap']['approval_id']}): quoted in another currency; shown but NOT in the totals" for r in built["rows"] if not r["in_total"]]
    diffs = []
    on_sheet = {(r["sku"].upper(), listing_key(r["sku"], r["url"], "")) for r in built["rows"]}
    for res in results:
        rec_url = str(res.get("Recommended URL") or "").strip()
        if not rec_url:
            continue
        sku = str(res["SKU"]).strip()
        mine = [ap for ap in approved if ap["sku"].upper() == sku.upper()]
        if not mine:
            diffs.append(f"{sku}: the results file recommends {res.get('Recommended Manufacturer') or 'a seller'}, but there is no approved line for it.")
        elif not any(listing_key(sku, ap["listing_url"], "") == listing_key(sku, rec_url, "") for ap in mine):
            diffs.append(f"{sku}: the results file recommends {res.get('Recommended Manufacturer') or 'a seller'}; the approved line is a different listing ({mine[0].get('seller') or mine[0]['listing_url']}).")
    for ap in approved:
        res = next((r for r in results if str(r["SKU"]).strip().upper() == ap["sku"].upper()), None)
        if res is not None and not str(res.get("Recommended URL") or "").strip():
            diffs.append(f"{ap['sku']}: the results file has no recommended seller; the approved line is {ap.get('seller') or ap['listing_url']}.")
    return [
        ("Cart link", "No cart link is included. A cart belongs to one logged-in account, so a link cannot carry items into someone else's cart. "
                      "Use the 'Open listing' links with the quantities in the Order List. How each platform takes orders is an UNVERIFIED "
                      f"ASSUMPTION from a small table in order_sheet.py, not a fact: {shared}."),
        ("Quantities", "'Qty to order' starts at the quantity Neeraj approved in the approved-orders file. Change the yellow cells to see other "
                       f"costs. 'Cost at {SCENARIOS[0]} pcs' and 'Cost at {SCENARIOS[1]} pcs' show those two quantities, using the larger of that "
                       "quantity and the MOQ (Neeraj mentioned a minimum of 200 each and that 500 can be ordered). Where the quantity is below a "
                       "listing's MOQ, 'Qty vs MOQ' shows BELOW MOQ."),
        ("Samples", "Whether samples are free is not stated in the files this sheet was built from. The seller inquiry drafts ask each seller; "
                    "the email agent's `drafts` output has them."),
        ("Differences vs. results file", " ".join(diffs) if diffs else "None: every recommended seller in the results file matches an approved line."),
        ("Price caveats", "Prices are the listing prices recorded in the results file, not confirmed seller quotes, unless a row says it used a "
                          "seller quote (only USD, evidence-verified quotes are ever used). Listing pages were not re-read when this sheet was built "
                          "and prices may move. A price shown as a range uses the high end. L-Com reference prices marked unverified are not confirmed."),
        ("Duty", "The duty rate in the yellow cell (C4) is an estimate, not a quote. " + DUTY_SOURCE),
        ("Sources", f"Results file: {os.path.basename(results_path)} (sha256 {sha256(results_path)}). Approved-orders file: "
                    f"{os.path.basename(approved_path)} (sha256 {sha256(approved_path)})."
                    + "".join(f" {name} (sha256 {digest})." for name, digest in extra_files)),
        ("Before ordering", "Approvals can be revoked after the approved-orders file was written, and this script does not see that. Re-run "
                            "`python decision_agent.py export` (it only includes valid approvals) if there is any doubt."),
        ("Left out of this sheet", "; ".join(left) if left else "Nothing: every approved line is on the Order List."),
    ]


# ---------- recalculation and cross-checks ----------

class FormulaError(Exception):
    pass


class Evaluator:
    """Recalculates the formulas this script writes (IF, MAX, SUM, + - * / comparisons, cell refs, ranges) in Decimal.
    Any formula it doesn't recognise is an error, so the check can't pass by skipping something."""
    TOKEN = re.compile(r'\s*(?:(?P<num>\d+(?:\.\d+)?)|(?P<str>"(?:[^"]|"")*")|(?P<fn>[A-Z]+)(?=\()|'
                       r'(?P<ref>\$?[A-Z]{1,3}\$?\d+(?::\$?[A-Z]{1,3}\$?\d+)?)|(?P<op><=|>=|<>|[-+*/(),<>=]))')

    def __init__(self, wb):
        self.wb, self.cache, self.stack = wb, {}, set()

    def value(self, ws, coord):
        key = (ws.title, coord)
        if key in self.cache:
            return self.cache[key]
        if key in self.stack:
            raise FormulaError(f"circular reference at {ws.title}!{coord}")
        raw = ws[coord].value
        if isinstance(raw, str) and raw.startswith("="):
            self.stack.add(key)
            try:
                result = self.run(ws, raw[1:])
            except FormulaError as e:
                raise FormulaError(f"{ws.title}!{coord} {raw}: {e}")
            self.stack.discard(key)
        elif isinstance(raw, bool):
            result = raw
        elif isinstance(raw, (int, float)):
            result = Decimal(repr(raw))
        else:
            result = raw
        self.cache[key] = result
        return result

    def run(self, ws, text):
        tokens, pos = [], 0
        while pos < len(text):
            m = self.TOKEN.match(text, pos)
            if not m:
                raise FormulaError(f"cannot parse near {text[pos:pos + 12]!r}")
            kind = m.lastgroup
            tokens.append((kind, m.group(kind)))
            pos = m.end()
        self.tokens, self.i, self.ws = tokens, 0, ws
        result = self.compare()
        if self.i != len(self.tokens):
            raise FormulaError("unexpected trailing tokens")
        return result

    def peek(self):
        return self.tokens[self.i] if self.i < len(self.tokens) else (None, None)

    def take(self, expect=None):
        kind, val = self.peek()
        if expect and val != expect:
            raise FormulaError(f"expected {expect!r}")
        self.i += 1
        return kind, val

    def compare(self):
        left = self.add()
        while self.peek()[1] in ("<", ">", "=", "<=", ">=", "<>"):
            op = self.take()[1]
            right = self.add()
            left = self.cmp(op, left, right)
        return left

    def cmp(self, op, a, b):
        a = "" if a is None else a
        b = "" if b is None else b
        if isinstance(a, str) != isinstance(b, str):
            a, b = (a, b) if op in ("=", "<>") else (str(a), str(b))
        if op == "=":
            return a == b
        if op == "<>":
            return a != b
        return {"<": a < b, ">": a > b, "<=": a <= b, ">=": a >= b}[op]

    def add(self):
        left = self.mul()
        while self.peek()[1] in ("+", "-"):
            op = self.take()[1]
            right = self.mul()
            left = self.num(left) + self.num(right) if op == "+" else self.num(left) - self.num(right)
        return left

    def mul(self):
        left = self.unary()
        while self.peek()[1] in ("*", "/"):
            op = self.take()[1]
            right = self.unary()
            if op == "*":
                left = self.num(left) * self.num(right)
            else:
                if self.num(right) == 0:
                    raise FormulaError("division by zero")
                left = self.num(left) / self.num(right)
        return left

    def unary(self):
        if self.peek()[1] == "-":
            self.take()
            return -self.num(self.unary())
        return self.primary()

    @staticmethod
    def num(v):
        if v is None:
            return Decimal(0)
        if isinstance(v, bool) or not isinstance(v, Decimal):
            raise FormulaError(f"{v!r} is not a number")
        return v

    def args(self):
        out = []
        self.take("(")
        while self.peek()[1] != ")":
            out.append(self.compare)
            break
        return out

    def primary(self):
        kind, val = self.take()
        if kind == "num":
            return Decimal(val)
        if kind == "str":
            return val[1:-1].replace('""', '"')
        if kind == "ref":
            if ":" in val:
                return self.range_values(val)
            return self.value(self.ws, val.replace("$", ""))
        if val == "(":
            inner = self.compare()
            self.take(")")
            return inner
        if kind == "fn":
            return self.function(val)
        raise FormulaError(f"unexpected {val!r}")

    def range_values(self, ref):
        a, b = ref.replace("$", "").split(":")
        (c1, r1), (c2, r2) = self.split(a), self.split(b)
        cells = []
        for col in range(openpyxl.utils.column_index_from_string(c1), openpyxl.utils.column_index_from_string(c2) + 1):
            for row in range(r1, r2 + 1):
                cells.append(self.value(self.ws, f"{openpyxl.utils.get_column_letter(col)}{row}"))
        return cells

    @staticmethod
    def split(a):
        m = re.fullmatch(r"([A-Z]+)(\d+)", a)
        return m[1], int(m[2])

    def function(self, name):
        self.take("(")
        if name == "IF":
            cond = self.compare()
            self.take(",")
            start = self.i
            self.skip_arg()
            a_end = self.i
            self.take(",")
            self.skip_arg()
            b_end = self.i
            self.take(")")
            after = self.i
            self.i = start if cond is True else a_end + 1
            result = self.compare()
            self.i = after
            return result
        values = []
        while True:
            v = self.compare()
            values += v if isinstance(v, list) else [v]
            if self.peek()[1] == ",":
                self.take()
                continue
            self.take(")")
            break
        numbers = [v for v in values if isinstance(v, Decimal)]
        if name == "MAX":
            return max(numbers) if numbers else Decimal(0)
        if name == "SUM":
            return sum(numbers, Decimal(0))
        raise FormulaError(f"unsupported function {name}")

    def skip_arg(self):
        depth = 0
        while self.i < len(self.tokens):
            val = self.tokens[self.i][1]
            if depth == 0 and val in (",", ")"):
                return
            depth += (val == "(") - (val == ")")
            self.i += 1
        raise FormulaError("unbalanced parentheses")


def recalc_builtin(path: str) -> tuple:
    """(values {(sheet, coord): value}, formula count). Raises FormulaError on any formula it can't evaluate."""
    wb = openpyxl.load_workbook(path)
    ev, values, count = Evaluator(wb), {}, 0
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for c in row:
                if isinstance(c.value, str) and c.value.startswith("="):
                    count += 1
                    values[(ws.title, c.coordinate)] = ev.value(ws, c.coordinate)
    return values, count


def find_recalc_script():
    explicit = os.environ.get("ORDER_SHEET_RECALC_PY")
    if explicit and os.path.exists(explicit):
        return explicit
    hits = glob.glob(os.path.join(os.path.expanduser("~"), ".claude", "skills", "**", "xlsx", "scripts", "recalc.py"), recursive=True)
    return hits[0] if hits else None


def libreoffice_check(path: str, expected: dict) -> str:
    """Independent check with the xlsx skill's recalc.py (LibreOffice) on a COPY. Returns a status line; raises on a real failure."""
    script = find_recalc_script()
    if not script:
        return "skipped (recalc.py not found)"
    if not shutil.which("soffice"):
        return "skipped (LibreOffice 'soffice' not found)"
    with tempfile.TemporaryDirectory() as tmp:
        copy = os.path.join(tmp, "check.xlsx")
        shutil.copy(path, copy)
        proc = subprocess.run([sys.executable, script, copy, "90"], capture_output=True, text=True, timeout=240, cwd=os.path.dirname(script))
        try:
            report = json.loads(proc.stdout)
        except ValueError:
            raise OrderSheetError(f"LibreOffice recalculation gave no readable result: {(proc.stdout + proc.stderr)[:300]}")
        if report.get("status") != "success" or report.get("total_errors", 0):
            raise OrderSheetError(f"LibreOffice recalculation found formula errors: {json.dumps(report)[:500]}")
        ws = openpyxl.load_workbook(copy, data_only=True)["Order List"]
        for coord, want in expected.items():
            got = to_decimal(ws[coord].value)
            if got is None or money(got) != money(want):
                raise OrderSheetError(f"LibreOffice total {coord} = {ws[coord].value} but the Decimal cross-check is {money(want)}")
    return f"LibreOffice via recalc.py: {report.get('total_formulas', '?')} formulas, 0 errors, totals match"


def expected_cells(rows: list, duty: Decimal) -> dict:
    """What every computed Order List cell must equal, from Decimal arithmetic done here (never from the workbook)."""
    exp, totals = {}, {c: Decimal(0) for c in "KLMN"}
    for i, r in enumerate(rows):
        x = FIRST_ROW + i
        moq = r["moq"]
        exp[f"J{x}"] = "MOQ NOT STATED" if moq is None else ("BELOW MOQ" if r["qty"] < moq else "ok")
        if r["in_total"]:
            line = r["unit"] * r["qty"]
            vals = {"K": line, "L": line * (1 + duty)}
            for letter, s in zip("MN", SCENARIOS):
                vals[letter] = r["unit"] * max(Decimal(s), Decimal(moq or 0))
            for letter, v in vals.items():
                exp[f"{letter}{x}"] = v
                totals[letter] += v
    total_row = FIRST_ROW + len(rows)
    for letter, v in totals.items():
        exp[f"{letter}{total_row}"] = v
    return exp


def verify_workbook(path: str, rows: list, duty: Decimal) -> tuple:
    """Fails loudly unless the saved workbook recalculates with zero errors and every computed cell and total equals the
    Decimal cross-check. Money cells must be formulas, never typed numbers."""
    try:
        values, count = recalc_builtin(path)
    except FormulaError as e:
        raise OrderSheetError(f"Recalculation failed: {e}")
    wb = openpyxl.load_workbook(path)
    ws = wb["Order List"]
    exp = expected_cells(rows, duty)
    problems = []
    for coord, want in exp.items():
        cell = ws[coord].value
        if not (isinstance(cell, str) and cell.startswith("=")):
            problems.append(f"{coord} is not a formula ({cell!r})")
            continue
        got = values[("Order List", coord)]
        ok = (got == want) if isinstance(want, str) else (isinstance(got, Decimal) and money(got) == money(want))
        if not ok:
            problems.append(f"{coord}: workbook says {got!r}, Decimal cross-check says {want!r}")
    for key, v in values.items():
        if isinstance(v, str) and v.startswith("#"):
            problems.append(f"{key} is an error value {v}")
    if problems:
        raise OrderSheetError("Verification failed, no sheet was kept: " + "; ".join(problems[:8]))
    money_cells = {c: exp[c] for c in exp if c[0] in "KLMN"}
    lo = libreoffice_check(path, {c: v for c, v in money_cells.items() if c.endswith(str(FIRST_ROW + len(rows)))})
    return count, lo, exp


# ---------- build ----------

def build(approved_path: str, results_path: str, duty: Decimal = DEFAULT_DUTY, out_dir: str = None, email_state_path: str = None,
          exclusions_path: str = None, lcom_path: str = None, at: datetime = None) -> dict:
    at = at or now()
    if not (Decimal(0) <= duty < Decimal(1)):
        raise OrderSheetError("--duty must be a rate between 0 and 1 (for example 0.35).")
    approved, invalid = read_approved(approved_path)
    results = read_results(results_path)
    exclusions = load_exclusions(exclusions_path or REVIEWER_EXCLUSIONS_CSV)
    quotes = load_quotes(email_state_path or EMAIL_STATE_PATH)
    built = build_rows(approved, results, results_path, quotes, exclusions, lcom_path or LCOM_PRICES_CSV, at)
    built["invalid"] = invalid
    mismatch = sorted({ap["source_results_file"] for ap in approved if ap["source_results_file"] != os.path.basename(results_path)})
    if mismatch:
        built["warnings"].append(f"RESULTS FILE DIFFERS: the approvals were made against {', '.join(mismatch)}, but --results is {os.path.basename(results_path)}")
    if not built["rows"]:
        raise OrderSheetError("Nothing to put on the Order List: " + "; ".join(
            [f"{ap['sku']} {why}" for ap, why in built["expired"]] + [f"{ap['sku']} {why}" for ap, why in built["could_not_build"]]
            + [f"{ap['sku']} excluded by reviewer" for ap, _, _ in built["excluded"]]) + ". No sheet was written.")
    others = other_options(built, results, exclusions, results_path)
    platforms = {r["platform"]: r["how"] for r in built["rows"]}
    extra = [(os.path.basename(p), sha256(p)) for p in (email_state_path or EMAIL_STATE_PATH,) if os.path.exists(p)]
    notes = make_notes(built, others, results_path, approved_path, approved, results, platforms, extra)
    out_dir = out_dir or DEFAULT_OUT
    os.makedirs(out_dir, exist_ok=True)
    final = os.path.join(out_dir, f"order_sheet_{at.strftime('%Y%m%d_%H%M%S')}.xlsx")
    temp = final[:-5] + ".verifying.xlsx"
    write_workbook(temp, built["rows"], others, duty, notes, at)
    try:
        count, lo, exp = verify_workbook(temp, built["rows"], duty)
        os.replace(temp, final)
    except Exception:
        if os.path.exists(temp):
            os.remove(temp)
        raise
    total_row = FIRST_ROW + len(built["rows"])
    built.update(path=final, formulas=count, libreoffice=lo, others=others,
                 totals={c: money(exp[f"{c}{total_row}"]) for c in "KLMN"})
    return built


def print_summary(b: dict, approved_path: str, results_path: str) -> None:
    print(f"Approved-orders file: {approved_path} (sha256 {sha256(approved_path)[:16]}...)")
    print(f"Results file:         {results_path} (sha256 {sha256(results_path)[:16]}...)")
    print(f"Order List rows: {len(b['rows'])}")
    for ap, why in b["expired"]:
        print(f"  EXPIRED, skipped: {ap['sku']} ({ap['approval_id']}): {why}")
    print("Could not build:" + (" none" if not b["could_not_build"] else ""))
    for ap, why in b["could_not_build"]:
        print(f"  {ap['sku']} ({ap['approval_id']}): {why}")
    for ap, _, why in b["excluded"]:
        print(f"  EXCLUDED BY REVIEWER (only in 'Other options'): {ap['sku']} ({ap['approval_id']}): {why}")
    for ap, cur in b["non_usd"]:
        print(f"  NOT IN TOTALS, quoted in {cur}: {ap['sku']} ({ap['approval_id']})")
    for n, why in b["invalid"]:
        print(f"  INVALID ROW at line {n} of the approved file, skipped: {why}")
    for w in b["warnings"]:
        print(f"WARNING: {w}")
    t = b["totals"]
    print(f"Totals (Decimal, matched to the recalculated workbook): line cost {fmt_money(t['K'])} | with duty {fmt_money(t['L'])} | "
          f"at {SCENARIOS[0]} pcs {fmt_money(t['M'])} | at {SCENARIOS[1]} pcs {fmt_money(t['N'])}")
    print(f"Recalculation: built-in evaluator OK ({b['formulas']} formulas, 0 errors, all computed cells and totals equal the Decimal cross-check); "
          f"{b['libreoffice']}")
    print(f"Written: {b['path']}")


def main(argv=None):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    if hasattr(sys.stdin, "reconfigure"):
        sys.stdin.reconfigure(encoding="utf-8")
    p = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    sub = p.add_subparsers(dest="cmd", required=True)
    s = sub.add_parser("build")
    s.add_argument("--approved", required=True, help="approved_orders.csv from `decision_agent.py export`")
    s.add_argument("--results", required=True, help="sourcing_results_*.xlsx")
    s.add_argument("--duty", default=str(DEFAULT_DUTY), help="import duty estimate as a rate, default 0.35")
    s.add_argument("--out", default=DEFAULT_OUT, help="output folder, default 'Order Sheets'")
    args = p.parse_args(argv)
    duty = to_decimal(args.duty)
    try:
        if duty is None:
            raise OrderSheetError(f"--duty '{args.duty}' is not a number.")
        b = build(args.approved, args.results, duty, args.out)
        print_summary(b, args.approved, args.results)
    except OrderSheetError as e:
        sys.exit(f"Order sheet NOT built: {e}")


if __name__ == "__main__":
    main()
