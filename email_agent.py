"""Email agent: drafts (and, after confirmation, sends) sample-order outreach for the sourcing agent's
recommended manufacturers. Reads a sourcing_results_*.xlsx; never touches sourcing_agent.py.

    python email_agent.py drafts  <results.xlsx>                        # show every draft; read-only, sends nothing
    python email_agent.py send    <results.xlsx> [--yes] [--test-recipient ADDR]
    python email_agent.py mark    <SKU> --results <results.xlsx>        # record a manual inquiry-form submission
    python email_agent.py reply   <SKU> --results <results.xlsx> [--file reply.txt]   # paste reply, summarize
    python email_agent.py report  <results.xlsx>                        # markdown + Excel report for Neeraj

State (email_state.json) is keyed by SKU + listing URL, so a re-run that recommends a different seller for a SKU
never shows the old seller's reply or sent mark under the new one.

Env (.env): SMTP_HOST, SMTP_PORT (587), SMTP_USER, SMTP_PASSWORD, SMTP_FROM (default SMTP_USER), SENDER_NAME,
ANTHROPIC_API_KEY, and SHIPPING_ADDRESS - one line in double quotes with \\n between lines
    SHIPPING_ADDRESS="Zync Technologies\\n123 Example St\\nPlano, TX 75024"
or a real multi-line value inside the quotes. Real sends are refused while it is unset or still the default.
"""
import argparse
import csv
import json
import os
import re
import smtplib
import sys
import unicodedata
from datetime import datetime
from decimal import Decimal, InvalidOperation
from email.message import EmailMessage
from urllib.parse import urlsplit

import openpyxl
from dotenv import load_dotenv
from openpyxl.styles import Font

HERE = os.path.dirname(os.path.abspath(__file__))
TEMPLATE_PATH = os.path.join(HERE, "email_template.txt")
STATE_PATH = os.path.join(HERE, "email_state.json")  # per SKU + listing: sent / reply / summary
REVIEWER_EXCLUSIONS_CSV = os.path.join(HERE, "reviewer_exclusions.csv")  # shared with the search agent; read only
NEAR_MISSES_CSV = os.path.join(HERE, "near_misses.csv")  # sku, supplier_or_url, concern, question; human-maintained, read only
MODEL = "claude-haiku-4-5"
DEFAULT_SHIPPING = "Zync Technologies\nPlano, TX"
REPLY_STATUSES = ("pricing_provided", "needs_info", "dead_end", "auto_reply_or_spam")


def cfg(name: str, default: str = "") -> str:
    return os.environ.get(name, default)


class SendRefused(Exception):
    """A real send was blocked before anything went out."""


def shipping_address() -> tuple:
    """(address text, is_default). Unset, blank, or still "Zync Technologies, Plano, TX" means the real
    address was never configured."""
    text = cfg("SHIPPING_ADDRESS").replace("\\n", "\n").strip()
    squash = lambda s: re.sub(r"[\s,]+", " ", s).strip().lower()
    return (text or DEFAULT_SHIPPING), (not text or squash(text) == squash(DEFAULT_SHIPPING))


def warnings_for(exclusions_loaded: bool = True) -> list:
    problems = []
    if shipping_address()[1]:
        problems.append("SHIPPING_ADDRESS is not set (or is still the default 'Zync Technologies, Plano, TX'). "
                        "Drafts use the default; real sends are REFUSED until it is changed in .env.")
    if not exclusions_loaded:
        problems.append(f"{os.path.basename(REVIEWER_EXCLUSIONS_CSV)} is missing or empty - known-bad sellers are "
                        "NOT being skipped.")
    return problems


def print_warnings(problems: list) -> None:
    if problems:
        bar = "!" * 72
        print("\n".join([bar, *(f"!!! WARNING: {p}" for p in problems), bar]))


# ---------- input: read the search agent's Results sheet ----------

def listing_key(sku: str, url: str, manufacturer: str) -> str:
    """State key: SKU + listing URL (host and path only), or the manufacturer's name when there is no URL."""
    ref = (url or "").strip()
    if ref:
        parts = urlsplit(ref)
        ref = f"{parts.netloc.lower()}{parts.path.rstrip('/')}"
    else:
        ref = (manufacturer or "").strip().lower()
    return f"{str(sku).strip().upper()}|{ref}"


def load_exclusions(path: str = None) -> list:
    """Rows of reviewer_exclusions.csv (sku, supplier_or_url, reason); [] when missing. Never written."""
    try:
        with open(path or REVIEWER_EXCLUSIONS_CSV, newline="", encoding="utf-8") as f:
            return [r for r in csv.DictReader(f) if (r.get("supplier_or_url") or "").strip()]
    except FileNotFoundError:
        return []


def excluded_reason(sku: str, manufacturer: str, url: str, exclusions: list) -> str:
    """The reviewer's reason when this SKU's seller or listing URL is in the exclusions file, else ""."""
    haystack = f"{manufacturer or ''} {url or ''}".lower()
    for e in exclusions:
        if e["sku"].strip().upper() == str(sku).strip().upper() and e["supplier_or_url"].strip().lower() in haystack:
            return e["reason"].strip()
    return ""


def load_results(path: str, exclusions: list = None) -> tuple:
    """(candidates, skipped). A candidate is a Results row with a recommended seller, ready to draft. Skipped rows
    are listed with the reason: no recommendation, the search agent errored, or a reviewer exclusion."""
    exclusions = load_exclusions() if exclusions is None else exclusions
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    rows = list(wb["Results"].iter_rows(values_only=True))
    wb.close()  # Windows keeps the file locked otherwise
    header, rows = list(rows[0]), rows[1:]
    out, skipped = [], []
    for values in rows:
        r = dict(zip(header, values))
        sku, product = str(r["SKU"]), r.get("Product") or ""
        skip = {"sku": sku, "product": product, "manufacturer": r.get("Recommended Manufacturer") or ""}
        if str(r.get("Recommendation") or "").startswith("Error researching"):
            skipped.append({**skip, "reason": "error researching this product", "kind": "error"})
            continue
        if not (r.get("Recommended Manufacturer") or r.get("Recommended URL")):
            skipped.append({**skip, "reason": "no recommended seller", "kind": "none"})
            continue
        url = r.get("Recommended URL") or ""
        why = excluded_reason(sku, r.get("Recommended Manufacturer"), url, exclusions)
        if why:
            skipped.append({**skip, "reason": f"excluded by reviewer: {why}", "kind": "reviewer"})
            continue
        moq, title = "", r.get("Recommended Listing Title") or ""
        for n in range(1, 10):  # per-candidate columns: find the recommended one by URL
            if f"Manufacturer {n} URL" not in r:
                break
            if url and r[f"Manufacturer {n} URL"] == url:
                moq = r.get(f"Manufacturer {n} MOQ") or ""
                title = title or r.get(f"Manufacturer {n} Listing Title") or ""
        out.append({
            "sku": sku, "product": product, "description": r.get("Keyword") or "",
            "manufacturer": r.get("Recommended Manufacturer") or "", "email": r.get("Recommended Email") or "",
            "url": url, "unit_price": r.get("Recommended Unit Price"), "moq": moq, "title": title,
            "note": r.get("Ordering Note") or "",
            "key": listing_key(sku, url, r.get("Recommended Manufacturer") or ""),
        })
    return out, skipped


def load_candidates(path: str) -> list:
    return load_results(path)[0]


def load_near_misses(path: str = None):
    """Rows of near_misses.csv (sku, supplier_or_url, concern, question), or None when the file is missing. Never written."""
    try:
        with open(path or NEAR_MISSES_CSV, newline="", encoding="utf-8") as f:
            return [r for r in csv.DictReader(f) if (r.get("sku") or "").strip() and (r.get("supplier_or_url") or "").strip()]
    except FileNotFoundError:
        return None


def near_miss_drafts(results_path: str, recommended: list, exclusions: list, near_rows: list) -> tuple:
    """(drafts, notes): one draft per near_misses.csv row whose candidate is in the results file's candidate columns.
    Opt-in only (drafts --near-miss). The search agent did NOT recommend these, so each draft carries the one concern a
    human must check (printed in the header, never in the text meant for the seller) and one polite question for the
    seller. Reviewer exclusions still apply, and a seller already recommended for that SKU is not drafted twice."""
    wb = openpyxl.load_workbook(results_path, read_only=True, data_only=True)
    rows = list(wb["Results"].iter_rows(values_only=True))
    wb.close()
    header = list(rows[0])
    by_sku = {}
    for values in rows[1:]:
        r = dict(zip(header, values))
        by_sku.setdefault(str(r["SKU"]).strip().upper(), r)
    have = {d["key"] for d in recommended}
    out, notes = [], []
    for near in near_rows:
        sku, wanted = near["sku"].strip(), near["supplier_or_url"].strip()
        r = by_sku.get(sku.upper())
        block = next((n for n in range(1, 10) if r and f"Manufacturer {n} URL" in r and (
            (r[f"Manufacturer {n} URL"] or "").strip() == wanted
            or wanted.lower() in (r.get(f"Manufacturer {n} Name") or "").lower())), None)
        if block is None:
            notes.append(f"near-miss {sku}: not found in the results file ({'no such SKU' if r is None else 'seller or URL not among its candidates'}) - skipped")
            continue
        url, maker = (r[f"Manufacturer {block} URL"] or "").strip(), (r.get(f"Manufacturer {block} Name") or "").strip()
        why = excluded_reason(sku, maker, url, exclusions)
        if why:
            notes.append(f"near-miss {sku} ({maker or url}): excluded by reviewer ({why}) - skipped")
            continue
        key = listing_key(sku, url, maker)
        if key in have:
            notes.append(f"near-miss {sku} ({maker or url}): already a recommended draft - not drafted twice")
            continue
        have.add(key)
        out.append(make_draft({
            "sku": sku, "product": r.get("Product") or "", "description": r.get("Keyword") or "", "manufacturer": maker,
            "email": (r.get(f"Manufacturer {block} Email") or "").strip(), "url": url, "unit_price": None,
            "moq": r.get(f"Manufacturer {block} MOQ") or "", "title": r.get(f"Manufacturer {block} Listing Title") or "",
            "note": "", "key": key, "check": near["concern"].strip(), "questions": [near["question"].strip()]}))
    return out, notes


def format_draft(d: dict) -> str:
    """One draft as printed by `drafts`. The NEEDS HUMAN CHECK and NO EMAIL lines are for us, above the text for the seller."""
    to = d["email"] or "NO EMAIL - inquiry form"
    check = f"NEEDS HUMAN CHECK: {d['check']}\n" if d.get("check") else ""
    label = f"{d['label']}\n" if d["label"] else ""
    return f"{'=' * 70}\n{d['sku']} | {d['manufacturer'] or '(unnamed)'} | To: {to}\n{check}{label}Subject: {d['subject']}\n\n{d['body']}"


def skip_summary(skipped: list) -> str:
    kinds = {"none": "no recommended seller", "error": "search agent error", "reviewer": "excluded by reviewer_exclusions.csv"}
    parts = [f"{sum(s['kind'] == k for s in skipped)} {label}" for k, label in kinds.items() if any(s["kind"] == k for s in skipped)]
    return f"Skipped {len(skipped)} row(s)" + (f": {', '.join(parts)}." if parts else ".")


# ---------- template ----------

# The search agent's Ordering Note is internal ("check before ordering", "target is crimp", the model's own
# caveats). A seller must never see it, so each known note becomes a polite confirmation question and
# anything unrecognised - including all the model's free text after "model:" - is dropped.
NOTE_QUESTIONS = [
    (re.compile(r'title names .+? \("(.+?)"\)'),
     lambda m: f'Your listing title mentions "{m[1]}". Could you confirm the exact variant matches the description above?'),
    (re.compile(r"one listing covers several models \((.+?)\)"),
     lambda m: f"Your listing covers several models ({m[1]}). Could you confirm which one the sample would be?"),
    (re.compile(r"listing offers (.+?) variants"),
     lambda m: f"Your listing offers {m[1]} options. Could you tell us which option to select for this product?"),
    (re.compile(r"price needs an order of ([\d,]+)\+ pieces"),
     lambda m: f"Is the listed price based on an order of {m[1]}+ pieces, and what would it be at sample quantity?"),
    (re.compile(r"price is for the (.+?) tier"),
     lambda m: f"Could you confirm which price tier applies to our sample quantity (the listing shows {m[1]})?"),
    (re.compile(r"wasn't tied to this price"),
     lambda m: "Could you confirm whether the listed price is per piece or per pack?"),
    (re.compile(r"promo price"),
     lambda m: "Could you confirm the standard price for this item?"),
    (re.compile(r"contact type isn't stated"),
     lambda m: "Could you confirm the contact type (crimp, solder cup, PCB or IDC)?"),
    (re.compile(r"several contact types"),
     lambda m: "Your listing mentions more than one contact type. Could you confirm which one applies?"),
    (re.compile(r"contradicts itself on category"),
     lambda m: "Could you confirm the category/rating of this item? The listing seems to mention more than one."),
    (re.compile(r"contradicts itself on gender"),
     lambda m: "Could you confirm the connector gender (male or female)?"),
    (re.compile(r"short cable rather than a plain adapter"),
     lambda m: "Could you confirm this is a standalone adapter/coupler and not a cable with connectors on the ends?"),
]
MAX_QUESTIONS = 4


def confirmation_questions(note: str) -> list:
    """Ordering note -> polite questions for the seller. Free text after "model:" is never used."""
    ours = re.split(r"(?:^|;\s*)model:", note or "", maxsplit=1)[0]
    questions = []
    for segment in ours.split("; "):
        for pattern, ask in NOTE_QUESTIONS:
            m = pattern.search(segment)
            if m:
                q = ask(m)
                if q not in questions:
                    questions.append(q)
                break
    return questions[:MAX_QUESTIONS]


# Names the search agent uses when a listing doesn't name its maker - never greet a seller with them.
PLACEHOLDER_NAME = re.compile(r"unknown|unnamed|generic|not stated|unbranded|seller|^none$", re.IGNORECASE)


def render_email(c: dict, template_path: str = TEMPLATE_PATH) -> tuple:
    """Returns (subject, body). Edit email_template.txt to change the wording."""
    with open(template_path, encoding="utf-8") as f:
        subject, _, body = f.read().partition("\n")
    name = (c["manufacturer"] or "").strip()
    title = (c.get("title") or "").strip()
    # The seller's own listing title when the results file has it; otherwise just the URL - never our description.
    if c["url"]:
        listing_ref = f'your listing "{title}" ({c["url"]})' if title else f"your listing ({c['url']})"
    else:
        listing_ref = "your products"
    values = {
        "sku": c["sku"], "product": c["description"] or c["product"],
        "greeting": "Hello," if not name or PLACEHOLDER_NAME.search(name) else f"Hello {name} team,",
        "listing_ref": listing_ref,
        "note_line": "".join(f"\n{i}. {q}" for i, q in enumerate(
            confirmation_questions(c["note"]) + list(c.get("questions") or []), start=4)),
        "sender_name": cfg("SENDER_NAME", "Swastik Raj"),
        "shipping_address": shipping_address()[0],
    }
    return subject.removeprefix("Subject:").strip().format(**values), body.strip().format(**values) + "\n"


def make_draft(c: dict) -> dict:
    """The body is always clean, ready to paste into a seller's inquiry form. The no-email label is a
    separate field for us: it never travels inside the text a seller receives."""
    subject, body = render_email(c)
    label = "" if c["email"] else f"NO EMAIL - submit via the inquiry form at {c['url']}"
    key = c.get("key") or listing_key(c["sku"], c["url"], c["manufacturer"])
    return {**c, "subject": subject, "body": body, "label": label, "key": key}


# ---------- state ----------

def load_state() -> dict:
    try:
        with open(STATE_PATH, encoding="utf-8") as f:
            return json.load(f)
    except FileNotFoundError:
        return {}


def save_state(state: dict) -> None:
    with open(STATE_PATH, "w", encoding="utf-8") as f:
        json.dump(state, f, indent=2)


def now() -> str:
    return datetime.now().strftime("%Y-%m-%d %H:%M")


def entry_for(state: dict, d: dict) -> dict:
    return state.setdefault(d["key"], {"sku": d["sku"], "seller": d["manufacturer"] or d["url"]})


# ---------- sending (always behind a confirmation) ----------

def smtp_send(draft: dict) -> None:
    msg = EmailMessage()
    msg["Subject"], msg["To"] = draft["subject"], draft["email"]
    msg["From"] = cfg("SMTP_FROM") or cfg("SMTP_USER")
    msg.set_content(draft["body"])
    with smtplib.SMTP(cfg("SMTP_HOST"), int(cfg("SMTP_PORT", "587"))) as s:
        s.starttls()
        s.login(cfg("SMTP_USER"), cfg("SMTP_PASSWORD"))
        s.send_message(msg)


def as_test(d: dict, test_recipient: str) -> dict:
    """A draft redirected to the test address: subject marked [TEST], body says where it would really have gone."""
    where = d["email"] or f"nobody - no email, inquiry form {d['url']}"
    return {**d, "email": test_recipient, "subject": f"[TEST] {d['subject']}",
            "body": f"[TEST - this would have gone to: {where}]\n\n{d['body']}"}


def duplicate_recipients(todo: list) -> dict:
    """{address: [SKUs]} for addresses that more than one email in this batch is going to."""
    groups = {}
    for d in todo:
        groups.setdefault(d["email"].strip().lower(), []).append(d["sku"])
    return {email: skus for email, skus in groups.items() if len(skus) > 1}


def confirm_and_send(drafts: list, state: dict, yes: bool = False, send=smtp_send, ask=input,
                     test_recipient: str = "") -> int:
    """Sends only after an explicit confirmation: 'y' for one email, or the word 'yes' for the whole batch.
    Anything else (Enter, 'n', 'q', typos) sends nothing. Real mode sends only drafts that have an email and
    aren't already sent, and refuses while SHIPPING_ADDRESS is unset or still the default. With test_recipient,
    EVERY draft (even those without an email) goes to that address only, marked [TEST], and is not recorded
    as sent."""
    if test_recipient:
        todo = [as_test(d, test_recipient) for d in drafts]
    else:
        if shipping_address()[1]:
            raise SendRefused("SHIPPING_ADDRESS is not set (or is still the default 'Zync Technologies, Plano, TX'). "
                              "Set the real address in .env before sending. Nothing was sent.")
        todo = [d for d in drafts if d["email"] and "sent" not in state.get(d["key"], {})]
    dups = {} if test_recipient else duplicate_recipients(todo)
    for email, skus in dups.items():
        print(f"WARNING: {len(skus)} separate emails ({', '.join(skus)}) are addressed to the same recipient {email}.")
    if yes:
        print(f"Batch send{' (TEST to ' + test_recipient + ')' if test_recipient else ''}: {len(todo)} email(s).")
        if ask("Type 'yes' to send all: ").strip().lower() != "yes":
            return 0
    sent = 0
    for d in todo:
        if test_recipient and d["email"] != test_recipient:
            raise RuntimeError("test mode would send to a real address - aborting")
        if not yes:
            same = f"\nNOTE: same address as {', '.join(s for s in dups[d['email'].lower()] if s != d['sku'])}" \
                if d["email"].lower() in dups else ""
            print(f"\nTo: {d['email']}\nSubject: {d['subject']}{same}\n\n{d['body']}")
            answer = ask(f"Send {d['sku']} to {d['email']}? [y/N/q] ").strip().lower()
            if answer == "q":
                break
            if answer != "y":
                continue
        send(d)
        if not test_recipient:  # a test send must not make the real one look done
            entry_for(state, d)["sent"] = now()
            save_state(state)
        sent += 1
    return sent


# ---------- reply summarization ----------

# Structured quote fields the summarizer may fill. The model extracts; verify_quote() below checks every number against
# the reply text and nulls what it can't prove. Anything not stated in the reply is None - never guessed.
QUOTE_INT_FIELDS = ("sample_quantity", "moq", "branding_min_qty", "lead_time_days", "production_lead_time_days")
QUOTE_MONEY_FIELDS = ("sample_unit_price", "sample_total_price", "branding_fee", "shipping_cost")
QUOTE_EVIDENCE_FIELDS = QUOTE_INT_FIELDS + QUOTE_MONEY_FIELDS      # each non-null one needs a verbatim evidence quote
QUOTE_KEYS = ("sample_available", "sample_quantity", "sample_unit_price", "sample_total_price", "price_basis", "moq",
              "bulk_tiers", "branding_possible", "branding_fee", "branding_min_qty", "lead_time_days",
              "production_lead_time_days", "shipping_cost", "shipping_terms", "currency", "needs_from_us",
              "quote_valid_until", "evidence", "quote_warnings")
PRICE_BASES = ("per_unit", "total_for_quantity", "unclear")
# Currency counts as USD only if $, US$, USD, 美元 or 美金 is in the evidence; RMB, CNY, the yuan signs, 人民币 or a number
# followed by 元 mean CNY; EUR and the euro sign mean EUR; anything else is "unknown". One place to extend.
# ORDER MATTERS: _currency_codes() blanks each marker it finds before trying the next pattern, so USD (which includes
# 美元 and 美金, both of which contain 元) is matched first and a USD quote is never also read as CNY. A bare 元 only
# counts right after a digit ("22元"), so words like 单元 (unit) are not read as a currency, and 元件 (component) is skipped.
CURRENCY_PATTERNS = (("USD", re.compile(r"US\$|USD|\$|美元|美金", re.IGNORECASE)),
                     ("CNY", re.compile(r"RMB|CNY|¥|￥|人民币|(?<=\d)\s*元(?!件)", re.IGNORECASE)),
                     ("EUR", re.compile(r"EUR|€", re.IGNORECASE)))


def parse_summary(text: str) -> dict:
    data = json.loads(text[text.index("{"): text.rindex("}") + 1])
    if data.get("status") not in REPLY_STATUSES:
        raise ValueError(f"bad status: {data.get('status')!r}")
    extras = {k: data[k] for k in QUOTE_KEYS if k in data}       # the quote fields travel on, still unverified
    return {**extras, "status": data["status"], "summary": str(data.get("summary", ""))}


QUOTE_PROMPT = (
    "A manufacturer replied to our sample-order inquiry (possibly several replies, oldest first, in any language). "
    "Reply with JSON only. Keys: status, summary, and the quote fields below.\n"
    "status: pricing_provided = gave pricing/availability; needs_info = asks us for more; dead_end = a real person "
    "declined or the product is not available; auto_reply_or_spam = out-of-office, automated message or spam with no "
    "real answer.\nsummary: 2-3 sentences, in English.\n"
    "Quote fields - use null for anything the reply does not state. NEVER guess, infer, convert or calculate:\n"
    "  sample_available (yes|no|unknown); sample_quantity (integer); sample_unit_price and sample_total_price (strings, "
    "digits only, either may be null); price_basis (per_unit|total_for_quantity|unclear); moq (integer); "
    "bulk_tiers (list of {min_qty, unit_price, evidence}); branding_possible (yes|no|unknown); branding_fee; "
    "branding_min_qty; lead_time_days (sample) and production_lead_time_days (integers); shipping_cost; "
    "shipping_terms (as stated, e.g. FOB, EXW, DDP); currency (as stated); needs_from_us (list of short strings: what "
    "the seller asks us for); quote_valid_until (as stated).\n"
    "evidence: an object mapping each non-null price, quantity, fee or lead-time field name to a SHORT VERBATIM quote "
    "copied from the reply that contains that exact number. Each bulk tier carries its own evidence. If you cannot quote "
    "it, set the field to null.\n\n")


def summarize_reply(reply: str) -> dict:
    from anthropic import Anthropic
    msg = Anthropic(api_key=cfg("ANTHROPIC_API_KEY")).messages.create(
        model=MODEL, max_tokens=1500, messages=[{"role": "user", "content": f"{QUOTE_PROMPT}Reply:\n{reply}"}])
    return parse_summary(msg.content[0].text)


def _currency_codes(text) -> set:
    """Every currency whose marker appears in the text. A marker that matched is blanked before the next pattern runs,
    so 美元 / 美金 (USD) can never also be counted as the 元 of CNY."""
    left, found = str(text), set()
    for code, pattern in CURRENCY_PATTERNS:
        if pattern.search(left):
            found.add(code)
            left = pattern.sub(" ", left)
    return found


def _currency_code(text):
    """The currency a model says it quoted in (a bare "元" is accepted here), else None."""
    said = str(text).strip()
    if said == "元":
        return "CNY"
    codes = _currency_codes(said)
    return next((code for code, _ in CURRENCY_PATTERNS if code in codes), None)


def _norm(text) -> str:
    """Case, width and whitespace-insensitive form used to find a quote inside the reply."""
    return re.sub(r"\s+", " ", unicodedata.normalize("NFKC", str(text))).strip().lower()


def _numbers(text) -> set:
    return {Decimal(n.replace(",", "")) for n in re.findall(r"\d[\d,]*(?:\.\d+)?|\.\d+", unicodedata.normalize("NFKC", str(text)))}


def _whole_number(value):
    """Positive whole number from an int, an integer-valued float or a digit string; else None."""
    if isinstance(value, bool):
        return None
    try:
        d = Decimal(str(value).replace(",", "").strip())
    except (InvalidOperation, ValueError):
        return None
    return int(d) if d.is_finite() and d == d.to_integral_value() and d >= 1 else None


def _amount(value):
    """Non-negative finite Decimal from a number or a string like "$1,200.50"; else None. Money is never a float."""
    if isinstance(value, bool) or value is None:
        return None
    try:
        d = Decimal(str(value).replace(",", "").replace("$", "").strip())
    except (InvalidOperation, ValueError):
        return None
    return d if d.is_finite() and d >= 0 else None


def _prove(label: str, number, evidence, text_norm: str, warnings: list) -> bool:
    """The evidence must appear in the reply text and the number must appear inside the evidence."""
    if not isinstance(evidence, str) or not evidence.strip():
        warnings.append(f"{label}: no evidence quote given")
    elif _norm(evidence) not in text_norm:
        warnings.append(f"{label}: evidence quote not found in the reply text")
    elif Decimal(number) not in _numbers(evidence):
        warnings.append(f"{label}: the number {number} does not appear in its evidence quote")
    else:
        return True
    return False


def verify_quote(raw: dict, reply_text: str) -> dict:
    """Checks the summarizer's quote fields against the reply and returns the verified quote plus `evidence` and
    `quote_warnings`. A numeric field whose evidence isn't in the reply, or whose number isn't in its evidence, is set
    to None with a warning naming the field and the reason."""
    warnings, evidence, out = [], {}, {k: None for k in QUOTE_KEYS}
    text_norm, given = _norm(reply_text), raw.get("evidence") if isinstance(raw.get("evidence"), dict) else {}
    used = []                                                              # evidence strings that back a price
    for field in QUOTE_INT_FIELDS:
        value = raw.get(field)
        if value is None:
            continue
        n = _whole_number(value)
        if n is None:
            warnings.append(f"{field}: {value!r} is not a positive whole number")
        elif _prove(field, n, given.get(field), text_norm, warnings):
            out[field], evidence[field] = n, given[field]
    for field in QUOTE_MONEY_FIELDS:
        value = raw.get(field)
        if value is None:
            continue
        d = _amount(value)
        if d is None:
            warnings.append(f"{field}: {value!r} is not an amount")
        elif _prove(field, d, given.get(field), text_norm, warnings):
            out[field], evidence[field] = format(d, "f"), given[field]
            used.append(given[field])
    tiers = []
    for i, tier in enumerate(raw.get("bulk_tiers") or []):
        label = f"bulk_tiers[{i}]"
        qty, price = (_whole_number(tier.get("min_qty")), _amount(tier.get("unit_price"))) if isinstance(tier, dict) else (None, None)
        if qty is None or price is None:
            warnings.append(f"{label}: needs a positive whole min_qty and an amount unit_price")
            continue
        ev = tier.get("evidence")
        if _prove(f"{label}.min_qty", qty, ev, text_norm, warnings) and _prove(f"{label}.unit_price", price, ev, text_norm, warnings):
            tiers.append({"min_qty": qty, "unit_price": format(price, "f"), "evidence": ev})
            used.append(ev)
    out["bulk_tiers"] = tiers or None
    for field in ("sample_available", "branding_possible"):
        value = str(raw.get(field) or "unknown").strip().lower()
        if value not in ("yes", "no", "unknown"):
            warnings.append(f"{field}: {value!r} is not yes/no/unknown")
            value = "unknown"
        out[field] = value
    basis = raw.get("price_basis")
    if basis is not None and basis not in PRICE_BASES:
        warnings.append(f"price_basis: {basis!r} is not per_unit/total_for_quantity/unclear")
        basis = None
    if used and basis is None:
        basis = "unclear"                                    # a price without a stated basis is not guessed at
    out["price_basis"] = basis if used else None
    terms = raw.get("shipping_terms")
    if terms:
        if _norm(terms) in text_norm:
            out["shipping_terms"] = str(terms).strip()
        else:
            warnings.append("shipping_terms: not found in the reply text")
    asks = raw.get("needs_from_us")
    out["needs_from_us"] = [str(a).strip() for a in asks if str(a).strip()] if isinstance(asks, list) else None
    out["needs_from_us"] = out["needs_from_us"] or None
    out["quote_valid_until"] = str(raw["quote_valid_until"]).strip() if raw.get("quote_valid_until") else None
    # Currency comes from the evidence, not from the model's say-so.
    if used:
        found = set().union(*(_currency_codes(u) for u in used))
        out["currency"] = next(iter(found)) if len(found) == 1 else "unknown"
        if len(found) > 1:
            warnings.append("currency: the evidence mixes currencies")
        said = raw.get("currency")
        if said and (_currency_code(said) or str(said).strip().upper()) != out["currency"]:
            warnings.append(f"currency: the model said {said!r} but the evidence shows {out['currency']}")
    out["evidence"] = evidence
    out["quote_warnings"] = warnings
    return out


def quote_lines(result: dict) -> list:
    """What was recorded from a reply's quote, for the terminal: the verified fields and any warnings."""
    shown = [f"{k}={result[k]}" for k in QUOTE_KEYS if k not in ("evidence", "quote_warnings", "bulk_tiers", "needs_from_us")
             and result.get(k) not in (None, "unknown")]
    out = ["Quote recorded: " + (", ".join(shown) if shown else "nothing quantified in this reply")]
    if result.get("bulk_tiers"):
        out.append("  bulk tiers: " + "; ".join(f">={t['min_qty']} pcs at {t['unit_price']}" for t in result["bulk_tiers"]))
    if result.get("needs_from_us"):
        out.append("  the seller asks us for: " + "; ".join(result["needs_from_us"]))
    out += [f"  WARNING (field set to null): {w}" for w in result.get("quote_warnings", [])]
    return out


def record_reply(state: dict, draft: dict, text: str, summarize=None) -> dict:
    """Files a reply under this draft's SKU + listing. A second reply is appended to the first, never replacing
    it, and the whole thread is summarized. The summarizer's quote fields are verified against the thread before
    they are stored. Nothing is stored if summarizing fails."""
    entry = state.get(draft["key"], {})
    replies = list(entry.get("replies", []))
    if not replies or replies[-1]["text"].strip() != text.strip():   # pasting the same reply twice adds nothing
        replies.append({"at": now(), "text": text})
    combined = replies[0]["text"] if len(replies) == 1 else "\n\n".join(
        f"--- reply {i} ({r['at']}) ---\n{r['text']}" for i, r in enumerate(replies, start=1))
    raw = (summarize or summarize_reply)(combined)
    if raw.get("status") not in REPLY_STATUSES:
        raise ValueError(f"bad status: {raw.get('status')!r}")
    result = {"status": raw["status"], "summary": str(raw.get("summary", "")), **verify_quote(raw, combined)}
    entry = entry_for(state, draft)
    entry.update(replies=replies, reply=combined, **result)
    return result


# ---------- report ----------

def contact_method(d: dict, entry: dict) -> str:
    if d["email"]:
        return "email sent" if "sent" in entry else "email drafted, not sent"
    return "inquiry form submitted" if "sent" in entry else "inquiry form not yet submitted"


def build_report(drafts: list, state: dict, skipped: list = ()) -> list:
    rows = []
    for d in drafts:
        s = state.get(d["key"], {})
        channel = "email" if d["email"] else "inquiry form"
        sent = f"yes ({channel}, {s['sent']})" if "sent" in s else f"no ({channel})"
        rows.append([d["sku"], d["product"], d["manufacturer"], d["email"] or d["url"], contact_method(d, s), sent,
                     "yes" if "reply" in s else "no", s.get("status", ""), s.get("summary", "")])
    for s in skipped:  # rows with no recommended seller, an error, or a reviewer exclusion: nothing was sent
        rows.append([s["sku"], s["product"], s["manufacturer"], "", "not contacted", "no", "no", "", s["reason"]])
    return rows


def orphaned(state: dict, drafts: list) -> list:
    """State records whose SKU + listing is no longer a recommended seller (the pick changed on a re-run)."""
    current = {d["key"] for d in drafts}
    return [f"{v.get('sku', k.split('|')[0])} ({v.get('seller', k)})" for k, v in state.items()
            if not k.startswith("_") and k not in current and isinstance(v, dict)]


REPORT_HEADER = ["SKU", "Product", "Manufacturer", "Contact", "Contact method", "Sent", "Reply received",
                 "Reply status", "Summary"]


def write_report(rows: list, stamp: str, notes: list = ()) -> tuple:
    os.makedirs(os.path.join(HERE, "Reports"), exist_ok=True)
    md_path = os.path.join(HERE, "Reports", f"email_report_{stamp}.md")
    xlsx_path = os.path.join(HERE, "Reports", f"email_report_{stamp}.xlsx")
    cell = lambda v: str(v).replace("|", "/").replace("\n", " ")
    contacted = sum(r[4] in ("email sent", "inquiry form submitted") for r in rows)
    lines = ["# Email outreach report", "", f"Generated {now()} | {len(rows)} product(s) | {contacted} contacted | "
             f"{sum(r[6] == 'yes' for r in rows)} replies", "", *[f"- {n}" for n in notes], *([""] if notes else []),
             "| " + " | ".join(REPORT_HEADER) + " |", "|" + "---|" * len(REPORT_HEADER)]
    lines += ["| " + " | ".join(cell(v) for v in r) + " |" for r in rows]
    with open(md_path, "w", encoding="utf-8") as f:
        f.write("\n".join(lines) + "\n")
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Email Outreach"
    ws.append(REPORT_HEADER)
    for r in rows:
        ws.append(r)
    for c in ws[1]:
        c.font = Font(bold=True)
    wb.save(xlsx_path)
    return md_path, xlsx_path


# ---------- CLI ----------

def find_draft(drafts: list, skipped: list, sku: str, results: str) -> dict:
    """The draft for `sku` in this results file, or a clear error (nothing is stored on error)."""
    for d in drafts:
        if d["sku"].upper() == sku.strip().upper():
            return d
    for s in skipped:
        if s["sku"].upper() == sku.strip().upper():
            sys.exit(f"{sku}: row found in {results} but it has no recommended seller to attach this to "
                     f"({s['reason']}). Nothing stored.")
    known = ", ".join(d["sku"] for d in drafts) or "none"
    sys.exit(f"{sku}: not a SKU in {results}. SKUs with a recommended seller: {known}. Nothing stored.")


def main():
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")  # piped output defaults to cp1252 on Windows
    if hasattr(sys.stdin, "reconfigure"):
        sys.stdin.reconfigure(encoding="utf-8")                 # so does piped/pasted input: Chinese replies broke
    load_dotenv(os.path.join(HERE, ".env"))
    p = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    sub = p.add_subparsers(dest="cmd", required=True)
    for name in ("drafts", "send", "report"):
        s = sub.add_parser(name)
        s.add_argument("results", help="sourcing_results_*.xlsx")
        if name == "drafts":
            s.add_argument("--near-miss", action="store_true",
                           help="also show drafts for the near-miss candidates listed in near_misses.csv (not recommended by the "
                                "search agent; each is labelled NEEDS HUMAN CHECK). Read only; send and report never include them.")
        if name == "send":
            s.add_argument("--yes", action="store_true", help="one confirmation for the whole batch")
            s.add_argument("--test-recipient", default="", metavar="ADDRESS",
                           help="send every email ONLY to this address, subject prefixed [TEST], nothing marked as sent")
    for name in ("mark", "reply"):
        s = sub.add_parser(name)
        s.add_argument("sku")
        s.add_argument("--results", required=True, help="sourcing_results_*.xlsx the SKU must be in")
        if name == "reply":
            s.add_argument("--file", help="UTF-8 text file with the reply (default: read stdin)")
    args = p.parse_args()

    exclusions = load_exclusions()
    results = args.results
    drafts, skipped = load_results(results, exclusions)
    drafts = [make_draft(c) for c in drafts]
    if args.cmd in ("drafts", "send", "report"):
        print_warnings(warnings_for(bool(exclusions)))
        print(skip_summary(skipped))

    if args.cmd == "drafts":  # read-only: never loads or writes email_state.json
        for d in drafts:
            print(format_draft(d))
        print(f"{len(drafts)} draft(s); {sum(bool(d['email']) for d in drafts)} with an email address. Nothing sent.")
        if args.near_miss:
            near_rows = load_near_misses()
            if near_rows is None:
                print_warnings([f"{os.path.basename(NEAR_MISSES_CSV)} is missing - no near-miss drafts were added."])
            else:
                extra, notes = near_miss_drafts(results, drafts, exclusions, near_rows)
                print(f"\n{'#' * 70}\nNEAR-MISS DRAFTS: the search agent did NOT recommend these. Each needs a human check "
                      f"before anything is sent or submitted.\n{'#' * 70}")
                for d in extra:
                    print(format_draft(d))
                for note in notes:
                    print(f"NOTE: {note}")
                print(f"{len(extra)} near-miss draft(s), {sum(bool(d['email']) for d in extra)} with an email address. Nothing sent.")
    elif args.cmd == "send":
        missing = [k for k in ("SMTP_HOST", "SMTP_USER", "SMTP_PASSWORD") if not cfg(k)]
        if missing:
            sys.exit(f"Set {', '.join(missing)} in .env first (nothing sent).")
        if args.test_recipient:
            real = {d["email"].lower() for d in drafts if d["email"]}
            if "@" not in args.test_recipient or args.test_recipient.lower() in real:
                sys.exit("--test-recipient must be a test address, not a manufacturer's.")
        try:
            sent = confirm_and_send(drafts, load_state(), args.yes, test_recipient=args.test_recipient)
            print(f"Sent {sent} email(s).")
        except SendRefused as e:
            sys.exit(f"Refusing to send: {e}")
    elif args.cmd == "mark":
        d = find_draft(drafts, skipped, args.sku, results)
        state = load_state()
        entry_for(state, d)["sent"] = now()
        save_state(state)
        print(f"{d['sku']} ({d['manufacturer'] or d['url']}): marked as submitted via the inquiry form.")
    elif args.cmd == "reply":
        d = find_draft(drafts, skipped, args.sku, results)
        text = open(args.file, encoding="utf-8").read() if args.file else sys.stdin.read()
        state = load_state()
        result = record_reply(state, d, text)
        save_state(state)
        print(f"{d['sku']} ({d['manufacturer'] or d['url']}): {result['status']}\n{result['summary']}")
        print("\n".join(quote_lines(result)))
    elif args.cmd == "report":
        state = load_state()
        lost = orphaned(state, drafts)
        notes = [skip_summary(skipped), *warnings_for(bool(exclusions))]
        if lost:
            notes.append(f"{len(lost)} earlier record(s) belong to a seller that is no longer the recommended one "
                         f"for that SKU and are not shown: {', '.join(lost)}.")
        md, xlsx = write_report(build_report(drafts, state, skipped), datetime.now().strftime("%Y%m%d_%H%M%S"), notes)
        print(f"Wrote {md}\nWrote {xlsx}")
        if lost:
            print(notes[-1])


if __name__ == "__main__":
    main()
