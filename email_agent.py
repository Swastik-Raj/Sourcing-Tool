"""Email agent: drafts (and, after confirmation, sends) sample-order outreach for the sourcing agent's
recommended manufacturers. Reads a sourcing_results_*.xlsx; never touches sourcing_agent.py.

    python email_agent.py drafts  <results.xlsx>             # show every draft, send nothing
    python email_agent.py send    <results.xlsx> [--yes]     # confirm per email (--yes = whole batch), then SMTP send
    python email_agent.py mark    <SKU>                      # record a manual inquiry-form submission
    python email_agent.py reply   <SKU> [--file reply.txt]   # paste reply (stdin, Ctrl-Z/Ctrl-D to end), summarize
    python email_agent.py report  [<results.xlsx>]           # markdown + Excel report for Neeraj

Env (.env): SMTP_HOST, SMTP_PORT (587), SMTP_USER, SMTP_PASSWORD, SMTP_FROM (default SMTP_USER),
SENDER_NAME, SHIPPING_ADDRESS, ANTHROPIC_API_KEY.
"""
import argparse
import json
import os
import smtplib
import sys
from datetime import datetime
from email.message import EmailMessage

import openpyxl
from dotenv import load_dotenv
from openpyxl.styles import Font

HERE = os.path.dirname(os.path.abspath(__file__))
TEMPLATE_PATH = os.path.join(HERE, "email_template.txt")
STATE_PATH = os.path.join(HERE, "email_state.json")  # per-SKU: sent / reply / summary
MODEL = "claude-haiku-4-5"
DEFAULT_SHIPPING = "Zync Technologies\nPlano, TX"
REPLY_STATUSES = ("pricing_provided", "needs_info", "dead_end")


def cfg(name: str, default: str = "") -> str:
    return os.environ.get(name, default)


# ---------- input: read the search agent's Results sheet ----------

def load_candidates(path: str) -> list:
    """One dict per Results row that has a recommended manufacturer or URL."""
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    rows = list(wb["Results"].iter_rows(values_only=True))
    wb.close()  # Windows keeps the file locked otherwise
    header, rows = list(rows[0]), rows[1:]
    out = []
    for values in rows:
        r = dict(zip(header, values))
        if str(r.get("Recommendation") or "").startswith("Error researching"):
            continue  # the search agent crashed on this product: nothing to send
        if not (r.get("Recommended Manufacturer") or r.get("Recommended URL")):
            continue
        url = r.get("Recommended URL") or ""
        moq = ""  # MOQ lives only in the per-candidate columns; find the recommended one by URL
        for n in range(1, 10):
            if f"Manufacturer {n} URL" not in r:
                break
            if url and r[f"Manufacturer {n} URL"] == url:
                moq = r.get(f"Manufacturer {n} MOQ") or ""
        out.append({
            "sku": str(r["SKU"]), "product": r.get("Product") or "", "description": r.get("Keyword") or "",
            "manufacturer": r.get("Recommended Manufacturer") or "", "email": r.get("Recommended Email") or "",
            "url": url, "unit_price": r.get("Recommended Unit Price"), "moq": moq,
            "note": r.get("Ordering Note") or "",
        })
    return out


# ---------- template ----------

def render_email(c: dict, template_path: str = TEMPLATE_PATH) -> tuple:
    """Returns (subject, body). Edit email_template.txt to change the wording."""
    with open(template_path, encoding="utf-8") as f:
        subject, _, body = f.read().partition("\n")
    price = f" (listed around ${c['unit_price']:.2f}/unit" if isinstance(c["unit_price"], (int, float)) else ""
    if price and c["moq"]:
        price += f", MOQ {c['moq']}"
    values = {
        "sku": c["sku"], "product": c["description"] or c["product"], "url": c["url"],
        "manufacturer": c["manufacturer"] or "Sales",
        "price_line": price + ")" if price else "",
        "note_line": f"\nOne thing to confirm: {c['note']}\n" if c["note"] else "",
        "sender_name": cfg("SENDER_NAME", "Swastik Raj"),
        "shipping_address": cfg("SHIPPING_ADDRESS", DEFAULT_SHIPPING).replace("\\n", "\n"),
    }
    return subject.removeprefix("Subject:").strip().format(**values), body.strip().format(**values) + "\n"


def make_draft(c: dict) -> dict:
    subject, body = render_email(c)
    if not c["email"]:
        body = f"[No email available - submit via the listing's inquiry form at {c['url']}]\n\n{body}"
    return {**c, "subject": subject, "body": body}


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


def confirm_and_send(drafts: list, state: dict, yes: bool = False, send=smtp_send, ask=input) -> int:
    """Sends only drafts that have an email and aren't already sent, after y/N per email (or one batch yes)."""
    todo = [d for d in drafts if d["email"] and "sent" not in state.get(d["sku"], {})]
    if yes:
        print(f"Batch send: {len(todo)} email(s).")
        if ask("Type 'yes' to send all: ").strip().lower() != "yes":
            return 0
    sent = 0
    for d in todo:
        if not yes:
            print(f"\nTo: {d['email']}\nSubject: {d['subject']}\n\n{d['body']}")
            answer = ask(f"Send {d['sku']} to {d['email']}? [y/N/q] ").strip().lower()
            if answer == "q":
                break
            if answer != "y":
                continue
        send(d)
        state.setdefault(d["sku"], {})["sent"] = now()
        save_state(state)
        sent += 1
    return sent


# ---------- reply summarization ----------

def parse_summary(text: str) -> dict:
    data = json.loads(text[text.index("{"): text.rindex("}") + 1])
    if data.get("status") not in REPLY_STATUSES:
        raise ValueError(f"bad status: {data.get('status')!r}")
    return {"status": data["status"], "summary": str(data.get("summary", ""))}


def summarize_reply(reply: str) -> dict:
    from anthropic import Anthropic
    msg = Anthropic(api_key=cfg("ANTHROPIC_API_KEY")).messages.create(
        model=MODEL, max_tokens=400,
        messages=[{"role": "user", "content": (
            "A manufacturer replied to our sample-order inquiry. Reply with JSON only: "
            '{"status": "pricing_provided" | "needs_info" | "dead_end", "summary": "<2-3 sentences: sample price/qty, '
            'branding/engraving, lead time, anything they need from us>"}.\n'
            "pricing_provided = gave pricing/availability; needs_info = asks us for more; dead_end = declined/irrelevant.\n\n"
            f"Reply:\n{reply}")}])
    return parse_summary(msg.content[0].text)


# ---------- report ----------

def build_report(drafts: list, state: dict) -> list:
    rows = []
    for d in drafts:
        s = state.get(d["sku"], {})
        channel = "email" if d["email"] else "inquiry form"
        sent = f"yes ({channel}, {s['sent']})" if "sent" in s else f"no ({channel})"
        rows.append([d["sku"], d["product"], d["manufacturer"], d["email"] or d["url"], sent,
                     "yes" if "reply" in s else "no", s.get("status", ""), s.get("summary", "")])
    return rows


REPORT_HEADER = ["SKU", "Product", "Manufacturer", "Contact", "Sent", "Reply received", "Reply status", "Summary"]


def write_report(rows: list, stamp: str) -> tuple:
    os.makedirs(os.path.join(HERE, "Reports"), exist_ok=True)
    md_path = os.path.join(HERE, "Reports", f"email_report_{stamp}.md")
    xlsx_path = os.path.join(HERE, "Reports", f"email_report_{stamp}.xlsx")
    cell = lambda v: str(v).replace("|", "/").replace("\n", " ")
    lines = ["# Email outreach report", "", f"Generated {now()} | {len(rows)} product(s) | "
             f"{sum(r[4].startswith('yes') for r in rows)} sent | {sum(r[5] == 'yes' for r in rows)} replies", "",
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

def main():
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")  # piped output defaults to cp1252 on Windows
    load_dotenv(os.path.join(HERE, ".env"))
    p = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    sub = p.add_subparsers(dest="cmd", required=True)
    for name in ("drafts", "send", "report"):
        s = sub.add_parser(name)
        s.add_argument("results", nargs=None if name != "report" else "?", help="sourcing_results_*.xlsx")
        if name == "send":
            s.add_argument("--yes", action="store_true", help="one confirmation for the whole batch")
    for name in ("mark", "reply"):
        s = sub.add_parser(name)
        s.add_argument("sku")
        if name == "reply":
            s.add_argument("--file", help="text file with the reply (default: read stdin)")
    args = p.parse_args()
    state = load_state()

    if args.cmd in ("drafts", "send", "report") and args.results:
        drafts = [make_draft(c) for c in load_candidates(args.results)]
        state["_source"] = {"results": os.path.abspath(args.results)}
    if args.cmd == "report" and not args.results:
        drafts = [make_draft(c) for c in load_candidates(state["_source"]["results"])]

    if args.cmd == "drafts":
        for d in drafts:
            to = d["email"] or "NO EMAIL - inquiry form"
            print(f"{'=' * 70}\n{d['sku']} | {d['manufacturer'] or '(unnamed)'} | To: {to}\nSubject: {d['subject']}\n\n{d['body']}")
        print(f"{len(drafts)} draft(s); {sum(bool(d['email']) for d in drafts)} with an email address. Nothing sent.")
    elif args.cmd == "send":
        missing = [k for k in ("SMTP_HOST", "SMTP_USER", "SMTP_PASSWORD") if not cfg(k)]
        if missing:
            sys.exit(f"Set {', '.join(missing)} in .env first (nothing sent).")
        print(f"Sent {confirm_and_send(drafts, state, args.yes)} email(s).")
    elif args.cmd == "mark":
        state.setdefault(args.sku, {})["sent"] = now()
        print(f"{args.sku}: marked as submitted via inquiry form.")
    elif args.cmd == "reply":
        text = open(args.file, encoding="utf-8").read() if args.file else sys.stdin.read()
        result = summarize_reply(text)
        state.setdefault(args.sku, {}).update(reply=text, **result)
        print(f"{args.sku}: {result['status']}\n{result['summary']}")
    elif args.cmd == "report":
        md, xlsx = write_report(build_report(drafts, state), datetime.now().strftime("%Y%m%d_%H%M%S"))
        print(f"Wrote {md}\nWrote {xlsx}")
    save_state(state)


if __name__ == "__main__":
    main()
