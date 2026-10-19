"""Page data and gate checks for the tabs after Search. Nothing here decides what an agent does: rows are read with the
email agent's own functions (workflows.draft_rows), everything else comes straight from the agents' own files, and every
action is a command from the allowlist that the agent checks again itself."""
import ast
import csv
import json
import re
from datetime import datetime
from pathlib import Path

from . import workflows
from .commands import FACTS_RE, REQUEST_ID, JobRejected
from .config import Settings, is_set

SMTP_SETTINGS = ("SMTP_HOST", "SMTP_USER", "SMTP_PASSWORD")
MAX_MODEL_CALLS_PER_PRODUCT = 3      # content_agent.py MAX_CALLS_PER_SKU: first try, one invalid-JSON retry, one repair


def _state(path: Path):
    """(parsed JSON object or None, problem text). A missing file is not a problem."""
    try:
        data = json.loads(path.read_text(encoding="utf-8"))
        return (data, "") if isinstance(data, dict) else (None, f"{workflows.record_name(path)} is not in the expected format.")
    except FileNotFoundError:
        return None, ""
    except (OSError, ValueError):
        return None, f"{workflows.record_name(path)} could not be read."


# ---------- Emails ----------

def emails(s: Settings, results_name: str) -> dict:
    """What the Emails tab shows for one results file. 'drafts' are the rows the email agent would draft, each with what
    email_state.json records under its key; 'to_send' is what a real send would send (the agent's rule: has an address,
    not already marked sent)."""
    rows = workflows.draft_rows(s, s.results_dir / results_name)
    state, problem = _state(s.project / "email_state.json")
    drafts = []
    for d in (rows or {}).get("drafts", []):
        entry = (state or {}).get(d["key"])
        entry = entry if isinstance(entry, dict) else {}
        drafts.append({"sku": str(d["sku"]), "seller": str(d["manufacturer"] or ""), "email": str(d["email"] or "").strip(), "url": str(d["url"] or ""),
                       "sent": entry.get("sent"), "status": entry.get("status"), "summary": entry.get("summary"), "replied": "reply" in entry})
    need = [n for n in SMTP_SETTINGS + ("SHIPPING_ADDRESS",) if not is_set(s, n)]
    return {"known": rows is not None, "drafts": drafts, "skipped": (rows or {}).get("skipped", []), "problem": problem,
            "to_send": [d for d in drafts if d["email"] and not d["sent"]], "send_missing": need,
            "settings": [(n, is_set(s, n)) for n in SMTP_SETTINGS + ("SHIPPING_ADDRESS", "ANTHROPIC_API_KEY")],
            "near_miss_file": (s.project / "near_misses.csv").is_file()}


def send_plan(s: Settings, results_name: str, test_recipient: str = "") -> dict:
    """Who a send would write to, for the confirmation page and the typed count. A test send goes ONLY to the test
    address, once per draft (the agent's rule), and marks nothing as sent."""
    view = emails(s, results_name)
    if not view["known"]:
        raise JobRejected("This search's results could not be read, so nothing can be sent from here.")
    if view["send_missing"]:
        raise JobRejected(f"Nothing was sent: email sending is not set up ({', '.join(view['send_missing'])} "
                          f"{'is' if len(view['send_missing']) == 1 else 'are'} not set).")
    rows = view["drafts"] if test_recipient else view["to_send"]
    return {"test": bool(test_recipient), "test_recipient": test_recipient, "count": len(rows),
            "recipients": [{"sku": d["sku"], "seller": d["seller"], "to": test_recipient or d["email"], "really": d["email"]} for d in rows]}


def draft_skus(s: Settings, results_name: str) -> set:
    return {d["sku"].upper() for d in emails(s, results_name)["drafts"]}


# ---------- Approvals ----------

def approvals(s: Settings, results_name: str) -> dict:
    """Requests written for this results file (with their text from Reports\\), the decisions recorded against them with
    each approval's expiry, and the export if it was made from this file. All from the decision agent's own files."""
    state, problem = _state(s.project / "decision_state.json")
    requests = []
    for r in ((state or {}).get("requests") or {}).values():
        if not isinstance(r, dict) or r.get("results_file") != results_name or not re.fullmatch(REQUEST_ID, str(r.get("id", ""))):
            continue
        text = None
        try:
            text = (s.project / "Reports" / f"approval_request_{r['id']}.txt").read_text(encoding="utf-8")
        except OSError:
            pass
        requests.append({"id": r["id"], "created": _shown(r.get("created_at")), "lines": len(r.get("lines") or []),
                         "answered": bool(r.get("reply")), "text": text})
    requests.sort(key=lambda r: r["id"])
    ids, now, decisions = {r["id"] for r in requests}, datetime.now().astimezone(), []
    for d in (state or {}).get("decisions") or []:
        if not isinstance(d, dict) or d.get("request_id") not in ids:
            continue
        expires = workflows._when(d.get("expires_at"))
        status = "Rejected" if d.get("decision") != "approve" else "Withdrawn" if d.get("revoked_at") else \
            "Expired" if expires and expires <= now else "Valid"
        decisions.append({"id": d.get("id"), "request": d.get("request_id"), "sku": d.get("sku"), "seller": d.get("seller"), "qty": d.get("qty"),
                          "max": d.get("max_unit_price"), "approver": d.get("approver"), "channel": d.get("channel"),
                          "recorded": _shown(d.get("recorded_at")), "expires": _shown(d.get("expires_at")), "status": status})
    export, export_problem = _state(s.project / "approved_orders.json")
    return {"problem": problem or export_problem, "requests": requests, "open": [r for r in requests if not r["answered"]],
            "decisions": decisions, "valid_skus": sorted({str(d["sku"]) for d in decisions if d["status"] == "Valid"}),
            "no_email_record": not (s.project / "email_state.json").exists(), "export": export_for(s, results_name, export)}


def export_for(s: Settings, results_name: str, export=None) -> dict:
    """The export on disk and whether it was made from this results file: {'exists', 'mine', 'source', 'count', 'at',
    'orders': [{sku, seller, qty, expires, expired}]}. Both names are kept so a refusal can show them."""
    if export is None:
        export, _ = _state(s.project / "approved_orders.json")
    if not isinstance(export, dict):
        return {"exists": False, "mine": False, "source": "", "count": 0, "at": "", "orders": []}
    now, orders = datetime.now().astimezone(), []
    for o in export.get("orders") or []:
        if isinstance(o, dict):
            expires = workflows._when(o.get("expires_at"))
            orders.append({"sku": o.get("sku"), "seller": o.get("seller"), "qty": o.get("qty"), "expires": _shown(o.get("expires_at")),
                           "expired": bool(expires and expires <= now)})
    source = str(export.get("source_results_file") or "")
    return {"exists": True, "mine": source == results_name, "source": source, "count": export.get("valid_approvals", len(orders)),
            "at": _shown(export.get("generated_at")), "orders": orders}


def export_gate(s: Settings, results_name: str) -> dict:
    """The export an order sheet or listings would be built from, or a refusal. approved_orders.csv is one file: if it
    was exported from another results file, building from it here would mix two searches, so it is refused with both
    names. Whether each approval is still valid stays the agent's call: expiry is shown, not enforced here."""
    export = export_for(s, results_name)
    if not export["exists"] or not (s.project / "approved_orders.csv").is_file():
        raise JobRejected("There is no approved list yet. Save it in step 3, Get approval, first. Nothing was started.")
    if not export["mine"]:
        raise JobRejected(f"The saved approved list belongs to a different search: it was saved from '{export['source'] or 'an unnamed file'}', "
                          f"not from the one you are working on, '{results_name}'. Save the approved list again in step 3, Get approval. "
                          "Nothing was started.")
    return export


# ---------- Order sheet ----------

def outputs(s: Settings, folder: str, pattern: str, limit: int = 5) -> list:
    """Newest files first: [{'name', 'at'}]."""
    root = s.project / folder
    found = sorted(root.glob(pattern), key=lambda p: p.stat().st_mtime, reverse=True) if root.is_dir() else []
    return [{"name": p.name, "at": f"{datetime.fromtimestamp(p.stat().st_mtime):%Y-%m-%d %H:%M}", "mtime": p.stat().st_mtime} for p in found[:limit]]


def order(s: Settings, results_name: str) -> dict:
    return {"export": export_for(s, results_name), "sheets": outputs(s, "Order Sheets", "order_sheet_*.xlsx")}


# ---------- Listing content ----------

_columns: dict = {}


def facts_columns(s: Settings) -> list:
    """The facts file's columns, read from the content agent's own FACT_COLUMNS by parsing its source (nothing is imported
    or run). [] if they cannot be read: uploading is then refused rather than guessed."""
    path = s.project / "content_agent.py"
    try:
        key = (str(path), path.stat().st_mtime)
        if key not in _columns:
            tree = ast.parse(path.read_text(encoding="utf-8"))
            node = next(n for n in tree.body if isinstance(n, ast.Assign) and any(isinstance(t, ast.Name) and t.id == "FACT_COLUMNS" for t in n.targets))
            _columns[key] = [str(c) for c in ast.literal_eval(node.value)]
        return list(_columns[key])
    except Exception:  # noqa: BLE001
        return []


def facts_files(s: Settings) -> list:
    """Names of the filled facts files a person uploaded, newest first."""
    folder = s.uploads_dir
    names = [p.name for p in folder.iterdir() if FACTS_RE.fullmatch(p.name) and p.is_file() and not p.is_symlink()] if folder.is_dir() else []
    return sorted(names, reverse=True)


def facts_summary(s: Settings, name: str) -> dict:
    """Display only: how many filled-in facts the file holds per status, per product. The file is opened for reading and
    nothing is ever written back: a status is only what the person who filled the file wrote."""
    by_status, skus = {}, set()
    try:
        with open(s.uploads_dir / name, encoding="utf-8-sig", newline="") as f:
            for row in csv.DictReader(f):
                row = {(k or "").strip(): (v or "").strip() for k, v in row.items()}
                if row.get("value"):
                    by_status[row.get("status") or "(no status)"] = by_status.get(row.get("status") or "(no status)", 0) + 1
                    skus.add(row.get("sku", "").upper())
    except (OSError, UnicodeDecodeError, csv.Error):
        return {"name": name, "readable": False, "by_status": [], "facts": 0, "products": 0}
    return {"name": name, "readable": True, "by_status": sorted(by_status.items()), "facts": sum(by_status.values()), "products": len(skus - {""})}


def content(s: Settings, results_name: str, facts: str = "") -> dict:
    """What the Listing content tab shows. 'count' is the number of products in the export: the number a person types
    before listings are generated."""
    export, files = export_for(s, results_name), facts_files(s)
    chosen = facts if facts in files else (files[0] if files else "")
    skus = sorted({str(o["sku"]) for o in export["orders"] if o.get("sku")}) if export["mine"] else []
    return {"export": export, "facts_files": files, "facts": chosen, "summary": facts_summary(s, chosen) if chosen else None,
            "templates": outputs(s, "Uploads", "facts_template_*.csv", 3), "listings": outputs(s, "Walmart Listings", "walmart_listings_*.*", 6),
            "skus": skus, "count": len(skus), "calls_each": MAX_MODEL_CALLS_PER_PRODUCT, "columns": facts_columns(s),
            "missing": [n for n in ("ANTHROPIC_API_KEY", "BRAND_NAME") if not is_set(s, n)]}


def _shown(when) -> str:
    """An agent's ISO time for a page, in local time; the raw text if it is not a time."""
    t = workflows._when(when)
    return f"{t.astimezone():%Y-%m-%d %H:%M}" if t else str(when or "")


def zone_label(t: datetime = None) -> str:
    """The time zone the pages show times in: its offset from UTC, then the system's own name for it when it has one.
    For example 'UTC+05:30 (IST)'. `t` is an aware time; the default is now, in this computer's (or container's) zone."""
    t = t or datetime.now().astimezone()
    z = t.strftime("%z") or "+0000"
    label, name = f"UTC{z[:3]}:{z[3:5]}", t.tzname() or ""
    return label if not name or name.startswith("UTC") else f"{label} ({name})"
