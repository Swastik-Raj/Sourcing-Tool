"""Workflow overview, read-only. A workflow is one results file in 'Excel Output Sheets'. URLs use the plain file stem; the
internal id (observability.workflow_id, the agents' own function) is only mapped back on the server. Step statuses only say
what the files show."""
import json
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path

import observability as obs

from .config import Settings

STEPS = ("Search", "Emails", "Approvals", "Order sheet", "Listing content")
STEP_TEXT = {
    "Search": "Finds alternative suppliers for each product and compares their prices with the reference price.",
    "Emails": "Writes sample-order emails to the suppliers the search recommended, for a person to check and send.",
    "Approvals": "Turns the recommendations into a request for a manager and records the manager's reply.",
    "Order sheet": "Builds the Excel sheet a person uses to place the approved orders by hand.",
    "Listing content": "Drafts the product listing text for a person to review and upload by hand.",
}
# Badge vocabulary. "Files found" means files exist, nothing more: it does not mean they are correct or approved.
NOT_STARTED, FOUND, BLOCKED = "Not started", "Files found", "Blocked"


@dataclass
class Step:
    name: str
    status: str
    detail: str = ""
    blocked: str = ""

    @property
    def slug(self) -> str:
        return self.name.lower().replace(" ", "-")


def internal_id(stem: str) -> str:
    """The agents' workflow id for a results file with this stem."""
    return obs.workflow_id(stem + ".xlsx")


_cache: dict = {}      # (path, mtime) -> value, so the pages stay fast


def _cached(key, make):
    if key not in _cache:
        _cache[key] = make()
    return _cache[key]


def _count_rows(path: Path) -> tuple:
    """(products, rows carrying a recommended URL), read-only."""
    import openpyxl
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    try:
        ws = wb["Results"] if "Results" in wb.sheetnames else wb[wb.sheetnames[0]]
        rows = ws.iter_rows(values_only=True)
        head = [str(c or "") for c in next(rows)]
        sku, url = head.index("SKU"), head.index("Recommended URL")
        n = rec = 0
        for r in rows:
            if r[sku]:
                n += 1
                rec += bool(r[url])
        return n, rec
    finally:
        wb.close()


def products_in(path: Path):
    try:
        return _cached(("n", str(path), path.stat().st_mtime), lambda: _count_rows(path)[0])
    except Exception:  # noqa: BLE001
        return None


def friendly_label(modified: datetime, products) -> str:
    hour = modified.hour % 12 or 12
    when = f"{modified:%b} {modified.day}, {hour}:{modified:%M} {'AM' if modified.hour < 12 else 'PM'}"
    return when + (f", {products} product{'s' if products != 1 else ''}" if products is not None else "")


def list_workflows(s: Settings) -> list:
    """Newest first: [{stem, id (internal), file, modified, label}]."""
    if not s.results_dir.is_dir():
        return []
    files = sorted((p for p in s.results_dir.glob("*.xlsx") if not p.name.startswith("~$")),
                   key=lambda p: p.stat().st_mtime, reverse=True)
    out, seen = [], {}
    for p in files:
        when = datetime.fromtimestamp(p.stat().st_mtime)
        label = friendly_label(when, products_in(p))
        seen[label] = seen.get(label, 0) + 1
        out.append({"stem": p.stem, "id": internal_id(p.stem), "file": p.name, "modified": when.strftime("%Y-%m-%d %H:%M"),
                    "label": label if seen[label] == 1 else f"{label} ({seen[label]})"})
    return out


def _json(path: Path):
    try:
        return json.loads(path.read_text(encoding="utf-8")), ""
    except FileNotFoundError:
        return None, ""
    except (OSError, ValueError):
        return None, f"{path.name} could not be read"


def draft_counts(s: Settings, path: Path):
    """What the Emails step can draft for this file, from the email agent's OWN row rules (its load_results, read-only), not
    a copy of them. None if that cannot be run: the page then shows only what the file says."""
    def make():
        try:
            import email_agent as ea
            excl = ea.load_exclusions(str(s.project / "reviewer_exclusions.csv"))
            drafts, skipped = ea.load_results(str(path), excl)
            kinds = {k: sum(x["kind"] == k for x in skipped) for k in ("none", "reviewer", "error")}
            return {"draftable": len(drafts), "no_seller": kinds["none"], "excluded": kinds["reviewer"], "errors": kinds["error"]}
        except Exception:  # noqa: BLE001
            return None
    return _cached(("d", str(path), path.stat().st_mtime, _mtime(s.project / "reviewer_exclusions.csv")), make)


def _mtime(p: Path):
    try:
        return p.stat().st_mtime
    except OSError:
        return 0


def search_detail(s: Settings, path: Path) -> tuple:
    """(products, text). Counts come from the file and, for what can be drafted, from the email agent's own rules."""
    n, rec = _cached(("c", str(path), path.stat().st_mtime), lambda: _count_rows(path))
    d = draft_counts(s, path)
    plural = f"{n} product{'s' if n != 1 else ''}"
    if d is None:
        return n, (f"{plural} in the results file. {rec} row{'s' if rec != 1 else ''} carry a recommendation; the Emails step shows how "
                   "many can be drafted. Recommendations are the search's reading of seller pages and are unverified until a person checks them.")
    parts = [f"{d['draftable']} can be drafted"]
    if d["no_seller"]:
        parts.append(f"{d['no_seller']} have no recommended seller")
    if d["excluded"]:
        parts.append(f"{d['excluded']} are excluded by the reviewer's exclusion list")
    if d["errors"]:
        parts.append(f"{d['errors']} could not be searched")
    return n, (f"{plural} in the results file: " + "; ".join(parts) + ". Recommendations are unverified until a person checks them.")


def steps_for(s: Settings, results_name: str) -> list:
    path = s.results_dir / results_name
    out = []
    try:
        n, text = search_detail(s, path)
        d = draft_counts(s, path)
        search = Step("Search", FOUND, text)
        can_draft = d["draftable"] if d else None
    except Exception:  # noqa: BLE001
        n = 0
        can_draft = None
        search = Step("Search", BLOCKED, "", "The results file could not be read. It may be open in Excel or damaged.")
    out.append(search)
    no_results = "There is no readable results file for this workflow." if search.status == BLOCKED else ""

    email, err = _json(s.project / "email_state.json")
    entries = [v for k, v in (email or {}).items() if not k.startswith("_") and isinstance(v, dict)]
    sent, replies = sum("sent" in e for e in entries), sum("reply" in e for e in entries)
    if err:
        emails = Step("Emails", BLOCKED, "", err)
    elif not entries:
        why = no_results or ("Nothing in this file can be drafted." if can_draft == 0 else "")
        emails = Step("Emails", BLOCKED if why else NOT_STARTED, "No emails recorded yet.", why)
    else:
        emails = Step("Emails", FOUND, f"The email record has {len(entries)} seller(s): {sent} marked sent, {replies} with a reply. "
                      "That record covers every results file, not only this one.")
    out.append(emails)

    dec, err = _json(s.project / "decision_state.json")
    reqs = [r for r in (dec or {}).get("requests", {}).values() if r.get("results_file") == results_name]
    if err:
        approvals = Step("Approvals", BLOCKED, "", err)
    elif not reqs:
        approvals = Step("Approvals", BLOCKED if no_results else NOT_STARTED, "No approval request has been written for this file.", no_results)
    else:
        ids = {r["id"] for r in reqs}
        decs = [d for d in dec.get("decisions", []) if d.get("request_id") in ids]
        now = datetime.now().isoformat(timespec="seconds")
        live = sum(d["decision"] == "approve" and not d.get("revoked_at") and str(d.get("expires_at", "")) > now for d in decs)
        waiting = sum(not r.get("reply") for r in reqs)
        approvals = Step("Approvals", FOUND, f"{len(reqs)} request(s); {waiting} still waiting for a recorded reply; "
                         f"{live} approval(s) currently valid (approvals expire).")
    out.append(approvals)

    exp, err = _json(s.project / "approved_orders.json")
    mine = exp if exp and exp.get("source_results_file") == results_name else None
    sheet = _newest(s.project / "Order Sheets", "order_sheet_*.xlsx")
    if err:
        order = Step("Order sheet", BLOCKED, "", err)
    elif mine is None:
        why = "Needs a valid approval, exported from the Approvals step." if approvals.status != FOUND else ""
        order = Step("Order sheet", BLOCKED if why else NOT_STARTED, "No approved orders have been exported for this file.", why)
    else:
        extra = f" The newest order sheet is {sheet.name} (sheets are not linked to a results file)." if sheet else " No order sheet file yet."
        order = Step("Order sheet", FOUND, f"{mine.get('valid_approvals', 0)} valid approval(s) exported at {mine.get('generated_at', '?')}.{extra}")
    out.append(order)

    facts = (s.project / "product_facts.csv").exists()
    listing = _newest(s.project / "Walmart Listings", "walmart_listings_*.xlsx")
    if listing:
        content = Step("Listing content", FOUND, f"Newest listing file: {listing.name} (not linked to a results file). Drafts need a person's review before upload.")
    else:
        why = "Needs an exported approval first." if mine is None else ""
        content = Step("Listing content", BLOCKED if why else NOT_STARTED,
                       "The product facts file exists." if facts else "The product facts file has not been created yet.", why)
    out.append(content)
    return out


def _newest(folder: Path, pattern: str):
    found = sorted(folder.glob(pattern), key=lambda p: p.stat().st_mtime, reverse=True) if folder.is_dir() else []
    return found[0] if found else None
