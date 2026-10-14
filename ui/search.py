"""Step 1, Search: which products can be chosen, the estimate, the allowlisted command and its argument rules, and a
display-only reading of the search agent's output lines. Nothing here calls a network or runs the agent."""
import ast
import os
import re
import shutil
from datetime import datetime
from pathlib import Path
from typing import Callable, Optional

from .commands import CANCEL_WARNINGS, GENERIC_EXITS, SKU_RE, Command, JobRejected, Param
from .config import MAX_PRODUCT_ROWS, UPLOADS_FOLDER, Settings

AGENT = "sourcing_agent.py"
REQUIRED_KEYS = (("ANTHROPIC_API_KEY", "The search cannot ask the AI service to read supplier pages"),
                 ("NIMBLE_API_KEY", "The search cannot look up supplier pages"))
PRESET_SKUS = ("HDFF", "FOA-020C", "LCSP1050")

# Rough figures from one 15-product batch: about $1.20 and 14 minutes, so about $0.08 and 56 seconds per product.
# They are an estimate for the confirmation text, not a measurement and not a limit.
EST_COST_PER_PRODUCT_USD = 0.08
EST_SECONDS_PER_PRODUCT = 56

MAX_LIMIT = MAX_PRODUCT_ROWS    # the largest --limit the UI will pass, and the most products one file search may hold
UPLOAD_RE = re.compile(r"input_\d{8}_\d{6}_[0-9a-f]{6}\.xlsx")     # the only names an uploaded sheet is ever saved under
UPLOAD_PREFIX = UPLOADS_FOLDER + "/"
SCAN_ROWS = 10000               # rows of a sheet that are read at most
OUTPUT_RE = re.compile(r"sourcing_results_\d{8}_\d{6}\.xlsx")
RESULTS_FOLDER = "Excel Output Sheets"


# ---------- what can be chosen ----------

_builtin_cache: dict = {}


def _catalog_is_still_a_stub(tree: ast.Module) -> bool:
    """True while the agent's read_lcom_catalog() does nothing but `return []` (or does not exist)."""
    fn = next((n for n in tree.body if isinstance(n, ast.FunctionDef) and n.name == "read_lcom_catalog"), None)
    if fn is None:
        return True
    body = [n for n in fn.body if not (isinstance(n, ast.Expr) and isinstance(n.value, ast.Constant) and isinstance(n.value.value, str))]
    return len(body) == 1 and isinstance(body[0], ast.Return) and isinstance(body[0].value, ast.List) and not body[0].value.elts


def builtin_products(s: Settings) -> list:
    """[(sku, description)] of the list the agent's --sku looks up when there is no input file. The agent uses
    `read_lcom_catalog() or PRODUCTS`, and that catalog reader returns [] today, so this is PRODUCTS, read by parsing
    the agent's source: nothing is imported or run. If the catalog reader ever does real work (it may fetch pages), this
    returns [] instead of guessing, and the page then offers no built-in products until this function is taught the
    new source."""
    path = s.project / AGENT
    try:
        key = (str(path), path.stat().st_mtime)
        if key not in _builtin_cache:
            tree = ast.parse(path.read_text(encoding="utf-8"))
            if not _catalog_is_still_a_stub(tree):
                _builtin_cache[key] = []
            else:
                node = next(n for n in tree.body if isinstance(n, ast.Assign)
                            and any(isinstance(t, ast.Name) and t.id == "PRODUCTS" for t in n.targets))
                _builtin_cache[key] = [(str(p["sku"]).strip(), str(p.get("description") or "").strip())
                                       for p in ast.literal_eval(node.value)]
        return list(_builtin_cache[key])
    except Exception:  # noqa: BLE001 - no readable list means nothing can be chosen
        return []


def input_files(s: Settings) -> list:
    """Names of the .xlsx files in the project folder itself: the only input files that can be picked."""
    if not s.project.is_dir():
        return []
    return sorted(p.name for p in s.project.glob("*.xlsx") if p.is_file() and not p.name.startswith("~$"))


def upload_files(s: Settings) -> list:
    """Uploaded sheets, newest first, as the values the forms use ('Uploads/<server-built name>')."""
    folder = s.uploads_dir
    if not folder.is_dir():
        return []
    names = sorted((p.name for p in folder.iterdir() if UPLOAD_RE.fullmatch(p.name) and p.is_file() and not p.is_symlink()), reverse=True)
    return [UPLOAD_PREFIX + n for n in names]


def input_path(s: Settings, value: str) -> Path:
    """The path (relative to the project folder) of an input file picked from the two listings. A value is only ever
    matched against those listings: it is never used as a path a person typed."""
    if isinstance(value, str) and value in input_files(s):
        return Path(value)
    if isinstance(value, str) and value in upload_files(s):
        return Path(UPLOADS_FOLDER) / value[len(UPLOAD_PREFIX):]
    raise JobRejected("That input file is not in the project folder or among the uploaded sheets. Choose one from the list.")


def products_from_rows(rows) -> tuple:
    """([(sku, description)], {optional column: found}) the way the agent's read_products reads a sheet: a header row
    with a cell reading 'SKU' in the first 5 rows, then every row with a SKU; the description is Keyword, else Product."""
    for i, row in enumerate(rows[:5]):
        head = {str(c).strip().lower(): n for n, c in enumerate(row) if c}
        if "sku" in head:
            col = lambda r, key: r[head[key]] if key in head and head[key] < len(r) else None
            products = [(str(col(r, "sku")).strip(), str(col(r, "keyword") or col(r, "product") or "").strip())
                        for r in rows[i + 1:] if col(r, "sku")]
            return products, {"keyword": "keyword" in head, "product": "product" in head, "price": "lcom sale price" in head}
    raise JobRejected("That file has no 'SKU' column heading in its first 5 rows.")


def sheet_rows(source) -> list:
    """The first sheet's rows (values only) of a workbook path or file object, read-only, at most SCAN_ROWS of them."""
    import openpyxl
    try:
        wb = openpyxl.load_workbook(source, read_only=True, data_only=True)
    except Exception:  # noqa: BLE001
        raise JobRejected("That file could not be opened as an Excel workbook. It may be open in Excel or damaged.")
    try:
        rows = []
        for row in wb.worksheets[0].iter_rows(values_only=True):
            rows.append(row)
            if len(rows) > SCAN_ROWS:
                raise JobRejected(f"That sheet has more than {SCAN_ROWS} rows.")
        return rows
    finally:
        wb.close()


def input_products(s: Settings, name: str) -> list:
    """[(sku, description)] of an input file (see products_from_rows). Read-only."""
    return products_from_rows(sheet_rows(s.project / input_path(s, name)))[0]


def resolve_selection(s: Settings, choice: str, skus=(), input_file: str = "", limit: str = "", add_builtin: str = "") -> dict:
    """The validated selection: {'params': arguments for the search command, 'products': [(sku, description)],
    'label': text}. The product count is always len(products), computed here on the server."""
    builtin = builtin_products(s)
    by_upper = {sku.upper(): (sku, desc) for sku, desc in builtin}
    if choice == "preset":
        missing = [k for k in PRESET_SKUS if k not in by_upper]
        if missing:
            raise JobRejected(f"The 3-product test cannot run: {', '.join(missing)} is not in the built-in product list. "
                              "Choose products another way.")
        picked = [by_upper[k] for k in PRESET_SKUS]
        return {"params": {"sku": [p[0] for p in picked]}, "products": picked, "label": "3-product test"}
    if choice == "all":
        if not builtin:
            raise JobRejected("The built-in product list could not be read.")
        return {"params": {}, "products": builtin, "label": "All built-in products"}
    if choice == "choose":
        picked = []
        for v in skus or ():
            if not isinstance(v, str) or not SKU_RE.fullmatch(v) or v.upper() not in by_upper:
                raise JobRejected(f"'{str(v)[:40]}' is not in the built-in product list.")
            if by_upper[v.upper()] not in picked:
                picked.append(by_upper[v.upper()])
        if not picked:
            raise JobRejected("No products are ticked. Tick at least one product.")
        return {"params": {"sku": [p[0] for p in picked]}, "products": picked, "label": "Chosen products"}
    if choice == "file":
        rows = input_products(s, input_file)
        params = {"input": input_file}
        shown = "Uploaded sheet " + input_file[len(UPLOAD_PREFIX):] if input_file.startswith(UPLOAD_PREFIX) else f"Input file {input_file}"
        if skus:                                      # the agent's own --input FILE --sku A B: only SKUs that are in this file
            if str(limit or "").strip() or add_builtin:
                raise JobRejected("Chosen products cannot be combined with a row limit or the built-in products option.")
            in_file, picked = {}, []
            for row in rows:
                in_file.setdefault(row[0].upper(), row)                # the agent takes the first match too
            for v in skus:
                if not isinstance(v, str) or not SKU_RE.fullmatch(v) or v.upper() not in in_file:
                    raise JobRejected(f"'{str(v)[:40]}' is not a product in that file.")
                if in_file[v.upper()] not in picked:
                    picked.append(in_file[v.upper()])
            return {"params": {"input": input_file, "sku": [p[0] for p in picked]}, "products": picked,
                    "label": f"{shown}, chosen products"}
        if str(limit or "").strip():
            params["limit"] = limit_rule(str(limit).strip())
            rows = rows[:int(params["limit"])]
        if not rows:
            raise JobRejected("That input file has no products under its 'SKU' heading.")
        if len(rows) > MAX_PRODUCT_ROWS:
            raise JobRejected(f"That file has {len(rows)} products. One search takes at most {MAX_PRODUCT_ROWS}: "
                              "use the first-rows box or choose products.")
        label = shown + (f", first {params['limit']} rows" if "limit" in params else "")
        if add_builtin:                               # as the agent does it: appended after the limit, nothing removed
            params["add_builtin"] = switch_rule(add_builtin)
            if not builtin:
                raise JobRejected("The built-in product list could not be read, so it cannot be included.")
            rows = rows + builtin
            label += f", plus the {len(builtin)} built-in product{'s' if len(builtin) != 1 else ''}"
        return {"params": params, "products": rows, "label": label}
    raise JobRejected("Choose which products to search.")


def estimate(count: int) -> dict:
    return {"count": count, "cost": f"{count * EST_COST_PER_PRODUCT_USD:.2f}",
            "minutes": max(1, round(count * EST_SECONDS_PER_PRODUCT / 60)), "seconds_each": EST_SECONDS_PER_PRODUCT,
            "cost_each": f"{EST_COST_PER_PRODUCT_USD:.2f}"}


# ---------- the allowlisted command ----------

def sku_shape_rule(v: str) -> str:
    """Shape only. Which list the SKU must be in depends on the other arguments, so _combination checks that."""
    if not isinstance(v, str) or not SKU_RE.fullmatch(v):
        raise JobRejected(f"'{str(v)[:40]}' is not a valid SKU (letters, digits, . _ & - only, up to 40 characters).")
    return v


def input_file_rule(s: Settings) -> Callable[[str], str]:
    """An input file is picked by exact value from the project folder listing or the uploaded sheets; it is never
    accepted as a typed path. The argument passed on is that file's path inside the project folder."""
    def check(v: str) -> str:
        return str(input_path(s, v))
    return check


def limit_rule(v: str) -> str:
    if not isinstance(v, str) or not re.fullmatch(r"[0-9]{1,4}", v) or not 1 <= int(v) <= MAX_LIMIT:
        raise JobRejected(f"The row limit must be a whole number from 1 to {MAX_LIMIT}.")
    return str(int(v))


def switch_rule(v: str) -> str:
    if v != "1":
        raise JobRejected("That option can only be switched on or left off.")
    return v


def output_rule(s: Settings) -> Callable[[str], str]:
    """Only a server-built name is accepted, only inside the results folder, and never over an existing file."""
    def check(v: str) -> str:
        if not isinstance(v, str) or not OUTPUT_RE.fullmatch(v):
            raise JobRejected("The results file name is not one this tool builds.")
        if (s.results_dir / v).exists():
            raise JobRejected("A results file with that name already exists. Wait a second and start again.")
        return str(Path(RESULTS_FOLDER) / v)
    return check


def new_output_name(now: Optional[datetime] = None) -> str:
    return f"sourcing_results_{now or datetime.now():%Y%m%d_%H%M%S}.xlsx"


def _combination(s: Settings) -> Callable[[dict], dict]:
    def check(given: dict) -> dict:
        if "sku" in given and any(k in given for k in ("limit", "add_builtin")):
            raise JobRejected("Chosen products cannot be combined with a row limit or the built-in products option.")
        if "input" not in given and any(k in given for k in ("limit", "add_builtin")):
            raise JobRejected("A row limit and the built-in products option are only used with an input file.")
        if "sku" in given:                      # every SKU must be in the list the agent will look it up in, spelled as it is there
            skus = given["sku"] if isinstance(given["sku"], (list, tuple)) else []
            rows = input_products(s, given["input"]) if "input" in given else builtin_products(s)
            known, where = {}, "that file" if "input" in given else "the built-in product list"
            for sku, _ in rows:
                known.setdefault(sku.upper(), sku)
            for v in skus:
                if not isinstance(v, str) or v.upper() not in known:
                    raise JobRejected(f"'{str(v)[:40]}' is not in {where}.")
            given = {**given, "sku": [known[v.upper()] for v in skus]}
        return given
    return check


def search_command(s: Settings) -> Command:
    """`python sourcing_agent.py [--input FILE [--limit N] [--add-builtin]] [--sku A B ...] --output "Excel Output Sheets\\<built name>"`.
    --inspect-tools is not a parameter, so it can never be passed. Tests swap the script for the stub."""
    script, fixed = AGENT, ()
    if s.enable_stub:
        script, fixed = str(Path(__file__).with_name("stub_agent.py")), ("search",)
    params = {"sku": Param("--sku", sku_shape_rule, False, many=True),
              "input": Param("--input", input_file_rule(s), False),
              "limit": Param("--limit", limit_rule, False),
              "add_builtin": Param("--add-builtin", switch_rule, False, switch=True),
              "output": Param("--output", output_rule(s))}
    exits = {**GENERIC_EXITS, 0: "The search agent finished. Its results are unverified until a person checks them."}
    return Command("search", "Search for suppliers", script, fixed, params, True,
                   CANCEL_WARNINGS["sourcing_agent search"], exits, _combination(s))


# ---------- reading the agent's output (display only) ----------

# The formats are the agent's own print lines (sourcing_agent.py: log(), run_batch, run_search).
_SKU = r"\[(.{1,60}?)\]"
_START = re.compile(rf"^{_SKU} Product (\d{{1,5}}) of (\d{{1,5}}) \.\.\.$")
_TOKENS = re.compile(rf"^{_SKU}   Tokens:     in=\d{{1,12}} out=\d{{1,12}}  \(~\$(\d{{1,6}}\.\d{{1,6}})\)$")
_ERROR = re.compile(rf"^{_SKU}   Error researching this product: ")
_HEADER = re.compile(r"^Model: .{1,80}?  \|  Products: (\d{1,5})  \|  Sites: ")
_TOTAL = re.compile(r"^Total tokens: in=\d{1,15} out=\d{1,15}  \|  Estimated total cost: ~\$(\d{1,7}\.\d{1,6})$")
_MISSING = re.compile(r"^Missing required environment variable\(s\): ([A-Z_]{1,40}(?:, [A-Z_]{1,40}){0,5})$")
_REPORT = re.compile(r"(sourcing_report_\d{8}_\d{6}\.md)")
_BILLING = "!!! BILLING / CREDIT / AUTH ERROR - stopping the batch"


def parse_progress(log_text: str) -> dict:
    """What the log says so far. Display only: a line that is not recognised is ignored, and nothing here decides
    what the UI does next."""
    out = {"total": None, "started": 0, "done": 0, "errors": 0, "current": [], "cost": 0.0, "final_cost": None,
           "billing": False, "missing": [], "report": None}
    started, finished = [], []
    for line in (log_text or "").splitlines():
        try:
            if m := _START.match(line):
                started.append(m[1])
                out["started"] = max(out["started"], int(m[2]))
                out["total"] = int(m[3])
            elif m := _TOKENS.match(line):
                finished.append(m[1])
                out["cost"] += float(m[2])
            elif m := _ERROR.match(line):
                finished.append(m[1])
                out["errors"] += 1
            elif m := _HEADER.match(line):
                out["total"] = int(m[1])
            elif m := _TOTAL.match(line):
                out["final_cost"] = float(m[1])
            elif m := _MISSING.match(line):
                out["missing"] = m[1].split(", ")
            elif _BILLING in line:
                out["billing"] = True
            elif line.startswith(("Saving results after every product to ", "Report written to: ")) and (m := _REPORT.search(line)):
                out["report"] = m[1]
        except (ValueError, OverflowError):
            continue
    out["done"] = len(finished)
    if out["total"] is not None:
        out["done"] = min(out["done"], out["total"])
        out["started"] = min(out["started"], out["total"])
    left = list(finished)
    for sku in started:                       # still being worked on = started and not yet finished
        if sku in left:
            left.remove(sku)
        else:
            out["current"].append(sku)
    out["cost"] = round(out["cost"], 4)
    return out


_REPORT_LINE = re.compile(r"^Report written to: (.+)$", re.M)
REPORT_NAME = re.compile(r"sourcing_report_\d{8}_\d{6}\.md")


def copy_report(s: Settings, log_text: str) -> dict:
    """Copy the report the agent says it wrote into Reports\\ so the page can link it: {'name', 'copied', 'problem'}.
    The path comes from the agent's own 'Report written to:' line and must be an existing regular file, directly inside
    the project folder, named sourcing_report_<date>_<time>.md. The original is only read: never moved, never changed;
    a copy that is already there is left as it is."""
    found = _REPORT_LINE.findall(log_text or "")
    if not found:
        return {"name": None, "copied": False, "problem": "The agent's log has no 'Report written to' line, so no report was copied."}
    bad = {"name": None, "copied": False, "problem": "The report path in the agent's log did not pass the checks, so no report was copied."}
    try:
        src = Path(found[-1].strip())
        name = src.name
        if not REPORT_NAME.fullmatch(name):
            return bad
        bad["name"] = name
        real, project = src.resolve(strict=True), s.project.resolve()
        if not src.is_absolute() or src.is_symlink() or not real.is_file() or real.parent != project or real.name != name:
            return bad
        folder = project / "Reports"
        folder.mkdir(exist_ok=True)
        if os.path.normcase(folder.resolve()) != os.path.normcase(folder):   # Reports must be a real folder of this project, not a link elsewhere
            return bad
        dest = folder / name
        if not dest.exists():
            shutil.copyfile(real, dest)
        return {"name": name, "copied": dest.is_file(), "problem": ""}
    except (OSError, ValueError):
        return bad


def reading(status: str, exit_code, progress: dict) -> list:
    """Plain-language sentences for a run that did not simply finish, from what the agent printed."""
    lines = []
    if progress["missing"]:
        lines.append(f"The search agent stopped before searching because a required setting is missing: {', '.join(progress['missing'])}. "
                     "Add it to the .env file in the project folder and start again. Nothing was spent.")
    if progress["billing"]:
        lines.append("The search agent reported a billing, credit or sign-in problem with the AI service and stopped the batch. "
                     "Products already finished are saved; the others were not searched and are listed as errors.")
    if status == "failed" and not lines:
        lines.append("The search agent stopped with an error. The last lines of its log are below.")
    if status == "interrupted":
        lines.append("The UI stopped while this search was running, so it is not known how far the search got.")
    return lines
