"""Host-side helper for docker\\verify.ps1. It is not part of the image. Each sub-command prints one JSON object.

    python docker\\verify_helper.py scan    <image.tar> runtime|test <needle> [<needle> ...]
    python docker\\verify_helper.py envhash <env file> <NAME> [<NAME> ...]
    python docker\\verify_helper.py envvalues <env file of made-up values> <NAME> [<NAME> ...]
    python docker\\verify_helper.py pagescan <port> <results stem> <needle> [<needle> ...]
    python docker\\verify_helper.py plantjob <port> <data folder> <fake user> <fake password> <fake proxy url>
    python docker\\verify_helper.py fixture <data folder>
    python docker\\verify_helper.py hosts   <port> <container name>
    python docker\\verify_helper.py flow    <port> <data folder> <results stem>
    python docker\\verify_helper.py pages   <port> <results stem>
    python docker\\verify_helper.py atomic  <container name> <data folder>

It talks only to 127.0.0.1 and to the local docker command, uses fake values only, and never prints a setting's value.
The UI is driven the way a browser drives it: the session cookie, the CSRF token from the page, the Origin header."""
import hashlib
import http.client
import io
import json
import os
import re
import socket
import subprocess
import sys
import tarfile
import threading
import time
from pathlib import Path
from urllib.parse import quote, urlencode

ROOT = Path(__file__).resolve().parent.parent
STATE_NAMES = {".env", "email_state.json", "decision_state.json", "approvals_log.jsonl", "approved_orders.json", "approved_orders.csv",
               "product_facts.csv"}
DATA_FOLDERS = {"Excel Output Sheets", "Reports", "Order Sheets", "Walmart Listings", "Uploads", "ui_data", ".git"}
REPLY = "approve VIC00001 qty 100 max 1.30\nreject HDFF\ncap $200\n"       # the reply test_gates.py uses on the same rows
RESULTS_NAME = "sourcing_results_20260101_100000.xlsx"


def sha(data: bytes) -> str:
    return hashlib.sha256(data).hexdigest()


# ---------- scan: every layer of a saved image ----------

def scan(tar_path: str, kind: str, needles: list) -> dict:
    """Looks inside every layer (so also at files a later layer deleted) and at the image's own settings and history."""
    wanted = [n.encode() for n in needles]
    bad, hits, layers, files = [], [], 0, 0
    with tarfile.open(tar_path) as outer:
        for member in outer.getmembers():
            if not member.isfile():
                continue
            blob = outer.extractfile(member).read()
            try:
                inner = tarfile.open(fileobj=io.BytesIO(blob), mode="r:*")
            except tarfile.TarError:                     # not a layer: the image's settings, history or index
                hits += [f"image settings {member.name}: needle {i}" for i, n in enumerate(wanted) if n in blob]
                continue
            layers += 1
            with inner:
                for f in inner.getmembers():
                    path = f.name.lstrip("./")
                    parts = path.split("/")
                    if parts[0] in ("app", "data", "tmp", "root", "home") or len(parts) == 1:
                        if parts[-1] in STATE_NAMES or DATA_FOLDERS & set(parts):
                            bad.append(path)
                        if kind == "runtime" and (parts[-1] in ("stub_agent.py", ".env.example") or {"tests", "test_fixtures", "scratch"} & set(parts)
                                                  or re.fullmatch(r"test_.*\.py", parts[-1])) and parts[0] == "app":
                            bad.append(path)
                    if f.isfile():
                        files += 1
                        data = inner.extractfile(f).read()
                        hits += [f"file {path}: needle {i}" for i, n in enumerate(wanted) if n in data]
    return {"ok": not bad and not hits and layers > 0, "layers": layers, "files": files, "bad_names": sorted(set(bad)), "needle_hits": hits}


# ---------- envhash: what python-dotenv reads from a settings file, as hashes ----------

def envhash(path: str, names: list) -> dict:
    from dotenv import dotenv_values
    values = dotenv_values(path)
    return {n: (sha(values[n].encode("utf-8")) if values.get(n) is not None else None) for n in names}


def envvalues(path: str, names: list) -> dict:
    """python-dotenv's reading of a file of made-up test values (never used on a real settings file)."""
    from dotenv import dotenv_values
    values = dotenv_values(path)
    return {n: values.get(n) for n in names}


# ---------- fixture: a results workbook in the search agent's own layout ----------

def fixture(data: str) -> dict:
    os.environ["OBS_ENABLED"] = "0"
    sys.path.insert(0, str(ROOT))
    import test_decision_agent as td                    # its helper writes the columns the agents read
    folder = Path(data) / "Excel Output Sheets"
    folder.mkdir(parents=True, exist_ok=True)
    rows = td.standard_rows()
    for r in rows:                                      # the order sheet reads the price from the seller's own block of columns
        r["Manufacturer 1 Unit Price"], r["Manufacturer 1 Listed Price"] = r["Recommended Unit Price"], f"${r['Recommended Unit Price']}"
    td.make_results(str(folder / RESULTS_NAME), rows)
    return {"ok": True, "stem": RESULTS_NAME[:-5], "file": RESULTS_NAME}


# ---------- hosts: the Host, Origin and CSRF checks, through the published port ----------

def raw(port: int, method: str, path: str, headers: dict, body: bytes = b"") -> tuple:
    """One request written by hand, so the Host header reaches the server exactly as typed here."""
    lines = [f"{method} {path} HTTP/1.1"] + [f"{k}: {v}" for k, v in headers.items()] + ["Connection: close"]
    if method == "POST":
        lines += [f"Content-Length: {len(body)}", "Content-Type: application/x-www-form-urlencoded"]
    with socket.create_connection(("127.0.0.1", port), timeout=15) as s:
        s.sendall(("\r\n".join(lines) + "\r\n\r\n").encode("latin-1") + body)
        data = b""
        while chunk := s.recv(65536):
            data += chunk
    head, _, rest = data.partition(b"\r\n\r\n")
    return int(head.split(b" ", 2)[1]), head.decode("latin-1"), rest.decode("utf-8", "replace")


def hosts(port: int, container: str) -> dict:
    good, origin = f"127.0.0.1:{port}", f"http://127.0.0.1:{port}"
    failures, seen = [], []

    def expect(what, got, want, text="", must=""):
        seen.append(f"{what}: {got}")
        if got != want or (must and must not in text):
            failures.append(f"{what}: got {got}, wanted {want}" + (f" with '{must}'" if must else ""))

    for host in (good, f"localhost:{port}"):
        expect(f"GET Host {host}", raw(port, "GET", "/health", {"Host": host})[0], 200)
    wrong = [f"0.0.0.0:{port}", f"host.docker.internal:{port}", f"127.0.0.1:{port + 1}", "127.0.0.1", f"LOCALHOST:{port}", f"Localhost:{port}",
             f"localhost.:{port}", f"127.0.0.1.:{port}", f"{container}:{port}", f"evil.example:{port}", "evil.example", f"[::1]:{port}"]
    if port != 8000:
        wrong.append("127.0.0.1:8000")
    for host in wrong:
        status, _, text = raw(port, "GET", "/health", {"Host": host})
        expect(f"GET Host {host}", status, 400, text, "Invalid Host header.")
        status, _, text = raw(port, "POST", "/jobs/run", {"Host": host, "Origin": origin}, b"command=help.email_agent")
        expect(f"POST Host {host}", status, 400, text, "Invalid Host header.")
    status, head, html = raw(port, "GET", "/step/search", {"Host": good})
    cookie = re.search(r"(?i)set-cookie: (ui_sid=[^;\r\n]+)", head)
    token = re.search(r'name="csrf_token" value="([0-9a-f]+)"', html)
    if status != 200 or not cookie or not token:
        return {"ok": False, "failures": failures + ["could not read a session cookie and token from /step/search"], "seen": seen}
    form = urlencode({"csrf_token": token[1], "command": "help.email_agent"}).encode()
    for bad in ("http://evil.example", f"http://0.0.0.0:{port}", f"http://host.docker.internal:{port}", f"http://127.0.0.1:{port + 1}",
                f"http://LOCALHOST:{port}", f"http://localhost.:{port}", f"https://127.0.0.1:{port}", "null", None):
        headers = {"Host": good, "Cookie": cookie[1], **({"Origin": bad} if bad else {})}
        status, _, text = raw(port, "POST", "/jobs/run", headers, form)
        expect(f"POST Origin {bad}", status, 403, text, "Cross-site")
    for what, body in (("no token", b"command=help.email_agent"), ("a wrong token", urlencode({"csrf_token": "0" * 64, "command": "help.email_agent"}).encode())):
        status, _, _ = raw(port, "POST", "/jobs/run", {"Host": good, "Origin": origin, "Cookie": cookie[1]}, body)
        expect(f"POST with {what}", status, 403)
    status, _, _ = raw(port, "POST", "/jobs/run", {"Host": good, "Origin": origin}, form)                 # the token belongs to a session
    expect("POST with the token but no session cookie", status, 403)
    status, _, _ = raw(port, "POST", "/jobs/run", {"Host": good, "Origin": origin, "Cookie": cookie[1]}, form)
    expect("POST with the right host, origin, cookie and token", status, 303)
    return {"ok": not failures, "checked": len(seen), "failures": failures, "seen": seen}


# ---------- flow and pages: real work through the UI, landing in the host folder ----------

class UI:
    def __init__(self, port: int):
        self.port, self.cookie, self.origin = port, "", f"http://127.0.0.1:{port}"

    def request(self, method: str, path: str, body=None, ctype: str = "") -> tuple:
        conn = http.client.HTTPConnection("127.0.0.1", self.port, timeout=120)
        headers = {"Origin": self.origin} if method == "POST" else {}
        if self.cookie:
            headers["Cookie"] = self.cookie
        if ctype:
            headers["Content-Type"] = ctype
        try:
            conn.request(method, path, body=body, headers=headers)
            r = conn.getresponse()
            data = r.read()
            for k, v in r.getheaders():
                if k.lower() == "set-cookie" and v.startswith("ui_sid="):
                    self.cookie = v.split(";")[0]
            return r.status, {k.lower(): v for k, v in r.getheaders()}, data
        finally:
            conn.close()

    def get(self, path: str) -> tuple:
        status, _, data = self.request("GET", path)
        return status, data.decode("utf-8", "replace")

    def token(self) -> str:
        return re.search(r'name="csrf_token" value="([0-9a-f]+)"', self.get("/step/search")[1])[1]

    def post(self, path: str, **fields) -> tuple:
        return self.request("POST", path, urlencode({"csrf_token": self.token(), **fields}), "application/x-www-form-urlencoded")

    def upload(self, path: str, field: str, filename: str, data: bytes) -> tuple:
        mark = "----lcomverify" + sha(data)[:16]
        body = (f"--{mark}\r\nContent-Disposition: form-data; name=\"csrf_token\"\r\n\r\n{self.token()}\r\n"
                f"--{mark}\r\nContent-Disposition: form-data; name=\"{field}\"; filename=\"{filename}\"\r\n"
                "Content-Type: application/octet-stream\r\n\r\n").encode() + data + f"\r\n--{mark}--\r\n".encode()
        return self.request("POST", path, body, f"multipart/form-data; boundary={mark}")

    def job(self, headers: dict, timeout: float = 180) -> tuple:
        """Follows a started command to its end. Returns (status, log)."""
        found = re.search(r"job=([0-9a-f-]+)", headers.get("location", ""))
        if not found:
            return "not started", ""
        end = time.monotonic() + timeout
        while time.monotonic() < end:
            status, data = self.get(f"/jobs/{found[1]}/log?offset=0")
            body = json.loads(data) if status == 200 else {}
            if body.get("status") and body["status"] != "running":
                return body["status"], body.get("text", "")
            time.sleep(0.3)
        return "timed out", ""


def flow(port: int, data: str, stem: str) -> dict:
    ui, root, steps = UI(port), Path(data), []

    def step(name, ok, detail=""):
        steps.append({"step": name, "ok": bool(ok), "detail": str(detail)[-600:]})
        return bool(ok)

    def run(name, path, expect_files, **fields):
        status, headers, body = ui.post(path, **fields)
        if status != 303:
            return step(name, False, f"HTTP {status}: {body.decode('utf-8', 'replace')[-400:]}")
        state, log = ui.job(headers)
        missing = [f for f in expect_files if not list(root.glob(f))]
        return step(name, state == "succeeded" and not missing, f"job {state}; missing on the host: {missing}; log end: {log[-300:]}" if state != "succeeded" or missing else "")

    import openpyxl
    book = openpyxl.Workbook()
    for row in (["SKU", "Product", "Keyword"], ["VER-1", "Verify cable", "Cat6 patch cable"], ["VER-2", "Verify adapter", "USB adapter"]):
        book.active.append(row)
    buf = io.BytesIO()
    book.save(buf)
    status, headers, body = ui.upload("/search/upload", "sheet", "verify products.xlsx", buf.getvalue())
    name = re.search(r"Uploads%2F(input_[0-9a-f_]+\.xlsx)", headers.get("location", ""))
    saved = root / "Uploads" / name[1] if name else None
    step("upload a product sheet", status == 303 and saved and saved.is_file() and saved.read_bytes() == buf.getvalue(),
         f"HTTP {status}; file on the host: {bool(saved and saved.is_file())}")

    run("write an approval request", "/approvals/request", ["decision_state.json", "Reports/approval_request_A-0001.txt"], wf=stem, ack="1")
    refused, _, _ = ui.post("/approvals/record", wf=stem, request_id="A-0001", reply=REPLY, confirm_request="A-0002", channel="verify channel",
                            approver="Verify Person")
    step("a record with the wrong typed request id is refused", refused == 400, f"HTTP {refused}")
    run("record the reply with typed fields", "/approvals/record", ["approvals_log.jsonl"], wf=stem, request_id="A-0001", reply=REPLY,
        confirm_request="A-0001", channel="verify channel", approver="Verify Person")
    try:
        decisions = {d["sku"]: d["decision"] for d in json.loads((root / "decision_state.json").read_text(encoding="utf-8"))["decisions"]}
    except (OSError, ValueError, KeyError) as e:
        decisions = {"error": type(e).__name__}
    step("the approval is in the state file on the host", decisions.get("VIC00001") == "approve" and decisions.get("HDFF") == "reject", decisions)
    run("export the approvals", "/approvals/export", ["approved_orders.json", "approved_orders.csv"], wf=stem)
    try:
        source = json.loads((root / "approved_orders.json").read_text(encoding="utf-8")).get("source_results_file")
    except (OSError, ValueError) as e:
        source = type(e).__name__
    step("the export names this results file", source == stem + ".xlsx", source)
    status, _, served = ui.request("GET", "/files/project/approved_orders.csv")
    step("the export downloads byte for byte", status == 200 and (root / "approved_orders.csv").is_file() and served == (root / "approved_orders.csv").read_bytes(), f"HTTP {status}")
    run("build the order sheet", "/order/build", ["Order Sheets/order_sheet_*.xlsx"], wf=stem)
    left = sorted(p.name for p in root.iterdir() if p.name.endswith((".tmp", ".lock", ".takeover")))
    step("no temp or lock file is left beside the state files", not left, left)
    return {"ok": all(s["ok"] for s in steps), "steps": steps}


PLANT_NAME = "sourcing_results_20260101_100200.xlsx"


def plantjob(port: int, data: str, user: str, password: str, proxy_url: str) -> dict:
    """A results file whose seller names and listing addresses carry a user name and password (and the container's own
    proxy setting), then the email agent's `drafts` through the UI: it prints those names and addresses. The stored log,
    the job record and the job's pages must show them masked. Nothing here needs the network."""
    os.environ["OBS_ENABLED"] = "0"
    sys.path.insert(0, str(ROOT))
    import test_decision_agent as td
    planted = f"http://{user}:{password}@seller.example/item"
    rows = [td.res_row("PLANT1", f"Maker {planted}", url=planted + "/1"), td.res_row("PLANT2", f"Maker via {proxy_url}", url=planted + "/2")]
    folder = Path(data) / "Excel Output Sheets"
    td.make_results(str(folder / PLANT_NAME), rows)
    ui = UI(port)
    status, headers, body = ui.post("/emails/drafts", wf=PLANT_NAME[:-5])
    found = re.search(r"job=([0-9a-f-]+)", headers.get("location", ""))
    if status != 303 or not found:
        return {"ok": False, "error": f"drafts was not started (HTTP {status})"}
    state, log = ui.job(headers)
    jid = found[1]
    stored = "".join(p.read_text(encoding="utf-8", errors="replace") for p in (Path(data) / "ui_data" / "jobs").glob(f"{jid}.*"))
    shown = ui.get(f"/jobs/{jid}")[1] + ui.get(f"/jobs/{jid}/log?offset=0")[1] + ui.get(f"/step/emails?wf={PLANT_NAME[:-5]}&job={jid}")[1]
    secrets_ = [user, password, proxy_url, proxy_url.split("://", 1)[-1].rsplit("@", 1)[0]]
    return {"ok": state == "succeeded" and stored.count("[REDACTED]@") >= 2 and not any(s in stored or s in shown for s in secrets_),
            "job": state, "masked_addresses_in_log": stored.count("[REDACTED]@"),
            "found_in_stored_log": [i for i, s in enumerate(secrets_) if s in stored], "found_on_pages": [i for i, s in enumerate(secrets_) if s in shown]}


def pagescan(port: int, stem: str, needles: list) -> dict:
    """Every page the UI serves for this run (each tab, Health, Jobs, every job's page and log): none may hold a needle."""
    ui, hits = UI(port), []
    paths = ["/", f"/?wf={stem}", "/health", "/jobs", "/step/search"] + [f"/step/{s}?wf={stem}" for s in ("emails", "approvals", "order-sheet", "listing-content")]
    status, html = ui.get("/jobs")
    for jid in sorted(set(re.findall(r'href="/jobs/(\d{8}-\d{6}-[0-9a-f]{6})"', html))):
        paths += [f"/jobs/{jid}", f"/jobs/{jid}/log?offset=0"]
    for path in paths:
        status, headers, body = ui.request("GET", path)
        text = body.decode("utf-8", "replace") + json.dumps(headers)
        hits += [f"{path}: needle {i}" for i, n in enumerate(needles) if n in text]
    return {"ok": not hits, "pages": len(paths), "hits": hits}


def pages(port: int, stem: str) -> dict:
    ui, problems = UI(port), []
    for path, wanted in ((f"/?wf={stem}", [stem]), (f"/step/approvals?wf={stem}", ["A-0001", "VIC00001", "Valid"]),
                         (f"/step/order-sheet?wf={stem}", ["Built: order_sheet_"]), ("/health", ["Times on these pages are shown in"])):
        status, html = ui.get(path)
        problems += [f"{path}: HTTP {status}"] if status != 200 else [f"{path}: '{w}' not on the page" for w in wanted if w not in html]
    status, html = ui.get("/health")
    zone = re.search(r"shown in <strong>([^<]+)</strong>", html)
    return {"ok": not problems, "problems": problems, "zone": zone[1] if zone else None,
            "set": sorted(set(re.findall(r"<code>([A-Z_]+)</code></th>\s*<td data-label=\"Status\"><span[^>]*>Set<", html)))}


# ---------- atomic: state writes on the mounted folder while the host reads ----------

WRITER = """
import json, sys
sys.path.insert(0, '/data')
import statefile
path = '/data/_verify_atomic.json'
for i in range(200):
    with statefile.locked(path):
        statefile.atomic_write(path, json.dumps({'n': i, 'pad': 'x' * 4000}))
print('written 200')
"""


def atomic(container: str, data: str) -> dict:
    target, counts, stop = Path(data) / "_verify_atomic.json", {"reads": 0, "busy": 0, "partial": 0}, threading.Event()

    def reader():
        while not stop.is_set():
            try:
                text = target.read_text(encoding="utf-8")
            except OSError:
                counts["busy"] += 1                      # not there yet, or held for the instant of the rename
                continue
            counts["reads"] += 1
            try:
                if len(json.loads(text)["pad"]) != 4000:
                    counts["partial"] += 1
            except (ValueError, KeyError):
                counts["partial"] += 1

    t = threading.Thread(target=reader, daemon=True)
    t.start()
    p = subprocess.run(["docker", "exec", "-i", container, "python", "-"], input=WRITER, capture_output=True, text=True, timeout=600)
    stop.set()
    t.join(5)
    try:
        last = json.loads(target.read_text(encoding="utf-8"))["n"]
    except (OSError, ValueError, KeyError) as e:
        last = type(e).__name__
    left = sorted(q.name for q in Path(data).iterdir() if q.name.startswith("_verify_atomic.json.") or q.name.endswith(".takeover"))
    target.unlink(missing_ok=True)
    return {"ok": p.returncode == 0 and counts["partial"] == 0 and last == 199 and not left, "writer_exit": p.returncode,
            "writer_said": (p.stdout.strip().splitlines() or p.stderr.strip().splitlines() or [""])[-1][:300], "last": last, "left_behind": left, **counts}


def main(argv: list) -> dict:
    cmd, a = argv[0], argv[1:]
    if cmd == "scan":
        return scan(a[0], a[1], a[2:])
    if cmd == "envhash":
        return envhash(a[0], a[1:])
    if cmd == "envvalues":
        return envvalues(a[0], a[1:])
    if cmd == "pagescan":
        return pagescan(int(a[0]), a[1], a[2:])
    if cmd == "plantjob":
        return plantjob(int(a[0]), a[1], a[2], a[3], a[4])
    if cmd == "fixture":
        return fixture(a[0])
    if cmd == "hosts":
        return hosts(int(a[0]), a[1])
    if cmd == "flow":
        return flow(int(a[0]), a[1], a[2])
    if cmd == "pages":
        return pages(int(a[0]), a[1])
    if cmd == "atomic":
        return atomic(a[0], a[1])
    raise SystemExit(__doc__)


if __name__ == "__main__":
    try:
        result = main(sys.argv[1:])
    except SystemExit:
        raise
    except Exception as e:  # noqa: BLE001 - verify.ps1 reads one JSON line whatever happens
        result = {"ok": False, "error": f"{type(e).__name__}: {e}"[:500]}
    print(json.dumps(result))
