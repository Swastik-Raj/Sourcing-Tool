"""Offline tests for the web UI: FastAPI TestClient, the stub agent only, no network, no real agent runs.
Run:  $env:OBS_ENABLED = "0"; python test_ui.py"""
import json
import os
import re
import shutil
import subprocess
import sys
import tempfile
import time
import traceback
from pathlib import Path
from types import SimpleNamespace

os.environ["OBS_ENABLED"] = "0"
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import openpyxl
from fastapi.testclient import TestClient

import observability as obs
from ui import commands, jobs as jobs_mod, security, workflows
from ui.app import create_app
from ui.commands import JobRejected
from ui.config import Settings, load_settings
from ui.jobs import JobBusy, JobManager, write_json_atomic

ORIGIN = {"origin": "http://127.0.0.1:8000"}
FAKE_KEY = "sk-ant-FAKEKEYVALUE1234567890"
TESTS = []


def test(f):
    TESTS.append(f)
    return f


def make_env(stub=True):
    root = Path(tempfile.mkdtemp(prefix="ui_test_"))
    project = root / "project"
    (project / "Excel Output Sheets").mkdir(parents=True)
    for name in ("sourcing_results_20260101_100000.xlsx", "sourcing_results_20260202_100000.xlsx"):
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Results"
        ws.append(["SKU", "Product", "Recommended URL"])
        ws.append(["AB-1", "Cable", "https://example.test/a"])
        ws.append(["CD-2", "Adapter", ""])
        wb.save(project / "Excel Output Sheets" / name)
        t = time.time() - (100000 if "0101" in name else 0)
        os.utime(project / "Excel Output Sheets" / name, (t, t))
    s = Settings(project=project, data=project / "ui_data", enable_stub=stub)
    app = create_app(s)
    client = TestClient(app, base_url="http://127.0.0.1:8000", raise_server_exceptions=False)
    return SimpleNamespace(root=root, project=project, settings=s, app=app, client=client, jobs=app.state.jobs)


def cleanup(env):
    for j in env.jobs.list():
        env.jobs.cancel(j["id"])
    time.sleep(0.3)
    shutil.rmtree(env.root, ignore_errors=True)


def token(env):
    html = env.client.get("/jobs").text
    return re.search(r'name="csrf_token" value="([0-9a-f]+)"', html)[1]


def raises(exc, fn, *a, **kw):
    try:
        fn(*a, **kw)
    except exc:
        return
    raise AssertionError(f"expected {exc.__name__} from {getattr(fn, '__name__', fn)}{a}{kw}")


# ---------- allowlist ----------

@test
def allowlist_rejects_unknown_commands_bad_skus_paths_and_flags(env):
    j = env.jobs
    raises(JobRejected, j.start, "rm.everything")
    raises(JobRejected, j.start, "email_agent.send")            # real commands are not in Stage 1
    for bad in ("x; calc", "-x", "a b", "../x", "A" * 41, "", "a\nb", "$(x)", "a|b"):
        raises(JobRejected, j.start, "stub.read", {"mode": "ok", "sku": bad})
    for bad in ("C:\\Windows\\win.ini", "..\\x.xlsx", "/etc/passwd", "Excel Output Sheets\\sourcing_results_20260101_100000.xlsx", "nope.xlsx"):
        raises(JobRejected, j.start, "stub.read", {"mode": "ok", "results": bad})
    raises(JobRejected, j.start, "stub.read", {"mode": "ok", "yes": "--yes"})         # extra flag/param
    raises(JobRejected, j.start, "stub.read", {"mode": "--help"})                       # not in the enum
    raises(JobRejected, j.start, "stub.read", {})                                       # missing mode
    argv = env.jobs.commands["stub.read"].argv({"mode": "ok", "sku": "C&P9M", "results": "sourcing_results_20260101_100000.xlsx"})
    assert argv == ["ok", "--sku", "C&P9M", "--results", "Excel Output Sheets\\sourcing_results_20260101_100000.xlsx"], argv
    assert not any(c.startswith("stub") for c in JobManager(Settings(project=env.project, data=env.project / "d2")).commands), "stub must be off by default"


# ---------- state-changing lock ----------

@test
def lock_refuses_second_state_job_and_releases(env):
    j = env.jobs
    first = j.start("stub.state", {"mode": "slow"})
    raises(JobBusy, j.start, "stub.state", {"mode": "ok"})
    assert j.running_state_job()["id"] == first["id"]
    ro = j.start("stub.read", {"mode": "ok"})                                         # read-only runs alongside
    assert j.wait(ro["id"])["status"] == "succeeded"
    assert j.cancel(first["id"])
    assert j.wait(first["id"])["status"] == "cancelled"
    done = j.start("stub.state", {"mode": "ok"})                                      # released after cancel
    assert j.wait(done["id"])["status"] == "succeeded"
    again = j.start("stub.state", {"mode": "fail"})                                   # released after success
    assert j.wait(again["id"])["status"] == "failed"
    last = j.start("stub.state", {"mode": "ok"})                                      # released after failure
    assert j.wait(last["id"])["status"] == "succeeded"
    assert j.running_state_job() is None


@test
def runner_uses_argument_list_no_shell_closed_stdin(env):
    seen, real = {}, subprocess.Popen
    def spy(argv, **kw):
        seen.update(argv=argv, **kw)
        return real(argv, **kw)
    jobs_mod.subprocess.Popen = spy
    try:
        m = env.jobs.wait(env.jobs.start("stub.read", {"mode": "ok", "sku": "C&P9M"})["id"])
    finally:
        jobs_mod.subprocess.Popen = real
    assert isinstance(seen["argv"], list) and seen["argv"][:3] == [sys.executable, "-X", "utf8"], seen["argv"]
    assert seen["argv"][-3:] == ["ok", "--sku", "C&P9M"], seen["argv"]
    assert seen["shell"] is False and seen["stdin"] == subprocess.DEVNULL and Path(seen["cwd"]) == env.project
    assert seen["env"]["PYTHONIOENCODING"] == "utf-8" and seen["env"]["PYTHONUNBUFFERED"] == "1"
    assert m["status"] == "succeeded" and "sku=C&P9M" in env.jobs.read_log(m["id"])[0]


@test
def failed_job_records_exit_code_meaning_and_tail(env):
    m = env.jobs.wait(env.jobs.start("stub.read", {"mode": "fail"})["id"])
    assert m["exit_code"] == 3 and "not one of the agent's documented codes" in m["exit_meaning"], m
    assert any("failing on purpose" in line for line in m["tail"]), m["tail"]


@test
def cancel_kills_the_whole_tree_and_warns(env):
    m = env.jobs.start("stub.state", {"mode": "slow"})
    child = None
    for _ in range(100):
        match = re.search(r"child pid (\d+)", env.jobs.read_log(m["id"])[0])
        if match:
            child = int(match[1])
            break
        time.sleep(0.1)
    assert child, "stub never reported its child"
    assert env.jobs.cancel(m["id"])
    done = env.jobs.wait(m["id"])
    assert done["status"] == "cancelled" and done["note"], done
    time.sleep(0.5)
    out = subprocess.run(["tasklist", "/FI", f"PID eq {child}"], capture_output=True, text=True).stdout
    assert str(child) not in out, "grandchild process survived the cancel"


@test
def interrupted_jobs_are_marked_on_startup(env):
    meta = {"id": "20260101-000000-abcdef", "command": "stub.read", "label": "x", "args": [], "state_changing": True,
            "status": "running", "started_at": "2026-01-01T00:00:00", "finished_at": None, "duration_s": None,
            "exit_code": None, "exit_meaning": "", "pid": 1, "note": "", "tail": []}
    write_json_atomic(env.settings.jobs / f"{meta['id']}.json", meta)
    fresh = JobManager(env.settings)
    got = fresh.get(meta["id"])
    assert got["status"] == "interrupted" and got["note"], got
    assert fresh.running_state_job() is None                                          # never resumed, never holds the lock


@test
def metadata_write_is_atomic(env):
    target = env.settings.jobs / "atomic.json"
    write_json_atomic(target, {"v": 1})
    calls, real = [], os.replace
    def spy(a, b):
        calls.append((str(a), str(b)))
        raise OSError("disk full")
    jobs_mod.os.replace = spy
    try:
        raises(OSError, write_json_atomic, target, {"v": 2})
    finally:
        jobs_mod.os.replace = real
    assert json.loads(target.read_text()) == {"v": 1}, "original must survive a failed write"
    assert calls and calls[0][0].endswith(".tmp") and calls[0][1] == str(target), calls
    write_json_atomic(target, {"v": 3})
    assert json.loads(target.read_text()) == {"v": 3}


# ---------- masking and escaping ----------

@test
def masking_removes_keys_emails_phones_from_logs(env):
    for raw in (f"key {FAKE_KEY}", "ANTHROPIC_API_KEY=abc123xyz789", "mail boss@example.com", "call +1 214 555 0100", "Bearer abcdefghijklmnop123"):
        out = jobs_mod.mask_line(raw)
        assert FAKE_KEY not in out and "boss@example.com" not in out and "555 0100" not in out and "abc123xyz789" not in out and "abcdefghijklmnop123" not in out, (raw, out)
    m = env.jobs.wait(env.jobs.start("stub.read", {"mode": "secret"})["id"])
    log = env.jobs.read_log(m["id"])[0]
    assert "sk-ant-abcdefgh12345678" not in log and "boss@example.com" not in log and "555 0100" not in log, log
    assert "[REDACTED]" in log and "[EMAIL]" in log, log


@test
def hostile_log_text_is_escaped(env):
    m = env.jobs.wait(env.jobs.start("stub.read", {"mode": "hostile"})["id"])
    html = env.client.get(f"/jobs/{m['id']}").text
    assert "<script>alert" not in html and "<img src=x" not in html, "log text must be escaped"
    assert "&lt;script&gt;alert" in html


# ---------- Host, Origin, CSRF ----------

@test
def host_header_is_checked(env):
    assert env.client.get("/health").status_code == 200
    assert env.client.get("/health", headers={"host": "evil.example"}).status_code == 400
    assert env.client.get("/health", headers={"host": "127.0.0.1:9999"}).status_code == 400
    assert env.client.get("/health", headers={"host": "localhost:8000"}).status_code == 200


@test
def origin_is_checked_on_post(env):
    t = token(env)
    body = {"command": "stub.read", "mode": "ok", "csrf_token": t}
    assert env.client.post("/jobs/run", data=body, headers={"origin": "http://evil.example"}, follow_redirects=False).status_code == 403
    assert env.client.post("/jobs/run", data=body, follow_redirects=False).status_code == 403            # no Origin, no Referer
    assert env.client.post("/jobs/run", data=body, headers={"origin": "null"}, follow_redirects=False).status_code == 403
    assert env.client.post("/jobs/run", data=body, headers=ORIGIN, follow_redirects=False).status_code == 303


@test
def csrf_token_is_required(env):
    t = token(env)
    body = {"command": "stub.read", "mode": "ok"}
    assert env.client.post("/jobs/run", data=body, headers=ORIGIN, follow_redirects=False).status_code == 403
    assert env.client.post("/jobs/run", data={**body, "csrf_token": "0" * 64}, headers=ORIGIN, follow_redirects=False).status_code == 403
    other = TestClient(env.app, base_url="http://127.0.0.1:8000", raise_server_exceptions=False)
    assert other.post("/jobs/run", data={**body, "csrf_token": t}, headers=ORIGIN, follow_redirects=False).status_code == 403   # token of another session
    r = env.client.post("/jobs/run", data={**body, "csrf_token": t}, headers=ORIGIN, follow_redirects=False)
    assert r.status_code == 303 and r.headers["location"].startswith("/jobs/")
    jid = r.headers["location"].rsplit("/", 1)[1]
    assert env.client.post(f"/jobs/{jid}/cancel", headers=ORIGIN, follow_redirects=False).status_code == 403
    assert env.client.post("/jobs/run", data={"command": "nope", "csrf_token": t}, headers=ORIGIN).status_code == 400


@test
def security_headers_are_sent(env):
    h = env.client.get("/").headers
    assert h["content-security-policy"] == "default-src 'self'"
    assert h["x-content-type-options"] == "nosniff" and h["x-frame-options"] == "DENY" and h["referrer-policy"] == "no-referrer"
    assert env.client.get("/health", headers={"host": "evil.example"}).headers["x-frame-options"] == "DENY"


# ---------- downloads ----------

def make_link(link: Path, target: Path) -> bool:
    try:
        os.symlink(target, link, target_is_directory=True)
        return True
    except OSError:
        r = subprocess.run(["cmd", "/c", "mklink", "/J", str(link), str(target)], capture_output=True)   # junction: no admin needed
        return r.returncode == 0


@test
def downloads_stay_inside_output_folders(env):
    s = env.settings
    (s.project / "Reports").mkdir()
    (s.project / "Reports" / "ok.csv").write_text("a,b\n")
    (s.project / "Reports" / "state.json").write_text("{}")
    (s.project / "Reports" / "tool.py").write_text("print(1)")
    (s.project / "email_state.json").write_text("{}")
    (env.root / "outside").mkdir()
    (env.root / "outside" / "secret.csv").write_text("top secret")
    (env.root / "secret.csv").write_text("top secret")
    ok = security.safe_output_file
    assert ok(s, "Reports", "ok.csv") == (s.project / "Reports" / "ok.csv").resolve()
    for folder, rel in [("Reports", "../secret.csv"), ("Reports", "..\\secret.csv"), ("Reports", "..\\..\\project\\email_state.json"),
                        ("Reports", str(env.root / "secret.csv")), ("Reports", "C:\\Windows\\win.ini"), ("Reports", "/etc/passwd"),
                        ("Reports", "\\\\server\\share\\x.csv"), ("Reports", "ok.csv\x00.txt"), ("Reports", "state.json"), ("Reports", "tool.py"),
                        ("Reports", "missing.csv"), ("Reports", ""), ("Reports", "."), ("..", "project/email_state.json"),
                        ("ui_data", "errors.log"), ("Reports", "ok.csv:stream")]:
        assert ok(s, folder, rel) is None, (folder, rel)
    if make_link(s.project / "Reports" / "link", env.root / "outside"):
        assert ok(s, "Reports", "link/secret.csv") is None, "link/junction escape must be refused"
        assert env.client.get("/files/Reports/link/secret.csv").status_code == 404
    else:
        raise AssertionError("could not create a symlink or junction to test escapes")
    assert env.client.get("/files/Reports/ok.csv").status_code == 200
    for url in ("/files/Reports/state.json", "/files/Reports/..%2fsecret.csv", "/files/Reports/%2e%2e/secret.csv", "/files/Nope/ok.csv"):
        assert env.client.get(url).status_code == 404, url


# ---------- secrets, pages, workflows ----------

@test
def secrets_never_appear_in_any_page(env):
    saved = {k: os.environ.get(k) for k in ("ANTHROPIC_API_KEY", "SMTP_PASSWORD", "SHIPPING_ADDRESS", "NIMBLE_API_KEY")}
    os.environ.update(ANTHROPIC_API_KEY=FAKE_KEY, SMTP_PASSWORD="hunter2-fake-password", SHIPPING_ADDRESS="9 Fake Rd\\nNowhere, ZZ 00000", NIMBLE_API_KEY="nimble-fake-secret-value")
    try:
        job = env.jobs.wait(env.jobs.start("stub.read", {"mode": "secret"})["id"])
        pages = ["/", "/health", "/jobs", f"/jobs/{job['id']}", f"/jobs/{job['id']}/log", "/step/search", "/step/order-sheet", "/nope", "/jobs/20260101-000000-000000"]
        text = "".join(env.client.get(p).text for p in pages)
        for secret in (FAKE_KEY, "hunter2-fake-password", "9 Fake Rd", "nimble-fake-secret-value"):
            assert secret not in text, f"{secret!r} leaked"
        health = env.client.get("/health").text
        assert "ANTHROPIC_API_KEY" in health and "SMTP_HOST" in health and "Not set" in health and "Needed for" in health
    finally:
        for k, v in saved.items():
            os.environ.pop(k, None) if v is None else os.environ.__setitem__(k, v)


@test
def unhandled_errors_show_only_a_reference(env):
    def boom():
        raise RuntimeError(f"db password {FAKE_KEY}")
    real, env.jobs.list = env.jobs.list, boom
    r = env.client.get("/jobs")
    env.jobs.list = real
    assert r.status_code == 500 and FAKE_KEY not in r.text and "RuntimeError" not in r.text and "quote this reference" in r.text
    log = (env.settings.data / "errors.log").read_text()
    assert FAKE_KEY not in log and "RuntimeError" in log


@test
def workflow_id_matches_the_agents_and_lists_newest_first(env):
    wfs = workflows.list_workflows(env.settings)
    assert [w["id"] for w in wfs] == ["lcom-sourcing_results_20260202_100000", "lcom-sourcing_results_20260101_100000"], wfs
    for w in wfs:
        assert w["id"] == obs.workflow_id(str(Path("Excel Output Sheets") / w["file"])) == obs.workflow_id(env.settings.results_dir / w["file"])
    here = Path(__file__).parent
    for agent in ("email_agent", "decision_agent", "order_sheet", "content_agent"):
        assert "obs.workflow_id(" in (here / f"{agent}.py").read_text(encoding="utf-8"), agent
    assert "obs.workflow_id(excel_path)" in (here / "sourcing_agent.py").read_text(encoding="utf-8")


@test
def overview_steps_only_claim_what_files_show(env):
    page = env.client.get("/").text
    assert "lcom-" not in page and "2 products" in page and "1 can be drafted" in page and "1 have no recommended seller" in page
    steps = {s.name: s for s in workflows.steps_for(env.settings, "sourcing_results_20260202_100000.xlsx")}
    assert [s for s in steps] == list(workflows.STEPS)
    assert steps["Search"].status == "Files found" and steps["Emails"].status == "Not started"
    assert steps["Approvals"].blocked == "" and steps["Order sheet"].blocked and steps["Listing content"].blocked
    (env.project / "decision_state.json").write_text(json.dumps({"next_request": 2, "decisions": [], "requests": {
        "A-0001": {"id": "A-0001", "results_file": "sourcing_results_20260202_100000.xlsx", "reply": None, "lines": []},
        "A-0002": {"id": "A-0002", "results_file": "sourcing_results_20260101_100000.xlsx", "reply": None, "lines": []}}}))
    (env.project / "email_state.json").write_text("{not json")
    steps = {s.name: s for s in workflows.steps_for(env.settings, "sourcing_results_20260202_100000.xlsx")}
    assert steps["Approvals"].status == "Files found" and "1 request(s)" in steps["Approvals"].detail and "still waiting" in steps["Approvals"].detail
    assert steps["Emails"].status == "Blocked" and steps["Emails"].blocked
    for slug in ("search", "emails", "approvals", "order-sheet", "listing-content"):
        assert "Not available yet" in env.client.get(f"/step/{slug}").text
    assert env.client.get("/step/bogus").status_code == 404


@test
def running_banner_and_job_pages(env):
    m = env.jobs.start("stub.state", {"mode": "slow"})
    page = env.client.get("/").text
    assert "A job is running" in page and "Cancel" in page
    assert m["id"] in env.client.get("/jobs").text
    log = env.client.get(f"/jobs/{m['id']}/log?offset=0").json()
    assert log["status"] == "running"
    assert env.client.get("/jobs/..%2f..%2fx").status_code in (404, 400)
    assert env.client.get("/jobs/not-an-id").status_code == 404
    env.jobs.cancel(m["id"])
    env.jobs.wait(m["id"])


@test
def loopback_only_startup(env):
    import run_ui
    saved = {k: os.environ.get(k) for k in ("UI_HOST", "UI_ALLOW_NON_LOOPBACK", "UI_DATA_DIR")}
    os.environ["UI_DATA_DIR"] = str(env.root / "data")         # create_app makes this folder: keep it out of the project
    try:
        calls = []                                                  # uvicorn is faked so a broken guard cannot start a server
        sys.modules["uvicorn"], real = SimpleNamespace(run=lambda *a, **k: calls.append(k)), sys.modules.get("uvicorn")
        os.environ.pop("UI_ALLOW_NON_LOOPBACK", None)
        for h in ("0.0.0.0", "192.168.1.5", "example.com", "::"):
            os.environ["UI_HOST"] = h
            raises(SystemExit, run_ui.main)
        assert not calls, "server must not start on a non-loopback host"
        os.environ["UI_HOST"] = "127.0.0.1"
        run_ui.main()
        assert calls and calls[0]["host"] == "127.0.0.1", calls
    finally:
        if real is not None:
            sys.modules["uvicorn"] = real
        for k, v in saved.items():
            os.environ.pop(k, None) if v is None else os.environ.__setitem__(k, v)

# ================= Stage 1.5: wording, labels, contrast, Home counts =================

import colorsys
from datetime import datetime

from ui import wording

TERM_VARIANTS = ["L-Com", "LCom", "lcom", "l com", "L-COM", "L_Com", "lcom-"]
STEM_A, STEM_B = "sourcing_results_20260101_100000", "sourcing_results_20260202_100000"


def has_term(text: str) -> bool:
    return wording.contains_reference_name(text)


def add_results(env, name, rows, mtime=None):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Results"
    ws.append(["SKU", "Product", "Recommended Manufacturer", "Recommended URL"])
    for r in rows:
        ws.append(r)
    path = env.project / "Excel Output Sheets" / name
    wb.save(path)
    if mtime:
        os.utime(path, (mtime, mtime))
    return path


def urls_in(html: str) -> list:
    return re.findall(r'(?:href|src|action)="([^"]*)"', html)


@test
def wording_table(env):
    cases = {"L-Com price": "Reference price", "L-Com Prices": "Reference prices", "LCom unit price": "Reference unit price",
             "L-Com SKU": "SKU", "lcom_prices.csv": "reference_prices.csv", "lcom-sourcing_results_20261004_161054": "sourcing_results_20261004_161054",
             "l com": "reference", "L_COM": "reference", "the L-Com catalogue": "the reference catalogue",
             "Telecom": "Telecom", "Lcommerce": "Lcommerce", "plain text 12.00": "plain text 12.00"}
    for raw, want in cases.items():
        assert wording.display(raw) == want, (raw, wording.display(raw))
    assert wording.display(None) is None and wording.display(5) == 5
    assert wording.strip_prefix("lcom-abc") == "abc" and wording.strip_prefix("abc") == "abc"
    from markupsafe import Markup
    assert isinstance(wording.display(Markup("<b>L-Com</b>")), Markup)       # already-rendered markup is left alone


@test
def reference_name_never_reaches_a_page_title_or_url(env):
    p = env.project
    (p / "Order Sheets").mkdir()
    (p / "Order Sheets" / "order_sheet_LCom_1.xlsx").write_bytes(b"x")
    (p / "Walmart Listings").mkdir()
    (p / "Walmart Listings" / "walmart_listings_L-Com.xlsx").write_bytes(b"x")
    (p / "approved_orders.json").write_text(json.dumps({"source_results_file": STEM_B + ".xlsx", "valid_approvals": 1, "generated_at": "L-Com time"}))
    add_results(env, "sourcing_results_20260303_100000.xlsx", [["L-Com SKU 1", "LCom price cable", "L-Com Maker", "http://x/lcom"]])
    job = env.jobs.wait(env.jobs.start("stub.read", {"mode": "term"})["id"])
    assert "L-Com" in env.jobs.read_log(job["id"])[0], "the stored log must stay verbatim"
    pages = ["/", "/jobs", "/health", f"/jobs/{job['id']}", f"/jobs/{job['id']}/log?offset=0", "/jobs/not-an-id", "/nope", "/step/bogus",
             "/?wf=nonsense"] + [f"/step/{s}" for s in ("search", "emails", "approvals", "order-sheet", "listing-content")]
    for w in workflows.list_workflows(env.settings):
        pages += [f"/?wf={w['stem']}", f"/step/order-sheet?wf={w['stem']}", f"/step/listing-content?wf={w['stem']}"]
    for page in pages:
        r = env.client.get(page, follow_redirects=True)
        body = r.text
        title = re.search(r"<title>(.*?)</title>", body, re.S)
        assert not has_term(body), (page, [m for m in re.findall(r"(?i).{20}l[\s_-]?com.{20}", body)][:3])
        assert not title or not has_term(title[1])
        for u in urls_in(body):
            assert not has_term(u), (page, u)
    assert "Reference price 12.00" in env.client.get(f"/jobs/{job['id']}").text
    for variant in TERM_VARIANTS:
        assert not has_term(wording.display(f"x {variant} y")) or variant == "lcom-", variant


@test
def plain_stem_urls_and_old_prefixed_urls_redirect(env):
    html = env.client.get("/").text
    assert f"?wf={STEM_B}" in html and "lcom" not in html.lower()
    for old, new in [(f"/?wf=lcom-{STEM_A}", f"/?wf={STEM_A}"), (f"/step/search?wf=LCOM-{STEM_A}", f"/step/search?wf={STEM_A}")]:
        r = env.client.get(old, follow_redirects=False)
        assert r.status_code == 303 and r.headers["location"] == new, (old, r.status_code, r.headers.get("location"))
    page = env.client.get(f"/?wf={STEM_A}").text
    assert f'value="{STEM_A}" selected' in page
    assert workflows.internal_id(STEM_A) == "lcom-" + STEM_A == obs.workflow_id(STEM_A + ".xlsx")      # the internal id is unchanged


@test
def friendly_labels_recent_five_and_show_all(env):
    assert workflows.friendly_label(datetime(2026, 10, 4, 16, 10), 15) == "Oct 4, 4:10 PM, 15 products"
    assert workflows.friendly_label(datetime(2026, 10, 4, 0, 5), 1) == "Oct 4, 12:05 AM, 1 product"
    assert workflows.friendly_label(datetime(2026, 10, 4, 12, 0), None) == "Oct 4, 12:00 PM"
    base = datetime(2026, 3, 1, 9, 30).timestamp()
    for i in range(5):                                          # 2 existing + 5 = 7 workflows
        add_results(env, f"sourcing_results_2026030{i}_090000.xlsx", [["A", "p", "m", "u"]] * (i + 1), base + i * 86400 + 99999999)
    wfs = workflows.list_workflows(env.settings)
    assert len(wfs) == 7 and all("lcom" not in w["label"].lower() and "_" not in w["label"] for w in wfs)
    assert wfs[0]["label"].endswith("5 products") and len({w["label"] for w in wfs}) == 7
    html = env.client.get("/").text
    recent = html.split("Recent workflows")[1].split("<details>")[0]
    assert recent.count("<li>") == 5 and "Show all (7)" in html
    for f in list((env.project / "Excel Output Sheets").glob("sourcing_results_2026030*.xlsx")):
        f.unlink()
    for f in list((env.project / "Excel Output Sheets").glob("*.xlsx"))[2:]:
        f.unlink()
    assert "Show all" not in env.client.get("/").text                      # 2 workflows: no cut-off, no details


@test
def home_count_separates_carrying_a_recommendation_from_being_draftable(env):
    for f in (env.project / "Excel Output Sheets").glob("*.xlsx"):
        f.unlink()
    add_results(env, "sourcing_results_20260404_100000.xlsx",
                [["A1", "p", "Acme", "http://a/1"], ["B2", "p", "Beta", "http://b/2"], ["C3", "p", "", ""], ["D4", "p", "", ""]])
    (env.project / "reviewer_exclusions.csv").write_text("sku,supplier_or_url,reason\nB2,beta,not acceptable\n", encoding="utf-8")
    text = env.client.get("/").text
    assert "4 products" in text and "1 can be drafted" in text and "2 have no recommended seller" in text, text
    assert "1 are excluded by the reviewer" in text and "unverified" in text
    assert "7 with a recommended seller" not in text and "recommended sellers" not in text
    d = workflows.draft_counts(env.settings, env.project / "Excel Output Sheets" / "sourcing_results_20260404_100000.xlsx")
    assert d == {"draftable": 1, "no_seller": 2, "excluded": 1, "errors": 0}, d
    # when the agent's rules cannot be run, only what the file says is shown, labelled precisely
    real = workflows.draft_counts
    workflows.draft_counts = lambda s, p: None
    try:
        text = env.client.get("/").text
    finally:
        workflows.draft_counts = real
    assert "rows carry a recommendation; the Emails step shows how many can be drafted" in text and "can be drafted;" not in text


@test
def displayed_log_has_a_note_wording_and_whole_lines_only(env):
    job = env.jobs.wait(env.jobs.start("stub.read", {"mode": "term"})["id"])
    page = env.client.get(f"/jobs/{job['id']}").text
    assert page.count("Wording differs from the raw log file.") == 1 and not has_term(page)
    assert "Reference price 12.00" in page and "reference_prices.csv" in page
    assert "L-Com price 12.00" in env.jobs.log_path(job["id"]).read_text(encoding="utf-8"), "the stored log is never altered"
    log = env.client.get(f"/jobs/{job['id']}/log?offset=0").json()
    assert not has_term(log["text"])
    f = env.settings.jobs / "20260101-000000-aaaaaa.log"
    f.write_text("first line\nsecond half-writ", encoding="utf-8")
    assert env.jobs.read_log("20260101-000000-aaaaaa", 0, whole_lines_only=True) == ("first line\n", 11)
    assert env.jobs.read_log("20260101-000000-aaaaaa", 11, whole_lines_only=True) == ("", 11)
    assert env.jobs.read_log("20260101-000000-aaaaaa", 0)[0].endswith("half-writ")


@test
def app_name_is_configurable_and_neutral_by_default(env):
    html = env.client.get("/").text
    assert "<title>Workflow - Sourcing Tool</title>" in html and 'class="brand" href="/">Sourcing Tool<' in html
    other = Settings(project=env.project, data=env.project / "d3", app_name="Acme Sourcing")
    c = TestClient(create_app(other), base_url="http://127.0.0.1:8000")
    html = c.get("/health").text
    assert "<title>Health - Acme Sourcing</title>" in html and "Sourcing Tool" not in html
    saved = os.environ.get("UI_APP_NAME")
    os.environ["UI_APP_NAME"] = "From Env"
    try:
        assert load_settings().app_name == "From Env"
    finally:
        os.environ.pop("UI_APP_NAME", None) if saved is None else os.environ.__setitem__("UI_APP_NAME", saved)
    assert load_settings().app_name == "Sourcing Tool"
    assert "Sourcing Tool" not in "".join(p.read_text(encoding="utf-8") for p in (Path(__file__).parent / "ui" / "templates").glob("*.html"))


def parse_tokens():
    css = (Path(__file__).parent / "ui" / "static" / "app.css").read_text(encoding="utf-8")
    light_block = re.search(r":root\s*\{(.*?)\}", css, re.S)[1]
    dark_block = re.search(r"prefers-color-scheme:\s*dark\)\s*\{\s*:root\s*\{(.*?)\}", css, re.S)[1]
    grab = lambda block: dict(re.findall(r"--([\w-]+):\s*(#[0-9a-fA-F]{6})", block))
    light = grab(light_block)
    return light, {**light, **grab(dark_block)}


def luminance(hexcolor: str) -> float:
    rgb = [int(hexcolor[i:i + 2], 16) / 255 for i in (1, 3, 5)]
    lin = [c / 12.92 if c <= 0.03928 else ((c + 0.055) / 1.055) ** 2.4 for c in rgb]
    return 0.2126 * lin[0] + 0.7152 * lin[1] + 0.0722 * lin[2]


def contrast(a: str, b: str) -> float:
    hi, lo = sorted((luminance(a), luminance(b)), reverse=True)
    return (hi + 0.05) / (lo + 0.05)


TEXT_PAIRS = [("text", "bg"), ("text", "surface"), ("muted", "bg"), ("muted", "surface"), ("accent", "bg"), ("accent", "surface"),
              ("accent-contrast", "accent"), ("danger-contrast", "danger"), ("success", "surface"), ("warning", "surface"),
              ("danger", "surface"), ("info", "surface"), ("success", "bg"), ("warning", "bg"), ("danger", "bg"), ("info", "bg")]
UI_PAIRS = [("border", "surface"), ("border", "bg"), ("success", "surface"), ("warning", "surface"), ("danger", "surface"),
            ("info", "surface"), ("accent", "surface")]


@test
def colour_tokens_meet_wcag_contrast_in_both_themes(env):
    light, dark = parse_tokens()
    for theme, tokens in (("light", light), ("dark", dark)):
        for fg, bg in TEXT_PAIRS:
            assert contrast(tokens[fg], tokens[bg]) >= 4.5, (theme, fg, bg, round(contrast(tokens[fg], tokens[bg]), 2))
        for fg, bg in UI_PAIRS:
            assert contrast(tokens[fg], tokens[bg]) >= 3.0, (theme, fg, bg, round(contrast(tokens[fg], tokens[bg]), 2))
    css = (Path(__file__).parent / "ui" / "static" / "app.css").read_text(encoding="utf-8")
    body = css.split("* { box-sizing")[1]                          # everything after the token block uses var(), no literal colours
    assert not re.search(r"#[0-9a-fA-F]{3,6}\b", body) and not re.search(r"\brgba?\(", body), "components must use tokens only"
    assert "http" not in css and "@import" not in css and "url(" not in css


@test
def every_status_has_a_distinct_badge_with_text(env):
    css = (Path(__file__).parent / "ui" / "static" / "app.css").read_text(encoding="utf-8")
    kinds = ["not-started", "files-found", "blocked", "running", "done", "failed", "cancelled", "interrupted"]
    markers = []
    for k in kinds:
        m = re.search(rf"\.badge-{k}::before[^{{]*\{{\s*content:\s*\"([^\"]+)\"", css)
        assert m, k
        markers.append(m[1])
    assert len(set(markers)) == len(kinds), markers
    from ui.app import JOB_BADGES
    assert set(JOB_BADGES.values()) <= {"Running", "Done", "Failed", "Cancelled", "Interrupted"}
    for j in ("ok", "fail"):
        env.jobs.wait(env.jobs.start("stub.read", {"mode": j})["id"])
    html = env.client.get("/jobs").text
    assert "badge-done" in html and ">Done<" in html and ">Failed<" in html and "No jobs yet" not in html


def main():
    failed = []
    for f in TESTS:
        env = make_env()
        try:
            f(env)
            print(f"ok   {f.__name__}")
        except Exception:  # noqa: BLE001
            failed.append(f.__name__)
            print(f"FAIL {f.__name__}\n{traceback.format_exc()}")
        finally:
            cleanup(env)
    print(f"\n{len(TESTS) - len(failed)}/{len(TESTS)} passed")
    if failed:
        sys.exit(1)
    print("ui tests passed")


if __name__ == "__main__":
    main()
