"""Offline tests for the agent gate changes: decision_agent record flags, email_agent send --confirm-count, the shared
state-file lock/atomic write, and the observability masking pattern. No network, fake SMTP, no live calls.
Run:  $env:OBS_ENABLED = "0"; python test_gates.py"""
import builtins
import contextlib
import importlib.util
import io
import json
import os
import shutil
import subprocess
import sys
import tempfile
import time
import traceback
from types import SimpleNamespace

os.environ["OBS_ENABLED"] = "0"
HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)

import decision_agent as da
import email_agent as ea
import observability as obs
import statefile
import test_decision_agent as td
import test_email_agent as te

TESTS = []


def test(f):
    TESTS.append(f)
    return f


def raises(exc, fn, *a, **kw):
    try:
        fn(*a, **kw)
    except exc as e:
        return e
    raise AssertionError(f"expected {exc.__name__}")


@contextlib.contextmanager
def no_console():
    """Fails on any attempt to open the Windows console or a tty device, through open() or os.open()."""
    real_open, real_os_open, hits = builtins.open, os.open, []

    def bad(p):
        return str(p).strip().upper().rstrip(":") in ("CON", "CONIN$", "CONOUT$", "/DEV/TTY") or str(p).upper().endswith(("\\CON", "/DEV/TTY"))

    def guard_open(file, *a, **k):
        if isinstance(file, (str, bytes, os.PathLike)) and bad(os.fsdecode(file)):
            hits.append(file)
            raise AssertionError(f"console opened: {file!r}")
        return real_open(file, *a, **k)

    def guard_os_open(path, *a, **k):
        if bad(os.fsdecode(path)):
            hits.append(path)
            raise AssertionError(f"console opened: {path!r}")
        return real_os_open(path, *a, **k)
    builtins.open, os.open = guard_open, guard_os_open
    try:
        yield hits
    finally:
        builtins.open, os.open = real_open, real_os_open
    assert not hits, hits


REPLY = "approve VIC00001 qty 100 max 1.30\nreject HDFF\ncap $200\n"


def reply_file(tmp):
    f = os.path.join(tmp, "reply.txt")
    open(f, "w", encoding="utf-8").write(REPLY)
    return f


def flags(**over):
    base = {"--confirm-request": "A-0001", "--channel": "Slack DM", "--approver": "Neeraj K."}
    base.update(over)
    return [x for k, v in base.items() if v is not None for x in (k, v)]


def snapshot():
    return open(da.STATE_PATH, "rb").read(), open(da.LOG_PATH, "rb").read() if os.path.exists(da.LOG_PATH) else b""


# ================= Change 1: decision_agent record =================

@test
def record_with_all_three_flags_records_once_and_matches_interactive():
    with no_console(), td.sandbox() as tmp:
        path, _ = td.request_for(td.standard_rows(), tmp)
        f = reply_file(tmp)
        td.da.ask_user = lambda p: (_ for _ in ()).throw(AssertionError("prompted: " + p))   # the flag path must never ask
        out = td.run_cli(["record", "--results", path, "--request", "A-0001", "--file", f, *flags()])
        assert "Recorded. Nothing has been ordered." in out
        flag_state, flag_log = td.state(), td.log()
        assert [e["event"] for e in flag_log].count("reply_recorded") == 1
        assert sum(d["decision"] == "approve" for d in flag_state["decisions"]) == 1 and len(flag_state["decisions"]) == 2
        r = flag_state["requests"]["A-0001"]["reply"]
        assert r["channel"] == "Slack DM" and r["approver"] == "Neeraj K." and r["sender_verified"] is False
    with td.sandbox() as tmp2:                                       # the same thing typed at the prompts
        path2, _ = td.request_for(td.standard_rows(), tmp2)
        td.record(path2, REPLY, channel="Slack DM", approver="Neeraj K.")
        assert td.state()["decisions"] == flag_state["decisions"] and td.state()["requests"]["A-0001"]["reply"] == r
        assert [e["event"] for e in td.log()] == [e["event"] for e in flag_log]
        for a, b in zip(td.log(), flag_log):
            a, b = dict(a), dict(b)
            for k in ("results_sha256",):
                a.pop(k, None), b.pop(k, None)
            assert a == b, (a, b)


@test
def mismatched_confirm_request_refuses_and_changes_nothing():
    with td.sandbox() as tmp:
        path, _ = td.request_for(td.standard_rows(), tmp)
        f = reply_file(tmp)
        before = snapshot()
        for argv in (["--request", "A-0001", *flags(**{"--confirm-request": "A-0002"})],
                     ["--request", "A-0001", *flags(**{"--confirm-request": "A-1"})],
                     ["--request", "A-0001", *flags(**{"--confirm-request": "yes"})],
                     flags()):                                         # no --request at all
            msg = td.cli_fails(["record", "--results", path, "--file", f, *argv])
            assert "Nothing was recorded" in msg, msg
            assert snapshot() == before
        assert "is not the same as --request" in td.cli_fails(["record", "--results", path, "--file", f, "--request", "A-0001", *flags(**{"--confirm-request": "A-0002"})])


@test
def partial_flags_refuse_and_name_the_missing_ones():
    with td.sandbox() as tmp:
        path, _ = td.request_for(td.standard_rows(), tmp)
        f = reply_file(tmp)
        before = snapshot()
        for given in (["--channel"], ["--approver"], ["--confirm-request"], ["--channel", "--approver"], ["--confirm-request", "--channel"]):
            argv = [x for k in given for x in (k, {"--channel": "Slack", "--approver": "Neeraj", "--confirm-request": "A-0001"}[k])]
            msg = td.cli_fails(["record", "--results", path, "--request", "A-0001", "--file", f, *argv])
            missing = {"--channel", "--approver", "--confirm-request"} - set(given)
            assert all(m in msg for m in missing) and not any(g in msg.split("missing")[0] for g in given), (given, msg)
            assert "give --confirm-request, --channel and --approver together" in msg
            assert snapshot() == before
        assert td.state()["decisions"] == []


@test
def empty_or_whitespace_values_refuse():
    with td.sandbox() as tmp:
        path, _ = td.request_for(td.standard_rows(), tmp)
        f = reply_file(tmp)
        before = snapshot()
        for bad in ("", "   ", "\t", "Neeraj\nK", "A\rB"):
            for flag in ("--approver", "--channel"):
                msg = td.cli_fails(["record", "--results", path, "--request", "A-0001", "--file", f, *flags(**{flag: bad})])
                assert "non-empty single line" in msg and snapshot() == before, (flag, bad, msg)
        env_has = {"USERNAME": os.environ.get("USERNAME")}                       # no default from the environment
        msg = td.cli_fails(["record", "--results", path, "--request", "A-0001", "--file", f, "--confirm-request", "A-0001", "--channel", "x"])
        assert "--approver" in msg and snapshot() == before and env_has


@test
def closed_stdin_without_flags_errors_writes_nothing_and_never_opens_console():
    with no_console(), td.sandbox() as tmp:
        path, _ = td.request_for(td.standard_rows(), tmp)
        f = reply_file(tmp)
        da.ask_user = da.read_console                                             # the real prompt function
        before = snapshot()
        msg = td.cli_fails(["record", "--results", path, "--request", "A-0001", "--file", f])      # stdin: a BytesIO pipe, not a tty
        assert msg == da.NO_TERMINAL and snapshot() == before
        old = sys.stdin
        try:
            for fake in (None, SimpleNamespace(closed=True, isatty=lambda: False),
                         SimpleNamespace(closed=False, isatty=lambda: True, readline=lambda *a: "")):   # a tty that is used up (EOF)
                sys.stdin = fake
                err = raises(da.AgentError, da.cmd_record, path, REPLY, "A-0001")
                assert str(err) == da.NO_TERMINAL
        finally:
            sys.stdin = old
        assert snapshot() == before and td.state()["decisions"] == []
    src = open(os.path.join(HERE, "decision_agent.py"), encoding="utf-8").read()
    assert '"CON"' not in src and "/dev/tty" not in src


@test
def interactive_path_still_works_and_trace_hides_approver_and_channel():
    import test_observability as tob
    with td.sandbox() as tmp, tob.tracing("decision_agent") as run:
        path = os.path.join(tmp, tob.RESULTS)
        td.make_results(path, td.standard_rows())
        td.run_cli(["request", path])
        f = reply_file(tmp)
        td.run_cli(["record", "--results", path, "--request", "A-0001", "--file", f, *flags(**{"--channel": "SecretChannelXyz", "--approver": "Zed Quillfeather"})])
        assert td.state()["requests"]["A-0001"]["reply"]["approver"] == "Zed Quillfeather"
        td.answers("A-0001", "SecretChannelXyz", "Zed Quillfeather", "yes")
        td.run_cli(["request", path])
        td.run_cli(["record", "--results", path, "--request", "A-0002", "--file", f])               # interactive, scripted prompts
    text = tob.never_leaks(run, "SecretChannelXyz", "Zed Quillfeather", "Quillfeather")
    assert "decision_agent.record" in text


# ================= Change 2: email_agent send --confirm-count =================

def two_drafts(tmp):
    path = os.path.join(tmp, "r.xlsx")
    te.make_xlsx(path, [te.row("A1", email="a@x.com"), te.row("B2", email="b@y.com"), te.row("C3", email="", url="http://form")])
    return path


@test
def confirm_count_wrong_sends_nothing_right_sends_exactly_n():
    with te.sandbox() as tmp:
        path = two_drafts(tmp)
        for n in ("0", "1", "3", "5"):
            e = raises(SystemExit, te.run_cli, ["send", path, "--confirm-count", n])
            assert "Refusing to send: --confirm-count is " + n + " but 2 email(s) would be sent" in str(e.code), e.code
            assert te.FakeSMTP.created == 0 and not os.path.exists(ea.STATE_PATH)
        out = te.run_cli(["send", path, "--confirm-count", "2"])                  # stdin is an empty pipe: nothing may be asked
        assert "Sent 2 email(s)" in out and te.FakeSMTP.created == 2
        assert sorted(m["To"] for m in te.FakeSMTP.sent) == ["a@x.com", "b@y.com"]
        assert sum("sent" in v for v in json.load(open(ea.STATE_PATH)).values()) == 2
        out = te.run_cli(["send", path, "--confirm-count", "0"])                  # both already sent: the count after skips is 0
        assert "Sent 0 email(s)" in out and te.FakeSMTP.created == 2
        e = raises(SystemExit, te.run_cli, ["send", path, "--confirm-count", "2"])
        assert "but 0 email(s) would be sent" in str(e.code) and te.FakeSMTP.created == 2


@test
def confirm_count_does_not_bypass_any_gate_and_test_sends_are_not_recorded():
    with te.sandbox() as tmp:
        path = two_drafts(tmp)
        os.environ.pop("SHIPPING_ADDRESS")
        e = raises(SystemExit, te.run_cli, ["send", path, "--confirm-count", "2"])
        assert "Refusing to send: SHIPPING_ADDRESS is not set" in str(e.code) and te.FakeSMTP.created == 0
        os.environ["SHIPPING_ADDRESS"] = "Zync Technologies\\n100 Test Way\\nPlano, TX 75024"
        os.environ.pop("SMTP_PASSWORD")
        e = raises(SystemExit, te.run_cli, ["send", path, "--confirm-count", "2"])
        assert "Set SMTP_PASSWORD in .env first" in str(e.code) and te.FakeSMTP.created == 0
        os.environ["SMTP_PASSWORD"] = "x"
        e = raises(SystemExit, te.run_cli, ["send", path, "--confirm-count", "2", "--test-recipient", "a@x.com"])   # a seller's address
        assert "must be a test address" in str(e.code) and te.FakeSMTP.created == 0
        out = te.run_cli(["send", path, "--confirm-count", "3", "--test-recipient", "me@test.example"])      # test mode counts every draft
        assert "Sent 3 email(s)" in out and {m["To"] for m in te.FakeSMTP.sent} == {"me@test.example"}
        assert not os.path.exists(ea.STATE_PATH)                                                            # never recorded as sent
        e = raises(SystemExit, te.run_cli, ["send", path, "--confirm-count", "2", "--test-recipient", "me@test.example"])
        assert "but 3 email(s) would be sent" in str(e.code)
        assert raises(SystemExit, te.run_cli, ["send", path, "--near-miss", "--confirm-count", "2"]).code == 2   # drafts-only flag
        assert raises(SystemExit, te.run_cli, ["send", path, "--confirm-count", "-1"]).code == 2
        assert raises(SystemExit, te.run_cli, ["send", path, "--confirm-count", "two"]).code == 2


def copy_agents(tmp, files):
    for f in files:
        shutil.copy(os.path.join(HERE, f), tmp)


@test
def closed_stdin_without_the_flag_is_a_clean_exit_1_not_a_traceback():
    with tempfile.TemporaryDirectory() as tmp:
        copy_agents(tmp, ("email_agent.py", "observability.py", "statefile.py", "email_template.txt"))
        te.make_xlsx(os.path.join(tmp, "r.xlsx"), [te.row("A1", email="a@x.com"), te.row("B2", email="b@x.com")])
        env = {k: v for k, v in os.environ.items() if k not in ("ANTHROPIC_API_KEY",)}
        env.update(SMTP_HOST="127.0.0.1", SMTP_PORT="1", SMTP_USER="u@z.example", SMTP_PASSWORD="x",
                   SHIPPING_ADDRESS="Zync Technologies\\n100 Test Way\\nPlano, TX 75024", OBS_ENABLED="0")
        run = lambda *a: subprocess.run([sys.executable, "-X", "utf8", "email_agent.py", *a], cwd=tmp, stdin=subprocess.DEVNULL,
                                        text=True, capture_output=True, env=env, timeout=60)
        for extra in ([], ["--yes"]):
            p = run("send", "r.xlsx", *extra)
            assert p.returncode == 1 and "Traceback" not in p.stderr and "no terminal to confirm on" in p.stderr, (p.returncode, p.stderr)
        p = run("send", "r.xlsx", "--confirm-count", "9")                                    # wrong count: exit 1, nothing sent
        assert p.returncode == 1 and "Refusing to send: --confirm-count is 9 but 2" in p.stderr
        assert not os.path.exists(os.path.join(tmp, "email_state.json"))


# ================= Change 3: state files =================

WORKER = '''
import json, sys
sys.path.insert(0, sys.argv[1])
import statefile
path, n, who = sys.argv[2], int(sys.argv[3]), sys.argv[4]
for i in range(n):
    with statefile.locked(path):
        try:
            data = json.load(open(path))
        except FileNotFoundError:
            data = {}
        data[who + str(i)] = i
        statefile.atomic_write(path, json.dumps(data))
'''


@test
def two_processes_racing_never_lose_an_update():
    with tempfile.TemporaryDirectory() as tmp:
        w = os.path.join(tmp, "worker.py")
        open(w, "w").write(WORKER)
        path = os.path.join(tmp, "state.json")
        env = dict(os.environ, STATE_LOCK_WAIT_SECONDS="60")
        procs = [subprocess.Popen([sys.executable, "-X", "utf8", w, HERE, path, "25", who], env=env, stderr=subprocess.PIPE, text=True) for who in ("a", "b", "c")]
        for p in procs:
            assert p.wait(timeout=120) == 0, p.stderr.read()
        assert len(json.load(open(path))) == 75, "an update was lost"
        assert sorted(os.listdir(tmp)) == ["state.json", "worker.py"], os.listdir(tmp)         # no lock or temp file left behind


@test
def atomic_write_keeps_the_old_file_when_it_fails():
    with tempfile.TemporaryDirectory() as tmp:
        path = os.path.join(tmp, "s.json")
        statefile.atomic_write(path, "old")
        raises(TypeError, statefile.atomic_write, path, ["not", "text"])                       # fails half way, inside the write
        assert open(path).read() == "old" and os.listdir(tmp) == ["s.json"]
        real, calls = os.replace, []

        def flaky(a, b):
            calls.append(1)
            if len(calls) < 3:
                raise PermissionError("in use")
            return real(a, b)
        statefile.os.replace = flaky
        try:
            statefile.atomic_write(path, "new")                                                  # Windows: retried until it works
            assert open(path).read() == "new" and len(calls) == 3
            statefile.os.replace = lambda a, b: (_ for _ in ()).throw(PermissionError("still in use"))
            raises(PermissionError, statefile.atomic_write, path, "newer")
        finally:
            statefile.os.replace = real
        assert open(path).read() == "new" and os.listdir(tmp) == ["s.json"]
        statefile.atomic_write(path, "a\nb", newline="")
        assert open(path, "rb").read() == b"a\nb"                                               # newline="" writes the text as given


def dead_pid():
    p = subprocess.Popen([sys.executable, "-c", "pass"])
    p.wait()
    return p.pid


def write_lock(path, pid, age_s):
    json.dump({"pid": pid, "at": time.time() - age_s, "nonce": "other"}, open(path + ".lock", "w"))


@test
def lock_is_exclusive_respects_live_takes_over_stale_and_always_releases():
    with tempfile.TemporaryDirectory() as tmp:
        path = os.path.join(tmp, "s.json")
        with statefile.locked(path):
            assert os.path.exists(path + ".lock")
            e = raises(statefile.StateBusy, statefile.locked(path, wait_s=0.2).__enter__)
            assert "s.json is busy" in str(e)
        assert not os.path.exists(path + ".lock")
        raises(RuntimeError, _boom, path)
        assert not os.path.exists(path + ".lock"), "lock must be released after an exception"
        write_lock(path, os.getpid(), 3600)                                       # old, but its process is alive: respected
        raises(statefile.StateBusy, statefile.locked(path, wait_s=0.2).__enter__)
        write_lock(path, dead_pid(), 60)                                          # dead process but not yet 15 minutes: respected
        raises(statefile.StateBusy, statefile.locked(path, wait_s=0.2).__enter__)
        write_lock(path, dead_pid(), 3600)                                        # dead and old: taken over, one line logged
        err = io.StringIO()
        with contextlib.redirect_stderr(err), statefile.locked(path, wait_s=0.5):
            assert json.load(open(path + ".lock"))["pid"] == os.getpid()
        assert err.getvalue().count("taking over a stale lock") == 1 and not os.path.exists(path + ".lock")
        assert statefile.pid_running(os.getpid()) and not statefile.pid_running(dead_pid())


def _boom(path):
    with statefile.locked(path):
        raise RuntimeError("boom")


@test
def busy_state_gives_exit_code_3_and_read_only_commands_do_not_lock():
    with td.sandbox() as tmp:
        path, _ = td.request_for(td.standard_rows(), tmp)
        os.environ["STATE_LOCK_WAIT_SECONDS"] = "0.3"
        try:
            with statefile.locked(da.STATE_PATH):
                for argv in (["request", path], ["revoke", "VIC00001", "--results", path], ["export", "--results", path]):
                    e = raises(SystemExit, td.run_cli, argv)
                    assert e.code == 3, (argv, e.code)
                assert raises(SystemExit, da.main, ["record", "--results", path, "--request", "A-0001", "--file", reply_file(tmp), *flags()]).code == 3
                assert "VIC00001" in td.run_cli(["status", "--results", path])             # read-only: runs while the lock is held
        finally:
            del os.environ["STATE_LOCK_WAIT_SECONDS"]
        assert not os.path.exists(da.STATE_PATH + ".lock")
    with te.sandbox() as tmp:
        path = two_drafts(tmp)
        os.environ["STATE_LOCK_WAIT_SECONDS"] = "0.3"
        try:
            with statefile.locked(ea.STATE_PATH):
                for argv in (["mark", "C3", "--results", path], ["send", path, "--confirm-count", "2"]):
                    assert raises(SystemExit, te.run_cli, argv).code == 3, argv
                assert te.FakeSMTP.created == 0
                assert "3 draft(s)" in te.run_cli(["drafts", path])
                assert "Wrote" in te.run_cli(["report", path])
        finally:
            del os.environ["STATE_LOCK_WAIT_SECONDS"]


@test
def exit_code_3_from_a_real_process_and_lock_released_by_commands():
    with tempfile.TemporaryDirectory() as tmp:
        copy_agents(tmp, ("decision_agent.py", "email_agent.py", "observability.py", "statefile.py", "email_template.txt"))
        td.make_results(os.path.join(tmp, "r.xlsx"), td.standard_rows())
        env = dict(os.environ, OBS_ENABLED="0", STATE_LOCK_WAIT_SECONDS="0.5")
        run = lambda *a: subprocess.run([sys.executable, "-X", "utf8", *a], cwd=tmp, stdin=subprocess.DEVNULL, text=True, capture_output=True, env=env, timeout=60)
        assert run("decision_agent.py", "request", "r.xlsx").returncode == 0
        assert not os.path.exists(os.path.join(tmp, "decision_state.json.lock"))               # released after a normal command
        with statefile.locked(os.path.join(tmp, "decision_state.json")):
            p = run("decision_agent.py", "revoke", "VIC00001", "--results", "r.xlsx")
            assert p.returncode == 3 and "is busy" in p.stderr, (p.returncode, p.stderr)
            assert run("decision_agent.py", "status", "--results", "r.xlsx").returncode == 0
        with statefile.locked(os.path.join(tmp, "email_state.json")):
            p = run("email_agent.py", "mark", "VIC00001", "--results", "r.xlsx")
            assert p.returncode == 3 and "is busy" in p.stderr, (p.returncode, p.stderr)
        # no terminal, no flags: exit 1, nothing written; with all three flags and DEVNULL stdin: records
        open(os.path.join(tmp, "reply.txt"), "w").write(REPLY)
        before = open(os.path.join(tmp, "decision_state.json"), "rb").read()
        p = run("decision_agent.py", "record", "--results", "r.xlsx", "--request", "A-0001", "--file", "reply.txt")
        assert p.returncode == 1 and "no terminal" in p.stderr and open(os.path.join(tmp, "decision_state.json"), "rb").read() == before, p.stderr
        p = run("decision_agent.py", "record", "--results", "r.xlsx", "--request", "A-0001", "--file", "reply.txt", *flags())
        assert p.returncode == 0 and "Recorded." in p.stdout, (p.stdout, p.stderr)
        assert len(open(os.path.join(tmp, "approvals_log.jsonl")).read().splitlines()) == 1 + 1 + 2     # request, reply_recorded, 2 decisions


@test
def export_writes_are_atomic():
    with td.sandbox() as tmp:
        path, _ = td.request_for(td.standard_rows(), tmp)
        td.record(path, REPLY)
        calls, real = [], statefile.atomic_write
        statefile.atomic_write = lambda p, d, newline=None: (calls.append(os.path.basename(p)), real(p, d, newline))[1]
        try:
            td.run_cli(["export", "--results", path])
        finally:
            statefile.atomic_write = real
        assert calls == ["approved_orders.json", "approved_orders.csv"], calls
        good = (open(da.EXPORT_JSON, "rb").read(), open(da.EXPORT_CSV, "rb").read())
        real_replace = os.replace
        statefile.os.replace = lambda a, b: (_ for _ in ()).throw(OSError("disk full")) if str(b).endswith(".csv") else real_replace(a, b)
        try:
            raises(OSError, da.cmd_export, path)
        finally:
            statefile.os.replace = real_replace
        assert open(da.EXPORT_CSV, "rb").read() == good[1], "a failed write must leave the old csv intact"
        assert not [f for f in os.listdir(tmp) if f.endswith(".tmp")], os.listdir(tmp)


# ================= Change 4: masking =================

NAMES = ["ANTHROPIC_API_KEY", "NIMBLE_API_KEY", "SMTP_PASSWORD", "LANGFUSE_SECRET_KEY", "LANGFUSE_PUBLIC_KEY", "SHIPPING_ADDRESS",
         "api_key", "secret", "password", "token", "x-api-key"]
BARE = ["Zq7xKv2mP", "hunter2-pass", "p@ss/w0rd!#9", "A" * 30 + "9", "0123456789", "aB3-_.~Qz"]
FORMS = ["{n}={v}", "{n}: {v}", "{n} = {v}", '"{n}": "{v}"', "'{n}': '{v}'", "{n}='{v}'", 'error while running: {n}={v} (retrying)',
         'failed with {{"{n}": "{v}", "other": 1}}', r'\"{n}\": \"{v}\"']


@test
def masking_covers_every_name_value_shape_and_form():
    for n in NAMES:
        for v in BARE:
            for form in FORMS:
                s = form.format(n=n, v=v)
                out = obs.mask_text(s)
                assert v not in out, (s, out)
                assert obs.mask_exc(RuntimeError(s)).count(v) == 0, s
        spaced = f'{n}="two words 9 Fake Rd\\nPlano"'
        assert "Fake Rd" not in obs.mask_text(spaced) and "two words" not in obs.mask_text(spaced), spaced
        assert "Fake Rd" not in obs.mask_text(f'"{n}": "two words 9 Fake Rd"')
    d = obs.mask({"detail": "ANTHROPIC_API_KEY=abc123xyz789", "list": ["NIMBLE_API_KEY: n1m613xyz"]})
    assert "abc123xyz789" not in str(d) and "n1m613xyz" not in str(d)
    from opentelemetry.sdk.trace import Event
    span = SimpleNamespace(_events=[Event("exception", {"exception.type": "RuntimeError", "exception.message": "SMTP_PASSWORD=Zq7xKv2mP then"}, 1)], _status=None)
    obs._masking_processor().on_end(span)
    assert "Zq7xKv2mP" not in str(span._events[0].attributes), span._events[0].attributes


@test
def masking_leaves_ordinary_prose_alone():
    for s in ["the API key is missing", "Missing required environment variable(s): ANTHROPIC_API_KEY, NIMBLE_API_KEY",
              "Set SMTP_HOST, SMTP_USER, SMTP_PASSWORD in .env first (nothing sent).", "Tokens: in=1200 out=300  (~$0.0012)",
              "the password was wrong", "token count is 5", "no secret here", "Add your key: see the README", "Stopped: model stopped with 1 candidates",
              "authorization was not requested", "SHIPPING_ADDRESS is not set (or is still the default)", "2026-10-08 12:00 done"]:
        assert obs.mask_text(s) == s, (s, obs.mask_text(s))


# ================= disabled-mode guarantee =================

def load_backup(name):
    spec = importlib.util.spec_from_file_location("pre_" + name, os.path.join(HERE, "scratch", "pre_gates", name + ".py"))
    m = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(m)
    return m


def run_decision_flow(mod, tmp, results_bytes):
    names = {"STATE_PATH": "decision_state.json", "LOG_PATH": "approvals_log.jsonl", "EXPORT_JSON": "approved_orders.json",
             "EXPORT_CSV": "approved_orders.csv", "REPORTS_DIR": "Reports", "EMAIL_STATE_PATH": "email_state.json",
             "REVIEWER_EXCLUSIONS_CSV": "reviewer_exclusions.csv"}
    for n, f in names.items():
        setattr(mod, n, os.path.join(tmp, f))
    open(mod.REVIEWER_EXCLUSIONS_CSV, "w", newline="").write("sku,supplier_or_url,reason\r\n")
    mod.now = lambda: td.T0
    path = os.path.join(tmp, "r.xlsx")
    open(path, "wb").write(results_bytes)
    ans = iter(["A-0001", "Slack DM", "Neeraj", "yes"])
    mod.ask_user = lambda p: next(ans)
    out = io.StringIO()
    with contextlib.redirect_stdout(out):
        mod.cmd_request(path)
        mod.cmd_record(path, REPLY, None)
        mod.cmd_export(path)
        mod.cmd_revoke("VIC00001", path)
    def norm(raw: bytes) -> bytes:                                   # the two runs use different temp folders: only that may differ
        return raw.replace(tmp.replace("\\", "\\\\").encode(), b"<TMP>").replace(tmp.encode(), b"<TMP>")
    files = {f: norm(open(os.path.join(tmp, f), "rb").read()) for f in ("decision_state.json", "approvals_log.jsonl", "approved_orders.json", "approved_orders.csv")}
    files.update({"Reports/" + f: norm(open(os.path.join(tmp, "Reports", f), "rb").read()) for f in sorted(os.listdir(os.path.join(tmp, "Reports")))})
    return out.getvalue().replace(tmp, "<TMP>"), files


@test
def with_no_new_flags_outputs_and_state_files_are_byte_identical_to_the_backups():
    old_da = load_backup("decision_agent")
    with tempfile.TemporaryDirectory() as seed:
        src = os.path.join(seed, "r.xlsx")
        td.make_results(src, td.standard_rows())
        data = open(src, "rb").read()
    saved = {n: getattr(da, n) for n in ("STATE_PATH", "LOG_PATH", "EXPORT_JSON", "EXPORT_CSV", "REPORTS_DIR", "EMAIL_STATE_PATH", "REVIEWER_EXCLUSIONS_CSV", "now", "ask_user")}
    try:
        with tempfile.TemporaryDirectory() as a, tempfile.TemporaryDirectory() as b:
            out_old, files_old = run_decision_flow(old_da, a, data)
            out_new, files_new = run_decision_flow(da, b, data)
    finally:
        for n, v in saved.items():
            setattr(da, n, v)
    assert out_old == out_new, "decision_agent console output differs"
    assert files_old.keys() == files_new.keys()
    for k in files_old:
        assert files_old[k] == files_new[k], f"{k} differs"
    # email_agent: the state file bytes (including a non-ASCII reply) and a mark run
    old_ea = load_backup("email_agent")
    state = {"k|1": {"sku": "A1", "seller": "Acme", "sent": "2026-10-05 12:00", "reply": "你好 café \"quoted\"", "replies": [{"at": "x", "text": "你好"}]}}
    outs, real_path = [], ea.STATE_PATH
    for mod in (old_ea, ea):
        with tempfile.TemporaryDirectory() as t:
            mod.STATE_PATH = os.path.join(t, "email_state.json")
            mod.save_state(state)
            outs.append(open(mod.STATE_PATH, "rb").read())
    ea.STATE_PATH = real_path
    assert outs[0] == outs[1], "email_state.json bytes differ"


def main():
    failed = []
    for f in TESTS:
        try:
            f()
            print(f"ok   {f.__name__}")
        except Exception:  # noqa: BLE001
            failed.append(f.__name__)
            print(f"FAIL {f.__name__}\n{traceback.format_exc()}")
    print(f"\n{len(TESTS) - len(failed)}/{len(TESTS)} passed")
    if failed:
        sys.exit(1)
    print("gate tests passed")


if __name__ == "__main__":
    main()
