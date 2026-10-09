"""Job runner: runs allowlisted commands as subprocesses (argument list, shell=False, stdin closed), streams masked
output to ui_data/jobs/<id>.log, keeps metadata in <id>.json (atomic writes), allows one state-changing job at a time."""
import json
import os
import re
import secrets
import subprocess
import sys
import threading
import time
from datetime import datetime
from pathlib import Path
from typing import Optional

import observability as obs   # mask_text: keys, emails, phone numbers, shipping address, protected values

from .commands import Command, JobRejected, build_allowlist
from .config import Settings

JOB_ID_RE = re.compile(r"\d{8}-\d{6}-[0-9a-f]{6}")
RUNNING, OK, FAILED, CANCELLED, INTERRUPTED = "running", "succeeded", "failed", "cancelled", "interrupted"


class JobBusy(Exception):
    """A state-changing job is already running; the request is refused, not queued."""


# observability's own rule needs a word boundary before 'api_key', so NAME_API_KEY=value slips past it; this catches it.
_NAMED_SECRET = re.compile(r"(?i)([A-Za-z0-9_]*(?:api_?key|secret|password|token)\s*[=:]\s*)(?:Bearer\s+)?\S+")


def mask_line(line: str) -> str:
    return _NAMED_SECRET.sub(lambda m: m[1] + "[REDACTED]", obs.mask_text(line))


def write_json_atomic(path: Path, data: dict) -> None:
    tmp = path.with_name(path.name + f".{secrets.token_hex(3)}.tmp")
    tmp.write_text(json.dumps(data, indent=2), encoding="utf-8")
    for attempt in range(40):
        try:
            os.replace(tmp, path)
            return
        except PermissionError:          # Windows: the target is open in a reader for a moment (a polling page)
            if attempt == 39:
                raise
            time.sleep(0.025)


class JobManager:
    def __init__(self, settings: Settings):
        self.settings = settings
        self.commands = build_allowlist(settings)
        self.jobs_dir = settings.jobs
        self.jobs_dir.mkdir(parents=True, exist_ok=True)
        self._lock = threading.Lock()
        self._state_job: Optional[str] = None      # id of the running state-changing job
        self._procs: dict = {}                     # id -> Popen
        self._cancelled: set = set()
        self.recover()

    # ---- metadata ----
    def _meta_path(self, jid: str) -> Path:
        return self.jobs_dir / f"{jid}.json"

    def log_path(self, jid: str) -> Path:
        return self.jobs_dir / f"{jid}.log"

    def get(self, jid: str) -> Optional[dict]:
        if not JOB_ID_RE.fullmatch(jid or ""):
            return None
        try:
            return json.loads(self._meta_path(jid).read_text(encoding="utf-8"))
        except (OSError, ValueError):
            return None

    def _save(self, meta: dict) -> None:
        write_json_atomic(self._meta_path(meta["id"]), meta)

    def list(self) -> list:
        metas = [m for p in self.jobs_dir.glob("*.json") if (m := self.get(p.stem))]
        return sorted(metas, key=lambda m: m["started_at"], reverse=True)

    def recover(self) -> None:
        """A job left 'running' by a previous process is marked interrupted. Never resumed."""
        for m in self.list():
            if m["status"] == RUNNING:
                m.update(status=INTERRUPTED, finished_at=datetime.now().isoformat(timespec="seconds"),
                         note="The UI stopped while this job was running. Check what state it may have left before running it again.")
                self._save(m)

    def running_state_job(self) -> Optional[dict]:
        with self._lock:
            jid = self._state_job
        return self.get(jid) if jid else None

    # ---- start ----
    def start(self, cmd_id: str, params: Optional[dict] = None) -> dict:
        cmd: Optional[Command] = self.commands.get(cmd_id)
        if cmd is None:
            raise JobRejected("That command is not allowed.")
        args = cmd.argv(params or {})                   # validates; raises JobRejected
        script = cmd.script if os.path.isabs(cmd.script) else str(self.settings.project / cmd.script)
        argv = [sys.executable, "-X", "utf8", script, *args]
        jid = f"{datetime.now():%Y%m%d-%H%M%S}-{secrets.token_hex(3)}"
        with self._lock:                                # check and claim the state-changing slot atomically
            if cmd.state_changing:
                if self._state_job is not None:
                    raise JobBusy("Another job that changes data is running. Wait for it to finish or cancel it.")
                self._state_job = jid
        meta = {"id": jid, "command": cmd.id, "label": cmd.label, "args": args, "state_changing": cmd.state_changing,
                "status": RUNNING, "started_at": datetime.now().isoformat(timespec="seconds"), "finished_at": None,
                "duration_s": None, "exit_code": None, "exit_meaning": "", "pid": None, "note": "", "tail": []}
        try:
            self._save(meta)
            self.log_path(jid).write_text("", encoding="utf-8")
            threading.Thread(target=self._run, args=(cmd, argv, meta), daemon=True).start()
        except BaseException:
            self._release(jid)
            raise
        return meta

    def _release(self, jid: str) -> None:
        with self._lock:
            if self._state_job == jid:
                self._state_job = None

    def _run(self, cmd: Command, argv: list, meta: dict) -> None:
        jid, t0 = meta["id"], time.monotonic()
        try:
            with self._lock:
                if cmd.state_changing and self._state_job != jid:   # checked again at the moment of starting
                    raise JobBusy("lost the state-changing slot")
            env = {**os.environ, "PYTHONIOENCODING": "utf-8", "PYTHONUNBUFFERED": "1"}
            flags = (subprocess.CREATE_NO_WINDOW | subprocess.CREATE_NEW_PROCESS_GROUP) if os.name == "nt" else 0
            proc = subprocess.Popen(argv, cwd=self.settings.project, env=env, shell=False, stdin=subprocess.DEVNULL,
                                    stdout=subprocess.PIPE, stderr=subprocess.STDOUT, creationflags=flags,
                                    encoding="utf-8", errors="replace")
            self._procs[jid] = proc
            meta["pid"] = proc.pid
            self._save(meta)
            with open(self.log_path(jid), "a", encoding="utf-8") as log:
                for line in proc.stdout:
                    log.write(mask_line(line.rstrip("\r\n")) + "\n")
                    log.flush()
            code = proc.wait()
            cancelled = jid in self._cancelled
            meta.update(exit_code=code, status=CANCELLED if cancelled else (OK if code == 0 else FAILED))
            meta["exit_meaning"] = "Stopped by you." if cancelled else cmd.exits.get(code, f"Exit code {code} (not one of the agent's documented codes).")
            if cancelled:
                meta["note"] = cmd.cancel_warning or "This command was stopped part-way. Check its output files before running it again."
        except Exception as e:   # noqa: BLE001 - never leave a job 'running' because the runner itself failed
            meta.update(status=FAILED, note=mask_line(f"The runner could not run this command: {type(e).__name__}"))
            with open(self.log_path(jid), "a", encoding="utf-8") as log:
                log.write(mask_line(f"runner error: {type(e).__name__}: {e}") + "\n")
        finally:
            self._procs.pop(jid, None)
            meta.update(finished_at=datetime.now().isoformat(timespec="seconds"), duration_s=round(time.monotonic() - t0, 1))
            if meta["status"] in (FAILED,):
                meta["tail"] = self.tail(jid, 15)
            self._release(jid)          # before the final save, so a viewer who sees 'finished' can already start the next job
            self._save(meta)

    # ---- cancel / log ----
    def cancel(self, jid: str) -> bool:
        proc = self._procs.get(jid)
        if proc is None or proc.poll() is not None:
            return False
        self._cancelled.add(jid)
        if os.name == "nt":
            subprocess.run(["taskkill", "/PID", str(proc.pid), "/T", "/F"], stdin=subprocess.DEVNULL,
                           stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, check=False)
        else:
            proc.kill()
        return True

    def read_log(self, jid: str, offset: int = 0, whole_lines_only: bool = False) -> tuple:
        """(text from offset, new offset). Offsets are characters of the already-masked log. With whole_lines_only (a job that
        is still running) a half-written last line is held back, so display wording is never applied to a cut-off word."""
        if not JOB_ID_RE.fullmatch(jid or ""):
            return "", offset
        try:
            text = self.log_path(jid).read_text(encoding="utf-8", errors="replace")
        except OSError:
            return "", offset
        if whole_lines_only:
            cut = text.rfind("\n") + 1
            text = text[:max(cut, offset)]
        return text[offset:], len(text)

    def tail(self, jid: str, n: int = 15) -> list:
        text, _ = self.read_log(jid)
        return text.splitlines()[-n:]

    def wait(self, jid: str, timeout: float = 30) -> dict:
        """For tests: block until the job is no longer running."""
        end = time.monotonic() + timeout
        while time.monotonic() < end:
            m = self.get(jid)
            if m and m["status"] != RUNNING:
                return m
            time.sleep(0.05)
        raise TimeoutError(jid)
