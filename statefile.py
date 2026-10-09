"""Shared by email_agent.py and decision_agent.py: atomic writes and a lock file for read-modify-write of state files.

    with statefile.locked(STATE_PATH):          # exclusive: waits up to STATE_LOCK_WAIT_SECONDS (default 10), else StateBusy
        state = load(); ...; statefile.atomic_write(STATE_PATH, json.dumps(state, indent=2))

The agents turn StateBusy into exit code 3 ("state file busy"). Read-only commands never take the lock."""
import contextlib
import json
import os
import secrets
import sys
import time

STALE_AFTER_S = 15 * 60                  # a lock this old whose process is gone may be taken over
DEFAULT_WAIT_S = 10.0


class StateBusy(Exception):
    """Another command holds the lock on this state file."""


def atomic_write(path: str, data: str, newline=None) -> None:
    """Temp file in the same folder, flush + fsync, then os.replace. A failure leaves the old file intact and no temp file
    behind. `newline` is passed to open(): None keeps the platform's text-mode behaviour, "" writes `data` as it is."""
    tmp = f"{path}.{os.getpid()}.{secrets.token_hex(3)}.tmp"
    try:
        with open(tmp, "w", encoding="utf-8", newline=newline) as f:
            f.write(data)
            f.flush()
            os.fsync(f.fileno())
        for attempt in range(20):
            try:
                os.replace(tmp, path)
                return
            except PermissionError:      # Windows: another process has the target open for a moment
                if attempt == 19:
                    raise
                time.sleep(0.05)
    except BaseException:
        with contextlib.suppress(OSError):
            os.remove(tmp)
        raise


def pid_running(pid: int) -> bool:
    if os.name == "nt":                  # never os.kill here: on Windows it terminates the process
        import ctypes
        handle = ctypes.windll.kernel32.OpenProcess(0x1000, False, int(pid))    # PROCESS_QUERY_LIMITED_INFORMATION
        if not handle:
            return False
        code = ctypes.c_ulong()
        ok = ctypes.windll.kernel32.GetExitCodeProcess(handle, ctypes.byref(code))
        ctypes.windll.kernel32.CloseHandle(handle)
        return bool(ok) and code.value == 259                                   # STILL_ACTIVE
    try:
        os.kill(pid, 0)
    except ProcessLookupError:
        return False
    except PermissionError:
        return True
    return True


def _remove(path: str) -> None:
    """Delete, retrying: on Windows a file another process has open (a waiter reading the lock) cannot be deleted for a moment."""
    for attempt in range(100):
        try:
            os.remove(path)
            return
        except FileNotFoundError:
            return
        except PermissionError:
            if attempt == 99:
                raise
            time.sleep(0.02)


def _read_lock(lock: str):
    try:
        with open(lock, encoding="utf-8") as f:
            return json.load(f)
    except (OSError, ValueError):
        return None


def _is_stale(info) -> bool:
    if not isinstance(info, dict):
        return False                     # unreadable or half-written: treat as held, it will resolve or time out
    try:
        return time.time() - float(info["at"]) > STALE_AFTER_S and not pid_running(int(info["pid"]))
    except (KeyError, TypeError, ValueError):
        return False


def _create(lock: str, nonce: str) -> bool:
    try:
        fd = os.open(lock, os.O_CREAT | os.O_EXCL | os.O_WRONLY)
    except FileExistsError:
        return False
    with os.fdopen(fd, "w", encoding="utf-8") as f:
        json.dump({"pid": os.getpid(), "at": time.time(), "nonce": nonce}, f)
        f.flush()
        os.fsync(f.fileno())
    return True


@contextlib.contextmanager
def locked(path: str, wait_s: float = None):
    lock, nonce = path + ".lock", secrets.token_hex(8)
    if wait_s is None:
        try:
            wait_s = float(os.environ.get("STATE_LOCK_WAIT_SECONDS", DEFAULT_WAIT_S))
        except ValueError:
            wait_s = DEFAULT_WAIT_S
    deadline = time.monotonic() + wait_s
    while not _create(lock, nonce):
        info = _read_lock(lock)
        if _is_stale(info) and _take_over(lock, info, nonce):
            break
        if time.monotonic() >= deadline:
            who = f" (process {info['pid']}, since {time.strftime('%H:%M:%S', time.localtime(info['at']))})" if isinstance(info, dict) and "pid" in info else ""
            raise StateBusy(f"{os.path.basename(path)} is busy: another command is using it{who}. Wait for it to finish and try again.")
        time.sleep(0.05)
    try:
        yield
    finally:
        info = _read_lock(lock)
        if isinstance(info, dict) and info.get("nonce") == nonce:       # never remove a lock somebody else holds now
            _remove(lock)


def _take_over(lock: str, seen, nonce: str) -> bool:
    """Replace a stale lock. A guard file makes sure only one process does this; the lock is re-read under the guard."""
    guard = lock + ".takeover"
    try:
        os.close(os.open(guard, os.O_CREAT | os.O_EXCL | os.O_WRONLY))
    except FileExistsError:
        with contextlib.suppress(OSError):                               # a crashed taker-over: the guard goes stale too
            if time.time() - os.path.getmtime(guard) > 60:
                os.remove(guard)
        return False
    try:
        if _read_lock(lock) != seen:                                     # somebody else already replaced it
            return False
        print(f"statefile: taking over a stale lock {os.path.basename(lock)} (process {seen['pid']} is gone).", file=sys.stderr)
        _remove(lock)
        return _create(lock, nonce)
    finally:
        _remove(guard)
