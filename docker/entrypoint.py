"""Container start-up for the sourcing tool: the image's entrypoint, run before the UI.

It checks the data folder mounted at /data, puts this version's program files into it, and starts the UI with that folder
as the project. The agents keep their state beside their own script, so they have to run from the data folder; that is why
their files are copied there at every start. Nothing else in the data folder is opened, moved or deleted.
It prints one line per action and never prints the value of a setting."""
import contextlib
import os
import re
import secrets
import sys
from pathlib import Path

APP = Path(__file__).resolve().parent.parent            # /app in the image
if str(APP) not in sys.path:
    sys.path.insert(0, str(APP))

from ui.config import PROGRAM_FILES, REFERENCE_FILES    # noqa: E402  (one list, shared with the Health page)

DATA = Path("/data")
ZONES = Path("/usr/share/zoneinfo")
SOURCE_MARKS = (".git", "ui", "run_ui.py", "Dockerfile")    # any of these in the data folder means it is the source folder
PROXY_EXCEPTIONS = ("localhost", "127.0.0.1", "host.docker.internal")    # reached directly, never through a proxy
PROXY_NAMES = ("HTTP_PROXY", "HTTPS_PROXY", "ALL_PROXY", "http_proxy", "https_proxy", "all_proxy")


class Refused(Exception):
    """The container must not start. The message says why and what to do."""


def is_mounted(path: Path) -> bool:
    """Whether a folder from outside the container is mounted exactly here."""
    try:
        with open("/proc/self/mountinfo", encoding="utf-8") as f:
            return any(len(parts := line.split()) > 4 and parts[4] == str(path) for line in f)
    except OSError:
        return os.path.ismount(path)


def check_zone(name, zones: Path = ZONES) -> None:
    """The pages show local time, and an unknown zone name silently means UTC, so a missing or misspelt TZ stops the start.
    The compose file passes TZ through as it is (empty when .env has none): this is the one place that refuses."""
    if not name or not name.strip():
        raise Refused("TZ is not set. Add your time zone to .env, for example TZ=Asia/Kolkata, then start again with "
                      "`docker compose up -d`. The pages show local time, so the tool does not start without it.")
    if not re.fullmatch(r"[A-Za-z0-9_+\-]+(/[A-Za-z0-9_+\-]+){0,2}", name) or not (zones / name).is_file():
        raise Refused("TZ in .env is not a time zone name this system knows. Set it to a name such as Asia/Kolkata or "
                      "America/Chicago, then start again with `docker compose up -d`.")


def check_data(data: Path, mounted=is_mounted) -> None:
    if not data.is_dir() or not mounted(data):
        raise Refused("no data folder is mounted at /data. Start the tool with `docker compose up -d`, which mounts one. "
                      "Without it, your work would be lost when the container is removed.")
    marks = [m for m in SOURCE_MARKS if (data / m).exists()]
    if marks:
        raise Refused(f"the folder mounted at /data looks like the source folder (it holds {', '.join(marks)}), not a data folder. "
                      "Starting would overwrite the program files in it. Mount an empty folder or your own data folder.")
    probe = data / f".lcom_write_test_{secrets.token_hex(4)}"
    try:
        probe.write_bytes(b"")
        probe.unlink()
    except OSError:
        raise Refused("the data folder is not writable by this container.")


def _put(src: Path, dst: Path) -> None:
    """Temp file beside the target, then a rename: a start that is cut off never leaves half a file under the real name."""
    tmp = dst.with_name(f"{dst.name}.{secrets.token_hex(3)}.tmp")
    try:
        tmp.write_bytes(src.read_bytes())
        os.replace(tmp, dst)
    except BaseException:
        with contextlib.suppress(OSError):
            tmp.unlink()
        raise


def _same_text(a: bytes, b: bytes) -> bool:
    return a.replace(b"\r\n", b"\n") == b.replace(b"\r\n", b"\n")       # line endings aside, as the Health page compares


def install(app: Path, data: Path, say=print) -> None:
    """Program files are replaced when they differ from this version's. A reference file is added only when the data folder
    has none; a person's own copy is kept. No other file is touched."""
    for name in PROGRAM_FILES:
        if not (app / name).is_file():
            raise Refused(f"this image is incomplete: {name} is missing from it.")
    for name in PROGRAM_FILES + REFERENCE_FILES:                         # checked for every file before anything is written
        dst = data / name
        if dst.is_symlink() or (dst.exists() and not dst.is_file()):
            raise Refused(f"{name} in the data folder is not an ordinary file. Move it away and start again.")
    for name in PROGRAM_FILES:
        src, dst = app / name, data / name
        if not dst.exists() or dst.read_bytes() != src.read_bytes():
            _put(src, dst)
            say(f"program file written: {name}")
    for name in REFERENCE_FILES:
        src, dst = app / name, data / name
        if not src.is_file():
            continue
        if not dst.exists():
            _put(src, dst)
            say(f"reference file added: {name}")
        elif not _same_text(dst.read_bytes(), src.read_bytes()):
            say(f"NOTE: your {name} differs from the one in this version. Yours is kept and used. The Health page lists these files.")


def merge_no_proxy(environ) -> tuple:
    """A proxy set for the internet must not catch this container's own addresses (the UI itself, a Langfuse on this
    computer). Both spellings of NO_PROXY get one list: every entry the person set, in their order (NO_PROXY's first),
    then the names in PROXY_EXCEPTIONS that are missing, each once. Returns (entries kept, names added): counts only,
    because the entries are the person's own settings."""
    kept, seen = [], set()
    for name in ("NO_PROXY", "no_proxy"):
        for entry in (environ.get(name) or "").split(","):
            entry = entry.strip()
            if entry and entry.lower() not in seen:
                kept.append(entry)
                seen.add(entry.lower())
    added = [h for h in PROXY_EXCEPTIONS if h not in seen]
    environ["NO_PROXY"] = environ["no_proxy"] = ",".join(kept + added)
    return len(kept), len(added)


def main(data: Path = DATA, mounted=is_mounted, zones: Path = ZONES, start=None) -> None:
    try:
        check_zone(os.environ.get("TZ"), zones)
        check_data(data, mounted)
        install(APP, data)
    except Refused as e:
        sys.exit(f"NOT STARTED: {e}")
    kept, added = merge_no_proxy(os.environ)
    proxies = sorted({n.upper() for n in PROXY_NAMES if os.environ.get(n)})      # names only: a proxy address can hold a password
    print(f"proxy settings present: {', '.join(proxies) or 'none'}; reached directly: {kept} name(s) from your settings "
          f"and {added} added ({', '.join(PROXY_EXCEPTIONS)})", flush=True)
    # These five belong to the container, whatever the settings file says: where the data is, and how the UI listens.
    os.environ.update(PROJECT_DIR=str(data), UI_DATA_DIR=str(data / "ui_data"), UI_SHIPPED_DIR=str(APP), UI_HOST="0.0.0.0",
                      UI_ALLOW_NON_LOOPBACK="1")
    os.environ.pop("UI_ENABLE_STUB", None)                               # the test stub is not in this image
    print("data folder ready, starting the UI", flush=True)
    if start is None:      # the UI runs as a fresh process started with exactly this environment (so /proc shows what it really has)
        os.execv(sys.executable, [sys.executable, str(APP / "run_ui.py")])
    else:
        start()


if __name__ == "__main__":
    main()
