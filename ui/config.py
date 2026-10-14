"""Folders and settings. PROJECT_DIR (default: the folder holding the agents) and UI_DATA_DIR are configurable so a
later container can mount them."""
import os
from dataclasses import dataclass
from pathlib import Path
from typing import Optional

LOOPBACK = ("127.0.0.1", "localhost", "::1")

# Folders the UI may serve files from, and the only extensions it will serve. State files (json) are never served.
OUTPUT_FOLDERS = ("Excel Output Sheets", "Reports", "Order Sheets", "Walmart Listings", "Uploads")
DOWNLOAD_EXTENSIONS = (".xlsx", ".csv", ".md", ".txt")

# Uploads: product sheets (and, later, the facts file) a person hands to the UI. Names are always built by the server.
UPLOADS_FOLDER = "Uploads"
MAX_UPLOAD_BYTES = 5 * 1024 * 1024
MAX_PRODUCT_ROWS = 200          # the one cap on products per search: rows in an uploaded sheet and the largest --limit
MAX_FACTS_BYTES = 1024 * 1024   # a filled facts file (.csv)

# What a container start (docker/entrypoint.py) puts into the data folder. Program files are replaced at every start;
# a reference file is copied only when the data folder has none, so a person's own copy is kept.
PROGRAM_FILES = ("sourcing_agent.py", "email_agent.py", "decision_agent.py", "order_sheet.py", "content_agent.py",
                 "observability.py", "statefile.py")
REFERENCE_FILES = ("email_template.txt", "reviewer_exclusions.csv", "near_misses.csv", "lcom_prices.csv")

# Variable NAME -> what a missing setting blocks, in plain language. Values are never read by the pages, only "is it set".
ENV_VARS = (
    ("ANTHROPIC_API_KEY", "Searching, reading seller replies, and writing listing drafts"),
    ("NIMBLE_API_KEY", "Searching for suppliers"),
    ("SMTP_HOST", "Sending emails to sellers"),
    ("SMTP_USER", "Sending emails to sellers"),
    ("SMTP_PASSWORD", "Sending emails to sellers"),
    ("SHIPPING_ADDRESS", "Sending emails to sellers (sending is refused while it is unset or still the placeholder)"),
    ("BRAND_NAME", "Writing listing drafts"),
    ("SMTP_PORT", "Optional: mail server port, 587 if unset"),
    ("SMTP_FROM", "Optional: the From address, the mail login if unset"),
    ("SENDER_NAME", "Optional: the name signed on seller emails"),
    ("ORDER_SHEET_RECALC_PY", "Optional: an extra check of order sheet totals"),
    ("OBS_ENABLED", "Optional: turns run tracing on when set to 1"),
    ("LANGFUSE_PUBLIC_KEY", "Run tracing, only when it is on"),
    ("LANGFUSE_SECRET_KEY", "Run tracing, only when it is on"),
    ("LANGFUSE_BASE_URL", "Run tracing, only when it is on"),
)


@dataclass(frozen=True)
class Settings:
    project: Path
    data: Path
    host: str = "127.0.0.1"
    port: int = 8000
    enable_stub: bool = False   # the stub agent exists only for tests
    app_name: str = "Sourcing Tool"
    retention_days: int = 30    # finished job logs older than this are pruned; 0 disables
    shipped: Optional[Path] = None   # container only: the folder holding the reference files that came with this version

    @property
    def jobs(self) -> Path:
        return self.data / "jobs"

    @property
    def results_dir(self) -> Path:
        return self.project / "Excel Output Sheets"

    @property
    def uploads_dir(self) -> Path:
        return self.project / UPLOADS_FOLDER


def load_settings(**over) -> Settings:
    project = Path(os.environ.get("PROJECT_DIR") or Path(__file__).resolve().parent.parent).resolve()
    data = Path(os.environ.get("UI_DATA_DIR") or project / "ui_data").resolve()
    base = dict(project=project, data=data, host=os.environ.get("UI_HOST", "127.0.0.1"),
                port=int(os.environ.get("UI_PORT", "8000")), enable_stub=os.environ.get("UI_ENABLE_STUB") == "1",
                app_name=os.environ.get("UI_APP_NAME") or "Sourcing Tool", retention_days=_retention_days(),
                shipped=Path(os.environ["UI_SHIPPED_DIR"]).resolve() if os.environ.get("UI_SHIPPED_DIR") else None)
    base.update(over)
    return Settings(**base)


def changed_reference_files(s: Settings) -> list:
    """Names of the reference files in the project folder that differ from the copies that came with this version
    (line endings aside). Names only: the two files are compared, their contents are never returned. Empty when no
    shipped folder is set, which is every run outside a container."""
    if not s.shipped:
        return []
    out = []
    for name in REFERENCE_FILES:
        try:
            mine, theirs = (s.project / name).read_bytes(), (s.shipped / name).read_bytes()
        except OSError:
            continue                    # one of the two is missing: nothing to compare
        if mine.replace(b"\r\n", b"\n") != theirs.replace(b"\r\n", b"\n"):
            out.append(name)
    return out


def _retention_days() -> int:
    try:
        return max(int(os.environ.get("UI_LOG_RETENTION_DAYS", "30")), 0)
    except ValueError:
        return 30


def is_set(s: Settings, name: str) -> bool:
    """Whether a setting has a value, in the environment or in the project's .env (the agents load that file themselves).
    Only yes or no ever leaves this function; the value is never returned, logged or shown."""
    if os.environ.get(name):
        return True
    try:
        from dotenv import dotenv_values
        return bool(dotenv_values(s.project / ".env").get(name))
    except Exception:  # noqa: BLE001 - an unreadable .env counts as not set
        return False


def is_loopback(host: str) -> bool:
    return host.strip("[]") in LOOPBACK
