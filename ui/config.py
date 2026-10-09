"""Folders and settings. PROJECT_DIR (default: the folder holding the agents) and UI_DATA_DIR are configurable so a
later container can mount them."""
import os
from dataclasses import dataclass
from pathlib import Path

LOOPBACK = ("127.0.0.1", "localhost", "::1")

# Folders the UI may serve files from, and the only extensions it will serve. State files (json) are never served.
OUTPUT_FOLDERS = ("Excel Output Sheets", "Reports", "Order Sheets", "Walmart Listings")
DOWNLOAD_EXTENSIONS = (".xlsx", ".csv", ".md", ".txt")

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

    @property
    def jobs(self) -> Path:
        return self.data / "jobs"

    @property
    def results_dir(self) -> Path:
        return self.project / "Excel Output Sheets"


def load_settings(**over) -> Settings:
    project = Path(os.environ.get("PROJECT_DIR") or Path(__file__).resolve().parent.parent).resolve()
    data = Path(os.environ.get("UI_DATA_DIR") or project / "ui_data").resolve()
    base = dict(project=project, data=data, host=os.environ.get("UI_HOST", "127.0.0.1"),
                port=int(os.environ.get("UI_PORT", "8000")), enable_stub=os.environ.get("UI_ENABLE_STUB") == "1",
                app_name=os.environ.get("UI_APP_NAME") or "Sourcing Tool")
    base.update(over)
    return Settings(**base)


def is_loopback(host: str) -> bool:
    return host.strip("[]") in LOOPBACK
