"""The allowlist: the only commands the UI can run. Stage 1 holds read-only `--help` commands for the five agents and,
only when settings.enable_stub is on, a stub agent for tests. Real search, send, approval and order commands arrive in
Stage 2 together with their human gates."""
import re
from dataclasses import dataclass, field
from pathlib import Path
from typing import Callable, Optional

from .config import Settings


class JobRejected(Exception):
    """The request is not an allowed command, or an argument failed its rule. Safe to show to the user."""


SKU_RE = re.compile(r"[A-Za-z0-9][A-Za-z0-9._&-]{0,39}")   # no leading '-', no spaces, no path characters
STUB_MODES = ("ok", "fail", "slow", "secret", "hostile", "term")

# What an exit code means, from the agents' own behaviour (Stage 0): sys.exit("message") and sys.exit(1) both give 1.
GENERIC_EXITS = {0: "Finished normally.",
                 1: "The agent stopped with an error message (see the end of the log). Nothing further was done.",
                 2: "The command line was rejected by the agent (bad or missing argument)."}


def sku_rule(settings: Settings) -> Callable[[str], str]:
    def check(v: str) -> str:
        if not isinstance(v, str) or not SKU_RE.fullmatch(v):
            raise JobRejected(f"'{str(v)[:40]}' is not a valid SKU (letters, digits, . _ & - only, up to 40 characters).")
        return v
    return check


def results_file_rule(settings: Settings) -> Callable[[str], str]:
    """A results file is picked from the folder listing by exact name; it is never accepted as a typed path."""
    def check(v: str) -> str:
        names = {p.name for p in settings.results_dir.glob("*.xlsx")} if settings.results_dir.is_dir() else set()
        if v not in names:
            raise JobRejected("That results file is not in 'Excel Output Sheets'. Choose one from the list.")
        return str(Path("Excel Output Sheets") / v)
    return check


def enum_rule(options) -> Callable[[str], str]:
    def check(v: str) -> str:
        if v not in options:
            raise JobRejected(f"'{str(v)[:40]}' is not one of: {', '.join(options)}.")
        return v
    return check


@dataclass(frozen=True)
class Param:
    flag: Optional[str]          # None = positional
    rule: Callable[[str], str]
    required: bool = True


@dataclass(frozen=True)
class Command:
    id: str
    label: str
    script: str                  # path relative to the project folder, or absolute (stub)
    fixed: tuple = ()
    params: dict = field(default_factory=dict)
    state_changing: bool = False
    cancel_warning: str = ""
    exits: dict = field(default_factory=lambda: dict(GENERIC_EXITS))

    def argv(self, given: dict) -> list:
        unknown = set(given) - set(self.params)
        if unknown:
            raise JobRejected(f"Not allowed for this command: {', '.join(sorted(unknown))}.")
        out = list(self.fixed)
        for name, p in self.params.items():
            if name not in given:
                if p.required:
                    raise JobRejected(f"Missing: {name}.")
                continue
            value = p.rule(given[name])
            out += [value] if p.flag is None else [p.flag, value]
        return out


AGENTS = ("sourcing_agent", "email_agent", "decision_agent", "order_sheet", "content_agent")


def build_allowlist(settings: Settings) -> dict:
    cmds = {}
    for a in AGENTS:
        cmds[f"help.{a}"] = Command(f"help.{a}", f"{a}.py --help", f"{a}.py", ("--help",))
    if settings.enable_stub:
        stub = str(Path(__file__).with_name("stub_agent.py"))
        mode = Param(None, enum_rule(STUB_MODES))
        opt = dict(mode=mode, sku=Param("--sku", sku_rule(settings), False),
                   results=Param("--results", results_file_rule(settings), False))
        warn = "The stub agent was stopped part-way; it keeps no state."
        cmds["stub.read"] = Command("stub.read", "stub (read-only)", stub, (), opt, False, warn)
        cmds["stub.state"] = Command("stub.state", "stub (changes state)", stub, ("--changes-state",), opt, True, warn)
    return cmds


# Shown after a cancel, per agent command (from the Stage 0 reading). Stage 2 commands will look these up.
CANCEL_WARNINGS = {
    "email_agent send": "Some emails may already have been sent. email_state.json is written after each send, so check it before sending again.",
    "email_agent reply": "The reply may not have been stored. Check email_state.json before pasting it again.",
    "decision_agent record": "The approval may or may not have been stored. Run 'status' before recording again.",
    "decision_agent export": "approved_orders.json and .csv may be half-written. Run 'export' again before building an order sheet.",
    "order_sheet build": "Only a temporary '.verifying.xlsx' file can be left behind; a finished sheet is only created after checks pass.",
    "sourcing_agent search": "The results file holds every product finished so far (it is saved after each product); the rest were not searched.",
    "content_agent generate": "The listing file may be incomplete. Re-run generate.",
}
