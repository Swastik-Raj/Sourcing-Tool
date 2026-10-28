"""The allowlist: the only commands the UI can run. It holds read-only `--help` commands for the five agents, the search
(ui/search.py, Stage 2a) and, only when settings.enable_stub is on, a stub agent for tests. Send, approval and order
commands arrive later together with their human gates."""
import re
from dataclasses import dataclass, field
from pathlib import Path
from typing import Callable, Optional

from .config import UPLOADS_FOLDER, Settings


class JobRejected(Exception):
    """The request is not an allowed command, or an argument failed its rule. Safe to show to the user."""


SKU_RE = re.compile(r"[A-Za-z0-9][A-Za-z0-9._&-]{0,39}")   # no leading '-', no spaces, no path characters
STUB_MODES = ("ok", "fail", "slow", "secret", "hostile", "term", "url")

# What an exit code means, from the agents' own behaviour (Stage 0): sys.exit("message") and sys.exit(1) both give 1.
GENERIC_EXITS = {0: "Finished normally.",
                 1: "The tool stopped with an error (see the end of the log). Nothing further was done.",
                 2: "The tool refused the request (a value was missing or wrong)."}
# email_agent.py and decision_agent.py also exit 3 when their state file is locked by another command (statefile.py).
STATE_EXITS = {**GENERIC_EXITS, 3: "Another task was saving at the same moment, so nothing was changed. Try again in a moment."}


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
            raise JobRejected("That search's results are not in the Excel Output Sheets folder. Choose a search from the list.")
        return str(Path("Excel Output Sheets") / v)
    return check


def enum_rule(options) -> Callable[[str], str]:
    def check(v: str) -> str:
        if v not in options:
            raise JobRejected(f"'{str(v)[:40]}' is not one of: {', '.join(options)}.")
        return v
    return check


def pattern_rule(pattern: str, what: str) -> Callable[[str], str]:
    rx = re.compile(pattern)

    def check(v: str) -> str:
        if not isinstance(v, str) or not rx.fullmatch(v):
            raise JobRejected(f"'{str(v)[:40]}' is not {what}.")
        return v
    return check


def typed_line_rule(what: str) -> Callable[[str], str]:
    """A short line a person typed (an approver's name, a channel). Never empty; never something the agent's command
    line could read as an option; never more than one line."""
    def check(v: str) -> str:
        if not isinstance(v, str) or not v.strip():
            raise JobRejected(f"Type {what}. It is never filled in for you.")
        v = v.strip()
        if len(v) > 80 or v.startswith("-") or any(ord(c) < 32 or ord(c) == 127 for c in v):
            raise JobRejected(f"{what.capitalize()} must be one line of up to 80 characters and cannot start with '-'.")
        return v
    return check


def inputs_file_rule(settings: Settings) -> Callable[[str], str]:
    """Text a person pasted is saved by the server under ui_data\\inputs with a name the server builds; only such a file
    is accepted, and it is passed on by its full path."""
    def check(v: str) -> str:
        folder = (settings.data / "inputs").resolve()
        if not isinstance(v, str) or not re.fullmatch(r"[0-9a-f]{16}\.txt", v) or not (folder / v).is_file() or (folder / v).is_symlink():
            raise JobRejected("The pasted text was not saved. Paste it again.")
        return str(folder / v)
    return check


def switch_rule(v: str) -> str:
    if v != "1":
        raise JobRejected("That option can only be switched on or left off.")
    return v


APPROVED_CSV = "approved_orders.csv"
FACTS_RE = re.compile(r"facts_\d{8}_\d{6}_[0-9a-f]{6}\.csv")                 # a filled facts file a person uploaded
TEMPLATE_RE = re.compile(r"facts_template_\d{8}_\d{6}_[0-9a-f]{6}\.csv")     # a template the content agent wrote


def facts_file_rule(settings: Settings) -> Callable[[str], str]:
    """A facts file is one a person uploaded, picked by its server-built name; it is never a typed path."""
    def check(v: str) -> str:
        path = settings.uploads_dir / v if isinstance(v, str) and FACTS_RE.fullmatch(v) else None
        if path is None or not path.is_file() or path.is_symlink():
            raise JobRejected("Choose a product details sheet you uploaded. Nothing was started.")
        return str(Path(UPLOADS_FOLDER) / v)
    return check


def new_template_rule(settings: Settings) -> Callable[[str], str]:
    """The template's name is built by the server; the content agent refuses to write over an existing file, and so does this."""
    def check(v: str) -> str:
        if not isinstance(v, str) or not TEMPLATE_RE.fullmatch(v) or (settings.uploads_dir / v).exists():
            raise JobRejected("The template file name is not one this tool builds.")
        return str(Path(UPLOADS_FOLDER) / v)
    return check


REQUEST_ID = r"A-\d{4}"
EMAIL_ADDRESS = r"[A-Za-z0-9][A-Za-z0-9._%+-]{0,63}@[A-Za-z0-9](?:[A-Za-z0-9.-]{0,180}[A-Za-z0-9])?\.[A-Za-z]{2,24}"


@dataclass(frozen=True)
class Param:
    flag: Optional[str]          # None = positional
    rule: Callable[[str], str]
    required: bool = True
    many: bool = False           # the value is a list; every item passes the rule and follows the flag once
    switch: bool = False         # a flag with no value: the given value must pass the rule, then only the flag is passed
    private: bool = False        # typed by a person at a gate: passed to the agent, never kept in the UI's own job record


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
    check: Optional[Callable[[dict], Optional[dict]]] = None    # a rule across parameters (which may be combined)

    def argv(self, given: dict) -> list:
        unknown = set(given) - set(self.params)
        if unknown:
            raise JobRejected(f"Not allowed for this command: {', '.join(sorted(unknown))}.")
        if self.check:
            given = self.check(given) or given      # may hand back the same arguments in their checked spelling
        out = list(self.fixed)
        for name, p in self.params.items():
            if name not in given:
                if p.required:
                    raise JobRejected(f"Missing: {name}.")
                continue
            if p.switch:
                p.rule(given[name])
                out.append(p.flag)
                continue
            if p.many:
                values = given[name]
                if not isinstance(values, (list, tuple)) or not values:
                    raise JobRejected(f"Missing: {name}.")
                out += [p.flag, *[p.rule(v) for v in values]]
                continue
            value = p.rule(given[name])
            out += [value] if p.flag is None else [p.flag, value]
        return out

    def recorded(self, args: list) -> list:
        """The arguments as the UI's job record keeps them: a value typed at a gate is replaced, never stored."""
        hide = {p.flag for p in self.params.values() if p.private and p.flag}
        return ["(typed by a person, not stored)" if i and args[i - 1] in hide else a for i, a in enumerate(args)]


AGENTS = ("sourcing_agent", "email_agent", "decision_agent", "order_sheet", "content_agent")


def _record_check(given: dict) -> None:
    """The decision agent requires --confirm-request to equal --request; the UI refuses a mismatch before starting it."""
    if str(given.get("confirm_request", "")).strip().upper() != str(given.get("request", "")).strip().upper():
        raise JobRejected("The request number you typed is not the request you picked. Nothing was recorded.")


def pipeline_commands(settings: Settings) -> dict:
    """Every agent command the tabs after Search can start: (script, subcommand, allowed flags, value rules). A flag that
    is not listed here cannot be passed: --yes, --allow-listing-facts, --dry-run, --out, --duty are deliberately absent."""
    results, sku = results_file_rule(settings), sku_rule(settings)
    send_exits = {**STATE_EXITS, 0: "Finished. The last lines below say how many emails were sent.",
                  1: "Sending was refused or stopped. Nothing further was sent (see the end of the log)."}
    record_exits = {**STATE_EXITS, 0: "The answer was recorded. Nothing has been ordered.",
                    1: "The tool did not accept the answer or the confirmation. Nothing was recorded (see the end of the log)."}
    table = [
        Command("emails.drafts", "Preview the emails", "email_agent.py", ("drafts",),
                {"results": Param(None, results), "near_miss": Param("--near-miss", switch_rule, False, switch=True)},
                True, "Previews are only shown; nothing was sent or saved.", STATE_EXITS),
        Command("emails.send", "Send emails to sellers", "email_agent.py", ("send",),
                {"results": Param(None, results), "confirm_count": Param("--confirm-count", pattern_rule(r"\d{1,4}", "a whole number")),
                 "test_recipient": Param("--test-recipient", pattern_rule(EMAIL_ADDRESS, "an email address"), False)},
                True, CANCEL_WARNINGS["email_agent send"], send_exits),
        Command("emails.mark", "Mark a seller as contacted", "email_agent.py", ("mark",),
                {"sku": Param(None, sku), "results": Param("--results", results),
                 "followup": Param("--followup", switch_rule, False, switch=True)}, True,
                "The seller may or may not have been marked as contacted. Check the Status column on this page before marking again.", STATE_EXITS),
        Command("emails.reply", "Record a seller's reply", "email_agent.py", ("reply",),
                {"sku": Param(None, sku), "results": Param("--results", results), "file": Param("--file", inputs_file_rule(settings))},
                True, CANCEL_WARNINGS["email_agent reply"], STATE_EXITS),
        Command("approvals.request", "Create an approval request", "decision_agent.py", ("request",),
                {"results": Param(None, results), "budget": Param("--budget", pattern_rule(r"\d{1,9}(\.\d{1,2})?", "an amount such as 1500 or 1500.50"), False)},
                True, "The request may or may not have been created. Check your requests on this page before creating another.", STATE_EXITS),
        Command("approvals.record", "Record your manager's answer", "decision_agent.py", ("record",),
                {"results": Param("--results", results), "request": Param("--request", pattern_rule(REQUEST_ID, "a request number such as A-0001")),
                 "file": Param("--file", inputs_file_rule(settings)),
                 "confirm_request": Param("--confirm-request", pattern_rule(REQUEST_ID, "a request number such as A-0001"), private=True),
                 "channel": Param("--channel", typed_line_rule("how the answer reached you"), private=True),
                 "approver": Param("--approver", typed_line_rule("who gave the approval"), private=True)},
                True, CANCEL_WARNINGS["decision_agent record"], record_exits, _record_check),
        Command("approvals.revoke", "Withdraw an approval", "decision_agent.py", ("revoke",),
                {"sku": Param(None, sku), "results": Param("--results", results)}, True,
                "The approval may or may not have been withdrawn. Check the approvals on this page.", STATE_EXITS),
        Command("approvals.export", "Save the approved list", "decision_agent.py", ("export",),
                {"results": Param("--results", results)}, True, CANCEL_WARNINGS["decision_agent export"], STATE_EXITS),
        # The export is always the decision agent's own approved_orders.csv in the project folder: never a path from a form.
        Command("order.build", "Create the order sheet", "order_sheet.py", ("build", "--approved", APPROVED_CSV),
                {"results": Param("--results", results)}, True, CANCEL_WARNINGS["order_sheet build"]),
        Command("content.template", "Create the product details sheet", "content_agent.py", ("facts-template", "--approved", APPROVED_CSV),
                {"results": Param("--results", results), "facts": Param("--facts", new_template_rule(settings))}, True,
                "The product details sheet may be incomplete. Create a new one."),
        Command("content.generate", "Write the listings", "content_agent.py", ("generate", "--approved", APPROVED_CSV),
                {"results": Param("--results", results), "facts": Param("--facts", facts_file_rule(settings))}, True,
                CANCEL_WARNINGS["content_agent generate"]),
    ]
    return {c.id: c for c in table}


def build_allowlist(settings: Settings) -> dict:
    cmds = {}
    for a in AGENTS:
        cmds[f"help.{a}"] = Command(f"help.{a}", f"{a}.py --help", f"{a}.py", ("--help",))
    from .search import search_command      # imported here: search.py builds on this module
    cmds["search"] = search_command(settings)
    cmds.update(pipeline_commands(settings))
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
    "email_agent send": "Some emails may already have been sent. Each one is recorded as soon as it is sent, so check the Status column on "
                        "this page before sending again.",
    "email_agent reply": "The reply may not have been saved. Check the Status column on this page before pasting it again.",
    "decision_agent record": "The answer may or may not have been recorded. Check the approvals on this page before recording it again.",
    "decision_agent export": "The approved list may be half-saved. Save it again before creating an order sheet.",
    "order_sheet build": "No finished order sheet was created: one is only saved after its checks pass. A temporary file may be left in the "
                         "Order Sheets folder.",
    "sourcing_agent search": "The results keep every product finished so far (they are saved after each product); the rest were not searched.",
    "content_agent generate": "The listings file may be incomplete. Write the listings again.",
}
