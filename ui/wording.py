"""The only place that maps agent and data vocabulary to what a person reads. The UI is for Zync's own sourcing tool and is
not branded by the company whose products serve as the price reference, so that company's name never reaches a page.

`display()` is applied to every value a template prints (Jinja `finalize`), to displayed job logs, and to file names.
Stored logs and files are never changed."""
import re

REFERENCE = "reference"
INTERNAL_PREFIX = "lcom-"          # the workflow id prefix (the Langfuse session id): internal only, never displayed

# Specific phrases first, then the bare name. Case-insensitive; "L-Com", "LCom", "lcom", "l com", "L_Com" all match.
_NAME = r"(?<![A-Za-z])L[\s_-]?Com(?![a-z])"
_RULES = [
    (re.compile(_NAME + r"[\s_-]*prices?\.csv", re.I), "reference_prices.csv"),           # the reference-prices file name (alias)
    (re.compile(_NAME + r"[\s_-]*unit[\s_-]*price", re.I), "Reference unit price"),
    (re.compile(_NAME + r"[\s_-]*(price)", re.I), lambda m: "Reference " + m[1].lower()),
    (re.compile(_NAME + r"[\s_-]*(sku)s?", re.I), "SKU"),
    (re.compile(r"(?<![A-Za-z])lcom-(?=[A-Za-z0-9_])", re.I), ""),                       # lcom-<stem> workflow ids
    (re.compile(_NAME, re.I), REFERENCE),
]


def display(text):
    """Display wording for a string; anything else is returned unchanged."""
    if not isinstance(text, str) or not text or hasattr(text, "__html__"):   # not text, or already-rendered markup (its parts were converted)
        return text
    for pattern, repl in _RULES:
        text = pattern.sub(repl, text)
    return text


def contains_reference_name(text: str) -> bool:
    """True if any case variant of the name is in the text (used by the tests and as a last check)."""
    return bool(re.search(r"(?i)(?<![a-z])l[\s_-]?com(?![a-z])", text or ""))


def strip_prefix(workflow: str) -> str:
    """Internal workflow id -> the plain stem used in URLs."""
    return workflow[len(INTERNAL_PREFIX):] if (workflow or "").lower().startswith(INTERNAL_PREFIX) else (workflow or "")
