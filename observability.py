"""Shared Langfuse tracing for every agent. See OBSERVABILITY.md.

Off unless OBS_ENABLED=1 and LANGFUSE_PUBLIC_KEY / LANGFUSE_SECRET_KEY are set. Off means off: this module then imports
nothing but the standard library, opens no socket, starts no thread. On, a telemetry failure never reaches the agent:
every public call swallows its own errors, and the exit-time flush is bounded."""
import atexit
import contextvars
import json
import logging
import os
import re
import sys
import threading
import time
import traceback
from datetime import date, datetime
from decimal import Decimal

# ---------- config: edit here ----------
# Is the text of prompts and responses sent to Langfuse? ON for the three agents that call a model (traces are viewed
# only by the project owner on a self-hosted server; masking still applies). OFF, and not switchable, for approvals
# (decision_agent, order_sheet). OBS_CAPTURE_CONTENT=a,b REPLACES the ON set with those agents (an empty value turns
# all capture off); it can never switch on the ALWAYS_METADATA_ONLY ones. Revisit this table before pointing at a shared
# or cloud Langfuse: email replies and unreleased listing text would then be visible to everyone in the project.
CAPTURE_CONTENT = {"sourcing_agent": True, "email_agent": True, "content_agent": True,
                   "decision_agent": False, "order_sheet": False}
ALWAYS_METADATA_ONLY = ("decision_agent", "order_sheet")
FLUSH_TIMEOUT_S = 5.0                      # OBS_FLUSH_TIMEOUT overrides
SECRET_ENV = ("ANTHROPIC_API_KEY", "NIMBLE_API_KEY", "LANGFUSE_SECRET_KEY", "LANGFUSE_PUBLIC_KEY", "SMTP_PASSWORD",
              "SMTP_USER", "SHIPPING_ADDRESS")
DEFAULT_ADDRESS_LINES = ("Zync Technologies", "Plano, TX")        # email_agent.DEFAULT_SHIPPING, copied
REDACT_KEYS = {"approver", "approved_by", "channel", "password", "api_key", "apikey", "secret", "token", "authorization",
               "x-api-key", "email", "recipient", "recipients", "phone", "address", "shipping_address", "reply",
               "reply_text", "body"}
CONTENT_ATTR_PREFIXES = ("gen_ai.input", "gen_ai.output", "gen_ai.prompt", "gen_ai.completion",
                         "gen_ai.system_instructions", "gen_ai.tool", "langfuse.observation.input",
                         "langfuse.observation.output", "langfuse.trace.input", "langfuse.trace.output")
_DOTENV = os.path.join(os.path.dirname(os.path.abspath(__file__)), ".env")   # only LANGFUSE_* / OBS_* are read from it
_META = "langfuse.observation.metadata."

_S = {"agent": None, "on": False, "client": None, "capture": False, "root": None, "protected": set(), "dead": False,
      "provider": None, "instrumented": False}
_CUR = contextvars.ContextVar("obs_scope", default=None)
_test_exporter = None                      # tests only: an in-memory span exporter instead of the Langfuse server


def _warn(msg: str) -> None:
    try:
        print(f"[obs] {msg}", file=sys.stderr)
    except Exception:  # noqa: BLE001
        pass


# ---------- masking: the one place everything passes through ----------

_EMAIL = re.compile(r"[\w.+'\-]+@[\w\-]+(?:\.[\w\-]+)+")
_KEYLIKE = re.compile(r"\b(?:sk-ant-|sk-lf-|pk-lf-|sk-|pk-)[A-Za-z0-9_\-]{8,}")
# NAME=value, NAME: value, "NAME": "value" (also \"NAME\": \"value\" inside repr'd JSON) for any name ending in api_key / secret_key /
# password / public_key (so ANTHROPIC_API_KEY, NIMBLE_API_KEY, SMTP_PASSWORD, LANGFUSE_SECRET_KEY, LANGFUSE_PUBLIC_KEY), the bare names
# secret and token (not spec_token: an ordinary metadata key), plus x-api-key, authorization and SHIPPING_ADDRESS. The VALUE is masked whatever it looks like: quoted (may hold spaces) or bare.
# Prose without an = or : after the name ("the API key is missing") is untouched.
_ASSIGN = re.compile(r"""(?i)((?<!\w)(?:[\w-]*?(?:api[_-]?key|secret[_-]?key|passw(?:or)?d|public[_-]?key)|secret|token|authorization|shipping[_-]?address)"""
                     r"""\\?["']?\s*[=:]\s*)(?:Bearer\s+)?(?:"(?:[^"\\]|\\.)*"|'[^']*'|\S+)""")
_BEARER = re.compile(r"(?i)\bBearer\s+[A-Za-z0-9._~+/=\-]{8,}")
# Phone numbers: 3+ digit groups joined by space . - or parentheses, a +country number, or a Chinese mobile. 9-15 digits;
# dates ("2026-10-08") and money ("1500.00 2000.00") are not phone numbers.
_PHONE = re.compile(r"(?<![\w.])(?:\+?\(?\d{1,4}\)?(?:[\s.\-]\(?\d{2,5}\)?){2,5}|\+\d{9,14}|1[3-9]\d{9})(?![\w])")
_MAX_STR = 20000


def _exact_secrets() -> list:
    out = set(_S["protected"])
    for name in SECRET_ENV:
        v = os.environ.get(name, "")
        if v:
            out.add(v.replace("\\n", "\n").strip())
    out.update(DEFAULT_ADDRESS_LINES)
    out |= {line.strip() for v in list(out) for line in v.split("\n")}
    return sorted((v for v in out if len(v.strip()) >= 3), key=len, reverse=True)


def _phone(m) -> str:
    t = m[0]
    digits = sum(c.isdigit() for c in t)
    return t if not 9 <= digits <= 15 or re.match(r"\d{4}-\d{2}-\d{2}", t) or re.search(r"\.\d{2}(?=\s|$|[-,;])", t) else "[PHONE]"


def mask_text(text: str) -> str:
    text = text if len(text) <= _MAX_STR else text[:_MAX_STR] + "...[truncated]"
    for secret in _exact_secrets():
        text = re.sub(r"(?<!\w)" + re.escape(secret) + r"(?!\w)" if secret[0].isalnum() and secret[-1].isalnum()
                      else re.escape(secret), "[REDACTED]", text, flags=re.IGNORECASE)
    text = _ASSIGN.sub(lambda m: m[1] + "[REDACTED]", text)
    text = _BEARER.sub("[REDACTED]", text)
    text = _KEYLIKE.sub("[REDACTED]", text)
    text = _EMAIL.sub("[EMAIL]", text)
    return _PHONE.sub(_phone, text)


def mask(obj, _depth: int = 0):
    """Recursively masks keys, emails, phone numbers, the shipping address, protected values and secrets. Returns
    JSON-friendly data. If masking itself fails, nothing raw is returned."""
    try:
        if obj is None or isinstance(obj, (bool, int)):
            return obj
        if isinstance(obj, float):
            return obj
        if isinstance(obj, Decimal):
            return float(obj)
        if isinstance(obj, (datetime, date)):
            return obj.isoformat()
        if _depth > 8:
            return "[too deep]"
        if isinstance(obj, str):
            return mask_text(obj)
        if isinstance(obj, bytes):
            return f"[{len(obj)} bytes]"
        if isinstance(obj, dict):
            return {mask_text(str(k)): ("[REDACTED]" if _redact_key(k) else mask(v, _depth + 1)) for k, v in obj.items()}
        if isinstance(obj, (list, tuple, set, frozenset)):
            return [mask(v, _depth + 1) for v in obj]
        return mask_text(str(obj))
    except Exception:  # noqa: BLE001
        return "[mask error]"


def _redact_key(k) -> bool:
    k = str(k).lower()
    return k in REDACT_KEYS or any(w in k for w in ("password", "secret", "api_key", "apikey", "authorization"))


def mask_exc(exc) -> str:
    return mask_text(f"{type(exc).__name__}: {exc}"[:500])


def protect(*values) -> None:
    """Registers runtime values (a typed approver name, a channel) so they are masked wherever they appear."""
    _S["protected"].update(str(v).strip() for v in values if v and len(str(v).strip()) >= 2)


# ---------- init / flush ----------

def _load_env_file() -> None:
    try:
        if _DOTENV and os.path.exists(_DOTENV):
            from dotenv import dotenv_values
            for k, v in dotenv_values(_DOTENV).items():
                if v is not None and (k.startswith("LANGFUSE_") or k.startswith("OBS_")):
                    os.environ.setdefault(k, v)
    except Exception:  # noqa: BLE001
        pass


def _wanted() -> bool:
    return (os.environ.get("OBS_ENABLED") == "1" and bool(os.environ.get("LANGFUSE_PUBLIC_KEY"))
            and bool(os.environ.get("LANGFUSE_SECRET_KEY")))


def capture_for(agent: str) -> bool:
    if agent in ALWAYS_METADATA_ONLY:
        return False
    env = os.environ.get("OBS_CAPTURE_CONTENT")
    if env is not None:
        return agent in {a.strip() for a in env.split(",")}
    return bool(CAPTURE_CONTENT.get(agent, False))


def enabled() -> bool:
    return bool(_S["on"]) and not _S["dead"]


def init(agent: str) -> bool:
    """Idempotent. Returns whether tracing is on."""
    if _S["agent"] is not None:
        return enabled()
    _S["agent"] = agent
    try:
        _load_env_file()
        if not _wanted():
            return False
        _S["capture"] = capture_for(agent)
        _S["client"] = _build_client(agent, _S["capture"])
        _S["on"] = True
        atexit.register(flush)
        asked = os.environ.get("OBS_CAPTURE_CONTENT")
        note = " (OBS_CAPTURE_CONTENT ignored: this agent is metadata-only)" if agent in ALWAYS_METADATA_ONLY and asked else ""
        _warn(f"tracing ON for {agent}; prompt/response content capture: {'ON' if _S['capture'] else 'OFF'}{note}; "
              f"env: {os.environ.get('OBS_ENV', 'dev')}")
    except Exception as e:  # noqa: BLE001 - tracing must never stop an agent
        _S["on"] = False
        _warn(f"tracing disabled, setup failed: {mask_exc(e)}")
    return enabled()


def _masking_processor():
    """Runs before Langfuse's processor: masks span events (exception messages) and the status text, which the
    attribute masking in _mask_spans does not reach. Built lazily: the OpenTelemetry SDK is imported only when on."""
    from opentelemetry.sdk.trace import Event, SpanProcessor
    from opentelemetry.trace import Status

    class MaskingProcessor(SpanProcessor):
        def on_end(self, span):
            try:
                events = []
                for ev in span._events:
                    attrs = {"exception.type": mask_text(str(ev.attributes.get("exception.type", "")))} if ev.name == "exception"                         else mask(dict(ev.attributes or {}))
                    events.append(Event(ev.name, attrs, ev.timestamp))
                span._events = events
                st = span._status
                if st is not None and st.description:
                    span._status = Status(st.status_code, mask_text(st.description))
            except Exception:  # noqa: BLE001
                span._events = []
    return MaskingProcessor()


def _mask_spans(*, params):
    """Langfuse export-stage hook: mask every string attribute; drop prompt/response bodies unless capture is on."""
    from langfuse.types import MaskOtelSpansResult, OtelSpanPatch
    patches = {}
    for ident, span in params.spans.items():
        try:
            sets, drops = {}, []
            for k, v in span.attributes.items():
                if not _S["capture"] and k.startswith(CONTENT_ATTR_PREFIXES):
                    drops.append(k)
                elif isinstance(v, str):
                    m = "[REDACTED]" if k.startswith(_META) and _redact_key(k[len(_META):]) else mask_text(v)
                    if m != v:
                        sets[k] = m
                elif isinstance(v, (tuple, list)) and v and all(isinstance(x, str) for x in v):
                    m = [mask_text(x) for x in v]
                    if m != list(v):
                        sets[k] = m
            if sets or drops:
                patches[ident] = OtelSpanPatch(set_attributes=sets, delete_attributes=tuple(drops))
        except Exception:  # noqa: BLE001
            patches[ident] = OtelSpanPatch(delete_attributes=tuple(span.attributes))   # fail closed for this span
    return MaskOtelSpansResult(span_patches=patches)


class _MaskLogs(logging.Filter):
    def filter(self, record):
        try:
            record.msg, record.args = mask_text(record.getMessage()), None
            if record.exc_info:
                record.exc_text = mask_text("".join(traceback.format_exception(*record.exc_info)))
        except Exception:  # noqa: BLE001
            record.msg, record.args, record.exc_info = "[log line withheld]", None, None
        return True


def _mask_library_logs():
    """Langfuse and OpenTelemetry log exporter errors (which can quote headers or keys) straight to stderr: mask those too."""
    otel = logging.getLogger("opentelemetry")
    if not any(isinstance(f, _MaskLogs) for h in otel.handlers for f in h.filters):
        mine = logging.StreamHandler(sys.stderr)
        mine.addFilter(_MaskLogs())
        otel.addHandler(mine)
    for h in logging.getLogger("langfuse").handlers:
        if not any(isinstance(f, _MaskLogs) for f in h.filters):
            h.addFilter(_MaskLogs())


def _build_client(agent: str, capture: bool):
    from langfuse import Langfuse                                            # imported only when tracing is on
    from opentelemetry.sdk.trace import TracerProvider
    if not capture:
        os.environ["TRACELOOP_TRACE_CONTENT"] = "false"                      # the instrumentation then records no bodies
    provider = TracerProvider(shutdown_on_exit=False)                        # no hidden atexit that could block exit
    provider.add_span_processor(_masking_processor())
    kw = {}
    if _test_exporter is not None:
        kw["span_exporter"] = _test_exporter
    client = Langfuse(public_key=os.environ["LANGFUSE_PUBLIC_KEY"], secret_key=os.environ["LANGFUSE_SECRET_KEY"],
                      base_url=os.environ.get("LANGFUSE_BASE_URL") or os.environ.get("LANGFUSE_HOST"),
                      tracer_provider=provider, mask=lambda *, data, **_: mask(data), mask_otel_spans=_mask_spans,
                      environment=os.environ.get("OBS_ENV", "dev"), **kw)
    _S["provider"] = provider
    _mask_library_logs()
    return client


def instrument_anthropic() -> None:
    """Call right before an Anthropic client is used by messages.create / .stream: the instrumentation then records those calls.
    Lazy because importing it costs ~3 s, and most commands never call the API. The beta tool runner is NOT captured by it
    (see OBSERVABILITY.md); sourcing_agent records its own generations. Idempotent, never raises."""
    if not enabled() or _S.get("instrumented"):
        return
    try:
        from opentelemetry.instrumentation.anthropic import AnthropicInstrumentor
        AnthropicInstrumentor().instrument(tracer_provider=_S["provider"])
        _S["instrumented"] = True
    except Exception as e:  # noqa: BLE001
        _warn(f"Anthropic instrumentation not installed: {mask_exc(e)}")


def flush(timeout=None) -> bool:
    """Bounded flush: gives up after a few seconds with one stderr line. Never raises."""
    if not _S["on"] or _S["dead"] or _S["client"] is None:
        return True
    try:
        timeout = float(timeout if timeout is not None else os.environ.get("OBS_FLUSH_TIMEOUT", FLUSH_TIMEOUT_S))
    except ValueError:
        timeout = FLUSH_TIMEOUT_S
    done, err = threading.Event(), []

    def run():
        try:
            _S["client"].flush()
        except BaseException as e:  # noqa: BLE001
            err.append(e)
        finally:
            done.set()
    threading.Thread(target=run, daemon=True, name="obs-flush").start()
    if not done.wait(timeout):
        _S["dead"] = True
        try:
            atexit.unregister(_S["client"]._resources.shutdown)                # Langfuse's own exit hook would block again
        except Exception:  # noqa: BLE001
            pass
        _warn(f"flush did not finish in {timeout:g}s - giving up; some traces may be missing")
        return False
    if err:
        _warn(f"flush failed: {mask_exc(err[0])}")
        return False
    return True


def _reset_for_tests() -> None:
    _S.update(agent=None, on=False, client=None, capture=False, root=None, protected=set(), dead=False, provider=None, instrumented=False)


# ---------- workflow identity ----------

def workflow_id(results_path):
    """'lcom-<results file stem>': the Langfuse session id of every command that touches that results file."""
    if not results_path:
        return None
    return "lcom-" + os.path.splitext(os.path.basename(str(results_path)))[0]


def current_workflow():
    """The running command's workflow id, or None when tracing is off."""
    root = _S["root"]
    return root.session if enabled() and root is not None else None


def url_host(url):
    m = re.match(r"\s*(?:\w+://)?([^/?#\s]+)", str(url or ""))
    return m[1].lower() if m else None


# ---------- scopes: spans, generations, the per-command trace ----------

class _Noop:
    """What every helper returns when tracing is off (or broke): same calls, no effect."""
    def __enter__(self):
        return self

    def __exit__(self, *exc):
        return False

    def __getattr__(self, name):
        return lambda *a, **k: None


NOOP = _Noop()


class Scope:
    def __init__(self, name, kind="span", attrs=None, model=None, current=True, root=False, session=None, tags=(),
                 input=None, output=None, usage=None):
        self.name, self.kind, self.model, self.current, self.root = name, kind, model, current, root
        self.attrs, self.session, self.tags, self.input, self.output, self.usage = dict(attrs or {}), session, list(tags), input, output, usage
        self._cm = self._obs = self._propagate = self._after = self._token = None
        self._t0 = time.monotonic()

    # -- lifecycle (every step swallows its own errors) --
    def _open(self):
        try:
            client = _S["client"]
            if self.root:
                self._enter_propagate()
            kw = dict(name=self.name, as_type=self.kind, metadata=mask(self.attrs) or None)
            if self.kind == "generation":
                kw.update(model=self.model, usage_details=self.usage)
            if _S["capture"]:
                kw.update(input=mask(self.input), output=mask(self.output))
            if self.current:
                self._cm = client.start_as_current_observation(**kw)
                self._obs = self._cm.__enter__()
            else:
                self._obs = client.start_observation(**kw)
            self._token = _CUR.set(self)
            if self.root:
                _S["root"] = self
                self._root_attrs()
        except Exception as e:  # noqa: BLE001
            self._obs = None
            _warn(f"span '{self.name}' not recorded: {mask_exc(e)}")
        return self

    def _enter_propagate(self):
        if self._propagate is not None:
            return
        from langfuse import propagate_attributes
        self._propagate = propagate_attributes(session_id=self.session, tags=self.tags or None, trace_name=self.name)
        self._propagate.__enter__()

    def _root_attrs(self):
        from langfuse._client.attributes import LangfuseOtelSpanAttributes as A
        o = self._obs._otel_span
        o.set_attribute(A.TRACE_NAME, self.name)
        if self.session:
            o.set_attribute(A.TRACE_SESSION_ID, self.session)
        o.set_attribute(A.TRACE_TAGS, list(self.tags))

    def __enter__(self):
        return self._open() if enabled() else self

    def __exit__(self, et, ev, tb):
        if self._obs is not None:
            try:
                self.set(duration_s=round(time.monotonic() - self._t0, 3))
                if ev is not None:
                    code = ev.code if isinstance(ev, SystemExit) else None
                    ok = isinstance(ev, SystemExit) and code in (None, 0)
                    if self.root:
                        self.set(exit_status="ok" if ok else ("exit" if isinstance(ev, SystemExit) else "error"),
                                 exit_code=0 if ok else (code if isinstance(code, int) else 1))
                    if not ok:
                        self._obs.update(level="ERROR", status_message=_error_text(ev, code))
                elif self.root:
                    self.set(exit_status="ok", exit_code=0)
            except Exception:  # noqa: BLE001
                pass
        self.end()
        return False                                   # the caller's exception always propagates untouched

    def start(self):
        return self._open() if enabled() else self

    def end(self, **attrs):
        if attrs:
            self.set(**attrs)
        try:
            if self._token is not None:
                _CUR.reset(self._token)
        except Exception:  # noqa: BLE001
            pass
        self._token = None
        try:
            if self._after is not None:
                self._after.__exit__(None, None, None)
            if self._cm is not None:
                self._cm.__exit__(None, None, None)     # not given the exception: OTel would record its raw message
            elif self._obs is not None:
                self._obs.end()
            if self._propagate is not None:
                self._propagate.__exit__(None, None, None)
        except Exception:  # noqa: BLE001
            pass
        self._cm = self._obs = self._propagate = self._after = None

    # -- recording --
    def set(self, **attrs):
        if self._obs is None:
            return
        try:
            for k, v in attrs.items():
                m = "[REDACTED]" if _redact_key(k) else mask(v)
                self._obs._otel_span.set_attribute(_META + k, m if isinstance(m, str) else json.dumps(m, default=str))
        except Exception:  # noqa: BLE001
            pass

    def tag(self, *tags):
        """Trace-level tags (only the per-command trace carries them)."""
        root = _S["root"]
        if root is None or root._obs is None:
            return
        try:
            from langfuse._client.attributes import LangfuseOtelSpanAttributes as A
            root.tags += [t for t in map(mask_text, tags) if t not in root.tags]
            root._obs._otel_span.set_attribute(A.TRACE_TAGS, list(root.tags))
        except Exception:  # noqa: BLE001
            pass

    def set_workflow(self, wid):
        """For a command whose results file is only named part-way through (the search)."""
        if self._obs is None or not wid or self.session:
            return
        try:
            from langfuse import propagate_attributes
            from langfuse._client.attributes import LangfuseOtelSpanAttributes as A
            self.session = wid
            self._obs._otel_span.set_attribute(A.TRACE_SESSION_ID, wid)
            self._after = propagate_attributes(session_id=wid, tags=self.tags or None, trace_name=self.name)
            self._after.__enter__()
        except Exception:  # noqa: BLE001
            pass

    def usage(self, input_tokens=None, output_tokens=None):
        if self._obs is None:
            return
        try:
            self._obs.update(usage_details={"input": input_tokens or 0, "output": output_tokens or 0})
        except Exception:  # noqa: BLE001
            pass

    def content(self, input=None, output=None):
        """Prompt / response bodies: dropped unless this agent has content capture on."""
        if self._obs is None or not _S["capture"]:
            return
        try:
            self._obs.update(input=mask(input) if input is not None else None, output=mask(output) if output is not None else None)
        except Exception:  # noqa: BLE001
            pass


def _error_text(ev, code) -> str:
    """Masked error message. Metadata-only agents (decision, order sheet) export the exception class only: their
    messages quote SKUs, sellers and reply lines."""
    if _S["agent"] in ALWAYS_METADATA_ONLY:
        return type(ev).__name__
    return mask_text(str(code)) if isinstance(ev, SystemExit) else mask_exc(ev)


def _make(name, kind, attrs, **kw):
    return Scope(name, kind, attrs, **kw) if enabled() else NOOP


def trace_command(agent, command, workflow=None, sku=None, **attrs):
    """One CLI command = one Langfuse trace named '<agent>.<command>', session id = the workflow id."""
    init(agent)
    tags = [f"agent:{agent}", f"env:{os.environ.get('OBS_ENV', 'dev')}"] + ([f"sku:{sku}"] if sku else [])
    return _make(f"{agent}.{command}", "span", {"agent": agent, "command": command, "workflow_id": workflow, **attrs},
                 root=True, session=workflow, tags=tags)


def open_span(name, **attrs):
    """Manual span that IS current (call .end()): for a loop body with several exits. Same task only."""
    return _make(name, "span", attrs).start() if enabled() else NOOP


def span(name, **attrs):
    return _make(name, "span", attrs)


def start_span(name, **attrs):
    """Manual span (call .end()): for code with several exit paths. Not made current."""
    return _make(name, "span", attrs, current=False).start() if enabled() else NOOP


def generation(name, model, **attrs):
    return _make(name, "generation", attrs, model=model)


def record_generation(name, model, input_tokens=None, output_tokens=None, latency_s=None, stop_reason=None,
                      prompt=None, response=None, **attrs):
    """A model call the instrumentation cannot see (the beta tool runner), recorded after the fact."""
    if not enabled():
        return
    g = _make(name, "generation", {"stop_reason": stop_reason, "latency_s": latency_s, **attrs}, model=model, current=False,
              usage={"input": input_tokens or 0, "output": output_tokens or 0}).start()
    g.content(prompt, response)
    g.end()


def annotate(**attrs):
    """Adds attributes to the innermost open span (the command's trace when none is deeper)."""
    s = _CUR.get()
    if s is not None:
        s.set(**attrs)


def set_workflow(wid):
    if _S["root"] is not None:
        _S["root"].set_workflow(wid)


def tag(*tags):
    s = _S["root"]
    if s is not None:
        s.tag(*tags)


def score(name, value, comment=None):
    """Langfuse score on the current span: number, bool or category string."""
    if not enabled():
        return
    try:
        client = _S["client"]
        if client.get_current_trace_id() is None:        # no open span: Langfuse rejects a score with nothing to attach to
            return
        kind, val = (("BOOLEAN", 1.0 if value else 0.0) if isinstance(value, bool) else
                     ("CATEGORICAL", str(value)) if isinstance(value, str) else ("NUMERIC", float(value)))
        client.create_score(name=name, value=val, data_type=kind, comment=mask_text(comment) if comment else None,
                            trace_id=client.get_current_trace_id(), observation_id=client.get_current_observation_id())
    except Exception as e:  # noqa: BLE001
        _warn(f"score '{name}' not recorded: {mask_exc(e)}")
