"""TEST-ONLY shim. Python imports a module named sitecustomize at start-up when its folder is on PYTHONPATH; the offline
end-to-end test puts this folder there, so every agent the UI starts loads this file first. It does nothing unless
LCOM_TEST_SHIM=1. When active it:
  - answers the model clients (anthropic.Anthropic / AsyncAnthropic) from tests\\fixtures\\model_replies.json,
  - replaces the Nimble session (mcp.ClientSession, streamable_http_client) and smtplib.SMTP with fakes,
  - refuses every connection that is not to this computer, and logs the attempt,
  - refuses to open anything in the real project folder (LCOM_TEST_FORBID_ROOT), and logs the attempt,
  - fails closed: if any of this cannot be installed the process exits (code 97) before the agent's first line.
It lives only under tests\\ and no agent or UI file refers to it. Environment: LCOM_TEST_SHIM_DIR (a folder for its marker,
logs and the SMTP sink), LCOM_TEST_FIXTURES (the fixtures folder)."""
import os

if os.environ.get("LCOM_TEST_SHIM") == "1":
    try:
        import ipaddress
        import json
        import re
        import socket
        import sys
        import time

        OUT = os.environ["LCOM_TEST_SHIM_DIR"]
        FIXTURES = os.environ.get("LCOM_TEST_FIXTURES", "")

        def _log(name, line):
            with open(os.path.join(OUT, name), "a", encoding="utf-8") as f:
                f.write(f"{os.getpid()} {line}\n")

        # ---- the real project folder is never opened by an agent under test ----
        forbid = os.environ.get("LCOM_TEST_FORBID_ROOT", "")
        if forbid:
            _root = os.path.normcase(os.path.abspath(forbid)) + os.sep
            _tests = os.path.normcase(os.path.join(os.path.abspath(forbid), "tests")) + os.sep

            def _tripwire(event, args):
                if event not in ("open", "os.listdir", "os.scandir", "os.mkdir", "os.remove", "os.rename", "os.rmdir"):
                    return
                if not args or not isinstance(args[0], (str, bytes, os.PathLike)):
                    return
                path = os.path.normcase(os.path.abspath(os.fsdecode(args[0]))) + os.sep
                if path.startswith(_root) and not path.startswith(_tests):
                    _log("realroot.log", f"{event} {path}")
                    raise PermissionError(f"test shim: the real project folder was used: {event} {path}")
            sys.addaudithook(_tripwire)

        # ---- no connection leaves this computer ----
        class BlockedNetwork(OSError):
            pass

        def _loopback(host) -> bool:
            if host in (None, "", "localhost"):
                return True
            if isinstance(host, bytes):
                host = host.decode("ascii", "replace")
            try:
                return ipaddress.ip_address(str(host).split("%")[0]).is_loopback
            except ValueError:
                return False

        def _guard(path, real, host_of):
            def wrapped(*a, **kw):
                host = host_of(*a, **kw)
                if not _loopback(host):
                    _log("net.log", f"{path} {host!r}")
                    raise BlockedNetwork(f"test shim: {path} to {host!r} refused (not this computer)")
                return real(*a, **kw)
            return wrapped

        def _addr_host(self, address, *a, **kw):
            return address[0] if isinstance(address, tuple) else "localhost"      # pipes and local sockets

        socket.socket.connect = _guard("socket.connect", socket.socket.connect, _addr_host)
        socket.socket.connect_ex = _guard("socket.connect_ex", socket.socket.connect_ex, _addr_host)
        socket.socket.sendto = _guard("socket.sendto", socket.socket.sendto,
                                      lambda self, data, *a: (a[-1][0] if a and isinstance(a[-1], tuple) else "localhost"))
        socket.create_connection = _guard("socket.create_connection", socket.create_connection, lambda address, *a, **kw: address[0])
        socket.getaddrinfo = _guard("socket.getaddrinfo", socket.getaddrinfo, lambda host, *a, **kw: host)
        if sys.platform == "win32":        # the Proactor event loop connects through ConnectEx, never through socket.connect
            from asyncio import windows_events
            windows_events.IocpProactor.connect = _guard("asyncio.IocpProactor.connect", windows_events.IocpProactor.connect,
                                                         lambda self, conn, address: address[0])

        # ---- the model: the real SDK clients on a canned transport ----
        import anthropic
        import httpx

        def _text_of(message) -> str:
            content = message.get("content")
            return content if isinstance(content, str) else " ".join(str(b.get("text", "")) for b in content or [] if isinstance(b, dict))

        def _handler(request):
            body = json.loads(request.content)
            with open(os.path.join(FIXTURES, "model_replies.json"), encoding="utf-8") as f:
                replies = json.load(f)
            prompt = " ".join(_text_of(m) for m in body.get("messages", []) if m.get("role") == "user")
            if "A manufacturer replied to our sample-order inquiry" in prompt:
                kind, text = "reply", replies["reply"]
            elif body.get("tools") is not None or "competitive sourcing research analyst" in prompt:
                sku = re.search(r"SKU: (\S+)", prompt)
                kind, text = "search", replies["search"].get(sku[1] if sku else "", replies["search"]["*"])
            elif body.get("system"):
                kind, text = "listing", replies["listing"]
            else:
                _log("model.log", "UNRECOGNISED request")
                return httpx.Response(400, json={"type": "error", "error": {"type": "invalid_request_error", "message": "test shim: unrecognised request"}})
            _log("model.log", f"{kind} model={body.get('model')}")
            return httpx.Response(200, json={"id": "msg_test", "type": "message", "role": "assistant", "model": body.get("model", "test"),
                                             "stop_reason": "end_turn", "stop_sequence": None, "content": [{"type": "text", "text": text}],
                                             "usage": {"input_tokens": 500, "output_tokens": 50}})

        class AsyncAnthropic(anthropic.AsyncAnthropic):
            def __init__(self, *a, **kw):
                kw.update(http_client=httpx.AsyncClient(transport=httpx.MockTransport(_handler)), max_retries=0)
                super().__init__(*a, **kw)

        class Anthropic(anthropic.Anthropic):
            def __init__(self, *a, **kw):
                kw.update(http_client=httpx.Client(transport=httpx.MockTransport(_handler)), max_retries=0)
                super().__init__(*a, **kw)

        anthropic.AsyncAnthropic, anthropic.Anthropic = AsyncAnthropic, Anthropic

        # ---- Nimble: a session that lists the two tools and answers every call ----
        import asyncio

        import mcp
        import mcp.client.streamable_http as _sh
        from mcp import types as _t

        class _Streams:
            async def __aenter__(self):
                return (None, None)

            async def __aexit__(self, *a):
                return False

        class FakeSession:
            def __init__(self, read=None, write=None, *a, **kw):
                pass

            async def __aenter__(self):
                return self

            async def __aexit__(self, *a):
                return False

            async def initialize(self):
                pass

            async def list_tools(self):
                schema = {"type": "object", "properties": {"url": {"type": "string"}, "query": {"type": "string"}}}
                return type("Tools", (), {"tools": [type("Tool", (), {"name": n, "description": "", "inputSchema": schema})()
                                                    for n in ("nimble_search", "nimble_extract")]})()

            async def call_tool(self, name, arguments, read_timeout_seconds=None):
                _log("nimble.log", name)
                waited = 0.0
                while os.path.exists(os.path.join(OUT, "hold")) and waited < 60:     # lets a test keep a search running
                    await asyncio.sleep(0.1)
                    waited += 0.1
                return _t.CallToolResult(content=[_t.TextContent(type="text", text='{"content": "Test listing page. Widget $1.00 per piece."}')], is_error=False)

        mcp.ClientSession = FakeSession
        _sh.streamable_http_client = lambda url, http_client=None, **kw: _Streams()

        # ---- SMTP: every message goes to a file ----
        import smtplib

        class FakeSMTP:
            def __init__(self, host="", port=0, *a, **kw):
                self.host, self.port = host, port

            def __enter__(self):
                return self

            def __exit__(self, *a):
                return False

            def starttls(self, *a, **kw):
                pass

            def login(self, *a, **kw):
                pass

            def send_message(self, msg, *a, **kw):
                with open(os.path.join(OUT, "smtp_sink.jsonl"), "a", encoding="utf-8") as f:
                    f.write(json.dumps({"to": msg["To"], "from": msg["From"], "subject": msg["Subject"], "host": self.host,
                                        "body": msg.get_content(), "at": time.time()}) + "\n")

        smtplib.SMTP = smtplib.SMTP_SSL = FakeSMTP

        open(os.path.join(OUT, f"shim_{os.getpid()}.ok"), "w").close()       # written last: every patch above is in place
    except BaseException as e:   # noqa: BLE001 - Python would only print this and carry on; a half-installed shim must never run an agent
        os.write(2, f"test shim could not be installed: {type(e).__name__}: {e}\n".encode())
        os._exit(97)
