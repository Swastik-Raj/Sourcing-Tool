"""TEST-ONLY: run as a subprocess the way the UI runs an agent, to prove what the shim does. It imports the libraries the
way the agents do, tries every route out of this computer, and prints one JSON line. It never contacts a real service:
192.0.2.1 is the reserved documentation address range and *.invalid never resolves."""
import asyncio
import json
import os
import smtplib
import socket
import sys
from email.message import EmailMessage

from anthropic import AsyncAnthropic
from anthropic.lib.tools import beta_async_tool
from mcp import ClientSession
from mcp.client.streamable_http import streamable_http_client

FAR = ("192.0.2.1", 443)
OUT = os.environ.get("LCOM_TEST_SHIM_DIR", "")
out = {"pid": os.getpid(), "sitecustomize": "sitecustomize" in sys.modules, "blocked": [], "leaked": []}


def logged(name: str) -> int:
    path = os.path.join(OUT, name)
    return open(path, encoding="utf-8").read().count("\n") if OUT and os.path.exists(path) else 0


def judge(name, e, before):
    """Blocked = the shim logged a refusal during the attempt (libraries such as httpx wrap the error in their own type)."""
    if logged("net.log") > before:
        out["blocked"].append(name)
    else:
        out["leaked"].append(f"{name} ({type(e).__name__}: {e})")


def attempt(name, fn):
    before = logged("net.log")
    try:
        fn()
        out["leaked"].append(name)
    except Exception as e:  # noqa: BLE001
        judge(name, e, before)


def sock(kind=socket.SOCK_STREAM):
    s = socket.socket(socket.AF_INET, kind)
    s.settimeout(1)
    return s


attempt("socket.connect", lambda: sock().connect(FAR))
attempt("socket.connect_ex", lambda: sock().connect_ex(FAR))
attempt("socket.sendto", lambda: sock(socket.SOCK_DGRAM).sendto(b"x", FAR))
attempt("socket.create_connection", lambda: socket.create_connection(FAR, timeout=1))
attempt("socket.getaddrinfo", lambda: socket.getaddrinfo("api.example.invalid", 443))


async def main():
    out["loop"] = type(asyncio.get_running_loop()).__name__

    async def a(name, coro_fn):
        before = logged("net.log")
        try:
            await asyncio.wait_for(coro_fn(), 2)
            out["leaked"].append(name)
        except Exception as e:  # noqa: BLE001
            judge(name, e, before)

    await a("asyncio.open_connection ip", lambda: asyncio.open_connection(*FAR))
    await a("asyncio.open_connection name", lambda: asyncio.open_connection("api.example.invalid", 443))

    async def raw():
        s = socket.socket()
        s.setblocking(False)
        await asyncio.get_running_loop().sock_connect(s, FAR)
    await a("loop.sock_connect", raw)

    async def with_httpx():
        import httpx
        async with httpx.AsyncClient(timeout=2) as c:
            await c.get("https://192.0.2.1/")
    await a("httpx.AsyncClient", with_httpx)

    server = await asyncio.start_server(lambda r, w: w.close(), "127.0.0.1", 0)      # this computer must still be reachable
    port = server.sockets[0].getsockname()[1]
    _, w = await asyncio.open_connection("127.0.0.1", port)
    w.close()
    socket.create_connection(("127.0.0.1", port), timeout=2).close()
    server.close()
    out["loopback_ok"] = True

    # the model and Nimble, used the way sourcing_agent.py uses them
    async with streamable_http_client("https://mcp.example.invalid/mcp", http_client=None) as (read, write):
        async with ClientSession(read, write) as session:
            await session.initialize()
            tools = (await session.list_tools()).tools
            tool = tools[1]

            async def call(**kw):
                r = await session.call_tool(name=tool.name, arguments=kw, read_timeout_seconds=5)
                return " ".join(b.text for b in r.content)
            out["nimble"] = [t.name for t in tools] + [await call(url="https://example.invalid/x")]
            client = AsyncAnthropic(api_key="sk-ant-FAKE-not-a-key", max_retries=3)
            runner = client.beta.messages.tool_runner(
                model="claude-haiku-4-5", max_tokens=100, max_iterations=4,
                tools=[beta_async_tool(call, name=tool.name, description="", input_schema=tool.inputSchema)],
                output_config={"format": {"type": "json_schema", "schema": {"type": "object", "additionalProperties": False}}},
                messages=[{"role": "user", "content": "You are a competitive sourcing research analyst.\n  SKU: TST-100\n"}])
            turns, last = 0, None
            async for message in runner:
                turns, last = turns + 1, message
            out["tool_runner"] = {"turns": turns, "final": json.loads(" ".join(b.text for b in last.content if b.type == "text"))["product"]}

asyncio.run(main())

from anthropic import Anthropic   # noqa: E402  (the lazy import email_agent.py and content_agent.py do)
msg = Anthropic(api_key="sk-ant-FAKE-not-a-key").messages.create(
    model="claude-haiku-4-5", max_tokens=10, messages=[{"role": "user", "content": "A manufacturer replied to our sample-order inquiry. Reply:\nhello"}])
out["sync_client"] = json.loads(msg.content[0].text)["status"]

m = EmailMessage()
m["To"], m["From"], m["Subject"] = "test@example.invalid", "sender@example.invalid", "shim self-test"
m.set_content("body")
with smtplib.SMTP("smtp.example.invalid", 587) as s:
    s.starttls()
    s.login("u", "p")
    s.send_message(m)
out["smtp"] = "sink"

root = os.environ.get("LCOM_TEST_FORBID_ROOT", "")
out["real_root"] = []
for target, how in ((os.path.join(root, "README.md"), open), (os.path.join(root, "email_state.json"), open), (os.path.join(root, "Excel Output Sheets"), os.listdir),
                    (root, os.listdir)):
    try:
        how(target)
        out["real_root"].append("OPENED " + os.path.basename(target))
    except PermissionError:
        out["real_root"].append("refused")
    except OSError as e:
        out["real_root"].append(type(e).__name__)
print(json.dumps(out))
