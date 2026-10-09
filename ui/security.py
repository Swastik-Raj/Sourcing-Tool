"""Host/Origin checks, CSRF tokens, security headers and safe file serving."""
import hashlib
import hmac
import secrets
from pathlib import Path
from urllib.parse import urlsplit

from fastapi import HTTPException, Request
from starlette.responses import PlainTextResponse

from .config import DOWNLOAD_EXTENSIONS, OUTPUT_FOLDERS, Settings

HEADERS = {
    "Content-Security-Policy": "default-src 'self'",
    "X-Content-Type-Options": "nosniff",
    "X-Frame-Options": "DENY",
    "Referrer-Policy": "no-referrer",
}
SECRET = secrets.token_bytes(32)    # per process: a restart invalidates open forms, which then ask for a reload
COOKIE = "ui_sid"


def allowed_hosts(s: Settings) -> set:
    return {f"127.0.0.1:{s.port}", f"localhost:{s.port}"} | ({f"{s.host}:{s.port}"} if s.host not in ("0.0.0.0", "::") else set())


def csrf_for(sid: str) -> str:
    return hmac.new(SECRET, sid.encode(), hashlib.sha256).hexdigest()


def check_csrf(request: Request, token: str) -> None:
    sid = request.cookies.get(COOKIE, "")
    if not sid or not token or not hmac.compare_digest(token, csrf_for(sid)):
        raise HTTPException(403, "This page has expired or the form is not valid. Reload the page and try again.")


def install(app, s: Settings) -> None:
    hosts = allowed_hosts(s)
    origins = {f"http://{h}" for h in hosts}

    @app.middleware("http")
    async def guard(request: Request, call_next):
        if request.headers.get("host", "") not in hosts:               # DNS rebinding
            return PlainTextResponse("Invalid Host header.", 400, headers=HEADERS)
        if request.method not in ("GET", "HEAD", "OPTIONS"):
            origin = request.headers.get("origin") or ""
            if not origin or origin == "null":                          # some browsers omit Origin: fall back to Referer
                ref = urlsplit(request.headers.get("referer") or "")
                origin = f"{ref.scheme}://{ref.netloc}" if ref.scheme else ""
            if origin not in origins:
                return PlainTextResponse("Cross-site request refused.", 403, headers=HEADERS)
        sid = request.cookies.get(COOKIE) or secrets.token_urlsafe(24)
        new = COOKIE not in request.cookies
        request.state.sid = sid                                          # state lives in the ASGI scope, shared with the route's Request
        response = await call_next(request)
        for k, v in HEADERS.items():
            response.headers[k] = v
        if new:
            response.set_cookie(COOKIE, sid, httponly=True, samesite="strict", path="/")
        return response


def safe_output_file(s: Settings, folder: str, rel: str):
    """The file under one of the output folders, or None. No '..', no absolute paths, no drive letters, only the
    whitelisted extensions, and the fully resolved path (symlinks and junctions followed) must stay inside the folder."""
    if folder not in OUTPUT_FOLDERS or not rel or "\x00" in rel:
        return None
    if rel.startswith(("/", "\\")) or ":" in rel or any(part in ("..", "") for part in rel.replace("\\", "/").split("/")):
        return None
    root = (s.project / folder).resolve()
    try:
        target = (root / rel).resolve(strict=True)
    except (OSError, ValueError):
        return None
    if not target.is_relative_to(root) or not target.is_file() or target.suffix.lower() not in DOWNLOAD_EXTENSIONS:
        return None
    return target
