import os
import secrets
import sys
import urllib.request
from datetime import datetime
from pathlib import Path
from typing import Optional
from urllib.parse import urlencode

from fastapi import FastAPI, Form, HTTPException, Request
from fastapi.responses import FileResponse, HTMLResponse, JSONResponse, RedirectResponse
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates

from . import security, wording
from .commands import JobRejected
from .config import ENV_VARS, OUTPUT_FOLDERS, Settings, load_settings
from .jobs import JobBusy, JobManager, mask_line
from .workflows import STEP_TEXT, STEPS, list_workflows, steps_for

HERE = Path(__file__).resolve().parent
STEP_SLUGS = {name.lower().replace(" ", "-"): name for name in STEPS}
JOB_BADGES = {"running": "Running", "succeeded": "Done", "failed": "Failed", "cancelled": "Cancelled", "interrupted": "Interrupted"}


def badge_kind(label: str) -> str:
    return label.lower().replace(" ", "-")


def create_app(settings: Optional[Settings] = None) -> FastAPI:
    s = settings or load_settings()
    app = FastAPI(docs_url=None, redoc_url=None, openapi_url=None)
    jobs = JobManager(s)
    app.state.jobs, app.state.settings = jobs, s
    templates = Jinja2Templates(directory=str(HERE / "templates"))        # autoescape is on for .html
    templates.env.finalize = wording.display                               # every printed value gets display wording
    templates.env.globals.update(app_name=s.app_name, badge_kind=badge_kind, job_badge=lambda st: JOB_BADGES.get(st, st))
    app.mount("/static", StaticFiles(directory=str(HERE / "static")), name="static")
    security.install(app, s)

    def redirect_old_prefix(request: Request):
        """Old links carried the internal workflow prefix; send them to the plain-stem URL."""
        wf = request.query_params.get("wf", "")
        if wf.lower().startswith(wording.INTERNAL_PREFIX):
            q = {k: v for k, v in request.query_params.items() if k != "wf"}
            q["wf"] = wording.strip_prefix(wf)
            return RedirectResponse(f"{request.url.path}?{urlencode(q)}", 303)
        return None

    def render(request: Request, name: str, status: int = 200, active: str = "", **ctx):
        wfs = list_workflows(s)
        stems = [w["stem"] for w in wfs]
        wf = request.query_params.get("wf")
        wf = wf if wf in stems else (stems[0] if stems else None)
        current = next((w for w in wfs if w["stem"] == wf), None)
        steps = steps_for(s, current["file"]) if current else []
        base = dict(request=request, csrf=security.csrf_for(getattr(request.state, "sid", "")), workflows=wfs, wf=wf,
                    current=current, steps=steps, running=jobs.running_state_job(), active=active, page_title=ctx.pop("page_title", ""))
        return templates.TemplateResponse(request, name, {**base, **ctx}, status_code=status)

    def csrf(request: Request, token: str) -> None:
        security.check_csrf(request, token)

    @app.exception_handler(HTTPException)
    async def http_error(request: Request, exc: HTTPException):
        return render(request, "error.html", exc.status_code, message=str(exc.detail), error_id=None, page_title="Something went wrong")

    @app.exception_handler(Exception)
    async def crash(request: Request, exc: Exception):
        eid = secrets.token_hex(4)
        try:
            with open(s.data / "errors.log", "a", encoding="utf-8") as f:   # detail goes to the masked log only
                f.write(mask_line(f"{datetime.now():%Y-%m-%d %H:%M:%S} {eid} {request.method} {request.url.path} {type(exc).__name__}: {exc}") + "\n")
        except OSError:
            pass
        return render(request, "error.html", 500, message="Something went wrong. Nothing was changed by this page.", error_id=eid,
                      page_title="Something went wrong")

    # ---- pages ----
    @app.get("/", response_class=HTMLResponse)
    def home(request: Request):
        old = redirect_old_prefix(request)
        if old:
            return old
        wfs = list_workflows(s)
        return render(request, "home.html", active="workflow", page_title="Workflow", recent=wfs[:5], all_count=len(wfs))

    @app.get("/step/{slug}", response_class=HTMLResponse)
    def step(request: Request, slug: str):
        old = redirect_old_prefix(request)
        if old:
            return old
        if slug not in STEP_SLUGS:
            raise HTTPException(404, "No such step.")
        name = STEP_SLUGS[slug]
        return render(request, "step.html", active="workflow", page_title=name, step_name=name, step_text=STEP_TEXT[name])

    @app.get("/health", response_class=HTMLResponse)
    def health(request: Request):
        env = [{"name": n, "set": bool(os.environ.get(n)), "blocks": why} for n, why in ENV_VARS]
        folders = [(n, (s.project / n).is_dir()) for n in OUTPUT_FOLDERS]
        obs_on = os.environ.get("OBS_ENABLED") == "1"
        return render(request, "health.html", active="health", page_title="Health", env=env, folders=folders, obs_on=obs_on,
                      python=sys.version.split()[0], langfuse=langfuse_status() if obs_on else None,
                      project_ok=s.project.is_dir(), data_ok=s.data.is_dir())

    def langfuse_status() -> str:
        base = (os.environ.get("LANGFUSE_BASE_URL") or os.environ.get("LANGFUSE_HOST") or "").rstrip("/")
        if not base.startswith(("http://", "https://")):
            return "The tracing address is not set to an http(s) address."
        try:
            with urllib.request.urlopen(base + "/api/public/health", timeout=3) as r:
                return f"Reachable (HTTP {r.status})."
        except Exception as e:  # noqa: BLE001
            return f"Not reachable ({type(e).__name__})."

    # ---- jobs ----
    @app.get("/jobs", response_class=HTMLResponse)
    def job_list(request: Request):
        return render(request, "jobs.html", active="jobs", page_title="Jobs", jobs=jobs.list(),
                      commands=[c for c in jobs.commands.values() if not c.state_changing])

    @app.post("/jobs/run")
    def job_run(request: Request, command: str = Form(...), csrf_token: str = Form(""), mode: str = Form("")):
        csrf(request, csrf_token)
        try:
            meta = jobs.start(command, {"mode": mode} if mode else {})
        except (JobRejected, JobBusy) as e:
            raise HTTPException(409 if isinstance(e, JobBusy) else 400, str(e))
        return RedirectResponse(f"/jobs/{meta['id']}", 303)

    @app.get("/jobs/{jid}", response_class=HTMLResponse)
    def job_view(request: Request, jid: str):
        meta = jobs.get(jid)
        if meta is None:
            raise HTTPException(404, "No such job.")
        text, offset = jobs.read_log(jid, 0, whole_lines_only=meta["status"] == "running")
        return render(request, "job.html", active="jobs", page_title="Job", job=meta, log=text, offset=offset)

    @app.get("/jobs/{jid}/log")
    def job_log(jid: str, offset: int = 0):
        meta = jobs.get(jid)
        if meta is None:
            raise HTTPException(404, "No such job.")
        text, new = jobs.read_log(jid, max(offset, 0), whole_lines_only=meta["status"] == "running")
        return JSONResponse({"text": wording.display(text), "offset": new, "status": meta["status"]})

    @app.post("/jobs/{jid}/cancel")
    def job_cancel(request: Request, jid: str, csrf_token: str = Form("")):
        csrf(request, csrf_token)
        if jobs.get(jid) is None:
            raise HTTPException(404, "No such job.")
        jobs.cancel(jid)
        return RedirectResponse(f"/jobs/{jid}", 303)

    # ---- files ----
    @app.get("/files/{folder}/{rel:path}")
    def download(folder: str, rel: str):
        target = security.safe_output_file(s, folder, rel)
        if target is None:
            raise HTTPException(404, "File not found.")
        return FileResponse(target, filename=target.name)

    return app
