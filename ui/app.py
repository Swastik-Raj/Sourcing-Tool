import os
import secrets
import sys
import urllib.request
from datetime import datetime
from pathlib import Path
from typing import List, Optional
from urllib.parse import urlencode

from fastapi import FastAPI, File, Form, HTTPException, Request, UploadFile
from fastapi.responses import FileResponse, HTMLResponse, JSONResponse, RedirectResponse
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates

from . import pipeline, search, security, uploads, wording
from .commands import JobRejected
from .config import (CAPABILITIES, ENV_VARS, MAX_FACTS_BYTES, MAX_UPLOAD_BYTES, OUTPUT_FOLDERS, Settings, changed_reference_files, is_optional,
                     is_set, load_settings, plain_names, tracing_on)
from .jobs import JobBusy, JobManager, mask_line
from .workflows import STEP_LABELS, STEP_TEXT, STEPS, Step, list_workflows, next_step, result_rows, search_detail, steps_for, upload_time

HERE = Path(__file__).resolve().parent
STEP_SLUGS = {name.lower().replace(" ", "-"): name for name in STEPS}
JOB_BADGES = {"running": "Running", "succeeded": "Done", "failed": "Failed", "cancelled": "Cancelled", "interrupted": "Interrupted"}


def badge_kind(label: str) -> str:
    return label.lower().replace(" ", "-")


def create_app(settings: Optional[Settings] = None, popen=None) -> FastAPI:
    """`settings` is how tests point the UI at a throwaway project folder; `popen` lets them see a command instead of
    running it. run_ui.py passes neither."""
    s = settings or load_settings()
    app = FastAPI(docs_url=None, redoc_url=None, openapi_url=None)
    jobs = JobManager(s, popen)
    app.state.jobs, app.state.settings = jobs, s
    templates = Jinja2Templates(directory=str(HERE / "templates"))        # autoescape is on for .html
    templates.env.finalize = wording.display                               # every printed value gets display wording (attributes too)
    app.state.templates = templates                                        # tests render through this same environment
    templates.env.globals.update(app_name=s.app_name, badge_kind=badge_kind, job_badge=lambda st: JOB_BADGES.get(st, st))
    templates.env.filters.update(upload_time=upload_time, plain_names=plain_names)
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

    def running_search():
        r = jobs.running_state_job()
        return r if r and r.get("command") == "search" else None

    def render(request: Request, name: str, status: int = 200, active: str = "", wf_override: str = "", **ctx):
        wfs = list_workflows(s)
        stems = [w["stem"] for w in wfs]
        wf = wf_override or request.query_params.get("wf")
        wf = wf if wf in stems else (stems[0] if stems else None)
        current = next((w for w in wfs if w["stem"] == wf), None)
        steps = steps_for(s, current["file"]) if current else []
        rs = running_search()
        if rs and steps:                    # a running search shows on its own workflow, or on any until its file exists
            out = rs["search"]["output"]
            if current["file"] == out or not (s.results_dir / out).exists():
                steps[0] = Step("Search", "Running", "A search is running. Each product's result is saved as soon as it finishes.")
        here = STEPS.index(ctx["step_name"]) if ctx.get("step_name") in STEPS else None
        base = dict(request=request, csrf=security.csrf_for(getattr(request.state, "sid", "")), workflows=wfs, wf=wf,
                    current=current, steps=steps, running=jobs.running_state_job(), active=active, page_title=ctx.pop("page_title", ""),
                    next_step=next_step(steps),                      # Home's Continue: the first step whose output the files do not show
                    following=steps[here + 1] if here is not None and here + 1 < len(steps) else None)   # a step page's Continue
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
        back = {name: slug for slug, name in STEP_SLUGS.items()}.get(request.query_params.get("step", ""))
        if back and back != "search":       # the results file selector on a tab: stay on that tab (wf is checked again there)
            return RedirectResponse(f"/step/{back}?{urlencode({'wf': request.query_params.get('wf', '')})}", 303)
        return render(request, "home.html", active="workflow", page_title="Overview", recent=wfs[:5], all_count=len(wfs))

    @app.get("/step/{slug}", response_class=HTMLResponse)
    def step(request: Request, slug: str):
        old = redirect_old_prefix(request)
        if old:
            return old
        if slug not in STEP_SLUGS:
            raise HTTPException(404, "There is no such page.")
        if slug == "search":
            return search_page(request)
        return step_page(request, slug)

    # ---- the tabs after Search: one results file, the agents' own files, allowlisted commands ----
    STEP_PAGES = {"emails": "emails.html", "approvals": "approvals.html", "order-sheet": "order.html", "listing-content": "content.html"}

    def results_for(wf: str) -> str:
        """The results file a form names by its stem. Only a file listed from 'Excel Output Sheets' is ever accepted."""
        for w in list_workflows(s):
            if w["stem"] == wf:
                return w["file"]
        raise JobRejected("Choose a search from the list. Nothing was started.")

    def step_page(request: Request, slug: str, status: int = 200, error: str = "", wf: str = "", **ctx):
        name = STEP_SLUGS[slug]
        wfs = list_workflows(s)
        asked = wf or request.query_params.get("wf", "")
        chosen = next((w for w in wfs if w["stem"] == asked), wfs[0] if wfs else None)       # as render() picks it
        data = {}
        if chosen and slug == "emails":
            data["view"] = pipeline.emails(s, chosen["file"])
        elif chosen and slug == "approvals":
            data["view"] = pipeline.approvals(s, chosen["file"])
        elif chosen and slug == "order-sheet":
            data["view"] = pipeline.order(s, chosen["file"])
        elif chosen and slug == "listing-content":
            data["view"] = pipeline.content(s, chosen["file"], ctx.pop("facts", "") or request.query_params.get("facts", ""),
                                            count=request.method == "GET" and request.query_params.get("count") == "1")
        job = jobs.get(request.query_params.get("job", "")) if request.method == "GET" else None
        if job and job.get("step") == slug:
            data["job"] = job
            data["job_log"], data["job_offset"] = jobs.read_log(job["id"], 0, whole_lines_only=job["status"] == "running")
        return render(request, STEP_PAGES[slug], status, active="workflow", page_title=STEP_LABELS[name], step_name=name,
                      step_text=STEP_TEXT[name], error=error, wf_override=chosen["stem"] if chosen else "", **{**data, **ctx})

    def act(request: Request, slug: str, wf: str, cmd: str, params: dict, input_file: str = ""):
        """Start one allowlisted command for a tab, then show that tab with the command's live log."""
        try:
            meta = jobs.start(cmd, params, extra={"step": slug, "wf": wf, **({"input_file": input_file} if input_file else {})})
        except (JobBusy, JobRejected) as e:
            jobs.discard_input(input_file)
            return step_page(request, slug, 409 if isinstance(e, JobBusy) else 400, f"Nothing was started. {e}", wf)
        return RedirectResponse(f"/step/{slug}?{urlencode({'wf': wf, 'job': meta['id']})}", 303)

    def refused(request: Request, slug: str, wf: str, why, status: int = 400, **ctx):
        return step_page(request, slug, status, str(why), wf, **ctx)

    # ---- step 2: emails ----
    @app.post("/emails/drafts")
    def emails_drafts(request: Request, csrf_token: str = Form(""), wf: str = Form(""), near_miss: str = Form("")):
        csrf(request, csrf_token)
        try:
            file = results_for(wf)
        except JobRejected as e:
            return refused(request, "emails", wf, e)
        return act(request, "emails", wf, "emails.drafts", {"results": file, **({"near_miss": near_miss} if near_miss else {})})

    def send_plan(wf: str, test_recipient: str) -> tuple:
        file = results_for(wf)
        test_recipient = test_recipient.strip()
        if test_recipient:
            try:
                jobs.commands["emails.send"].params["test_recipient"].rule(test_recipient)
            except JobRejected as e:
                raise JobRejected(f"{e} Nothing was sent.")
            if test_recipient.lower() in {d["email"].lower() for d in pipeline.emails(s, file)["drafts"] if d["email"]}:
                raise JobRejected("A test address cannot be one of the sellers' own addresses. Nothing was sent.")
        return file, pipeline.send_plan(s, file, test_recipient)

    @app.post("/emails/send/review", response_class=HTMLResponse)
    def emails_send_review(request: Request, csrf_token: str = Form(""), wf: str = Form(""), test_recipient: str = Form("")):
        csrf(request, csrf_token)
        try:
            _, plan = send_plan(wf, test_recipient)
        except JobRejected as e:
            return refused(request, "emails", wf, e)
        return step_page(request, "emails", wf=wf, plan=plan)

    @app.post("/emails/send")
    def emails_send(request: Request, csrf_token: str = Form(""), wf: str = Form(""), test_recipient: str = Form(""),
                    confirm_count: str = Form("")):
        csrf(request, csrf_token)
        try:
            file, plan = send_plan(wf, test_recipient)
        except JobRejected as e:
            return refused(request, "emails", wf, e)
        typed = confirm_count.strip()
        if not plan["count"]:
            return refused(request, "emails", wf, "There is nothing to send for this search. Nothing was sent.")
        if not typed.isdigit() or int(typed) != plan["count"]:          # the server's own count; the agent checks it again with --confirm-count
            return refused(request, "emails", wf, "That is not the number of emails listed. Nothing was sent.", plan=plan)
        params = {"results": file, "confirm_count": str(plan["count"])}
        if plan["test"]:
            params["test_recipient"] = plan["test_recipient"]
        return act(request, "emails", wf, "emails.send", params)

    def known_sku(file: str, sku: str) -> str:
        if sku.strip().upper() not in pipeline.draft_skus(s, file):
            raise JobRejected("Choose a product from this search's sellers. Nothing was stored.")
        return sku.strip()

    @app.post("/emails/mark")
    def emails_mark(request: Request, csrf_token: str = Form(""), wf: str = Form(""), sku: str = Form(""), confirm: str = Form(""),
                    followup: str = Form("")):
        """Marks the inquiry as sent through the seller's website, or (followup=1) the drafted follow-up questions as sent by
        hand. Either needs the person's own tick."""
        csrf(request, csrf_token)
        try:
            file = results_for(wf)
            sku = known_sku(file, sku)
            if confirm != "1":
                raise JobRejected("Tick the box to confirm you sent the questions to the seller yourself. Nothing was stored." if followup else
                                  "Tick the box to confirm you contacted the seller through their website yourself. Nothing was stored.")
        except JobRejected as e:
            return refused(request, "emails", wf, e)
        return act(request, "emails", wf, "emails.mark", {"sku": sku, "results": file, **({"followup": followup} if followup else {})})

    def pasted(text: str, what: str) -> str:
        if not text.strip():
            raise JobRejected(f"Paste {what}. Nothing was stored.")
        return jobs.save_input(text)

    @app.post("/emails/reply")
    def emails_reply(request: Request, csrf_token: str = Form(""), wf: str = Form(""), sku: str = Form(""), reply: str = Form("")):
        csrf(request, csrf_token)
        try:
            file = results_for(wf)
            sku = known_sku(file, sku)
            if not is_set(s, "ANTHROPIC_API_KEY"):
                raise JobRejected("Reading a reply needs the AI service key (ANTHROPIC_API_KEY), which is not set up. Nothing was stored.")
            name = pasted(reply, "the seller's reply")
        except JobRejected as e:
            return refused(request, "emails", wf, e)
        return act(request, "emails", wf, "emails.reply", {"sku": sku, "results": file, "file": name}, input_file=name)

    # ---- step 3: approvals ----
    @app.post("/approvals/request")
    def approvals_request(request: Request, csrf_token: str = Form(""), wf: str = Form(""), budget: str = Form(""), ack: str = Form("")):
        csrf(request, csrf_token)
        try:
            file = results_for(wf)
            if pipeline.approvals(s, file)["no_email_record"] and ack != "1":
                raise JobRejected("No seller has been contacted yet, so the request would say that seller replies are unavailable. "
                                  "Tick the box to create the request anyway. Nothing was written.")
        except JobRejected as e:
            return refused(request, "approvals", wf, e)
        return act(request, "approvals", wf, "approvals.request", {"results": file, **({"budget": budget.strip()} if budget.strip() else {})})

    @app.post("/approvals/record")
    def approvals_record(request: Request, csrf_token: str = Form(""), wf: str = Form(""), request_id: str = Form(""), reply: str = Form(""),
                         confirm_request: str = Form(""), channel: str = Form(""), approver: str = Form("")):
        csrf(request, csrf_token)
        try:
            file = results_for(wf)
            if request_id not in {r["id"] for r in pipeline.approvals(s, file)["open"]}:
                raise JobRejected("Pick the request this answer is for, from those still waiting for an answer. Nothing was recorded.")
            given = {"results": file, "request": request_id, "confirm_request": confirm_request.strip().upper(), "channel": channel, "approver": approver}
            cmd = jobs.commands["approvals.record"]                       # every typed field is checked before the pasted text is saved
            try:
                for field in ("confirm_request", "channel", "approver"):
                    cmd.params[field].rule(given[field])
                cmd.check(given)
            except JobRejected as e:
                raise JobRejected(str(e) if "Nothing was recorded" in str(e) else f"{e} Nothing was recorded.")
            given["file"] = pasted(reply, "the manager's reply")
        except JobRejected as e:
            return refused(request, "approvals", wf, e)
        return act(request, "approvals", wf, "approvals.record", given, input_file=given["file"])

    @app.post("/approvals/revoke")
    def approvals_revoke(request: Request, csrf_token: str = Form(""), wf: str = Form(""), sku: str = Form("")):
        csrf(request, csrf_token)
        try:
            file = results_for(wf)
            if sku.strip() not in pipeline.approvals(s, file)["valid_skus"]:
                raise JobRejected("Type the product code (SKU) of an approval that is still valid for this search. Nothing was withdrawn.")
        except JobRejected as e:
            return refused(request, "approvals", wf, e)
        return act(request, "approvals", wf, "approvals.revoke", {"sku": sku.strip(), "results": file})

    @app.post("/approvals/export")
    def approvals_export(request: Request, csrf_token: str = Form(""), wf: str = Form("")):
        csrf(request, csrf_token)
        try:
            file = results_for(wf)
        except JobRejected as e:
            return refused(request, "approvals", wf, e)
        return act(request, "approvals", wf, "approvals.export", {"results": file})

    # ---- step 4: order sheet and listing content, both from the export made for THIS results file ----
    @app.post("/order/build")
    def order_build(request: Request, csrf_token: str = Form(""), wf: str = Form("")):
        csrf(request, csrf_token)
        try:
            file = results_for(wf)
            pipeline.export_gate(s, file)
        except JobRejected as e:
            return refused(request, "order-sheet", wf, e)
        return act(request, "order-sheet", wf, "order.build", {"results": file})

    @app.post("/content/template")
    def content_template(request: Request, csrf_token: str = Form(""), wf: str = Form("")):
        csrf(request, csrf_token)
        try:
            file = results_for(wf)
            pipeline.export_gate(s, file)
            s.uploads_dir.mkdir(parents=True, exist_ok=True)           # the content agent does not create the folder it writes into
        except JobRejected as e:
            return refused(request, "listing-content", wf, e)
        name = f"facts_template_{datetime.now():%Y%m%d_%H%M%S}_{secrets.token_hex(3)}.csv"
        return act(request, "listing-content", wf, "content.template", {"results": file, "facts": name})

    @app.post("/content/facts")
    def content_facts(request: Request, csrf_token: str = Form(""), wf: str = Form(""), facts_file: UploadFile = File(None)):
        csrf(request, csrf_token)
        if facts_file is None or not facts_file.filename:
            return refused(request, "listing-content", wf, "No file was chosen. Nothing was uploaded.")
        try:
            columns = pipeline.facts_columns(s)
            if not columns:
                raise JobRejected("The listing tool's column list could not be read, so the product details sheet cannot be checked.")
            name = uploads.save_facts(s, facts_file.filename, facts_file.file.read(MAX_FACTS_BYTES + 1), columns)
        except JobRejected as e:
            return refused(request, "listing-content", wf, f"Nothing was uploaded. {e}")
        return RedirectResponse(f"/step/listing-content?{urlencode({'wf': wf, 'facts': name, 'count': '1'})}", 303)   # count it right away

    def generate_plan(wf: str, facts: str) -> tuple:
        file = results_for(wf)
        pipeline.export_gate(s, file)
        view = pipeline.content(s, file, facts)
        if not facts or view["facts"] != facts:
            raise JobRejected("Choose a product details sheet you uploaded. Listings are only written from details a person filled in. Nothing was started.")
        if view["missing"]:
            raise JobRejected(f"Nothing was started: {plain_names(view['missing'])} {'is' if len(view['missing']) == 1 else 'are'} not set up "
                              f"({', '.join(view['missing'])}).")
        if not view["count"]:
            raise JobRejected("The approved list holds no product. Nothing was started.")
        return file, view

    @app.post("/content/review", response_class=HTMLResponse)
    def content_review(request: Request, csrf_token: str = Form(""), wf: str = Form(""), facts: str = Form("")):
        csrf(request, csrf_token)
        try:
            generate_plan(wf, facts)
        except JobRejected as e:
            return refused(request, "listing-content", wf, e, facts=facts)
        return step_page(request, "listing-content", wf=wf, facts=facts, confirm=True)

    @app.post("/content/generate")
    def content_generate(request: Request, csrf_token: str = Form(""), wf: str = Form(""), facts: str = Form(""), confirm_count: str = Form("")):
        csrf(request, csrf_token)
        try:
            file, view = generate_plan(wf, facts)
        except JobRejected as e:
            return refused(request, "listing-content", wf, e, facts=facts)
        typed = confirm_count.strip()
        if not typed.isdigit() or int(typed) != view["count"]:
            return refused(request, "listing-content", wf, "That is not the number of products listed. Nothing was started.", facts=facts, confirm=True)
        return act(request, "listing-content", wf, "content.generate", {"results": file, "facts": facts})

    # ---- step 1: search ----
    def search_job(jid: str):
        meta = jobs.get(jid)
        return meta if meta and meta.get("command") == "search" and isinstance(meta.get("search"), dict) else None

    def elapsed_s(meta: dict) -> int:
        try:
            return max(int((datetime.now() - datetime.fromisoformat(meta["started_at"])).total_seconds()), 0)
        except (KeyError, ValueError):
            return 0

    def progress_for(meta: dict) -> dict:
        text, _ = jobs.read_log(meta["id"], 0, whole_lines_only=meta["status"] == "running")
        p = search.parse_progress(text)
        p["expected"] = meta["search"].get("expected")
        p["total"] = p["total"] or p["expected"]
        p["elapsed_s"] = elapsed_s(meta) if meta["status"] == "running" else int(meta.get("duration_s") or 0)
        p["current"] = [wording.display(c) for c in p["current"]]
        return p

    def search_page(request: Request, status: int = 200, error: str = "", review: Optional[dict] = None):
        common = dict(active="workflow", page_title=STEP_LABELS["Search"], step_name="Search", step_text=STEP_TEXT["Search"], error=error)
        asked = request.query_params.get("job", "")
        job = search_job(asked) if asked else None
        if job is None and review is None and not error and running_search():
            return RedirectResponse(f"/step/search?job={running_search()['id']}", 303)
        if job:
            info, out = job["search"], s.results_dir / job["search"]["output"]
            progress = progress_for(job)
            log, offset = jobs.read_log(job["id"], 0, whole_lines_only=job["status"] == "running")
            table = result_rows(s, out) if out.is_file() else None
            in_file = len(table["rows"]) if table else 0
            done_ok = job["status"] == "succeeded" and not progress["billing"]
            report = None
            if job["status"] != "running":                         # a copy for the download link; the original is never touched
                report = search.copy_report(s, jobs.read_log(job["id"])[0])
                report["name"] = report["name"] or progress["report"]
            return render(request, "search.html", status, wf_override=out.stem if out.is_file() else "",
                          view="running" if job["status"] == "running" else "finished", job=job, info=info, progress=progress,
                          log=log, offset=offset, table=table, in_file=in_file,
                          partial=bool(table) and (not done_ok or in_file < (info.get("expected") or 0)),
                          counts=search_detail(s, out)[1] if table else "", reading=search.reading(job["status"], job["exit_code"], progress),
                          results_file=out.name if out.is_file() else "", report=report, **common)
        keys = [{"name": n, "set": is_set(s, n), "blocks": why} for n, why in search.REQUIRED_KEYS]
        builtin = search.builtin_products(s)
        have = {sku.upper() for sku, _ in builtin}
        view, preview, asked_file = "review" if review else "form", None, request.query_params.get("file", "")
        if asked_file and not review and not error and request.method == "GET":       # a sheet's preview and its own choice of rows
            try:
                view, preview = "preview", uploads.preview(s, asked_file)
            except JobRejected as e:
                status, common["error"] = 400, str(e)
        return render(request, "search.html", status, view=view, keys=keys,
                      missing=[k["name"] for k in keys if not k["set"]], busy=jobs.running_state_job(), builtin=builtin,
                      preset=search.PRESET_SKUS, preset_missing=[k for k in search.PRESET_SKUS if k not in have],
                      files=search.input_files(s), uploaded=search.upload_files(s), max_limit=search.MAX_LIMIT,
                      max_mb=MAX_UPLOAD_BYTES // (1024 * 1024), review=review, preview=preview, preview_file=asked_file,
                      preview_label=(f"your spreadsheet uploaded {upload_time(asked_file) or asked_file[len(search.UPLOAD_PREFIX):]}"
                                     if asked_file.startswith(search.UPLOAD_PREFIX) else asked_file),
                      **common)

    def selection(choice: str, sku: list, input_file: str, limit: str, add_builtin: str, rows: str = "") -> dict:
        if rows == "all":                    # the sheet preview form: one radio decides which of its fields count
            sku, limit = [], ""
        elif rows == "first":
            sku = []
        elif rows == "pick":
            limit, add_builtin = "", ""
            if not sku:
                raise JobRejected("No products are ticked. Tick at least one product.")
        sel = search.resolve_selection(s, choice, sku, input_file, limit, add_builtin)
        return {**sel, "choice": choice, "input_file": input_file, "limit": limit, "add_builtin": sel["params"].get("add_builtin", ""),
                "estimate": search.estimate(len(sel["products"]))}

    @app.post("/search/upload")
    def search_upload(request: Request, csrf_token: str = Form(""), sheet: UploadFile = File(None)):
        csrf(request, csrf_token)
        if sheet is None or not sheet.filename:
            return search_page(request, 400, error="No file was chosen. Nothing was uploaded.")
        try:
            value = uploads.save_workbook(s, sheet.filename, sheet.file.read(MAX_UPLOAD_BYTES + 1))
        except JobRejected as e:
            return search_page(request, 400, error=f"Nothing was uploaded. {e}")
        return RedirectResponse(f"/step/search?{urlencode({'file': value})}", 303)

    @app.post("/search/review", response_class=HTMLResponse)
    def search_review(request: Request, csrf_token: str = Form(""), choice: str = Form(""), sku: List[str] = Form(default=[]),
                      input_file: str = Form(""), limit: str = Form(""), add_builtin: str = Form(""), rows: str = Form("")):
        csrf(request, csrf_token)
        try:
            return search_page(request, review=selection(choice, sku, input_file, limit, add_builtin, rows))
        except JobRejected as e:
            return search_page(request, 400, error=str(e))

    @app.post("/search/start")
    def search_start(request: Request, csrf_token: str = Form(""), choice: str = Form(""), sku: List[str] = Form(default=[]),
                     input_file: str = Form(""), limit: str = Form(""), add_builtin: str = Form(""), rows: str = Form(""),
                     confirm_count: str = Form("")):
        csrf(request, csrf_token)
        try:
            sel = selection(choice, sku, input_file, limit, add_builtin, rows)   # recomputed here: nothing the browser says about the count is used
        except JobRejected as e:
            return search_page(request, 400, error=str(e))
        missing = [n for n, _ in search.REQUIRED_KEYS if not is_set(s, n)]
        if missing:
            return search_page(request, 400, error=f"Nothing was started: {plain_names(missing)} {'is' if len(missing) == 1 else 'are'} "
                                                   f"not set up ({', '.join(missing)}).")
        count, typed = len(sel["products"]), confirm_count.strip()
        if not typed.isdigit() or int(typed) != count:
            return search_page(request, 400, review=sel, error="That is not the number of products selected. Nothing was started.")
        busy = jobs.running_state_job()                           # say so before anything else; jobs.start still takes the lock itself
        if busy:
            return search_page(request, 409, error="Nothing was started. Another task that changes data is running. Wait for it to finish or cancel it.")
        name = search.new_output_name()
        try:
            s.results_dir.mkdir(parents=True, exist_ok=True)       # the agent does not create the folder it saves into
            meta = jobs.start("search", {**sel["params"], "output": name},
                              extra={"search": {"output": name, "expected": count, "selection": sel["label"]}})
        except JobBusy as e:
            return search_page(request, 409, error=f"Nothing was started. {e}")
        except JobRejected as e:
            return search_page(request, 400, error=f"Nothing was started. {e}")
        return RedirectResponse(f"/step/search?job={meta['id']}", 303)

    @app.get("/health", response_class=HTMLResponse)
    def health(request: Request):
        env = [{"name": n, "set": is_set(s, n), "blocks": why, "optional": is_optional(n)} for n, why in ENV_VARS]
        filled = {e["name"]: e["set"] for e in env}
        capabilities = [(what, all(filled[n] for n in needs)) for what, needs in CAPABILITIES]
        folders = [(n, (s.project / n).is_dir()) for n in OUTPUT_FOLDERS]
        obs_on = os.environ.get("OBS_ENABLED") == "1"          # only for the reachability check below, as before
        return render(request, "health.html", active="health", page_title="Setup", env=env, folders=folders, obs_on=obs_on,
                      capabilities=capabilities, all_set=all(ok for _, ok in capabilities), tracing=tracing_on(s),
                      python=sys.version.split()[0], langfuse=langfuse_status() if obs_on else None,
                      project_ok=s.project.is_dir(), data_ok=s.data.is_dir(), zone=pipeline.zone_label(),
                      changed=changed_reference_files(s))

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
        jobs.prune()
        return render(request, "jobs.html", active="jobs", page_title="Activity", jobs=jobs.list(),
                      commands=[c for c in jobs.commands.values() if not c.state_changing])

    @app.post("/jobs/run")
    def job_run(request: Request, command: str = Form(""), csrf_token: str = Form(""), mode: str = Form("")):
        csrf(request, csrf_token)
        try:
            if not command.startswith(("help.", "stub.")):      # every other command has its own route, with its gate
                raise JobRejected("That command is not run from this page.")
            meta = jobs.start(command, {"mode": mode} if mode else {})
        except (JobRejected, JobBusy) as e:
            raise HTTPException(409 if isinstance(e, JobBusy) else 400, str(e))
        return RedirectResponse(f"/jobs/{meta['id']}", 303)

    @app.get("/jobs/{jid}", response_class=HTMLResponse)
    def job_view(request: Request, jid: str):
        meta = jobs.get(jid)
        if meta is None:
            raise HTTPException(404, "No such task.")
        text, offset = jobs.read_log(jid, 0, whole_lines_only=meta["status"] == "running")
        return render(request, "job.html", active="jobs", page_title="Task", job=meta, log=text, offset=offset)

    @app.get("/jobs/{jid}/log")
    def job_log(jid: str, offset: int = 0):
        meta = jobs.get(jid)
        if meta is None:
            raise HTTPException(404, "No such task.")
        text, new = jobs.read_log(jid, max(offset, 0), whole_lines_only=meta["status"] == "running")
        body = {"text": text, "offset": new, "status": meta["status"]}
        if search_job(jid):
            body["progress"] = progress_for(meta)
        return JSONResponse(wording.display_all(body))          # every string in it, as a page prints it

    @app.post("/jobs/{jid}/cancel")
    def job_cancel(request: Request, jid: str, csrf_token: str = Form("")):
        csrf(request, csrf_token)
        meta = jobs.get(jid)                    # read once, before cancelling: the runner rewrites this record as the job stops
        if meta is None:
            raise HTTPException(404, "No such task.")
        jobs.cancel(jid)
        if meta.get("step") in STEP_SLUGS:
            return RedirectResponse(f"/step/{meta['step']}?{urlencode({'wf': meta.get('wf', ''), 'job': jid})}", 303)
        searching = meta.get("command") == "search" and isinstance(meta.get("search"), dict)       # search_job()'s own test
        return RedirectResponse(f"/step/search?job={jid}" if searching else f"/jobs/{jid}", 303)

    # ---- files ----
    @app.get("/files/project/{name}")
    def download_export(name: str):
        """The decision agent writes its export into the project folder itself. Exactly that one file is served from there."""
        target = security.safe_project_file(s, name)
        if target is None:
            raise HTTPException(404, "File not found.")
        return FileResponse(target, filename=target.name)

    @app.get("/files/{folder}/{rel:path}")
    def download(folder: str, rel: str):
        target = security.safe_output_file(s, folder, rel)
        if target is None:
            raise HTTPException(404, "File not found.")
        return FileResponse(target, filename=target.name)

    return app
