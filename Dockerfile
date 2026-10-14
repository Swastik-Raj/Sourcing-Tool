# Sourcing tool: the web UI and the agents in one image.
#
#   docker compose build      builds the runtime target only (see docker-compose.yml)
#   docker/verify.ps1         also builds the test target, to run the nine suites inside a container
#
# Nothing secret is in this image: keys and mail settings come from the environment at run time, and all data lives in
# the folder mounted at /data. Files are copied by name, never with "COPY . .", and .dockerignore lets through only the
# files named there.

# ---------- base: Python and the pinned packages ----------
# Python 3.13.5 is the version the suites were run on. The digest pins the exact image behind the tag.
FROM python:3.13.5-slim-bookworm@sha256:4c2cf9917bd1cbacc5e9b07320025bdb7cdf2df7b0ceaccb55e9dd7e30987419 AS base
LABEL com.zync.lcom="1"
ENV PYTHONUTF8=1 \
    PYTHONUNBUFFERED=1 \
    PIP_DISABLE_PIP_VERSION_CHECK=1 \
    HOME=/tmp
WORKDIR /app
# requirements.txt pins the packages the code imports; docker/constraints.txt pins everything those pull in.
COPY requirements.txt /app/requirements.txt
COPY docker/constraints.txt /app/docker/constraints.txt
RUN pip install --no-cache-dir -r requirements.txt -c docker/constraints.txt
# An ordinary user with no password and no shell, and an empty /data for the data folder to be mounted on.
# There is no VOLUME line on purpose: without a mounted folder the container must refuse to start, not hide the data.
RUN groupadd --gid 10001 lcom \
 && useradd --uid 10001 --gid 10001 --no-create-home --home-dir /tmp --shell /usr/sbin/nologin lcom \
 && mkdir /data \
 && chown 10001:10001 /data

# ---------- runtime: what colleagues run ----------
FROM base AS runtime
COPY run_ui.py \
     sourcing_agent.py email_agent.py decision_agent.py order_sheet.py content_agent.py \
     observability.py statefile.py \
     email_template.txt reviewer_exclusions.csv near_misses.csv lcom_prices.csv \
     /app/
COPY ui/__init__.py ui/app.py ui/commands.py ui/config.py ui/jobs.py ui/pipeline.py ui/search.py ui/security.py \
     ui/uploads.py ui/wording.py ui/workflows.py \
     /app/ui/
COPY ui/templates/ /app/ui/templates/
COPY ui/static/ /app/ui/static/
COPY docker/entrypoint.py /app/docker/entrypoint.py
# After the packages are installed (they keep their compiled files): the agents run from the data folder and must not
# leave __pycache__ folders in it.
ENV PYTHONDONTWRITEBYTECODE=1
USER 10001:10001
HEALTHCHECK --interval=30s --timeout=5s --start-period=20s --retries=3 \
    CMD ["python", "-c", "import os, urllib.request as u; u.urlopen('http://127.0.0.1:' + os.environ.get('UI_PORT', '8000') + '/static/app.css', timeout=3)"]
ENTRYPOINT ["python", "/app/docker/entrypoint.py"]

# ---------- test: the runtime image plus the suites (never shipped, built only by docker/verify.ps1) ----------
FROM runtime AS test
COPY ui/stub_agent.py /app/ui/stub_agent.py
COPY test_pricing.py test_email_agent.py test_decision_agent.py test_gates.py test_order_sheet.py test_content_agent.py \
     test_observability.py test_ui.py .env.example \
     /app/
COPY tests/test_ui_pipeline.py /app/tests/test_ui_pipeline.py
COPY tests/shim/ /app/tests/shim/
COPY tests/fixtures/ /app/tests/fixtures/
COPY test_fixtures/ /app/test_fixtures/
# test_gates.py compares today's decision and email agents with these two earlier copies.
COPY scratch/pre_gates/decision_agent.py scratch/pre_gates/email_agent.py /app/scratch/pre_gates/
