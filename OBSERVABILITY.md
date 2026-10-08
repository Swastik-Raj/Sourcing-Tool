# Observability (Langfuse)

One shared module, `observability.py`, traces every agent. Off by default. The agents themselves are unchanged: each has a few
`obs.*` calls around the code that was already there, and no agent depends on Langfuse or LangChain.

## Turn it on

```powershell
pip install -r requirements.txt     # the tracing pins are in the same file
$env:LANGFUSE_PUBLIC_KEY = "pk-lf-..."
$env:LANGFUSE_SECRET_KEY = "sk-lf-..."
$env:LANGFUSE_BASE_URL   = "https://cloud.langfuse.com"     # or https://us.cloud.langfuse.com, or your own server
$env:OBS_ENABLED         = "1"
```

The same variables work in `.env` (only `LANGFUSE_*` and `OBS_*` are read from it). `LANGFUSE_HOST` is accepted for
`LANGFUSE_BASE_URL`. **Off** = `OBS_ENABLED` is not `1`, or either key is missing. Then nothing is imported, no thread starts,
no socket opens, and every agent behaves exactly as before. To turn it off: `Remove-Item Env:OBS_ENABLED` (or set it to `0`).

| Variable | Meaning |
|---|---|
| `OBS_ENABLED` | `1` to trace. |
| `OBS_ENV` | `dev` (default). Becomes the `env:` tag and Langfuse's environment. Use `prod` for real runs. |
| `OBS_CAPTURE_CONTENT` | Comma list of agents whose prompt/response text is sent. **Replaces** the default list (empty = none). `decision_agent` and `order_sheet` can never be switched on. |
| `OBS_FLUSH_TIMEOUT` | Seconds the exit-time flush may take (default 5). |

When tracing is on, each run prints one line to stderr: `[obs] tracing ON for email_agent; prompt/response content capture: OFF; env: dev`.

## What links the runs together

`workflow_id = "lcom-" + <results file name without .xlsx>`, for example `lcom-sourcing_results_20261004_161054`. It is the
Langfuse **session id** of every command that touches that results file: `sourcing_agent` search (taken from the very path it
writes), `email_agent` drafts / send / mark / reply / report, `decision_agent` request / record / status / revoke / export,
`order_sheet build`, `content_agent` facts-template / generate. Each command is one trace named `<agent>.<command>`.

- Tags on a trace: `agent:<name>`, `env:<OBS_ENV>`, `sku:<SKU>` for each product it touched, `request:A-000N` (decision and
  order sheet), `outcome:<recommended|no_match|error>` (search), `rule_fail:<rule>` (content).
- Listing URLs are recorded as host only (`www.alibaba.com`), never the query string.
- `email_state.json` entries get an extra `workflow_id` key **only while tracing is on**. No existing key changes.
- `content_agent check <workbook>` is given no results file, so its session is `lcom-unlinked-<workbook name>`.

## What each agent records

| Agent | Spans and generations | Scores |
|---|---|---|
| `sourcing_agent` | trace `sourcing_agent.search` with run totals (matches the timing summary: wall time, Nimble calls and timeouts, tokens, cost, recommended / no_match / errored, Anthropic retries by status, Nimble rate-limit retries, billing stop). One span `product <SKU>` per product (keyword, candidates, exclusion rules by name with counts, which 80% bar candidates missed, tokens, cost by the agent's own price table as `agent_cost_usd`, stop reason, retries, billing stop). Child generation `claude.turn` per model call (model, input/output tokens, latency, stop reason, `agent_cost_usd`). Child span `nimble.<tool>` per Nimble call (ok/error/timeout, latency, response size, country and locale pins, rate-limit backoffs, URL host). | `sku_outcome` (categorical), `cost_usd` |
| `email_agent` | `email_agent.<cmd>`. drafts/send/report: drafts, with-email, skipped by kind, sent, blocked sends (never addresses). reply: span `summarize_reply` with the model call (tokens, latency) under it. | `verification_nulls`, `quote_warnings`, `currency_detected` |
| `decision_agent` | `decision_agent.<cmd>`: request id, lines, subtotal, lines left out, orphans, approvals, rejections, cap, decisions, expiries/revokes, rejected-reply **reason codes**. | none |
| `order_sheet` | `order_sheet.build`: rows, skipped, BELOW MOQ, large-total warnings, formula count, recalculation result, totals. | none |
| `content_agent` | `content_agent.<cmd>`; span `sku <SKU>` per product: status, model calls and cost, repair used, JSON retry used, pass/fail per validator rule, BLOCKED. The model call itself is captured by the Anthropic instrumentation. | `validator_failures` |

### Is the beta tool runner captured? No.
`sourcing_agent` calls `client.beta.messages.tool_runner(...)`. `opentelemetry-instrumentation-anthropic` 0.62.4 wraps
`Messages.create` / `.stream` (plain and beta) but a tool-runner run produced **no span** when tested (see
`test_tool_runner_is_not_captured_...`). So `sourcing_agent` records each model call itself (`claude.turn`). Its latency is the
time since the previous turn, so on turns after the first it includes the Nimble calls the model asked for (those appear as
child spans of the same product). The email and content agents use plain `messages.create`, which *is* captured. If a future
instrumentation version starts capturing the runner, that test fails on purpose, to stop double counting.

## What is sent

- **Prompt and response text**: `sourcing_agent`, `email_agent` and `content_agent` send it by default (masked as below).
  This assumes a self-hosted Langfuse that only the owner views: email replies and unreleased listing text are then visible
  to everyone in the project, so change `CAPTURE_CONTENT` at the top of `observability.py` (or set `OBS_CAPTURE_CONTENT=`
  to nothing) before using a shared or cloud project. `decision_agent` and `order_sheet` never send it, whatever the setting.
- **Always masked, before anything leaves the process**: API keys (Anthropic, Nimble, Langfuse, SMTP password), email
  addresses, phone numbers, the shipping address, and the approver name and channel typed into `decision_agent record`.
  This applies to attributes, inputs, outputs, exception messages, span events and Langfuse's own log lines. `.env` is never logged.
- `decision_agent` and `order_sheet` export only the exception **class** on an error (their messages quote SKUs and reply lines).

## Reading a workflow in Langfuse

1. **Sessions** → open `lcom-<results file name>`. All commands for that results file are listed in time order, even days apart.
2. Open a trace (`email_agent.reply`, `decision_agent.request`, ...) for its spans, attributes (Metadata) and scores.
3. Filter **Traces / Observations** by tag, e.g. `sku:HDFF` or `request:A-0003`, to follow one product across commands.
4. The search trace has one `product <SKU>` span per product; open it for `claude.turn` generations (tokens, latency) and `nimble.*` spans.
5. Cost: Langfuse computes cost from the model name and tokens. Compare with `agent_cost_usd` in the metadata (the agent's own
   price table). If Langfuse shows no cost, its model table has no price for `claude-haiku-4-5`; add a model definition
   (input $1 / output $5 per million tokens).

## If nothing appears

Check `[obs] tracing ON ...` printed on stderr (if absent, `OBS_ENABLED` / keys are not seen by that process). Confirm the keys
with `python -c "from langfuse import Langfuse; print(Langfuse().auth_check())"`. Check the base URL matches the project's region.
A line `[obs] flush did not finish in 5s` means the server was unreachable (firewall / proxy / wrong URL); the agent still finished normally.
Enabled tracing adds about 5 seconds to each command (importing Langfuse and starting its client); a command that calls the API
adds about 3 more for the Anthropic instrumentation.
