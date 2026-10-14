# Running the sourcing tool in Docker

> **Status:** `docker\verify.ps1` stages 1 and 2 pass, and stage 3 up to its proxy check (V14–V17). What it settled is
> written below; the rest has not been run against an image yet. Remove this note once all four stages have passed.

The container holds the web UI, the agents and the Python they need. Your keys and mail settings stay in a `.env` file on
your computer, and all your work stays in one folder on your computer (the **data folder**). Removing or rebuilding the
container never touches either.

All commands are PowerShell, run in the folder that holds `docker-compose.yml`.

## What you need

- Windows with Docker Desktop, started and running.
- This project's folder (the image is built from it; nothing is downloaded from a private registry).
- Internet access for the first build only (the Python base image and the pinned packages).

## First run

```powershell
Copy-Item .env.docker.example .env      # then open .env in a text editor and fill it in (see "Settings" below)
docker compose build                     # builds the image; several minutes the first time
docker compose up -d                     # starts the container in the background
docker compose ps                        # wait until STATUS says "healthy"
Start-Process "http://127.0.0.1:8010/"
```

Open the **Health** page first: it shows which settings are set (never their values), the time zone the pages use, and
whether any of your reference files differ from the ones that came with this version.

The UI answers only at `http://127.0.0.1:<port>/` and `http://localhost:<port>/`, and only from this computer. Any other
address (another name, another port, the computer's network name) gets "Invalid Host header." on purpose.

## Where your data lives

The data folder is `data` beside `docker-compose.yml`, unless `LCOM_DATA` in `.env` names another folder. Inside it:

| What | Names | Who writes it |
|---|---|---|
| Program files | `sourcing_agent.py`, `email_agent.py`, `decision_agent.py`, `order_sheet.py`, `content_agent.py`, `observability.py`, `statefile.py` | the container, at every start (this version's copies; do not edit them) |
| Reference files | `email_template.txt`, `reviewer_exclusions.csv`, `near_misses.csv`, `lcom_prices.csv` | the container adds one only when it is missing; after that the copy is yours |
| State | `email_state.json`, `decision_state.json`, `approvals_log.jsonl`, `approved_orders.json`, `approved_orders.csv` | the agents |
| Results and outputs | `Excel Output Sheets\`, `Reports\`, `Order Sheets\`, `Walmart Listings\`, `Uploads\`, `sourcing_report_*.md` | the agents and the UI |
| Job logs | `ui_data\jobs\` (secrets and addresses are masked) | the UI |

The agents keep their state beside their own script, which is why the program files are copied into the data folder.
Nothing else in the folder is opened, moved or deleted at start-up.

If you edited a reference file, your copy is kept and used, and the Health page lists it as different from this
version's. (The pages show `lcom_prices.csv` as `reference_prices.csv`.) To take this version's copy instead: stop the
tool, move your copy out of the data folder, start again.

## Start, stop, logs

```powershell
docker compose up -d          # start (or start again after a change to .env or a new version)
docker compose stop           # stop; the container is kept
docker compose start          # start the stopped container again
docker compose ps             # is it running and healthy?
docker compose logs --tail 100 ui     # the container's own messages
docker compose down           # remove the container (and its network); the data folder stays
```

Each job's log is a file in `data\ui_data\jobs\`, and the **Jobs** page shows it.

After you change `.env`, run `docker compose up -d`. Compose sees the change and recreates the container with the new
values; your data folder is not touched. (Tested: an added setting and a changed one both reached the container and the
pages after a plain `docker compose up -d`.) The container does not start by itself when Docker Desktop starts; run
`docker compose up -d` when you want it.

## If the container stops at once

`docker compose ps` shows it as exited, and `docker compose logs ui` says why, in one line that starts with
`NOT STARTED:`.

| The message says | What to do |
|---|---|
| TZ is not set | Add your time zone to `.env`, for example `TZ=Asia/Kolkata`, then `docker compose up -d`. |
| TZ in .env is not a time zone name this system knows | Correct the name (they look like `Asia/Kolkata` or `America/Chicago`), then `docker compose up -d`. |
| no data folder is mounted at /data | Start with `docker compose up -d`, not with `docker run`. |
| the folder mounted at /data looks like the source folder | `LCOM_DATA` points at the program's source folder. Point it at an empty folder or your data folder. Nothing in the source folder was changed. |
| the data folder is not writable | Check the folder's permissions, and that it is not on a read-only or synced drive. |
| … is not an ordinary file | A folder or a link has the name of a program or reference file. Move it out of the data folder. |

The time zone is required because every time on the pages is local time; without a zone they would quietly be UTC.

## Updating to a new version

```powershell
docker compose stop
# replace this project's folder with the new version (keep your .env and your data folder)
docker compose build
docker compose up -d
```

At the next start the program files in the data folder are replaced with the new version's. State, outputs, job logs
and your reference files are not touched. If a reference file changed in the new version and you have your own copy, the
Health page says so.

## Backing up and restoring

Stop first, so no job is writing while the copy is made.

```powershell
docker compose stop
Compress-Archive -Path .\data\* -DestinationPath ("lcom-data-{0:yyyyMMdd-HHmm}.zip" -f (Get-Date))
docker compose start
```

To restore into a new, empty data folder:

```powershell
docker compose stop
Expand-Archive -Path .\lcom-data-20261014-0930.zip -DestinationPath .\data-restored
# then set LCOM_DATA=./data-restored in .env, and:
docker compose up -d
```

The backup holds your state and outputs; it holds no keys, because `.env` is not in the data folder. Keep backups out of
git and out of anything you share.

## Moving existing work into the data folder

For work done earlier with `python run_ui.py` in the project folder. Do this with the container stopped and the native
UI closed, into a data folder that has no state files yet. The commands copy; the old folder is left as it was.

```powershell
docker compose stop
$from = "C:\path\to\the old project folder"      # where you used to run python run_ui.py
$to   = ".\data"
foreach ($file in "email_state.json", "decision_state.json", "approvals_log.jsonl", "approved_orders.json", "approved_orders.csv") {
    if (Test-Path -LiteralPath "$to\$file") { throw "$to already holds $file. Stop: moving would overwrite work. Use an empty data folder." }
}
New-Item -ItemType Directory -Force $to | Out-Null
foreach ($folder in "Excel Output Sheets", "Reports", "Order Sheets", "Walmart Listings", "Uploads") {
    if (Test-Path -LiteralPath "$from\$folder") {
        New-Item -ItemType Directory -Force "$to\$folder" | Out-Null
        Copy-Item -Path "$from\$folder\*" -Destination "$to\$folder" -Recurse
    }
}
foreach ($file in "email_state.json", "decision_state.json", "approvals_log.jsonl", "approved_orders.json", "approved_orders.csv", "product_facts.csv") {
    if (Test-Path -LiteralPath "$from\$file") { Copy-Item -LiteralPath "$from\$file" -Destination $to }
}
docker compose up -d
```

- Copy your edited reference files (`reviewer_exclusions.csv`, `near_misses.csv`, `lcom_prices.csv`, `email_template.txt`)
  the same way if you want to keep your edits; otherwise leave them and the container adds this version's.
- Do **not** copy: `.env` (your settings go in the `.env` beside `docker-compose.yml`), any `.py` file, the `ui`,
  `tests` or `.git` folders, `__pycache__`, `ui_data`, or any file ending in `.lock`, `.takeover` or `.tmp`.
- Set `TZ` to the time zone the old work was done in, so times written before and after the move read the same.

## Rules that keep your data safe

- **One data folder, one running copy.** Never run the native UI (`python run_ui.py`) and the container on the same
  folder, and never two containers on one folder. The tool's "one job at a time" lock lives inside one running copy, so
  two copies could run two jobs on the same files at once. Docker refuses a second container with the same name, but it
  cannot see a native UI or another computer.
- **Never in a synced folder.** Do not put the data folder in OneDrive, Dropbox, Google Drive or on a network share. Sync
  tools hold files open and copy half-written files, and the state files are replaced by renaming.
- **Never the source folder.** The data folder is not the program's folder. The container refuses to start if it looks
  like one.
- Close a state file or a results workbook in Excel before you start a step that writes it.

## After a hard stop: leftover lock files

If the computer lost power or the container was killed while an email or approval step was writing, a lock file can be
left in the data folder:

- `email_state.json.lock` and `email_state.json.lock.takeover`
- `decision_state.json.lock` and `decision_state.json.lock.takeover`
- temp files such as `email_state.json.<numbers>.tmp` or `decision_state.json.<numbers>.tmp`

The step then fails with a message that the state file is busy because another command is using it. The tool never
deletes these files by itself. To clear them:

1. Wait a minute and try once more: a job may really be running. The Jobs page shows it.
2. `docker compose stop`, and make sure no other copy of the tool uses this data folder.
3. Delete only the files named above, for example
   `Remove-Item -LiteralPath .\data\email_state.json.lock`
4. `docker compose start`

Never delete `email_state.json` or `decision_state.json` themselves.

## Never run

| Command or button | Why |
|---|---|
| `docker compose down -v`, `docker system prune`, `docker volume prune`, `docker container prune`, `docker image prune -a`, `docker volume rm` | They delete Docker objects beyond this tool, including those of other tools on this computer. This tool needs none of them. |
| Docker Desktop **Troubleshoot → Reset to factory defaults** and **Clean / Purge data** | They delete **all** containers, images and volumes in Docker Desktop, of every tool. Your data folder is a normal folder on disk and survives, but every image must be built again and other tools lose their Docker data. |
| `docker run --env-file .env ...` | Plain `docker run` keeps quotes as part of a value, so quoted keys stop working. Always use `docker compose`. |
| `docker compose config` or `docker inspect` on a shared screen | They print every setting, keys included. |
| A second container, or the native UI, on the same data folder | See "Rules that keep your data safe". |

## Settings

`.env.docker.example` lists every setting by name, with what it is for. Copy it to `.env` and fill it in; never commit
`.env` or send it to anyone.

- A line starting with `#` is switched off. An empty value is not the same as a switched-off line.
- A value with `$` in it **must** be in single quotes: `SMTP_PASSWORD='pa$word'`. Without quotes, compose treats `$word`
  as a reference to another setting, replaces it with nothing and only prints a warning; the tool would get `pa`.
- Also put a value in single quotes if it contains `#`, or begins or ends with a space.

What compose reads exactly as the tool's own reader (python-dotenv) does, tested:

| Shape | Example | Use it? |
|---|---|---|
| Plain value | `BRAND_NAME=TestBrand` | yes |
| Double quotes, with `\n` for a line break | `SHIPPING_ADDRESS="Company\n1 Street\nCity, ST 12345"` | yes |
| Single quotes, with `$`, `#` or spaces inside | `SMTP_PASSWORD='pa$$word with#hash'` | yes |
| A value over several lines inside double quotes | `NAME="line one` (new line) `line two"` | yes, but the one-line `\n` form is clearer |
| A reference to a setting earlier in the file | `B=${A}-suffix` | yes, both read it the same |
| An `export` in front | `export NAME=value` | yes, both ignore the `export` |
| An unquoted `$` | `NAME=pa$word` | **no**: compose and the tool read different values |
- Required for the tool to start: `TZ`. Required for the work named in the file: the API keys, the mail settings,
  `SHIPPING_ADDRESS` and `BRAND_NAME`. The Health page shows which are missing.
- `LCOM_PORT` changes the port (default 8010). The same number is used inside and outside the container.
- `LCOM_DATA` moves the data folder.

### A company proxy

If your network needs a proxy, set `HTTP_PROXY` and `HTTPS_PROXY` in `.env`, and `NO_PROXY` for any of your own names that
must be reached directly.

- **Proxy settings are secrets**: a proxy address can hold a user name and password. The container's start-up lines
  name the proxy settings that are set and count the entries, never their values; the Health page never shows them; job
  logs replace the user name and password in any address (`http://[REDACTED]@proxy:3128`) and the proxy settings
  themselves. Every page does the same for any address it shows, also when the address comes from a results file (a
  seller name, a listing address) and also inside form fields and the pages' data requests. Plain email addresses are
  shown in full: the page that sends emails lists every recipient on purpose.
- **Downloads are not masked.** A file you download (a results workbook, a report, an order sheet, a listing, the
  approved-orders export) is your own file and comes back exactly as it is in the data folder.
- At start the container adds `localhost`, `127.0.0.1` and `host.docker.internal` to `NO_PROXY` (and its lower-case
  spelling `no_proxy`), after every entry you set yourself, in your order, without repeats. The tool's own address and a
  Langfuse on this computer are then reached directly, while everything else still goes through your proxy.
- `::1` is not added on purpose: the three HTTP libraries the tool uses treat it differently in `NO_PROXY` (tested on
  Windows and in the image: one of them ignores `::1`, another fails on `[::1]`), so it would not work reliably. The
  tool does not use `::1`.

### Tracing (only for the people who maintain the tool)

Tracing is off in the container, whatever `OBS_ENABLED` says in `.env`. To switch it on, set `LCOM_TRACING=1` and the
Langfuse keys. The container reaches a Langfuse that runs on this computer at `http://host.docker.internal:3000`; set
`LCOM_LANGFUSE_URL` only for another address, and never use `localhost` there (inside the container that means the
container itself).

## For maintainers

`docker\verify.ps1` checks the image and the compose file in four stages, with fake keys and a throwaway data folder
under `%TEMP%`. Run one stage, read its report, then run the next:

```powershell
$env:OBS_ENABLED = "0"
powershell -NoProfile -ExecutionPolicy Bypass -File docker\verify.ps1 -Stage 1 -WriteConstraints   # first time only: -WriteConstraints
powershell -NoProfile -ExecutionPolicy Bypass -File docker\verify.ps1 -Stage 2
powershell -NoProfile -ExecutionPolicy Bypass -File docker\verify.ps1 -Stage 3
powershell -NoProfile -ExecutionPolicy Bypass -File docker\verify.ps1 -Stage 4                      # also cleans up
powershell -NoProfile -ExecutionPolicy Bypass -File docker\verify.ps1 -Cleanup                      # at any point: remove this run's objects
```

It only ever removes its own objects (names starting with `lcom-verify`, label `com.zync.lcom=1`), never your running
`lcom-ui`.

Keep these out of git (add them to `.gitignore`):

```
data/
lcom-data-*.zip
```
