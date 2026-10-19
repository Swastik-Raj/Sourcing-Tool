# The UI: running the pipeline from a web page

A local web page that drives the five tools in their real order. It starts each agent as a separate program with a fixed
list of allowed commands, shows the agent's own output, and reads status from the agents' own files. It has no login, so
it only listens on this computer.

## Start it

```powershell
$env:OBS_ENABLED = "0"      # optional: switch tracing off for this session
python run_ui.py            # then open http://127.0.0.1:8000/
```

`UI_HOST`, `UI_PORT`, `PROJECT_DIR`, `UI_DATA_DIR`, `UI_APP_NAME` and `UI_LOG_RETENTION_DAYS` change where it listens, which
folders it uses, the name in the header and how long job logs are kept. A host other than this computer is refused unless
`UI_ALLOW_NON_LOOPBACK=1`, and then a warning is printed.

## The tabs, in order

| Tab | What it runs | What a person must do before it runs |
|---|---|---|
| 1. Find suppliers | `sourcing_agent.py --input <sheet> [--limit N] [--sku ...] [--add-builtin] --output <results>` | Upload or pick a sheet, read the cost estimate, type the number of products |
| 2. Email sellers | `email_agent.py drafts`, `send --confirm-count N [--test-recipient ADDR]`, `mark`, `reply --file` | Read the list of recipients and type how many emails; tick a box to record a form submission; paste a seller's reply |
| 3. Get approval | `decision_agent.py request`, `record --confirm-request --channel --approver --file`, `revoke`, `export` | Paste the manager's reply and type the request number, how it reached you and who approved (never filled in for you); type a SKU to withdraw an approval |
| 4. Order sheet | `order_sheet.py build --approved approved_orders.csv --results <results>` | Nothing typed: it only creates a file. Ordering and payment stay manual |
| 5. Listings | `content_agent.py facts-template`, `generate --facts <uploaded file>` | Fill in the product details sheet (the facts file) yourself and upload it; type the number of products |

The pages call a results file "a search", the export "the approved list", and the facts file "the product details sheet".
The top bar also has Activity (every task started, with its log) and Setup (which settings are filled in).

Every step after Find suppliers works on one search, chosen in the "Working on the search of" box, and shows a progress box
for it. The box is read from files on disk and its "Next" link only opens a page.

Things the UI will not do:

- Start a second command that changes data while one is running. The second is refused, not queued.
- Send email, record an approval or write listings from a single click.
- Pass `--yes`, `--allow-listing-facts`, `--dry-run`, `--out`, `--duty` or `--inspect-tools` to an agent.
- Build an order sheet or listings from an export that was made from a different results file.
- Set or change a fact's status. The facts file you upload is saved byte for byte and only read.
- Show a key, a password, the shipping address or a phone number. Logs mask them, and mask every email address.

## Folders

| Folder | What is in it |
|---|---|
| `Uploads\` | Product sheets and filled facts files you uploaded, and facts templates the content agent wrote. Names are built by the UI |
| `Excel Output Sheets\`, `Reports\`, `Order Sheets\`, `Walmart Listings\` | The agents' own output, offered for download |
| `ui_data\` | The UI's job logs and records. Pasted replies are kept in `ui_data\inputs\` only while their command runs |

## Tests (offline: no network, no keys)

All of these run without any real key and never call a real service.

```powershell
$env:OBS_ENABLED = "0"

# the seven agent suites
foreach ($t in "test_pricing","test_email_agent","test_decision_agent","test_gates","test_order_sheet","test_content_agent","test_observability") { python "$t.py"; if ($LASTEXITCODE -ne 0) { "FAILED: $t" } }

# the UI suites: the original one, then the pipeline (unit tests, route tests and the offline end-to-end run)
python test_ui.py
python tests\test_ui_pipeline.py

# only the offline end-to-end run and the shim's self-test
python tests\test_ui_pipeline.py shim_self_test end_to_end
```

Each suite ends with `ok` or `... passed`. Run the UI suites on their own, not at the same time as other heavy work:
a few of their checks are timed.

The end-to-end test runs the real agents in a throwaway folder under `%TEMP%`. A test-only file,
`tests\shim\sitecustomize.py`, replaces the AI service, the page-fetching service and the mail server with fixtures from
`tests\fixtures\` and refuses any connection that is not to this computer. It is switched on only by `LCOM_TEST_SHIM=1`,
which only the test sets. The test fails if the shim is not active, if any connection is attempted, or if any file of the
real project folder is opened.

The mutation check breaks one rule at a time in a copy under `scratch\mutation_copy\` and expects a test to fail:

```powershell
$env:OBS_ENABLED = "0"; python scratch\mutate_pipeline.py
```
