# Sourcing and Listing Pipeline (POC)

Tested on **Python 3.13.5, Windows 11, PowerShell**.

This tool finds cheaper equivalents of L-Com products on Chinese sourcing sites, drafts the sample-order emails, turns a
manager's approval into an order sheet, and drafts Walmart listing text. Every step reads and writes local files, and a
person decides at each gate. Nothing sends email, places an order or publishes a listing on its own.

Pipeline order (run from the project folder, in PowerShell):

1. **Search** - `python sourcing_agent.py --sku HDFF FOA-020C` (or `--input <products.xlsx>`) writes `sourcing_results_<timestamp>.xlsx`.
2. **Emails** - `python email_agent.py drafts <results.xlsx>`, then `send`, `mark`, `reply`, `report` (same results file).
3. **Approvals** - `python decision_agent.py request <results.xlsx>`, then `record --results <results.xlsx> --request A-000N`, `export --results <results.xlsx>`.
4. **Order sheet** - `python order_sheet.py build --approved approved_orders.csv --results <results.xlsx>`.
5. **Listing content** - `python content_agent.py facts-template --approved approved_orders.csv --results <results.xlsx>`, fill `product_facts.csv`, then `generate ... --facts product_facts.csv`, then `check <walmart_listings_*.xlsx>`.

Each script prints its full usage at the top of its `.py` file.

## Setup (PowerShell)

```powershell
cd "C:\path\to\Search Tool POC"
python -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install -r requirements.txt
Copy-Item .env.example .env
notepad .env
$env:PYTHONIOENCODING = "utf-8"
```

`.env` is ignored by git; never commit it. Set `PYTHONIOENCODING` in each new PowerShell window (it stops Chinese text in
listings and replies from crashing redirected output). If script activation is blocked, run
`Set-ExecutionPolicy -Scope CurrentUser RemoteSigned` once. Variables set in the shell win over `.env`.

| Needed for | Variables in `.env` |
|---|---|
| Search (`sourcing_agent.py`) | `ANTHROPIC_API_KEY`, `NIMBLE_API_KEY` |
| Emails (`email_agent.py`) | `drafts`, `report`, `mark`: none. `reply`: `ANTHROPIC_API_KEY`. `send`: `SMTP_HOST`, `SMTP_USER`, `SMTP_PASSWORD`, `SHIPPING_ADDRESS` (`SMTP_PORT`, `SMTP_FROM`, `SENDER_NAME` optional) |
| Approvals (`decision_agent.py`) | none |
| Order sheet (`order_sheet.py`) | none (`ORDER_SHEET_RECALC_PY` optional) |
| Listing content (`content_agent.py`) | `generate`: `ANTHROPIC_API_KEY`, `BRAND_NAME` |
| Tracing (optional) | `OBS_ENABLED=1`, `LANGFUSE_PUBLIC_KEY`, `LANGFUSE_SECRET_KEY`, `LANGFUSE_BASE_URL` - see [OBSERVABILITY.md](OBSERVABILITY.md) |

`.env.example` lists every variable with a note on what it is. Tracing is on only if `OBS_ENABLED=1`.

## Human steps (nothing happens without them)

- **Search**: you choose the products and read the results; it only reads public pages and writes files.
- **Emails**: `drafts` only shows text. `send` shows each email and asks `y` per email, or you type `yes` once for the batch;
  it refuses while `SHIPPING_ADDRESS` is unset. Use `--test-recipient you@example.com` to send test copies to yourself only.
  Seller replies are pasted in by a person with `reply`. With no terminal (a program running it), `send --confirm-count N`
  sends without asking, but only if exactly N emails would go out after every check; any other number sends nothing. It does
  not skip any other check. Without a terminal and without that flag, `send` stops with exit code 1 and sends nothing.
- **Approvals**: the request is a text file for a manager. A person pastes the manager's reply into `record`, then types the
  request id, the channel, the approver's name and a final `yes`. Approvals expire after 7 days and can be revoked.
  With no terminal, give all three of `--confirm-request A-000N` (must equal `--request`), `--channel` and `--approver`
  (typed by a person, never defaulted); one or two of them is refused. The prompts need a real terminal: if there is none the
  command stops (it no longer falls back to the console device) and records nothing.
- **Order sheet**: a person places the orders by hand from the sheet. The tool never buys anything.
- **Listing content**: a person fills in the product facts, reviews the drafts and uploads them to Walmart by hand.

## Run the UI (Stage 1: skeleton only)

A local web page that shows each workflow's progress and runs allowlisted commands as jobs. It has no login, so it only listens on this computer.

```powershell
$env:OBS_ENABLED = "0"      # optional: switch tracing off for this session
python run_ui.py            # then open http://127.0.0.1:8000/
```

The pages use neutral wording (all display wording lives in `ui/wording.py`; job logs and files on disk are never changed, and a displayed log says so). Stage 1 can only run each agent's `--help`; the search, email, approval and order screens come in Stage 2. Job logs and metadata are written to `ui_data\` (add it to `.gitignore`). `UI_HOST`, `UI_PORT`, `PROJECT_DIR` and `UI_DATA_DIR` change where it listens and which folders it uses, and `UI_APP_NAME` changes the name shown in the header (default "Sourcing Tool"); a non-loopback host is refused unless `UI_ALLOW_NON_LOOPBACK=1`. Test with `python test_ui.py`.

## State files, locks and exit codes

`email_state.json`, `decision_state.json`, `approved_orders.json` and `approved_orders.csv` are written atomically (temp file,
then rename). Commands that change `email_state.json` or `decision_state.json` (`send`, `mark`, `reply`, `request`, `record`,
`revoke`, `export`) take a lock file next to it (`<state file>.lock`) and wait up to 10 seconds (`STATE_LOCK_WAIT_SECONDS`
changes that); a lock left by a stopped program is taken over after 15 minutes. `drafts`, `report` and `status` never lock.

| Exit code | Meaning |
|---|---|
| 0 | Finished |
| 1 | The command stopped with a message and did nothing further (also `content_agent check`: problems found in the listings) |
| 2 | Bad command line |
| 3 | State file busy: another command holds its lock. Nothing was changed; try again |

## Tests (offline)

No network, no keys needed. The one-line check (prints `FAILED: <name>` for anything that does not pass):

```powershell
$env:OBS_ENABLED = "0"; foreach ($t in "test_pricing","test_email_agent","test_decision_agent","test_order_sheet","test_content_agent","test_observability","test_gates") { python "$t.py"; if ($LASTEXITCODE -ne 0) { "FAILED: $t" } }
```

Each suite ends with `ok` or `... tests passed`. `python test_observability.py --mutations` runs a slow extra check (about 10 minutes).

## Files the tools create (all local to your machine, not shared between machines)

| Where | What |
|---|---|
| project folder | `sourcing_results_<timestamp>.xlsx` and `sourcing_report_<timestamp>.md` from a search (move the results to `Excel Output Sheets\` to keep them tidy) |
| `Reports\` | approval requests (`approval_request_A-000N.txt/.md`) and email reports |
| `Order Sheets\` | `order_sheet_<timestamp>.xlsx` |
| `Walmart Listings\` | `walmart_listings_<timestamp>.xlsx/.md` |
| state files | `email_state.json` (what was sent and replied), `decision_state.json`, `approvals_log.jsonl` (append-only audit log), `approved_orders.csv/.json` |
| you edit | `product_facts.csv` (facts the listing text may use), `reviewer_exclusions.csv`, `lcom_prices.csv`, `near_misses.csv` |

Copying the project to another machine does not copy this state unless you copy the files too.

## Known limits

- The independent recalculation check in `order_sheet.py` needs LibreOffice (`soffice`) and the xlsx skill's `recalc.py`. If either
  is missing the script prints `skipped (...)` for that check and carries on; the built-in formula evaluator still checks every
  formula and total, and the build stops if that fails.
- Walmart character limits in `content_agent.py` are unverified assumptions; check them against Seller Center.
- L-Com reference prices in `lcom_prices.csv` are mostly unverified (labelled "unverified" in the results).
- Tracing needs a Langfuse server you control; see [OBSERVABILITY.md](OBSERVABILITY.md). It adds about 5 seconds per command.
- Supplier sites are flaky; a "no match" can mean the pages did not load (see Known limitations below).

## Docker

The project is not containerised yet. If it is later, pass the keys at run time (never bake `.env` into an image) and mount the state files and output folders from the host.

## Usage

```bash
# Quick test - the 5 built-in products (hand-picked from L-Com's site), no input file needed
python sourcing_agent.py

# First 10 rows of the spreadsheet plus the 5 built-in products (15 total)
python sourcing_agent.py --input IE_Top_100_SKU_By_Brand_Results.xlsx --limit 10 --add-builtin

# Re-run specific products only (e.g. after a transient API error). SKUs are looked up in
# --input and the built-in products; quote any with special characters, e.g. "C&P9M"
python sourcing_agent.py --input IE_Top_100_SKU_By_Brand_Results.xlsx --sku FOA-020C HG2409U-PRO

# Run against a real product list
python sourcing_agent.py --input IE_Top_100_SKU_By_Brand_Results.xlsx

# Test on just the first N rows before committing to a full batch
python sourcing_agent.py --input IE_Top_100_SKU_By_Brand_Results.xlsx --limit 10

# Choose the output file name (otherwise timestamped automatically)
python sourcing_agent.py --input products.xlsx --output results.xlsx

# Inspect the Nimble tools' schemas without spending any Anthropic tokens
python sourcing_agent.py --inspect-tools
```

### Input file format

An `.xlsx` file with a header row (anywhere in the first 5 rows) containing at least a
`SKU` column. `Keyword` (falls back to `Product` if missing) is used as the product
description Claude searches against. Any other columns are ignored. Rows with a blank SKU
are skipped.

The built-in products live in `PRODUCTS` in `sourcing_agent.py`. **L-Com reference prices come
from `lcom_prices.csv`** (`sku, pack_size, pack_price, source, date_checked`; unit price =
pack_price / pack_size, to the cent). The tool reads it first and only falls back to the
price in the input sheet / `PRODUCTS` for a SKU that's missing, with a warning and an
"unverified" label. Rows checked by a person say so in `source` (e.g. "Srijan, live check")
with a date; everything else is "unverified" (C&P9M: "Newark distributor price,
unconfirmed"). L-Com's site blocks automated access, so update the CSV by hand after a
live check. `read_lcom_catalog()` is a stub for pulling
products and prices straight from L-Com's catalog later; while it returns nothing, the
tool falls back to `PRODUCTS`.

| SKU | Product | Keyword |
|---|---|---|
| ECF504-SC6 | Ethernet Cat6 Adapters | Cat6 RJ45 Coupler Shielded (8x8) Panel Mount Style |

## Output

Every run produces three things:

1. **Terminal output** - live progress (`Product N of M: <SKU> ...`), each tool call made
   and a preview of its result, then each product's ranked candidates as they complete.
2. **Markdown report** (`sourcing_report_<timestamp>.md`) - a human-readable writeup per
   product: full candidate details, rubric breakdown, and token/cost accounting. Good for
   a quick read or sharing a summary.
3. **Excel report** (`sourcing_results_<timestamp>.xlsx`, or `--output` path) - sheet
   `Results`: one row per product (SKU, Product, Keyword, L-Com Unit Price, L-Com Price
   Source, L-Com Price Date, Recommendation),
   then **Recommended Manufacturer / Email / URL / Unit Price** (filled green - who to buy
   from; blank and uncoloured when no candidate qualifies), then a `Manufacturer 1-5`
   block each (Name, Accuracy, Listed Price, Unit Price, Unit Price Confidence, vs. L-Com
   Price, MOQ, URL, Email, Match Tier, Comment). Sheet `Comparison`: candidates stacked
   against the L-Com benchmark per product, recommended row in green. Match Tier is
   derived from match_percent: **>=90 Auto-accepted, 80-89 Flagged for manual review, <80 Rejected.**

### Unit price, L-Com comparison & recommendation

- Claude reports `price_total`, `quantity_covered` and `unit_price_confidence`
  (`stated` / `inferred` / `ambiguous`). The code computes `unit_price = price_total /
  quantity_covered`. Ambiguous or non-USD prices get no unit price. They show as
  UNDETERMINED and are listed at the top of the Markdown report. The tool never guesses.
- `vs. L-Com` compares `unit_price` with the L-Com unit price from `lcom_prices.csv` (pack
  price divided by pack size; L-Com prices packs as one SKU, e.g. `Package/10`). The report
  header lists every SKU's price source and date and the SKUs still on unverified prices.
- Contact type is a hard exclude (`apply_form_rules`): when the target states crimp, solder,
  PCB/DIP or IDC contacts and a candidate states a different one, it's wrong form (accuracy 0).
  A listing stating both gets a "check before ordering" note; one that states none gets a
  "contact type isn't stated" note. Targets that state no contact type are unaffected, and
  nothing is inferred from words like "insertion" (C&P9M's description now says crimp).
- `reviewer_exclusions.csv` (`sku, supplier_or_url, reason`): a human call the tool honors. A
  candidate for that SKU whose maker name or URL contains `supplier_or_url` is excluded and
  shown as "excluded by reviewer: <reason>". Seeded with HDFF / Xiangtianzhong and VIC00001 /
  Xindaying (short cables per their product pages).
- A listing that contradicts itself (title says Female but its attribute table says male;
  "compatible with CAT5e" next to a Cat3-Cat6A table) gets an ordering note and a comment, not a
  zero score. A plain category range ("Cat5e Cat6 Cat6a") is left alone.
- IDC is its own contact type, separate from crimp. A listing that states both is kept with a note.
- A product that crashes (API error, encoding error, unusable response) is recorded as
  "Error researching this product: <reason> - re-run with --sku" in the Excel Results and
  Comparison sheets and the report - never as "No candidates found", never green, and the email
  agent skips it. The report header and the end-of-run summary give separate counts (recommended /
  no recommendation / errored) and the exact PowerShell `--sku` command to re-run the errored ones.
  A billing, credit or authentication error from the API stops the batch at once: finished products
  are saved, products not yet started are recorded as errored without being run, and the reason is
  printed loudly.
- A candidate with no URL is never recommended: the reason says "best candidate has no URL,
  locate it manually" and nothing is highlighted green.
- The report shows each candidate's **Spec lines seen** (type/jacket/length/category the model
  read on a detail page, or "none") so a spec-text cable rule can be tested on real text next round.
- A price range always uses its high end (`apply_price_rules`; the Listed Price column keeps
  the full range). A listing that ties prices to quantity tiers gets a note naming the tier used.
- Spec text that reads like a short cable ("PVC jacket", "24 AWG", "custom length") on a
  coupler/adapter target adds an ordering note, never an exclusion. The search reads titles, so
  this only fires when the model carries such text into its caveats.
- A unit price more than 10x cheaper than the next-cheapest candidate for the same product
  (`IMPLAUSIBLE_PRICE_RATIO`) is treated as an extraction error. It is marked IMPLAUSIBLE,
  excluded from the L-Com comparison and never recommended. The check needs at least 3
  priced candidates - with 2 there's no telling a parsing error from two real prices.
- Candidates from the same manufacturer (name compared case-insensitively) are collapsed to
  the highest-scoring one; the freed slot is left empty. Placeholder names ("Unknown",
  "unnamed", "Generic", "...seller") are exempt, since those are different anonymous sellers.
- A listing of a different product form (a cable when the target is a coupler, a multi-port
  variant, a different connector class) is excluded outright: Claude reports
  `same_product_form` / `listing_form` per candidate and its accuracy is zeroed. For coupler
  and adapter targets, code also excludes any listing whose title gives a cable length
  ("50cm", "1.5m") or says cable/cord without a part word (`apply_form_rules`), since the
  model's own call was inconsistent. Mount style (inline vs panel vs keystone) is scored in
  `mount_form`, not treated as a different product: for coupler targets, code keeps an inline
  coupler the model excluded, as long as it shows the target's connector type, has no cable
  signs and isn't a multi-port variant.
- Antennas: a dual-band, multi-band, "2way", MIMO or multi-port listing against a
  single-band, single-port target gets 0 on category/spec (prompt rule plus
  `apply_spec_rules` in code), so it can't win on one matching band.
- Each recommendation carries a **Check before ordering** note when the listing text calls
  for one (`ordering_notes`): the title names a different variant (VGA/DB15 for a DB9 target,
  RJ11 for RJ45, Micro/Mini/Type-C USB, single- vs multimode fiber), one listing covers
  several models or options (MC-6BP/MC-6BR, simplex/duplex, male/female), the price needs an
  order of `BULK_MOQ_NOTE` (500)+ pieces, a promo or unattached pack size, plus Claude's own
  `listing_caveats`. MOQ/tier remarks are taken out of those caveats, but any quantity in
  them still goes through the 500 check (the model sometimes states an MOQ only there). It appears under the Recommendation line in the report and in the
  amber **Ordering Note** column in Excel. The 500 threshold is a stopgap until we have a
  real order quantity per SKU.
- A quantity above 1 only divides the price when that number appears next to a pack word in
  the title or price text (`pack_size_supported`); otherwise the price is per piece. This
  stops MOQs being used as pack sizes (HDFF's $0.0032, Suzhou Bulovb's $0.0688).
- AliExpress new-shopper prices ("$1.09 $5.77 -81% New shoppers save…") are replaced by the
  regular price; if none is shown the price is marked `promo (regular price unknown)`. Code
  backs this up (`apply_price_rules`): when the price used is the promo amount the model
  itself described, it's swapped for the stated regular price or labelled promo.
- The implausible-price check compares only against stated/inferred prices (never promo), and
  also flags any candidate whose model caveat says the price looks unusually low.
- Cable detection also counts phrases that describe the item as a cable even next to
  "adapter"/"connector" ("printer data cable", "extension cable", "ferrite", "24AWG",
  "PVC jacket"); broader ones like "Ethernet Cable" are left out because couplers name the
  cable they connect to.
- A candidate is recommended only with >= 80% accuracy (`MIN_RECOMMEND_ACCURACY`) and a
  confirmed unit price >= 80% below L-Com (`MIN_MARGIN_PCT` - room for shipping, storage and
  import taxes). If several qualify, the most accurate wins; among those within 5 points of it
  (`TIE_BAND_POINTS`): a `stated` price first (then inferred, then promo), then a named
  maker, then the lowest price. Otherwise the message names the bar that failed:
  product form, accuracy only, price margin only, or both. Missing rubric scores are reported
  as "scoring failed", never as a margin problem.
- The per-product score line in the report is built from the same scores as the table; the
  model no longer writes a free-text note.
- `python test_pricing.py` runs an offline check of this logic.

## Cost & budget controls

- Model: `claude-haiku-4-5`
- `MAX_TOOL_CALLS_PER_PRODUCT` (currently 13, in `sourcing_agent.py`) hard-caps tool calls
  per product across search + extract combined - this is the main cost lever. It's a first
  guess for finding up to 5 candidates (`MAX_CANDIDATES`), scaled from 8 for 3; tune it
  based on real batch results.
- Observed cost with the old 3-candidate / 8-call setup: roughly **$0.02-0.07 per product** depending on how many tool calls
  it takes to find candidates (a clean match with 3 tool calls is cheap; a product needing
  the full 8-call budget costs more). A 100-SKU batch should land somewhere around $2-4.
- Cost is printed per-product and as a running total at the end of every run - check it on
  a small `--limit` run before committing to a large batch.

## Known limitations (search)

- Site extraction is occasionally flaky (JS rendering timing, geo/locale redirects,
  anti-bot throttling) - a "no match" result sometimes means the search pages didn't
  render usable content that run, not that no candidate genuinely exists. Re-running the
  same product can produce a different (often better) result.
- Every extract call is pinned to `country: "US"`, `locale: "en"` (`EXTRACT_GEO`), since
  unpinned calls sometimes got other countries' storefronts with prices in foreign currencies.
- Extracted pages are stripped of menus, filters and non-product link URLs
  (`strip_page_chrome`) before being cut to `MAX_TOOL_RESULT_CHARS`. Without this,
  Made-in-China's listings started ~30,000 characters in and Claude only ever saw menus.
- Individual product detail pages on these sites are unreliable to extract directly and
  often come back empty - Claude is instructed to score primarily off search-page listing
  titles instead, and treat a detail-page extract as optional bonus context only.
- Supplier email addresses are rarely available from these listing pages (most sites use
  in-platform contact/chat, not public email) - the Email column will usually be blank.
- Each tool call times out after `TOOL_CALL_TIMEOUT_S` (60s, no retry); the model is told and
  moves on, instead of one hung page freezing the batch. In the 2026-09-28 baseline, 145 of 147
  successful extracts finished within 60s.
- Up to `MAX_CONCURRENT_PRODUCTS` (3) products are researched at once; every progress line is
  prefixed with its `[SKU]`. Anthropic 429/529/5xx errors are retried by the SDK with
  exponential backoff (`ANTHROPIC_MAX_RETRIES`), Nimble rate limits by the tool wrapper
  (`NIMBLE_RATE_LIMIT_RETRIES`); both counts are printed at the end of a run. Lower the
  concurrency if they climb.
- With `PREFETCH_SITE_SEARCHES` on, each product's three site searches (description as the
  query) run at the same time in code before Claude starts; they count as 3 of the product's
  tool calls. `PREFETCH_STAGGER_S` spaces their starts out if a site throttles.
- Each run ends with a timing summary: wall-clock time, Nimble vs Anthropic share, timeouts,
  peak simultaneous Nimble calls, Nimble latency and timeouts per site, and why each product
  stopped (found 5 candidates / budget used up, with failed-call counts / model stopped early).
- The report and Excel file are saved after every product, so a stopped or crashed run keeps
  everything finished so far. If the .xlsx is open in Excel, that save is skipped and retried
  after the next product.
- Within one product, tool calls still run one after another.
