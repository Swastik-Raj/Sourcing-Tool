# Competitive Sourcing Research Agent

A POC tool that automates competitive sourcing research: for each product (SKU +
description), it searches Chinese sourcing/manufacturing sites (AliExpress, Alibaba,
Made-in-China) for equivalent products, scores each candidate against a 5-attribute rubric,
and returns up to 5 ranked candidates per product with pricing, MOQ, and supplier info.

It uses Claude (Haiku, for cost) with a hard tool-call budget per product, connected to the
[Nimble](https://nimbleway.com) MCP server for web search/extraction.

## How it works

For each product, Claude is given:
- The product's SKU and description
- A hard budget of tool calls (search + extract combined)
- Two tools: a site-search extractor (hits each site's own on-site search page directly,
  since generic web search barely indexes individual listings on these platforms) and a
  page extractor (for pulling detail from an individual listing when needed)

Claude scores each candidate it finds against a 5-attribute rubric (product type,
category/spec, shielding/material, mount/form factor, gender/pins — 20 points each, 100
total), awarding 0 for anything not *explicitly* stated in the listing text — no benefit of
the doubt. It returns up to 5 distinct candidates (different manufacturers, ranked by a
balance of match quality and price), or reports no match with a reason if nothing plausible
was found.

**Why not 1688.com?** Its search results sit behind a Taobao/Alibaba login wall regardless
of browser driver tier — confirmed unreachable, so it's excluded rather than silently
wasting tool-call budget on it.

## Setup

```bash
pip install "anthropic[mcp]" pydantic openpyxl python-dotenv
```

Create a `.env` file in the project folder (already gitignored):

```bash
ANTHROPIC_API_KEY=sk-ant-...   # your Anthropic API key
NIMBLE_API_KEY=...             # your Nimble API key (used as the MCP bearer token)
```

Environment variables that are already set in the shell take precedence over `.env`.

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

## Known limitations

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
