# Competitive Sourcing Research Agent

A POC tool that automates competitive sourcing research: for each product (SKU +
description), it searches Chinese sourcing/manufacturing sites (AliExpress, Alibaba,
Made-in-China) for equivalent products, scores each candidate against a 5-attribute rubric,
and returns up to 3 ranked candidates per product with pricing, MOQ, and supplier info.

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
the doubt. It returns up to 3 distinct candidates (different manufacturers, ranked by a
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
# Quick test - 5 built-in sample products, no input file needed
python sourcing_agent.py

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
   `Results`: one row per product (SKU, Product, Keyword, L-Com Price, Recommendation) plus
   a `Manufacturer 1/2/3` block each (Name, Accuracy, Listed Price, Unit Price, Unit Price
   Confidence, vs. L-Com Price, MOQ, URL, Email, Match Tier, Comment). Sheet `Comparison`:
   candidates stacked against the L-Com benchmark per product. The recommended candidate is
   filled green in both sheets; nothing is green when no candidate qualifies. Match Tier is
   derived from match_percent: **>=90 Auto-accepted, 80-89 Flagged for manual review, <80 Rejected.**

### Unit price, L-Com comparison & recommendation

- Claude reports `price_total`, `quantity_covered` and `unit_price_confidence`
  (`stated` / `inferred` / `ambiguous`). The code computes `unit_price = price_total /
  quantity_covered`. Ambiguous or non-USD prices get no unit price. They show as
  UNDETERMINED and are listed at the top of the Markdown report. The tool never guesses.
- `vs. L-Com` compares `unit_price` with the input's `Lcom sale Price` column.
- A candidate is recommended only with >= 80% accuracy (`MIN_RECOMMEND_ACCURACY`) and a
  confirmed unit price >= 30% below L-Com (`MIN_MARGIN_PCT`). Among those, the highest
  accuracy x margin wins. Otherwise the report says not to source from any of them.
- `python test_pricing.py` runs an offline check of this logic.

## Cost & budget controls

- Model: `claude-haiku-4-5`
- `MAX_TOOL_CALLS_PER_PRODUCT` (currently 8, in `sourcing_agent.py`) hard-caps tool calls
  per product across search + extract combined - this is the main cost lever. It's marked
  in the code as a first guess for finding/verifying 3 candidates instead of 1; tune it
  based on real batch results.
- Observed cost so far: roughly **$0.02-0.07 per product** depending on how many tool calls
  it takes to find candidates (a clean match with 3 tool calls is cheap; a product needing
  the full 8-call budget costs more). A 100-SKU batch should land somewhere around $2-4.
- Cost is printed per-product and as a running total at the end of every run - check it on
  a small `--limit` run before committing to a large batch.

## Known limitations

- Site extraction is occasionally flaky (JS rendering timing, geo/locale redirects,
  anti-bot throttling) - a "no match" result sometimes means the search pages didn't
  render usable content that run, not that no candidate genuinely exists. Re-running the
  same product can produce a different (often better) result.
- Individual product detail pages on these sites are unreliable to extract directly and
  often come back empty - Claude is instructed to score primarily off search-page listing
  titles instead, and treat a detail-page extract as optional bonus context only.
- Supplier email addresses are rarely available from these listing pages (most sites use
  in-platform contact/chat, not public email) - the Email column will usually be blank.
- This is sequential, one product at a time - a large batch will take a while; there's no
  parallelism yet.
