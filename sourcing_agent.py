"""
Competitive sourcing research POC.

For each product, connects directly to the Nimble MCP server (as an MCP client - not
Anthropic's server-side MCP connector) and exposes ONLY its search and extract tools to
Claude. Claude is given a hard budget of tool calls per product, must return a structured
(schema-enforced) result with up to 3 ranked candidates, and runs on Haiku for cost.

Setup:
    pip install "anthropic[mcp]" pydantic openpyxl
    set ANTHROPIC_API_KEY=...   (your Anthropic API key)
    set NIMBLE_API_KEY=...      (your Nimble API key, used as the MCP bearer token)

Run:
    python sourcing_agent.py                                    (5 built-in sample products)
    python sourcing_agent.py --input products.xlsx               (SKU/Product/Keyword columns)
    python sourcing_agent.py --input products.xlsx --output results.xlsx
"""

import argparse
import asyncio
import json
import os
import re
import sys
from datetime import datetime
from typing import Optional

import httpx2
import openpyxl
from mcp import ClientSession
from mcp.client.streamable_http import streamable_http_client
from pydantic import BaseModel, ValidationError

from anthropic import AsyncAnthropic
from anthropic.lib.tools import beta_async_tool

MODEL = "claude-haiku-4-5"
# Haiku 4.5 list pricing, used only to print an estimate - not billed exactly.
HAIKU_INPUT_PER_MTOK = 1.00
HAIKU_OUTPUT_PER_MTOK = 5.00

NIMBLE_MCP_URL = "https://mcp.nimbleway.com/mcp"

# 1688.com is deliberately excluded: its search results sit behind a Taobao/Alibaba
# login wall regardless of driver tier (confirmed with both vx8 and vx10 - neither
# reaches real content), so it's not reachable via extract at all.
SOURCING_SITES = ["aliexpress.com", "alibaba.com", "made-in-china.com"]

# Nimble's general web search barely indexes individual listings on these sites, and its
# "shopping" focus mode is hard-restricted to a fixed roster of Western retailers (Amazon,
# Best Buy, etc. - confirmed via the API rejecting any other subagent name) that doesn't
# cover any of these. Each site's own on-site search page, fetched directly with a
# JS-rendering driver, reliably returns real listings - so that's the primary discovery
# method instead. {query} is filled in per-search with URL-encoded search terms.
SITE_SEARCH_URLS = {
    "aliexpress.com": "https://www.aliexpress.com/wholesale?SearchText={query}",
    "alibaba.com": "https://www.alibaba.com/trade/search?SearchText={query}",
    "made-in-china.com": "https://www.made-in-china.com/products-search/hot-china-products/{query}.html",
}

MAX_TOOL_CALLS_PER_PRODUCT = 8  # searches + extracts combined, hard cap
# ^ first guess for finding/verifying 3 distinct candidates instead of 1 - tune based on
# actual cost/results once tested on a real batch.
MAX_ITERATIONS_PER_PRODUCT = MAX_TOOL_CALLS_PER_PRODUCT + 3  # backstop against loops
MAX_TOOL_RESULT_CHARS = 2500  # truncate every tool result before it goes back to Claude
MAX_TOKENS_PER_TURN = 3000

PRODUCTS = [
    {"sku": "ECF504-SC6", "description": "Cat6 RJ45 Coupler Shielded (8x8) Panel Mount Style"},
    {"sku": "TDG1026KS-C6", "description": "Cat6 Coupler - Shielded RJ45 (8x8) Keystone Feed-thru"},
    {"sku": "HDFF", "description": "HDMI Panel Mount Adapter, Female to Female"},
    {"sku": "CAPUSB-A", "description": "USB Protective Cover for Type A Jacks, ABS material"},
    {"sku": "VIC00001", "description": "DVI 24+5 female to female coupler, bulkhead mount"},
]

ATTRIBUTES = ["product_type", "category_spec", "shielding_material", "mount_form", "gender_pins"]
ATTRIBUTE_LABELS = {
    "product_type": "Product type",
    "category_spec": "Category/spec",
    "shielding_material": "Shielding/material",
    "mount_form": "Mount/form factor",
    "gender_pins": "Gender/pins",
}

RUBRIC_INSTRUCTIONS = """\
Score the candidate with this exact rubric (the same method used in our manual sourcing
audit). Five attributes, weighted 20 points each, for a 100-point total:

1. product_type (0-20) - is the listed item fundamentally the same kind of product?
2. category_spec (0-20) - category/performance rating match (e.g. Cat5e vs Cat6 vs Cat6a,
   or the equivalent spec/performance class for this product type).
3. shielding_material (0-20) - shielding and/or material match.
4. mount_form (0-20) - mount/form factor match (e.g. panel mount, keystone feed-thru,
   bulkhead mount).
5. gender_pins (0-20) - gender/pin configuration match (e.g. female-to-female), or the
   equivalent connector-orientation attribute for non-connector products. If a product
   type genuinely has no equivalent attribute, score it 0 - nothing was confirmed.

CRITICAL RULE: award 0 points for any attribute the candidate listing's own text does not
explicitly state. Do not infer or assume a match from category context, product photos, or
"typical" industry defaults - only credit what the listing text actually says. No benefit
of the doubt. match_percent must equal the sum of the five scores.
"""


class AttributeBreakdown(BaseModel):
    product_type: int
    category_spec: int
    shielding_material: int
    mount_form: int
    gender_pins: int


class Candidate(BaseModel):
    manufacturer: Optional[str] = None
    distributor_site: Optional[str] = None
    listing_title: Optional[str] = None
    price: Optional[str] = None
    moq: Optional[str] = None
    email: Optional[str] = None
    match_percent: Optional[int] = None
    attribute_breakdown: Optional[AttributeBreakdown] = None
    url: Optional[str] = None


class SourcingResult(BaseModel):
    product: str
    no_match: bool
    no_match_reason: Optional[str] = None
    candidates: list[Candidate] = []
    note: Optional[str] = None


def build_prompt(product: dict) -> str:
    sites = ", ".join(SOURCING_SITES)
    search_urls = "\n".join(f"  - {site}: {url}" for site, url in SITE_SEARCH_URLS.items())
    return f"""\
You are a competitive sourcing research analyst. Find this product for sale on Chinese
sourcing/manufacturing sites: {sites}.

Target product:
  SKU: {product['sku']}
  Description: "{product['description']}"

The SKU above is our own internal part number - it is not a manufacturer part number and
will never appear on a competitor's listing. You are looking for a DIFFERENT manufacturer's
equivalent product, judged only by the five rubric attributes below, never by whether the
listing happens to mention this SKU or any particular brand/part number.

The search tool's general web mode barely indexes individual listings on these sites, and
its "shopping" focus mode only covers a fixed set of Western retailers (Amazon, Best Buy,
etc.) - neither will find real listings here. Instead, use each site's own on-site search
page directly as your primary discovery method:
{search_urls}
Build the URL by substituting {{query}} with your search terms (spaces as "+" for the
aliexpress/alibaba forms, "_" for the made-in-china path form), then call extract on it
with driver="vx8" (these are JS-heavy pages - the default driver will return nothing
useful).

Score primarily from the listing titles/prices on that search page itself - they are
usually specific enough to score against the rubric (e.g. a title stating "Female to
Female Panel Mount" directly confirms two attributes). Individual product detail pages on
these sites are unreliable to extract directly (heavier anti-bot throttling than search
pages) and often come back empty even after a wait - treat one extract attempt on a
promising listing URL as an optional bonus for more detail, never as required, and never
let an empty/incomplete result there block you from scoring off the search listing text
you already have.

You have a HARD BUDGET of {MAX_TOOL_CALLS_PER_PRODUCT} tool calls total for this product -
more than before because you're now finding and verifying up to 3 candidates instead of 1.
Spend it deliberately: a handful of site-search extracts across different sites/queries to
surface multiple distinct candidates, plus the occasional individual listing page only if a
title is genuinely ambiguous on a rubric attribute. Once you've used your budget, or you
already have enough to score confidently, STOP calling tools immediately and give your
final answer. Do not narrate your process in free text at any point - every response you
give must be the final structured result, nothing else.

{RUBRIC_INSTRUCTIONS}

Return up to 3 candidates in `candidates`, ranked best first, balancing match quality
against price (a slightly lower match % at a much lower price can rank above a perfect
match at a high price) - report both numbers per candidate, never silently substitute one
for the other. Each candidate must be a DIFFERENT manufacturer/supplier - never list 2-3
listings from the same company just to fill the slots. If you only find 1 or 2 genuinely
distinct, plausible candidates, return only that many - do not force a weak extra pick just
to reach 3.

If you find no plausible candidate at all, or every candidate you found scores very low,
set no_match to true, leave candidates empty, and briefly say why in no_match_reason - do
not guess or force a weak candidate into looking like a real match.
"""


def truncate_text(text: str, limit: int = MAX_TOOL_RESULT_CHARS) -> str:
    text = re.sub(r"!\[[^\]]*\]\([^)]*\)", "", text)  # strip markdown images
    text = re.sub(r"(\]\(https?://[^)?\s]+)\?[^)]*\)", r"\1)", text)  # strip link tracking params
    text = re.sub(r"\s+", " ", text).strip()
    if len(text) > limit:
        return text[:limit] + " ...[truncated]"
    return text


class ToolBudget:
    def __init__(self, limit: int):
        self.limit = limit
        self.used = 0

    def take(self) -> bool:
        if self.used >= self.limit:
            return False
        self.used += 1
        return True


def pick_tool(tools, exact_names, contains, excludes=("agent", "template", "crawl", "map")):
    for t in tools:
        if t.name.lower() in exact_names:
            return t
    for t in tools:
        name = t.name.lower()
        if contains in name and not any(x in name for x in excludes):
            return t
    return None


def make_bounded_tool(tool, session: ClientSession, budget: ToolBudget):
    tool_name = tool.name
    input_schema = getattr(tool, "input_schema", None) or getattr(tool, "inputSchema", None)
    description = (tool.description or "")[:600]

    async def call(**kwargs):
        if not budget.take():
            return (
                f"Tool call budget exhausted ({budget.limit} calls used for this product). "
                "Stop searching and return your best final answer now."
            )
        print(f"    -> {tool_name}({kwargs})")
        result = await session.call_tool(name=tool_name, arguments=kwargs)
        text = " ".join(
            block.text for block in result.content if getattr(block, "type", None) == "text"
        )
        print(f"    <- {truncate_text(text, limit=300)}")
        return truncate_text(text)

    return beta_async_tool(call, name=tool_name, description=description, input_schema=input_schema)


def strict_json_schema(model: type[BaseModel]) -> dict:
    """model_json_schema() doesn't set additionalProperties, but Claude's strict
    output mode requires it explicitly false on every object (including $defs)."""
    schema = model.model_json_schema()

    def lock_down(node):
        if isinstance(node, dict):
            if node.get("type") == "object":
                node.setdefault("additionalProperties", False)
            for value in node.values():
                lock_down(value)
        elif isinstance(node, list):
            for value in node:
                lock_down(value)

    lock_down(schema)
    return schema


def usage_tokens(message) -> tuple:
    usage = getattr(message, "usage", None)
    if usage is None:
        return 0, 0
    return usage.input_tokens or 0, usage.output_tokens or 0


async def research_product(client: AsyncAnthropic, session: ClientSession, tools, product: dict):
    budget = ToolBudget(MAX_TOOL_CALLS_PER_PRODUCT)
    runner_tools = [make_bounded_tool(t, session, budget) for t in tools]

    runner = client.beta.messages.tool_runner(
        model=MODEL,
        max_tokens=MAX_TOKENS_PER_TURN,
        max_iterations=MAX_ITERATIONS_PER_PRODUCT,
        tools=runner_tools,
        output_config={
            "format": {"type": "json_schema", "schema": strict_json_schema(SourcingResult)}
        },
        messages=[{"role": "user", "content": build_prompt(product)}],
    )

    total_in, total_out = 0, 0
    last_message = None
    async for message in runner:
        last_message = message
        i, o = usage_tokens(message)
        total_in += i
        total_out += o

    parsed = None
    if last_message is not None:
        text = " ".join(block.text for block in last_message.content if block.type == "text")
        try:
            parsed = SourcingResult.model_validate_json(text)
        except ValidationError:
            parsed = None
    return parsed, total_in, total_out, budget.used


def candidate_score(candidate: Candidate) -> tuple:
    """Returns (computed_match_percent, {attr: score}) - recomputes from the breakdown
    when present so a miscounted match_percent from the model can't slip through."""
    breakdown = candidate.attribute_breakdown
    if breakdown is not None:
        scores = breakdown.model_dump()
        return sum(max(0, min(20, v)) for v in scores.values()), scores
    return candidate.match_percent, {}


def match_tier(match_percent: Optional[int]) -> str:
    if match_percent is None:
        return ""
    if match_percent >= 90:
        return "Auto-accepted"
    if match_percent >= 80:
        return "Flagged for manual review"
    return "Rejected"


def candidate_comment(scores: dict) -> str:
    if not scores:
        return ""
    zero_attrs = [ATTRIBUTE_LABELS[attr] for attr in ATTRIBUTES if not scores.get(attr)]
    if zero_attrs:
        return f"0 pts on: {', '.join(zero_attrs)} (not stated on listing)"
    return "All five rubric attributes matched something explicit on the listing."


def print_product_result(product: dict, result: Optional[SourcingResult], calls_used: int):
    print("-" * 72)
    print(f"{product['sku']} - \"{product['description']}\"  ({calls_used} tool calls used)")
    print()

    if result is None:
        print("  Could not get a structured result for this product (empty/refused response).")
        return

    if result.no_match or not result.candidates:
        reason = result.no_match_reason or "no plausible candidate found"
        print(f"  No match found - {reason}")
        return

    for i, candidate in enumerate(result.candidates, start=1):
        computed, scores = candidate_score(candidate)
        manufacturer = candidate.manufacturer or "Not stated on listing"
        source = candidate.distributor_site or "Unknown site"
        price = candidate.price or "Not stated"

        print(f"  #{i}  {manufacturer}  |  via {source}  |  Tier: {match_tier(computed)}")
        print(f"      Listing:  {candidate.listing_title or '?'}")
        print(f"      Price:    {price}   MOQ: {candidate.moq or 'Not stated'}")
        print(f"      Match:    {computed}%")
        print(f"      URL:      {candidate.url or 'Not stated'}")
        if scores:
            line = " | ".join(
                f"{ATTRIBUTE_LABELS[attr]} {max(0, min(20, scores.get(attr, 0)))}" for attr in ATTRIBUTES
            )
            print(f"      Why:      {line}  ->  {candidate_comment(scores)}")
        print()

    if result.note:
        print(f"  Note: {result.note}")


def format_result_markdown(
    product: dict, result: Optional[SourcingResult], calls_used: int, tin: int, tout: int, cost: float
) -> str:
    lines = [f"## {product['sku']} - {product['description']}", ""]

    if result is None:
        lines.append("**No structured result** (empty/refused response).")
    elif result.no_match or not result.candidates:
        reason = result.no_match_reason or "no plausible candidate found"
        lines.append(f"**No match found** - {reason}")
    else:
        for i, candidate in enumerate(result.candidates, start=1):
            computed, scores = candidate_score(candidate)
            lines += [
                f"### Candidate {i}: {computed}% - {match_tier(computed)}",
                "",
                "| Field | Value |",
                "|---|---|",
                f"| Manufacturer / Producer | {candidate.manufacturer or 'Not stated on listing'} |",
                f"| Supplier / Site | {candidate.distributor_site or 'Unknown site'} |",
                f"| Listing | {candidate.listing_title or '?'} |",
                f"| Price | {candidate.price or 'Not stated'} |",
                f"| MOQ | {candidate.moq or 'Not stated'} |",
                f"| Email | {candidate.email or 'Not stated'} |",
                f"| URL | {candidate.url or 'Not stated'} |",
                "",
            ]
            if scores:
                lines.append("| Rubric attribute | Score (of 20) |")
                lines.append("|---|---|")
                for attr in ATTRIBUTES:
                    lines.append(f"| {ATTRIBUTE_LABELS[attr]} | {max(0, min(20, scores.get(attr, 0)))} |")
                lines += ["", candidate_comment(scores), ""]
        if result.note:
            lines.append(f"*{result.note}*")

    lines += ["", f"*Tool calls used: {calls_used} | Tokens: in={tin} out={tout} | Cost: ~${cost:.4f}*", ""]
    return "\n".join(lines)


def write_report(
    sections: list, grand_in: int, grand_out: int, grand_cost: float, num_products: int, path: str
) -> None:
    header = [
        "# Competitive Sourcing Research Report",
        "",
        f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M')}",
        f"Model: {MODEL} | Sites: {', '.join(SOURCING_SITES)} | Products: {num_products}",
        "",
        "---",
        "",
    ]
    footer = [
        "---",
        "",
        f"**Total tokens:** in={grand_in} out={grand_out} | **Estimated total cost:** ~${grand_cost:.4f}",
    ]
    content = "\n".join(header + sections + footer) + "\n"
    with open(path, "w", encoding="utf-8") as f:
        f.write(content)


CANDIDATE_XLSX_FIELDS = ["Name", "MOQ", "Price", "URL", "Email", "Match Tier", "Comment"]


def write_excel_report(rows: list, path: str) -> None:
    """rows: list of (product: dict, result: Optional[SourcingResult]) tuples."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Results"

    header = ["SKU", "Product", "Keyword"]
    for n in (1, 2, 3):
        header += [f"Manufacturer {n} {field}" for field in CANDIDATE_XLSX_FIELDS]
    ws.append(header)

    for product, result in rows:
        row = [product["sku"], product.get("product_name", ""), product["description"]]
        candidates = result.candidates if result else []
        for i in range(3):
            if i < len(candidates):
                c = candidates[i]
                computed, scores = candidate_score(c)
                row += [
                    c.manufacturer or "",
                    c.moq or "",
                    c.price or "",
                    c.url or "",
                    c.email or "",
                    match_tier(computed),
                    candidate_comment(scores),
                ]
            else:
                row += [""] * len(CANDIDATE_XLSX_FIELDS)
        ws.append(row)

    for col_cells in ws.columns:
        width = max((len(str(cell.value)) for cell in col_cells if cell.value), default=10)
        ws.column_dimensions[col_cells[0].column_letter].width = min(max(width + 2, 10), 60)

    wb.save(path)


def read_products(path: str) -> list:
    """Reads SKU / Product / Keyword columns from an .xlsx (header row detected by
    scanning for a cell literally containing "SKU"), matching the manual spreadsheet's
    input layout. Description is taken from Keyword, falling back to Product."""
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb.worksheets[0]

    headers, header_row = {}, None
    for row in ws.iter_rows(min_row=1, max_row=5):
        cells = {str(c.value).strip().lower(): c.column for c in row if c.value}
        if "sku" in cells:
            headers, header_row = cells, row[0].row
            break
    if header_row is None:
        raise ValueError(f"Could not find a 'SKU' header row in the first 5 rows of {path}")

    sku_col = headers["sku"]
    keyword_col = headers.get("keyword")
    product_col = headers.get("product")

    products = []
    for row in ws.iter_rows(min_row=header_row + 1):
        sku = row[sku_col - 1].value
        if not sku:
            continue
        keyword = row[keyword_col - 1].value if keyword_col else None
        product_name = row[product_col - 1].value if product_col else None
        products.append({
            "sku": str(sku).strip(),
            "description": str(keyword or product_name or "").strip(),
            "product_name": str(product_name or "").strip(),
        })
    return products


async def connect_tools(nimble_api_key: str):
    http_client = httpx2.AsyncClient(headers={"Authorization": f"Bearer {nimble_api_key}"})
    return http_client


def parse_args():
    parser = argparse.ArgumentParser(description="Competitive sourcing research POC")
    parser.add_argument(
        "--input", help="Path to an .xlsx with SKU/Product/Keyword columns. Defaults to 5 built-in sample products."
    )
    parser.add_argument("--output", help="Path for the .xlsx results file. Defaults to a timestamped filename.")
    parser.add_argument("--limit", type=int, help="Only process the first N products (for a quick/cheap test run).")
    parser.add_argument(
        "--inspect-tools", action="store_true", help="Print Nimble tool schemas and exit (no Anthropic tokens spent)."
    )
    return parser.parse_args()


async def main_async():
    args = parse_args()

    anthropic_key = os.environ.get("ANTHROPIC_API_KEY")
    nimble_key = os.environ.get("NIMBLE_API_KEY")
    missing = [n for n, v in [("ANTHROPIC_API_KEY", anthropic_key), ("NIMBLE_API_KEY", nimble_key)] if not v]
    if missing:
        print(f"Missing required environment variable(s): {', '.join(missing)}")
        print("Set them and re-run, e.g.:")
        print("  set ANTHROPIC_API_KEY=sk-ant-...")
        print("  set NIMBLE_API_KEY=...")
        sys.exit(1)

    products = read_products(args.input) if args.input else PRODUCTS
    if args.limit:
        products = products[: args.limit]
    if not products:
        print(f"No products found in {args.input} - check it has a 'SKU' header column.")
        sys.exit(1)

    client = AsyncAnthropic(api_key=anthropic_key)

    print("Competitive Sourcing Research POC")
    print(f"Model: {MODEL}  |  Products: {len(products)}  |  Sites: {', '.join(SOURCING_SITES)}")
    print(f"Tool call budget: {MAX_TOOL_CALLS_PER_PRODUCT} per product")

    http_client = httpx2.AsyncClient(
        headers={"Authorization": f"Bearer {nimble_key}"}, timeout=60.0
    )
    async with http_client:
        async with streamable_http_client(NIMBLE_MCP_URL, http_client=http_client) as (read, write):
            async with ClientSession(read, write) as session:
                await session.initialize()
                tools_result = await session.list_tools()

                search_tool = pick_tool(tools_result.tools, {"nimble_search", "search"}, "search")
                extract_tool = pick_tool(tools_result.tools, {"nimble_extract", "extract"}, "extract")
                if search_tool is None or extract_tool is None:
                    available = ", ".join(t.name for t in tools_result.tools)
                    print(f"Could not find search/extract tools on the Nimble MCP server.")
                    print(f"Available tools: {available}")
                    sys.exit(1)

                tools = [search_tool, extract_tool]
                print(f"Using Nimble tools: {search_tool.name}, {extract_tool.name}")
                for t in (search_tool, extract_tool):
                    schema = getattr(t, "input_schema", None) or getattr(t, "inputSchema", None) or {}
                    print(f"  {t.name} params: {list(schema.get('properties', {}).keys())}")

                if args.inspect_tools:
                    for t in (search_tool, extract_tool):
                        schema = getattr(t, "input_schema", None) or getattr(t, "inputSchema", None) or {}
                        print("-" * 72)
                        print(f"{t.name} full schema:")
                        print(json.dumps(schema, indent=2))
                    print("-" * 72)
                    print("--inspect-tools: exiting before spending any Anthropic tokens.")
                    return

                grand_in, grand_out, grand_cost = 0, 0, 0.0
                report_sections = []
                excel_rows = []

                for idx, product in enumerate(products, start=1):
                    print(f"\nProduct {idx} of {len(products)}: {product['sku']} ...")
                    try:
                        result, tin, tout, calls_used = await research_product(
                            client, session, tools, product
                        )
                        print_product_result(product, result, calls_used)
                        cost = (tin / 1_000_000 * HAIKU_INPUT_PER_MTOK) + (
                            tout / 1_000_000 * HAIKU_OUTPUT_PER_MTOK
                        )
                        grand_in += tin
                        grand_out += tout
                        grand_cost += cost
                        print(f"  Tokens:     in={tin} out={tout}  (~${cost:.4f})")
                        report_sections.append(format_result_markdown(product, result, calls_used, tin, tout, cost))
                        excel_rows.append((product, result))
                    except Exception as exc:  # noqa: BLE001 - POC: keep going on any per-product failure
                        print("-" * 72)
                        print(f"{product['sku']} - \"{product['description']}\"")
                        print(f"  Error researching this product: {exc}")
                        report_sections.append(
                            f"## {product['sku']} - {product['description']}\n\n"
                            f"**Error researching this product:** {exc}\n"
                        )
                        excel_rows.append((product, None))

                print("-" * 72)
                print(
                    f"Total tokens: in={grand_in} out={grand_out}  |  "
                    f"Estimated total cost: ~${grand_cost:.4f}"
                )

                timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                report_path = f"sourcing_report_{timestamp}.md"
                write_report(report_sections, grand_in, grand_out, grand_cost, len(products), report_path)
                print(f"Report written to: {os.path.abspath(report_path)}")

                excel_path = args.output or f"sourcing_results_{timestamp}.xlsx"
                write_excel_report(excel_rows, excel_path)
                print(f"Excel results written to: {os.path.abspath(excel_path)}")


def main():
    asyncio.run(main_async())


if __name__ == "__main__":
    main()
