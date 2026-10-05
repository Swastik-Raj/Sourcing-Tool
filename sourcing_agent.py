"""
Competitive sourcing research POC.

For each product, connects directly to the Nimble MCP server (as an MCP client - not
Anthropic's server-side MCP connector) and exposes ONLY its search and extract tools to
Claude. Claude is given a hard budget of tool calls per product, must return a structured
(schema-enforced) result with up to 5 ranked candidates, and runs on Haiku for cost.

Setup:
    pip install "anthropic[mcp]" pydantic openpyxl python-dotenv
    .env file in this folder with:
        ANTHROPIC_API_KEY=...   (your Anthropic API key)
        NIMBLE_API_KEY=...      (your Nimble API key, used as the MCP bearer token)

Run:
    python sourcing_agent.py                                    (5 built-in sample products)
    python sourcing_agent.py --input products.xlsx               (SKU/Product/Keyword columns)
    python sourcing_agent.py --input products.xlsx --output results.xlsx
"""

import argparse
import asyncio
import collections
import csv
import contextvars
import json
import logging
import math
import os
import re
import sys
import time
from datetime import datetime
from urllib.parse import urlparse
from typing import Literal, Optional

import httpx2
import openpyxl
from dotenv import load_dotenv
from openpyxl.styles import Font, PatternFill
from mcp import ClientSession, MCPError
from mcp.client.streamable_http import streamable_http_client
from pydantic import BaseModel, ValidationError
from pydantic.json_schema import SkipJsonSchema

from anthropic import APIStatusError, AsyncAnthropic
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

# Pinned on every extract call. Unpinned, pages came back as other countries' storefronts
# (Korean won, Jamaican dollars, Polish zloty...), leaving prices unusable.
EXTRACT_GEO = {"country": "US", "locale": "en"}

MAX_CANDIDATES = 5  # distinct manufacturers returned (and shown in Excel) per product
MAX_TOOL_CALLS_PER_PRODUCT = 13  # searches + extracts combined, hard cap
# ^ first guess for 5 candidates: scaled up from 8 for 3 (8 x 5/3). Tune on real cost/results.
MAX_ITERATIONS_PER_PRODUCT = MAX_TOOL_CALLS_PER_PRODUCT + 3  # backstop against loops
# A single extract once hung for hours and froze a whole batch. Baseline (2026-09-28):
# 145 of 147 successful extracts finished within 60s, while the 8 calls that hit the old
# 120s limit cost 16 of 64 minutes. Past this the call fails (no retry) and Claude moves on.
TOOL_CALL_TIMEOUT_S = 60
# Products researched at once. Nimble waits were 86% of the baseline's time and everything
# ran one at a time. Lower this if the rate-limit retry counts at the end of a run climb.
MAX_CONCURRENT_PRODUCTS = 3
# Stage 2: run the three site searches for a product at the same time, in code, before Claude
# starts (they were always its first 3 calls, one after another). They go through the normal
# tool wrapper, so each counts against MAX_TOOL_CALLS_PER_PRODUCT (3 of 13 used up front).
PREFETCH_SITE_SEARCHES = True
PREFETCH_STAGGER_S = 0.0  # delay between the three starts, if Alibaba throttles simultaneous requests
ANTHROPIC_MAX_RETRIES = 5  # SDK retries 429/529/5xx itself with exponential backoff
NIMBLE_RATE_LIMIT_RETRIES = 3  # backoff 2s, 4s, 8s on a Nimble rate-limit response
MAX_TOOL_RESULT_CHARS = 2500  # truncate every tool result before it goes back to Claude
MAX_TOKENS_PER_TURN = 3000

# Sourcing only makes sense with real margin under L-Com's price, and only for candidates
# that aren't rubric-rejected. Both are business knobs - tune with the sourcing team.
MIN_MARGIN_PCT = 80  # 30% got eaten by shipping, storage and import taxes
# A unit price this many times cheaper than the next-cheapest candidate for the same
# product is treated as an extraction error (e.g. price divided by MOQ), not a bargain.
IMPLAUSIBLE_PRICE_RATIO = 10
IMPLAUSIBLE_NOTE = (
    "unit price implausible (far below the other candidates, or flagged as unusually low by the "
    "model) - likely extraction error or wrong listing, verify manually before ordering"
)
# The model's own "this price looks wrong" caveat. Fired once across every run through 29 Sep:
# the $0.11 UABS "adapter" that was really a printer cable.
LOW_PRICE_WARNING = re.compile(r"unusually low|suspicious|too (?:low|cheap|good)|verify (?:the |actual )?price|"
                               r"price (?:seems|looks|may be) (?:off|wrong|incorrect)", re.IGNORECASE)
MIN_RECOMMEND_ACCURACY = 80

# Hand-picked from L-Com's site. The first two are connector types like the original 10;
# the last three are untested categories (fiber optic, surge protection, antenna/RF).
# lcom_price: L-Com's per-unit USD price, looked up by hand from distributor listings.
PRODUCTS = [
    # L-Com's own distributor listing on Octopart.
    {"sku": "ECF504-AA", "description": "Flanged Panel Mounted USB 2.0 Coupler, Shielded, Type A/A Connectors", "lcom_price": 24.79},
    # LOWER CONFIDENCE: no confirmed L-Com listing found. This is Newark's qty-1 price standing in
    # for L-Com's own; other distributors range $2.59-$5.60 depending on source/quantity.
    {"sku": "C&P9M", "description": "Insertion Type D-Sub Connector, DB9 Male, Crimp Contacts", "lcom_price": 3.98},
    # Graybar's listed L-Com price for this exact SKU.
    {"sku": "FOA-020C", "description": "LC to SC Simplex Multimode Fiber Optic Adapter", "lcom_price": 63.56},
    # Newark's qty-1 listed price.
    {"sku": "LCSP1050", "description": "Coaxial Surge Protector, 18kA, 50 ohm, N-Type F/F Bulkhead, 1 Pole", "lcom_price": 33.04},
    # Direct reseller listing, consistent across two independent sources.
    {"sku": "HG2409U-PRO", "description": "2.4 GHz 9dBi Omnidirectional Antenna, N-Female Connector", "lcom_price": 99.00},
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
   or the equivalent spec/performance class for this product type). For antennas this
   covers frequency, gain AND band/port configuration: a dual-band, multi-band, "2way",
   MIMO or multi-port listing is a different product from a single-band, single-port target
   and scores 0 here, even if one of its bands matches.
3. shielding_material (0-20) - shielding and/or material match.
4. mount_form (0-20) - mount/form factor match. Mount wording differs by category:
   connectors - panel mount, bulkhead, keystone feed-thru; antennas - pole/mast mount,
   wall mount, magnetic base; surge protectors - bulkhead, DIN rail, ground plate.
5. gender_pins (0-20) - gender/pin configuration match (e.g. female-to-female), or the
   equivalent connector-orientation attribute for non-connector products.

CRITICAL RULE: award 0 points for any attribute the candidate listing's own text does not
explicitly state. Do not infer or assume a match from category context, product photos, or
"typical" industry defaults - only credit what the listing text actually says. No benefit
of the doubt. The converse also holds: if the listing explicitly states the attribute (e.g.
"Pole Mount") and it does not contradict the target, award the points - even when the target
description itself doesn't specify that attribute. match_percent must equal the sum of the
five scores, and every candidate MUST include attribute_breakdown.

PRODUCT FORM (separate from the score): set same_product_form to false when the listing is
a fundamentally different product form from the target - a cable, cord, lead or extension
(anything with a length of cable) when the target is an adapter/coupler/connector, or vice
versa; a multi-port variant of a single-port target; a different connector class (USB-B vs
USB-C, RJ45 vs RJ11, N vs SMA). A listing whose title says "cable", "cord" or gives a length
(50cm, 1.5m, 6ft) is a cable even if it also says "adapter" or "coupler". Mount style is NOT a
form difference: inline, keystone, bulkhead and panel-mount couplers are the same form -
score that difference in mount_form instead. Apply this the same way for every product. listing_form: 2-5
words for what the listing physically is (e.g. "patch cable with coupler end", "panel
mount coupler"). Candidates with same_product_form false are never recommended.

moq: whenever the listing shows a minimum order ("Min. order: 500 pieces", "100 Pieces (MOQ)"),
put it here - don't leave it "Not stated" and mention it only elsewhere.

listing_caveats: one short line on anything a buyer should check before ordering - the
listing covers several models or variants, its title names a different pin count, connector
or spec than the target, a material or mounting detail that differs from the target. Leave
MOQ and quantity-tier pricing out (that goes in moq and price). Also say so here when the
listing contradicts itself on an attribute - e.g. its title says "Female" but its attribute
table says "male", or the text says "compatible with CAT5e" while a table lists Cat3-Cat6A.
"" if none.

spec_lines: if you saw the listing's spec table or description (a detail-page extract), copy the
lines about type, jacket, length, category/rating, gender and contact/mount briefly and
verbatim, e.g. "Jacket: PVC; Length: Customized; Category: Cat5e Cat6". "" if you only saw the
search-result title. Never invent or paraphrase specs you did not see.
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
    price: Optional[str] = None  # raw listing price string, display only
    # Required and non-nullable on purpose: every nullable field is an anyOf union, and
    # strict output mode rejects the schema ("Schema is too complex") past ~11 of them.
    # 0 / "" mean "unknown".
    price_total: float  # numeric USD amount that `price` covers
    quantity_covered: int  # how many units price_total buys
    unit_price_confidence: Literal["stated", "inferred", "promo (regular price unknown)", "ambiguous"]
    unit_price_note: str
    same_product_form: bool  # false = different form (e.g. a cable vs a coupler): never recommended
    listing_form: str  # what the listing physically is, e.g. "patch cable with coupler end"
    listing_caveats: str  # the model's "check before ordering" line, "" if none
    spec_lines: str = ""  # spec-table lines the model actually saw (type/jacket/length/category), shown in the report
    # Filled by code after parsing (e.g. a score it overrode); kept out of the schema Claude sees.
    code_notes: SkipJsonSchema[list[str]] = []
    moq: Optional[str] = None
    email: Optional[str] = None
    match_percent: Optional[int] = None
    # Required: when it was optional the model once skipped it for every HDFF candidate,
    # which then read as "not enough margin" instead of "scoring failed".
    attribute_breakdown: AttributeBreakdown
    url: Optional[str] = None


class SourcingResult(BaseModel):
    # No free-text notes: the model's own note restated scores that drifted from the table.
    # The report builds its summary from the scores instead (score_summary).
    product: str
    no_match: bool
    no_match_reason: Optional[str] = None
    candidates: list[Candidate] = []


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

You have a HARD BUDGET of {MAX_TOOL_CALLS_PER_PRODUCT} tool calls total for this product, sized
for finding and verifying up to {MAX_CANDIDATES} distinct candidates.
Spend it deliberately: a handful of site-search extracts across different sites/queries to
surface multiple distinct candidates, plus the occasional individual listing page only if a
title is genuinely ambiguous on a rubric attribute. Once you've used your budget, or you
already have enough to score confidently, STOP calling tools immediately and give your
final answer. Do not narrate your process in free text at any point - every response you
give must be the final structured result, nothing else.

{RUBRIC_INSTRUCTIONS}

PRICING - for every candidate, work out what quantity the listed price actually buys:
- price: the raw price text exactly as shown on the listing, including any regular /
  struck-through price and promo wording next to it.
- price_total: that price as a number in USD. If a range is shown, use the HIGHER end.
  0 if no price is shown.
- PROMO PRICES: AliExpress often shows a one-time new-shopper price with the regular price
  right after it, e.g. "$1.09 $5.77 -81% New shoppers save $4.68" - $1.09 is the promo,
  $5.77 the real price. Whenever promo wording appears ("new shoppers", "new user", "welcome
  deal", "first order", "with coupon", "marked down from"), use the REGULAR price as
  price_total. If the regular price is not shown, set unit_price_confidence to
  "promo (regular price unknown)".
- quantity_covered: how many units price_total buys (e.g. "$0.33 / piece" -> 1;
  "$19.75 ... (10PCS)" or "pack of 10" -> 10). 0 if unknown. Only use a number above 1 when
  the listing states a pack/set/lot size attached to THAT price. The MOQ / minimum order is
  NOT quantity_covered: "$1.27-1.58, Min. order 500 pieces" and "US$0.99-6.88, 100 Pieces
  (MOQ)" are per-piece prices (-> 1).
- unit_price_note: one short line on how you read the quantity ("" if obvious).
- unit_price_confidence:
  "stated"    - the listing explicitly says per piece/unit, or explicitly states the pack size.
  "inferred"  - not explicit, but the listing text makes it clear (e.g. title says "10PCS").
  "ambiguous" - you cannot tell whether the price is per unit or for a pack/MOQ batch
                (e.g. "$19.75, MOQ 10 pcs" with nothing saying which), or the price is not
                in USD. DO NOT GUESS - mark it ambiguous and say why in unit_price_note.
Never divide the price yourself; just report price_total and quantity_covered.

Return up to {MAX_CANDIDATES} candidates in `candidates`, ranked best first, balancing match quality
against price (a slightly lower match % at a much lower price can rank above a perfect
match at a high price) - report both numbers per candidate, never silently substitute one
for the other. Each candidate must be a DIFFERENT manufacturer/supplier - never list several
listings from the same company just to fill the slots. If you only find 1, 2 or 3 genuinely
distinct, plausible candidates, return only that many - do not force a weak extra pick just
to reach {MAX_CANDIDATES}.

If you find no plausible candidate at all, or every candidate you found scores very low,
set no_match to true, leave candidates empty, and briefly say why in no_match_reason - do
not guess or force a weak candidate into looking like a real match.
"""


PRICE_PATTERN = re.compile(r"(?:US ?\$|\$)\s?\d")
LISTING_URL = re.compile(r"/item/|/product-detail/|/product/")  # AliExpress / Alibaba / Made-in-China
PRE_PRICE_CONTEXT = 400  # chars kept before the first price: the first listing's title/supplier


def is_chrome_line(line: str) -> bool:
    """Menu items, buttons and badges: bullets or 1-3 word lines with no digits (a listing
    line has a price, a quantity or a long title)."""
    words = re.sub(r"[*#>`_\\|-]", " ", line).split()
    return not words or (not re.search(r"\d", line) and (len(words) <= 3 or re.match(r"\s*[*+-]\s", line)))


def strip_page_chrome(text: str) -> str:
    """Cut an extracted search page down to its listings before truncation. Search pages
    put 1-30k chars of menus, filters and link URLs ahead of the first listing (Made-in-China:
    listings start ~30,000 chars in), so plain truncation handed Claude nothing but menus."""
    try:  # Nimble returns {"content": "...markdown..."} - decode so newlines are real
        text = json.JSONDecoder().raw_decode(text)[0]["content"]
    except (ValueError, KeyError, TypeError):
        pass
    text = re.sub(r"!\[[^\]]*\]\([^)]*\)", "", text)  # images
    text = re.sub(r"\[\s*\]\([^)]*\)", "", text)  # empty links
    # Keep URLs only for product listings (the candidate needs one); other links -> their text.
    text = re.sub(r"\[([^\]]*)\]\(([^)\s]*)\)",
                  lambda m: f"[{m[1]}]({m[2].split('?')[0]})" if LISTING_URL.search(m[2]) else m[1], text)
    text = re.sub(r"(?<!\()https?://[^\s)\]]+",
                  lambda m: m[0].split("?")[0] if LISTING_URL.search(m[0]) else "", text)
    text = " ".join(line.strip() for line in text.split("\n") if not is_chrome_line(line))
    # ponytail: "listings start near the first price" - fine for search pages; a page with a
    # stray $ in its header would keep some chrome, never lose listings.
    first_price = PRICE_PATTERN.search(text)
    return text[max(0, first_price.start() - PRE_PRICE_CONTEXT):] if first_price else text


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


# Products run concurrently, so every progress line is tagged with its product's SKU.
CURRENT_SKU = contextvars.ContextVar("sku", default="")


def log(message: str = "") -> None:
    sku = CURRENT_SKU.get()
    print(f"[{sku}] {message}" if sku else message)


class _RetryCounter(logging.Handler):
    """Counts the Anthropic SDK's own retry log lines ("Retrying due to status code 429")."""
    def emit(self, record):
        m = re.search(r"status code (\d+)", record.getMessage())
        if m:
            ANTHROPIC_RETRIES[m[1]] += 1


ANTHROPIC_RETRIES = collections.Counter()
NIMBLE_RATE_LIMITS = {"count": 0}
_sdk_log = logging.getLogger("anthropic")
_sdk_log.setLevel(logging.DEBUG)
_sdk_log.propagate = False  # count, don't print
_sdk_log.addHandler(_RetryCounter())


def is_rate_limited(message: str) -> bool:
    return bool(re.search(r"\b429\b|rate.?limit|too many requests", message, re.IGNORECASE))


# Timing instrumentation: one row per Nimble call, one per product (see timing_summary).
CALL_LOG = []  # {"sku", "tool", "site", "url", "start", "secs", "outcome"}
PRODUCT_LOG = []  # {"sku", "secs", "calls"}


def site_of(url: str) -> str:
    host = urlparse(url or "").netloc
    return next((s for s in SOURCING_SITES if s.split(".")[0] in host), host or "-")


def make_bounded_tool(tool, session: ClientSession, budget: ToolBudget, sku: str = "-"):
    tool_name = tool.name
    input_schema = getattr(tool, "input_schema", None) or getattr(tool, "inputSchema", None)
    description = (tool.description or "")[:600]

    async def call(**kwargs):
        if not budget.take():
            return (
                f"Tool call budget exhausted ({budget.limit} calls used for this product). "
                "Stop searching and return your best final answer now."
            )
        if "extract" in tool_name:
            kwargs.update(EXTRACT_GEO)  # override whatever the model passed
        log(f"    -> {tool_name}({kwargs})")
        row = {"sku": sku, "tool": tool_name, "site": site_of(kwargs.get("url")), "url": kwargs.get("url"),
               "start": time.monotonic(), "outcome": "ok"}
        try:
            for attempt in range(NIMBLE_RATE_LIMIT_RETRIES + 1):
                try:
                    result = await session.call_tool(
                        name=tool_name, arguments=kwargs, read_timeout_seconds=TOOL_CALL_TIMEOUT_S
                    )
                    error_text = " ".join(getattr(b, "text", "") for b in result.content) if result.is_error else ""
                except MCPError as exc:
                    if attempt == NIMBLE_RATE_LIMIT_RETRIES or not is_rate_limited(str(exc)):
                        raise
                    error_text = str(exc)
                if not (error_text and is_rate_limited(error_text)) or attempt == NIMBLE_RATE_LIMIT_RETRIES:
                    break
                NIMBLE_RATE_LIMITS["count"] += 1
                log(f"    Nimble rate-limited - retrying in {2 ** (attempt + 1)}s")
                await asyncio.sleep(2 ** (attempt + 1))
        except MCPError as exc:  # timeout or tool-side error: tell Claude, don't hang the batch
            row["outcome"] = "timeout" if "timed out" in str(exc).lower() else "error"
            log(f"    <- FAILED after {time.monotonic() - row['start']:.1f}s: {exc}")
            return f"This {tool_name} call failed ({exc}). Try a different page or query, or give your answer."
        except Exception as exc:
            # A bug in this wrapper would otherwise reach Claude as a silent "tool error" on every
            # call (a whole run once returned 0 candidates that way). Make it loud.
            row["outcome"] = "error"
            log(f"    <- WRAPPER BUG: {type(exc).__name__}: {exc}")
            raise
        finally:
            row["secs"] = time.monotonic() - row["start"]
            CALL_LOG.append(row)
        text = " ".join(
            block.text for block in result.content if getattr(block, "type", None) == "text"
        )
        if "extract" in tool_name:
            text = strip_page_chrome(text)
        log(f"    <- ({row['secs']:.1f}s) {truncate_text(text, limit=300)}")
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


def site_search_url(site: str, text: str) -> str:
    """The site's own search page for `text` (the same URL shapes the prompt describes)."""
    words = re.findall(r"[A-Za-z0-9]+", text)
    return SITE_SEARCH_URLS[site].format(query=("_" if site == "made-in-china.com" else "+").join(words))


async def prefetch_site_searches(runner_tools: list, product: dict) -> str:
    """All three site searches at once, through the budgeted extract wrapper."""
    extract = next(t for t in runner_tools if "extract" in t.name)

    async def one(i: int, site: str):
        await asyncio.sleep(i * PREFETCH_STAGGER_S)
        url = site_search_url(site, product["description"])
        return site, url, await extract.call({"url": url, "driver": "vx8"})

    pages = await asyncio.gather(*(one(i, site) for i, site in enumerate(SOURCING_SITES)))
    return "\n\n".join(f"=== {site} search results ({url}) ===\n{text}" for site, url, text in pages)


async def research_once(client: AsyncAnthropic, session: ClientSession, tools, product: dict):
    budget = ToolBudget(MAX_TOOL_CALLS_PER_PRODUCT)
    runner_tools = [make_bounded_tool(t, session, budget, product['sku']) for t in tools]
    prompt = build_prompt(product)
    if PREFETCH_SITE_SEARCHES:
        prompt += (
            f"\nALREADY FETCHED: the first search on each site was run for you with the product "
            f"description as the query - that used 3 of your {MAX_TOOL_CALLS_PER_PRODUCT} tool calls. "
            "Don't repeat these searches; spend the rest on different queries or listing pages.\n\n"
            + await prefetch_site_searches(runner_tools, product)
        )

    runner = client.beta.messages.tool_runner(
        model=MODEL,
        max_tokens=MAX_TOKENS_PER_TURN,
        max_iterations=MAX_ITERATIONS_PER_PRODUCT,
        tools=runner_tools,
        output_config={
            "format": {"type": "json_schema", "schema": strict_json_schema(SourcingResult)}
        },
        messages=[{"role": "user", "content": prompt}],
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
            # Strict mode can't enforce maxItems, so enforce it here: Excel has MAX_CANDIDATES blocks.
            parsed.candidates = dedupe_manufacturers(parsed.candidates)[:MAX_CANDIDATES]
            apply_form_rules(parsed, product)
            apply_spec_rules(parsed, product)
            apply_price_rules(parsed)
            apply_reviewer_exclusions(parsed, product["sku"])
        except ValidationError:
            parsed = None
    return parsed, total_in, total_out, budget.used


# Names the model uses when a listing doesn't name its maker. Two of these are different
# anonymous sellers, not one manufacturer, so they're exempt from dedup.
# ponytail: keyword list, extend when a new placeholder wording shows up.
PLACEHOLDER_MANUFACTURER = re.compile(r"unknown|unnamed|generic|not stated|unbranded|seller", re.IGNORECASE)


# The model's same_product_form call was inconsistent (2026-09-28: kept a "50cm ... Adapter
# Cable" for ECF504-UABS, excluded plain inline couplers for ECF504-SC6). A stated length,
# or cable/cord with no part word in the title, marks a cable - 9 of 9 hits across three past
# runs were real cables, 0 false positives. Only applied when the target is a coupler/adapter.
CABLE_LENGTH = re.compile(r"\b\d+(?:\.\d+)?\s*(?:cm|m|ft|feet|foot|meters?|metres?|inch(?:es)?)\b", re.IGNORECASE)
CABLE_WORD = re.compile(r"\b(?:cable|cord)s?\b", re.IGNORECASE)
PART_WORD = re.compile(r"\b(?:coupler|adapter|adaptor|connector|jack|socket|plug|coupling|joiner|keystone|feed.?thr(?:u|ough))s?\b", re.IGNORECASE)


# Phrases that describe the item itself as a cable, so they count even when "adapter" or
# "connector" is also in the title - the part-word exemption let "USB 2.0 ... Printer Data
# Cable Adapter with a Male to B Male ... Magnetic Ring" through as an adapter (UABS, 2026-09-29,
# $0.11). Checked on 286 past titles: 15 hits, all real cables. Broader phrases ("Ethernet
# Cable", "Patch Cord") also hit couplers that name the cable they connect, so they're left out.
STRONG_CABLE = re.compile(r"\b(?:data|printer|scanner|extension|charging|charger|sync|patch)\s+(?:data\s+)?cables?\b"
                          r"|\bferrite\b|\bmagnetic ring\b|\b\d+\s*awg\b|\bpvc jacket\b", re.IGNORECASE)


# Spec-sheet wording seen on HDFF's recommended page (xtz-tech, 2026-10-04: "Jacket PVC", "AWG 24/28/26",
# "Length (Customized)") - note only, never an exclusion. Candidate text rarely carries page specs (the
# search reads titles), and no saved report has any, so this can't be replayed yet: 0 hits on all 10 reports.
SPEC_CABLE = re.compile(r"\bpvc jacket\b|\bjacket\s*:?\s*pvc\b|\b\d+(?:/\d+)*\s*awg\b|\bawg\s*\d+|"
                        r"\b(?:custom(?:ized|ised)?|adjustable)\s+length\b|\blength\s*\(custom", re.IGNORECASE)


def cable_evidence(title: str) -> str:
    if CABLE_LENGTH.search(title):
        return f"length {CABLE_LENGTH.search(title)[0].strip()} in title"
    if STRONG_CABLE.search(title):
        return f"\"{STRONG_CABLE.search(title)[0]}\" in title"
    if CABLE_WORD.search(title) and not PART_WORD.search(title):
        return "titled as a cable"
    return ""


INLINE_COUPLER = re.compile(r"\bin[\s-]?line\b.*\bcoupl|\bcoupl\w*\b.*\bin[\s-]?line\b", re.IGNORECASE)
MULTI_PORT = re.compile(r"\b\d+\s*-?\s*(?:port|way|gang)s?\b|\b(?:dual|quad|multi)\b", re.IGNORECASE)
# Contact / termination type. A crimp-contact target is a different part from a solder-cup
# listing (C&P9M, Srijan's RS check, 2026-10-04: all four candidates were solder type).
CONTACT_TYPES = {
    "crimp": re.compile(r"\bcrimp(?:ed|ing)?\b", re.IGNORECASE),
    "solder": re.compile(r"\bsolder\b", re.IGNORECASE),
    "PCB/DIP": re.compile(r"\bpcb\b|\bdip\b", re.IGNORECASE),
    "IDC": re.compile(r"\bidc\b|insulation[\s-]displacement", re.IGNORECASE),
}


def contact_types(text: str) -> set:
    return {name for name, pat in CONTACT_TYPES.items() if pat.search(text)}


def candidate_contact_text(c: Candidate) -> str:
    return f"{c.listing_title or ''} {c.listing_form or ''} {c.listing_caveats or ''} {c.spec_lines or ''}"


# A listing that contradicts itself (Kabasi: "compatible with CAT5e" next to a "Cat.3-Cat.6A"
# table; LUNG KAY: title "Type A Female", attribute table "Type A male") is flagged, not zeroed:
# a category range alone is legitimate (Hyconnect's "Cat5e Cat6 Cat6a").
COMPATIBLE_CAT = re.compile(r"compatible with\s+cat\.?\s?(\d)", re.IGNORECASE)
CAT_RANGE = re.compile(r"cat\.?\s?(\d)\w*\s*[-–]\s*cat\.?\s?(\d)", re.IGNORECASE)
PORT_GENDER = re.compile(r"\btype[\s-]?([ab])\s+(male|female)\b", re.IGNORECASE)


def contradictions(c: Candidate) -> list:
    text = f"{c.listing_title or ''} {c.spec_lines or ''} {c.listing_caveats or ''}"
    found = []
    single = {int(n) for n in COMPATIBLE_CAT.findall(text)}
    top = max((int(hi) for _, hi in CAT_RANGE.findall(text)), default=0)
    if single and top > max(single):
        found.append(f"listing contradicts itself on category: \"compatible with CAT{max(single)}\" but also lists "
                     f"up to Cat.{top} - confirm the rating with the seller")
    genders = {}
    for port, gender in PORT_GENDER.findall(text):
        genders.setdefault(port.upper(), set()).add(gender.lower())
    for port, seen in sorted(genders.items()):
        if len(seen) > 1:
            found.append(f"listing contradicts itself on gender: Type {port} is both male and female - "
                         "confirm with the seller")
    return found


REVIEWER_EXCLUSIONS_CSV = os.path.join(os.path.dirname(os.path.abspath(__file__)), "reviewer_exclusions.csv")


def apply_reviewer_exclusions(result: SourcingResult, sku: str, path: str = REVIEWER_EXCLUSIONS_CSV) -> None:
    """Honor reviewer_exclusions.csv (sku, supplier_or_url, reason): a candidate for that SKU whose maker
    name or URL contains supplier_or_url (case-insensitive) is excluded as wrong product form."""
    try:
        with open(path, newline="", encoding="utf-8") as f:
            rows = [r for r in csv.DictReader(f) if r["sku"].strip().upper() == sku.strip().upper()]
    except FileNotFoundError:
        return
    for c in result.candidates:
        haystack = f"{c.manufacturer or ''} {c.url or ''}".lower()
        for r in rows:
            if r["supplier_or_url"].strip().lower() in haystack:
                c.same_product_form = False
                c.listing_form = f"excluded by reviewer: {r['reason'].strip()}"
                break


CONNECTOR_TOKENS = re.compile(r"\b(?:rj45|rj11|rj12|usb|hdmi|dvi|vga|displayport|db9|db25|d-?sub|lc|sc|st|fc|"
                              r"n-?type|sma|bnc|f-?type)\b", re.IGNORECASE)


def apply_form_rules(result: SourcingResult, product: dict) -> None:
    """Deterministic product-form calls where the model was inconsistent across similar
    products (2026-09-28 runs), applied only when the target is a coupler/adapter:
    - cables are excluded: a stated length or cable/cord with no part word in the title
      (kept a "50cm ... Adapter Cable" for ECF504-UABS);
    - inline couplers are the same form as a panel-mount/keystone coupler target, left to
      mount_form to score (SC5E excluded one, SC6 and TDG1026KS-C6 kept one). Only when
      the listing shows the target's connector type and isn't a multi-port variant."""
    target = product.get("description", "")
    target_contacts = contact_types(target)
    if target_contacts:  # applies to any target that states a contact type, not only couplers
        for c in result.candidates:
            found = contact_types(candidate_contact_text(c))
            if found and not found & target_contacts and c.same_product_form:
                c.same_product_form = False
                c.listing_form = (f"wrong contact type ({'/'.join(sorted(found))} vs target "
                                  f"{'/'.join(sorted(target_contacts))}; model said: {c.listing_form})")
    if not re.search(r"coupler|adapter|adaptor", target, re.IGNORECASE) or CABLE_WORD.search(target):
        return
    def connectors(text):  # Cat5e/Cat6/Cat6a/Cat7 couplers are RJ45 even when a title doesn't say so
        found = {t.lower().replace("-", "") for t in CONNECTOR_TOKENS.findall(text)}
        return found | ({"rj45"} if re.search(r"\bcat\s?[5-8]", text, re.IGNORECASE) else set())

    target_connectors = connectors(target)
    for c in result.candidates:
        title = c.listing_title or ""
        why = cable_evidence(title)
        if why:
            if c.same_product_form:
                c.same_product_form = False
                c.listing_form = f"cable ({why}; model said: {c.listing_form})"
            continue
        listing_connectors = connectors(title)
        if (not c.same_product_form and re.search(r"coupler", target, re.IGNORECASE)
                and INLINE_COUPLER.search(f"{title} {c.listing_form}") and not MULTI_PORT.search(title)
                and target_connectors & listing_connectors):
            c.same_product_form = True
            c.listing_form = f"{c.listing_form} (inline coupler: same form, mount scored separately)"


MULTI_BAND = re.compile(r"\b(?:dual|multi|tri|quad)[\s-]?band\b|\b\d[\s-]?way\b|\bmimo\b|"
                        r"\b(?:\d|dual|two|multi)[\s-]?ports?\b", re.IGNORECASE)
GHZ = re.compile(r"(\d+(?:\.\d+)?)\s*g(?:hz)?\b", re.IGNORECASE)


def band_port_mismatch(title: str) -> str:
    """Why a listing is multi-band / multi-port, or "" - e.g. "Dual Band", "2way", or two
    different GHz frequencies ("2.4GHz 5.8GHz", "2.4/5.8GHz")."""
    m = MULTI_BAND.search(title)
    if m:
        return m[0]
    slash = re.search(r"(\d+(?:\.\d+)?)\s*(?:g(?:hz)?)?\s*/\s*(\d+(?:\.\d+)?)\s*g(?:hz)?\b", title, re.IGNORECASE)
    freqs = {float(f) for f in GHZ.findall(title) if 0.3 <= float(f) <= 100}
    if slash:
        freqs |= {float(slash[1]), float(slash[2])}
    return f"{len(freqs)} bands ({', '.join(f'{f:g} GHz' for f in sorted(freqs))})" if len(freqs) > 1 else ""


def apply_spec_rules(result: SourcingResult, product: dict) -> None:
    """Antennas: a dual-band/MIMO/multi-port listing is a different product from a single-band,
    single-port target, but the rubric had nowhere to say so - Roho's "2way 2.4GHz 5.8GHz ...
    Dual Band MIMO" antenna won HG2409U-PRO at 85% and then 95%. Zero its category_spec,
    unless the target itself asks for multiple bands or ports."""
    target = product.get("description", "")
    if not re.search(r"\bantenna\b", target, re.IGNORECASE) or band_port_mismatch(target):
        return
    for c in result.candidates:
        why = band_port_mismatch(c.listing_title or "")
        if why and c.attribute_breakdown is not None and c.attribute_breakdown.category_spec:
            c.attribute_breakdown.category_spec = 0
            c.code_notes.append(f"Category/spec set to 0: multi-band/multi-port listing ({why}) vs "
                                "single-band target")


PROMO_WORDS = re.compile(r"new shoppers?|new user|welcome deal|first order|with coupon|promo(?:tional)?", re.IGNORECASE)
# The promo wording must follow the amount directly ("$1.09 new shopper", "$1.09 (promotional"):
# a looser gap read "$17.13 $22.84 -25% Regular price after promo discount" as a $22.84 promo.
PROMO_AMOUNT = re.compile(r"\$\s*(\d+(?:\.\d+)?)\s*\(?\s*(?:new shoppers?|new user|welcome|first order|promo)"
                          r"|(?:promo(?:tional)?(?:\s+price)?|new shoppers?(?:\s+price)?)[^$;]{0,20}\$\s*(\d+(?:\.\d+)?)", re.IGNORECASE)
REGULAR_AMOUNT = re.compile(r"(?:regular|base|original|list|full)(?:\s+price)?\s*(?:is|of|:)?\s*~?\s*\$\s*(\d+(?:\.\d+)?)",
                            re.IGNORECASE)


# The lookahead keeps "$22.84 -25%" (a discount) from reading as a $22.84-25 range.
PRICE_RANGE = re.compile(r"\$\s*(\d+(?:\.\d+)?)\s*[-–]\s*\$?\s*(?![\d.]+\s*%)(\d+(?:\.\d+)?)")
# "$2.68 100-999 pieces $2.58 >=1,000 pieces": each price tied to its own quantity tier
PRICE_TIER = re.compile(r"\$\s*(\d+(?:\.\d+)?)\s*\(?\s*((?:\d[\d,]*\s*[-–]\s*\d[\d,]*|[≥>]=?\s*\d[\d,]*)\s*"
                        r"(?:pieces|pcs|pc))", re.IGNORECASE)


def price_tier_note(c: Candidate) -> str:
    """When the listing ties different prices to quantity tiers, say which tier was used."""
    tiers = PRICE_TIER.findall(c.price or "")
    if len(tiers) < 2:
        return ""
    used = [t for p, t in tiers if abs(float(p) - c.price_total) < 0.005]
    others = ", ".join(f"${p} at {t}" for p, t in tiers if abs(float(p) - c.price_total) >= 0.005)
    return f"price is for the {used[0]} tier (others: {others})" if used else ""


def apply_price_rules(result: SourcingResult) -> None:
    """Backstop for new-shopper promo prices the model reported as the real price, despite the
    prompt (e.g. UABS 2026-09-29: "$1.09 new shopper discount from base price ~$2.31" priced at
    $1.09). If the price used is the promo amount: use the stated regular price, or label it
    promo when none is given. Also: a listed range ("$0.05-1.58") always uses its high end -
    CAPUSB-A candidate 3 took the low end while other rows took the high one (Srijan, 2026-10-04)."""
    for c in result.candidates:
        rng = PRICE_RANGE.search(c.price or "")
        if rng and c.price_total and abs(c.price_total - float(rng[1])) < 0.005 and float(rng[2]) > float(rng[1]):
            c.code_notes.append(f"listed range ${rng[1]}-{rng[2]}: low end ${c.price_total:g} replaced by the high end")
            c.price_total = float(rng[2])
    for c in result.candidates:
        text = f"{c.price or ''} {c.unit_price_note or ''}"
        if c.unit_price_confidence not in ("stated", "inferred") or not c.price_total or not PROMO_WORDS.search(text):
            continue
        regulars = {float(r) for r in REGULAR_AMOUNT.findall(text)}
        # "Regular price $2.28 (promo $1.09...)" would otherwise read $2.28 as a promo amount
        promos = {float(a or b) for a, b in PROMO_AMOUNT.findall(text)} - regulars
        if not any(abs(p - c.price_total) < 0.005 for p in promos):
            continue  # the model already used a non-promo price
        regular = [r for r in regulars if r > c.price_total]
        if regular:
            c.code_notes.append(f"promo ${c.price_total:g} replaced by the listing's regular price ${max(regular):g}")
            c.price_total = max(regular)
        else:
            c.code_notes.append(f"${c.price_total:g} is a promo price; regular price not shown")
            c.unit_price_confidence = "promo (regular price unknown)"


def dedupe_manufacturers(candidates: list) -> list:
    """Keep only the highest-scoring candidate per manufacturer (case-insensitive,
    trimmed), preserving rank order. No backfill - fewer candidates is the right answer."""
    def key(c):
        name = (c.manufacturer or "").strip().lower()
        return None if not name or PLACEHOLDER_MANUFACTURER.search(name) else name

    best = {}
    for i, c in enumerate(candidates):
        k = key(c)
        if k is not None and (k not in best or
                              (candidate_score(c)[0] or 0) > (candidate_score(candidates[best[k]])[0] or 0)):
            best[k] = i
    return [c for i, c in enumerate(candidates) if key(c) is None or best[key(c)] == i]


def is_garbled(result: Optional[SourcingResult]) -> bool:
    """Haiku occasionally emits a whole response with spaces sprayed mid-word
    ("al ib ab a.c om", "$2 .1 0"). URLs and domains never contain whitespace, so any
    in those fields means the response text can't be trusted."""
    return result is not None and any(
        re.search(r"\s", field or "")
        for c in result.candidates
        for field in (c.url, c.distributor_site)
    )


async def research_product(client: AsyncAnthropic, session: ClientSession, tools, product: dict):
    """research_once, retried once if the response comes back garbled. A second garbled
    response is dropped (reported as no structured result) rather than written out."""
    total_in = total_out = calls = 0
    for attempt in (1, 2):
        parsed, tin, tout, used = await research_once(client, session, tools, product)
        total_in, total_out, calls = total_in + tin, total_out + tout, calls + used
        if not is_garbled(parsed):
            return parsed, total_in, total_out, calls
        log(f"  Response text came back garbled (attempt {attempt}) - "
              + ("retrying." if attempt == 1 else "dropping it."))
    return None, total_in, total_out, calls


def candidate_score(candidate: Candidate) -> tuple:
    """Returns (computed_match_percent, {attr: score}) - recomputed from the breakdown so a
    miscounted match_percent from the model can't slip through. A wrong product form (a cable
    for a coupler) zeroes the match: a 20-point deduction still let one win on price."""
    breakdown = candidate.attribute_breakdown
    if breakdown is None:
        return None, {}
    scores = breakdown.model_dump()
    if not candidate.same_product_form:
        return 0, scores
    return sum(max(0, min(20, v)) for v in scores.values()), scores


def match_tier(match_percent: Optional[int]) -> str:
    if match_percent is None:
        return ""
    if match_percent >= 90:
        return "Auto-accepted"
    if match_percent >= 80:
        return "Flagged for manual review"
    return "Rejected"


def candidate_comment(candidate: Candidate) -> str:
    """Built only from the scores and flags - never restates numbers the table doesn't show."""
    _, scores = candidate_score(candidate)
    if not scores:
        return "Scoring failed: no rubric scores returned for this candidate."
    if not candidate.same_product_form:
        if (candidate.listing_form or "").startswith("excluded by reviewer"):
            return candidate.listing_form
        return f"Wrong product form ({candidate.listing_form or 'not stated'}) - excluded."
    zero_attrs = [ATTRIBUTE_LABELS[attr] for attr in ATTRIBUTES if not scores.get(attr)]
    text = (f"0 pts on: {', '.join(zero_attrs)}" if zero_attrs
            else "All five rubric attributes matched something explicit on the listing.")
    return "; ".join([text] + candidate.code_notes + contradictions(candidate))


PACK_WORDS = r"(?:pcs|pc|pieces?|packs?|sets?|lots?|units?|count)"


def pack_size_supported(candidate: Candidate) -> bool:
    """quantity_covered > 1 only counts when that number sits next to a pack word in the
    title or price text - not in an MOQ. Dividing by an MOQ produced HDFF's $0.0032 and
    Suzhou Bulovb's $0.0688 ("US$0.99-6.88", 100 Pieces MOQ)."""
    q = candidate.quantity_covered
    if q <= 1:
        return True
    text = f"{candidate.listing_title or ''} {candidate.price or ''}"
    pattern = rf"(?<!\d){q}\s*{PACK_WORDS}\b|\b(?:pack|set|lot|bag|box)\s*(?:of\s*)?{q}(?!\d)|\bx\s*{q}(?!\d)"
    for m in re.finditer(pattern, text, re.IGNORECASE):
        nearby = text[max(0, m.start() - 12): m.end() + 12].lower()
        if "moq" not in nearby and "min" not in nearby:
            return True
    # An explicit per-piece price that matches, e.g. "$23.68 ($4.74/pc)" for a 5-pack.
    for m in re.finditer(r"\$\s*(\d+(?:\.\d+)?)\s*/\s*(?:pc|pcs|piece|unit)\b", text, re.IGNORECASE):
        if candidate.price_total and abs(float(m[1]) - candidate.price_total / q) <= 0.02 * candidate.price_total / q + 0.01:
            return True
    return False


def unit_price(candidate: Candidate) -> Optional[float]:
    """True single-unit cost, or None when it can't be determined confidently. A quantity
    the listing didn't attach to the price is treated as 1 (see pack_size_supported)."""
    c = candidate
    if c.unit_price_confidence == "ambiguous" or not c.price_total or not c.quantity_covered:
        return None
    return c.price_total / (c.quantity_covered if pack_size_supported(c) else 1)


def price_vs_lcom(unit: Optional[float], lcom_price: Optional[float]) -> Optional[tuple]:
    """(dollars cheaper per unit, percent cheaper) vs L-Com. Negative = more expensive."""
    if unit is None or not lcom_price:
        return None
    diff = lcom_price - unit
    return diff, diff / lcom_price * 100


def format_vs_lcom(vs: Optional[tuple]) -> str:
    if vs is None:
        return "n/a"
    diff, pct = vs
    return f"${abs(diff):.2f} ({abs(pct):.0f}%) {'cheaper' if diff >= 0 else 'more expensive'}"


def format_unit(unit: Optional[float]) -> str:
    if unit is None:
        return "UNDETERMINED"
    return f"${unit:.2f}" if unit >= 0.1 else f"${unit:.4f}"  # don't round a $0.0032 bug to $0.00


def implausible_units(result: Optional[SourcingResult]) -> set:
    """Indices of candidates whose unit price is > IMPLAUSIBLE_PRICE_RATIO x cheaper than the
    next-cheapest candidate for the same product, or that the model itself flagged as
    unusually low. The comparison needs >= 3 priced candidates (with 2, "one is a parsing
    error" and "two different real prices" look the same) and only uses stated/inferred
    prices as the reference - a leaked $1.09 promo once hid a $0.11 outlier at 9.9x."""
    cands = result.candidates if result else []
    units = [unit_price(c) for c in cands]
    reliable = [u if u is not None and c.unit_price_confidence in ("stated", "inferred") else None
                for c, u in zip(cands, units)]
    suspect = {i for i, c in enumerate(cands) if LOW_PRICE_WARNING.search(c.listing_caveats or "")}
    # ponytail: a 2-candidate product with a real parsing error now goes unflagged; add an
    # absolute floor (e.g. vs L-Com's price) if that shows up in practice.
    if sum(u is not None for u in units) < 3:
        return suspect
    for i, u in enumerate(units):
        others = [x for j, x in enumerate(reliable) if j != i and x is not None]
        if u is not None and others and u * IMPLAUSIBLE_PRICE_RATIO < min(others):
            suspect.add(i)
    return suspect


def checked_unit(result: SourcingResult, i: int) -> Optional[float]:
    """unit_price() for comparisons against L-Com: None when implausible."""
    return None if i in implausible_units(result) else unit_price(result.candidates[i])


def confidence_label(result: SourcingResult, i: int) -> str:
    c = result.candidates[i]
    if i in implausible_units(result):
        return f"{c.unit_price_confidence} (IMPLAUSIBLE)"
    label = c.unit_price_confidence or "?"
    if not pack_size_supported(c):
        label += f" (qty {c.quantity_covered} ignored)"
    return label + (" (FLAGGED)" if unit_price(c) is None else "")


# Among candidates that clear both bars, the most accurate wins. Only candidates within
# TIE_BAND_POINTS of that top accuracy count as a tie, broken by: stated price (promo last, so it
# never beats a stated wholesale price), then a named maker, then the lowest price.
TIE_BAND_POINTS = 5
CONFIDENCE_RANK = {"stated": 0, "inferred": 1, "promo (regular price unknown)": 2}


def is_named(candidate: Candidate) -> bool:
    name = (candidate.manufacturer or "").strip()
    return bool(name) and not PLACEHOLDER_MANUFACTURER.search(name)


def recommend(result: Optional[SourcingResult], lcom_price: Optional[float]) -> tuple:
    """Returns (index of recommended candidate or None, one-line reason). The reason always
    names the rule that actually decided it."""
    if result is None:  # a crash or an unusable response is never a genuine "no match"
        return None, error_recommendation("no result for this product")
    if result.no_match or not result.candidates:
        return None, "No candidates found - nothing to source."
    if not lcom_price:
        return None, "No L-Com reference price for this product - cannot judge margin, not recommending."
    cands = result.candidates
    unscored = [i for i, c in enumerate(cands) if candidate_score(c)[0] is None]
    if len(unscored) == len(cands):
        return None, ("Scoring failed - the model returned these candidates without rubric scores; "
                      "see the run log. Not recommending.")
    suspect = implausible_units(result)
    rows = []
    for i, c in enumerate(cands):
        vs = price_vs_lcom(checked_unit(result, i), lcom_price)
        rows.append({"i": i, "c": c, "acc": candidate_score(c)[0], "pct": vs[1] if vs else None})
    qualifying = [r for r in rows if r["acc"] is not None and r["c"].same_product_form
                  and r["acc"] >= MIN_RECOMMEND_ACCURACY and r["pct"] is not None and r["pct"] >= MIN_MARGIN_PCT]
    if not qualifying:
        return None, no_qualifier_reason(rows, lcom_price, suspect, unscored)

    def pick(rs):  # most accurate; ties within TIE_BAND_POINTS broken by price confidence, named maker, price
        top = max(r["acc"] for r in rs)
        contenders = [r for r in rs if r["acc"] >= top - TIE_BAND_POINTS]
        return min(contenders, key=lambda r: (CONFIDENCE_RANK.get(r["c"].unit_price_confidence, 3),
                                              not is_named(r["c"]), unit_price(r["c"]))), contenders

    # A candidate with no URL can't be recommended: nobody can open it to check or order.
    with_url = [r for r in qualifying if (r["c"].url or "").strip()]
    if not with_url:
        nb, _ = pick(qualifying)
        return None, (f"Best candidate (Candidate {nb['i'] + 1}, {nb['c'].manufacturer or 'unnamed supplier'}) qualifies "
                      f"({nb['acc']}% accuracy, {format_unit(unit_price(nb['c']))}/unit) but has no URL - "
                      "locate it manually. Not recommending.")
    best, contenders = pick(with_url)
    c = best["c"]
    unit = unit_price(c)
    tie = (f" Tie-break over {len(contenders) - 1} other candidate(s) within {TIE_BAND_POINTS} accuracy points: "
           "stated price, then named maker, then lowest price." if len(contenders) > 1 else "")
    return best["i"], (
        f"Buy from Candidate {best['i'] + 1} ({c.manufacturer or 'unnamed supplier'}): {best['acc']}% accuracy at "
        f"{format_unit(unit)}/unit ({c.unit_price_confidence}), {format_vs_lcom(price_vs_lcom(unit, lcom_price))} "
        f"than L-Com.{tie}"
    )


def error_text(exc: BaseException) -> str:
    """Short, single-line reason for a product that crashed."""
    msg = " ".join(str(getattr(exc, "message", None) or exc).split())
    text = f"{type(exc).__name__}: {msg}" if msg else type(exc).__name__
    return text if len(text) <= 140 else text[:137] + "..."


class BatchStopped(Exception):
    """A product that never ran because an earlier one hit a billing/auth error."""


# A 400 with this wording is how Anthropic reports an empty balance ("credit balance is too low").
BILLING_WORDS = re.compile(r"credit balance|billing|payment|insufficient (?:funds|credit)|api key|authenticat", re.IGNORECASE)


def is_fatal_api_error(exc: BaseException) -> bool:
    """Billing, credit or authentication failure: every remaining product would fail the same way."""
    if not isinstance(exc, APIStatusError):
        return False
    return exc.status_code in (401, 402, 403) or bool(BILLING_WORDS.search(str(getattr(exc, "message", "") or exc)))


def error_recommendation(reason: str) -> str:
    return f"Error researching this product: {reason} - re-run with --sku"


def product_recommendation(product: dict, result: Optional[SourcingResult]) -> tuple:
    """recommend(), except that a product that errored reads as an error, never as a no-match."""
    if product.get("error"):
        return None, error_recommendation(product["error"])
    return recommend(result, product.get("lcom_price"))


def run_counts(rows: list) -> tuple:
    """(recommended, no recommendation, errored) over (product, result) rows."""
    errored = sum(bool(p.get("error")) for p, _ in rows)
    recommended = sum(not p.get("error") and product_recommendation(p, r)[0] is not None for p, r in rows)
    return recommended, len(rows) - errored - recommended, errored


def ps_quote(text: str) -> str:
    """PowerShell single-quoted literal: safe for SKUs like C&P9M."""
    return "'" + text.replace("'", "''") + "'"


def rerun_command(rows: list, input_path: Optional[str] = None) -> str:
    skus = [p["sku"] for p, _ in rows if p.get("error")]
    if not skus:
        return ""
    parts = ["python sourcing_agent.py"] + (["--input", ps_quote(input_path)] if input_path else [])
    return " ".join(parts + ["--sku"] + [ps_quote(s) for s in skus])


def summary_lines(rows: list, total: int, input_path: Optional[str] = None, stopped: str = "") -> list:
    """Counts (recommended / no recommendation / errored), why a batch stopped early, and the re-run command."""
    rec, norec, err = run_counts(rows)
    pending = total - len(rows)
    lines = [f"Results: **{rec} recommended** | {norec} no recommendation | **{err} errored**"
             + (f" | {pending} still running" if pending > 0 else "")]
    if stopped:
        lines.append(f"**BATCH STOPPED EARLY:** {stopped}. Finished products were saved; the rest were not run.")
    if err:
        lines.append(f"Re-run the errored products: `{rerun_command(rows, input_path)}`")
    return lines


def no_qualifier_reason(rows: list, lcom_price: float, suspect: set, unscored: list) -> str:
    """Says which bar failed: product form, accuracy, price margin, or both."""
    def name(r):
        return f"Candidate {r['i'] + 1}, {r['c'].manufacturer or 'unnamed'}"

    def price_text(r):
        if r["pct"] is None:
            why = "implausible" if r["i"] in suspect else r["c"].unit_price_confidence
            return f"price couldn't be confirmed ({why})"
        return f"{r['pct']:.0f}% cheaper" if r["pct"] >= 0 else f"{-r['pct']:.0f}% more expensive"

    right_form = [r for r in rows if r["acc"] is not None and r["c"].same_product_form]
    if not right_form:
        forms = sorted({r["c"].listing_form for r in rows if r["acc"] is not None and r["c"].listing_form})
        why = "product form: none is the right kind of product" + (f" (found: {', '.join(forms)})" if forms else "")
    elif not any(r["acc"] >= MIN_RECOMMEND_ACCURACY for r in right_form):
        best = max(right_form, key=lambda r: (r["acc"], r["pct"] if r["pct"] is not None else -1e9))
        price_ok = best["pct"] is not None and best["pct"] >= MIN_MARGIN_PCT
        why = (f"accuracy only: best reached {best['acc']}% ({name(best)}), needs {MIN_RECOMMEND_ACCURACY}%, "
               f"at a price that does qualify ({price_text(best)})" if price_ok else
               f"accuracy and price: best reached {best['acc']}% ({name(best)}), needs {MIN_RECOMMEND_ACCURACY}%, "
               f"and its price doesn't qualify either ({price_text(best)})")
    else:
        accurate = [r for r in right_form if r["acc"] >= MIN_RECOMMEND_ACCURACY]
        best = max(accurate, key=lambda r: r["pct"] if r["pct"] is not None else -1e9)
        why = (f"price margin only: best accurate candidate reached {best['acc']}% ({name(best)}) but "
               + ("its " if best["pct"] is None else "is ") + f"{price_text(best)}, needs >= {MIN_MARGIN_PCT}% cheaper")
    extra = "".join(f" Candidate {i + 1}: {IMPLAUSIBLE_NOTE}." for i in sorted(suspect))
    if unscored:
        extra += f" Scoring failed for candidate(s) {', '.join(str(i + 1) for i in unscored)} - see the run log."
    return f"No candidate qualifies - {why} (L-Com unit price ${lcom_price:.2f}). Do not source from any of these.{extra}"


# "Check before ordering" detection, from the listing text. Families of mutually exclusive
# variants: if the target names one and the listing title names another, flag it.
VARIANT_FAMILIES = [
    # (target pattern, conflicting listing pattern, label)
    (r"\bd[be]-?9\b|\b9[\s-]?pins?\b", r"\bvga\b|\b(?:hd|db|de)-?15\b|\b15[\s-]?pins?\b|\bd[bd]-?25\b|\b25[\s-]?pins?\b|"
                                      r"\bd[cb]-?37\b|\b37[\s-]?pins?\b", "a different D-sub pin count"),
    (r"\brj-?45\b|\bcat\s?[5-8]", r"\brj-?1[12]\b|\brj-?9\b|\b[46]p[246]c\b", "a different modular jack (RJ11/RJ12)"),
    (r"\busb\b", r"\bmicro[\s-]?usb\b|\bmicro[\s-]?b\b|\bmini[\s-]?usb\b|\bmini[\s-]?b\b|\btype[\s-]?c\b|\busb[\s-]?c\b",
     "a different USB connector (Micro/Mini/Type-C)"),
    (r"\bhdmi\b", r"\b(?:mini|micro)[\s-]?hdmi\b", "a different HDMI size (Mini/Micro)"),
    (r"\bmulti[\s-]?mode\b", r"\bsingle[\s-]?mode\b|\bSM\b", "single-mode fiber"),
    (r"\bsingle[\s-]?mode\b", r"\bmulti[\s-]?mode\b|\bMM\b", "multimode fiber"),
]
# Several variants sold under one listing / URL.
MULTI_VARIANT = [
    (r"\bsimplex\b.*\bduplex\b|\bduplex\b.*\bsimplex\b", "simplex and duplex"),
    (r"\bmale\s*(?:/|&|and|\s)\s*female\b", "male and female"),
    (r"\bsingle[\s-]?mode\b.*\bmulti[\s-]?mode\b|\bmulti[\s-]?mode\b.*\bsingle[\s-]?mode\b|\bSM\s*/\s*MM\b", "single- and multimode"),
    (r"\b(?:\d{1,2}[\s,/]+){2,}\d{1,2}\s*pins?\b", "several pin counts"),
]
BULK_MOQ_NOTE = 500  # flag only genuinely bulk-only pricing; replace with real order qty per SKU when we have it


def order_quantity(text: str) -> Optional[int]:
    """Minimum order from an MOQ or price string: "500 pieces", "Min. order: 1,000", "1,000 Pieces (MOQ)"."""
    m = re.search(r"([\d,]+)\s*(?:pieces?|pcs|pc|units?)?\s*\(?\s*moq\b|min\.?\s*order:?\s*([\d,]+)|"
                  r"^\s*([\d,]+)\s*(?:pieces?|pcs|pc|units?)?\s*$", text or "", re.IGNORECASE)
    digits = next((g for g in m.groups() if g), "").replace(",", "") if m else ""
    return int(digits) if digits.isdigit() else None


# The model's caveat line mixed useful notes with MOQ/tier remarks on nearly every pick, and
# on 2026-09-29 it put HDFF's "MOQ 500" and TDG1026KS-C6's "MOQ 1000" ONLY there (its moq field
# said "Not stated"). So MOQ/tier sentences are taken out of the caveat, but any quantity in
# them goes through the same BULK_MOQ_NOTE check as the structured MOQ instead of being lost.
MOQ_TIER_SENTENCE = re.compile(r"\bmoq\b|\bmin(?:imum)?\.?\s*order|\btier(?:ed)?\b|\b(?:lower|higher|larger)\s+volumes?\b|"
                               r"\bat volume\b|\bbulk\b", re.IGNORECASE)
MOQ_IN_TEXT = re.compile(r"(?:\bmoq\b|\bmin(?:imum)?\.?\s*order(?:\s*quantity)?)\s*(?:is|of|:)?\s*([\d,]+)", re.IGNORECASE)


def split_caveats(caveats: str) -> tuple:
    """(caveat text without MOQ/tier sentences, largest MOQ those sentences stated or None)."""
    kept, moqs = [], []
    for part in re.split(r";\s*|\.\s+", caveats or ""):
        part = part.strip(" .")
        if not part:
            continue
        if MOQ_TIER_SENTENCE.search(part):
            moqs += [int(q.replace(",", "")) for q in MOQ_IN_TEXT.findall(part) if q.replace(",", "").isdigit()]
        else:
            kept.append(part)
    return "; ".join(kept), max(moqs, default=None)


def ordering_notes(candidate: Candidate, product: dict) -> list:
    """What a buyer should check before ordering this candidate, from its listing text plus
    the model's own caveat. Empty when nothing is worth flagging."""
    c, target = candidate, product.get("description", "")
    title = c.listing_title or ""
    notes = []
    for target_pat, other_pat, label in VARIANT_FAMILIES:
        if re.search(target_pat, target, re.IGNORECASE):
            hit = re.search(other_pat, title, re.IGNORECASE)
            if hit and not re.search(other_pat, target, re.IGNORECASE):
                notes.append(f"title names {label} (\"{hit[0]}\") - confirm it is the target's variant")
    models = re.findall(r"\b([A-Z]{2,})-(\w*\d\w*)\b", title)
    by_prefix = {}
    for prefix, suffix in models:
        by_prefix.setdefault(prefix, set()).add(suffix)
    for prefix, suffixes in by_prefix.items():
        if len(suffixes) > 1:
            notes.append(f"one listing covers several models ({', '.join(f'{prefix}-{s}' for s in sorted(suffixes))}) "
                         "- confirm which one ships")
    for pat, label in MULTI_VARIANT:
        if re.search(pat, title, re.IGNORECASE):
            notes.append(f"listing offers {label} variants - pick the right option when ordering")
    target_contacts = contact_types(target)
    if target_contacts:
        found = contact_types(candidate_contact_text(c))
        if not found:
            notes.append(f"contact type isn't stated on the listing (target is {'/'.join(sorted(target_contacts))}) "
                         "- confirm before ordering")
        elif len(found) > 1:
            notes.append(f"listing mentions several contact types ({'/'.join(sorted(found))}; target is "
                         f"{'/'.join(sorted(target_contacts))}) - check before ordering")
    caveats, caveat_moq = split_caveats(c.listing_caveats)
    moq = max((q for q in (order_quantity(c.moq or ""), order_quantity(c.price or ""), caveat_moq) if q), default=None)
    if moq and moq >= BULK_MOQ_NOTE:
        notes.append(f"price needs an order of {moq:,}+ pieces (MOQ)")
    notes += contradictions(c)
    spec = SPEC_CABLE.search(f"{title} {c.listing_caveats or ''} {c.spec_lines or ''} {c.unit_price_note or ''}")
    if spec and re.search(r"coupler|adapter|adaptor", target, re.IGNORECASE):
        notes.append(f"listing text mentions \"{spec[0]}\" - may be a short cable rather than a plain adapter, verify the product page")
    tier = price_tier_note(c)
    if tier:
        notes.append(tier)
    if c.unit_price_confidence.startswith("promo"):
        notes.append("promo price - regular price not shown on the listing")
    if not pack_size_supported(c):
        notes.append(f"listing's quantity {c.quantity_covered} wasn't tied to this price - priced per piece")
    if caveats:
        notes.append(f"model: {caveats}")
    return notes


def ordering_note(result: Optional[SourcingResult], rec_idx: Optional[int], product: dict) -> str:
    """The recommended candidate's notes as one line, or "" (no recommendation / nothing to flag)."""
    if result is None or rec_idx is None:
        return ""
    return "; ".join(ordering_notes(result.candidates[rec_idx], product))


def score_summary(result: SourcingResult) -> str:
    """One line per product, templated from the same scores as the table so the two can't drift."""
    parts = []
    for i, c in enumerate(result.candidates, start=1):
        acc, scores = candidate_score(c)
        if acc is None:
            parts.append(f"{i}. {c.manufacturer or 'unnamed'}: scoring failed")
            continue
        detail = "/".join(str(max(0, min(20, scores.get(a, 0)))) for a in ATTRIBUTES)
        form = "" if c.same_product_form else f", wrong form ({c.listing_form})"
        parts.append(f"{i}. {c.manufacturer or 'unnamed'}: {acc}% ({detail}{form})")
    return "Scores (type/spec/shielding/mount/gender): " + "; ".join(parts)


def print_product_result(product: dict, result: Optional[SourcingResult], calls_used: int):
    log("-" * 72)
    log(f"{product['sku']} - \"{product['description']}\"  ({calls_used} tool calls used)")
    log()

    if result is None:
        log("  Could not get a structured result for this product (empty, refused or garbled response).")
        return

    if result.no_match or not result.candidates:
        reason = result.no_match_reason or "no plausible candidate found"
        log(f"  No match found - {reason}")
        return

    for i, candidate in enumerate(result.candidates, start=1):
        computed, scores = candidate_score(candidate)
        vs = price_vs_lcom(checked_unit(result, i - 1), product.get("lcom_price"))
        manufacturer = candidate.manufacturer or "Not stated on listing"
        source = candidate.distributor_site or "Unknown site"
        price = candidate.price or "Not stated"

        log(f"  #{i}  {manufacturer}  |  via {source}  |  Tier: {match_tier(computed)}")
        log(f"      Listing:  {candidate.listing_title or '?'}")
        log(f"      Price:    {price}   MOQ: {candidate.moq or 'Not stated'}")
        log(f"      Unit:     {format_unit(unit_price(candidate))} ({confidence_label(result, i - 1)})  "
              f"vs L-Com: {format_vs_lcom(vs)}")
        log(f"      Accuracy: {computed}%" + ("" if candidate.same_product_form else f"  (wrong product form: {candidate.listing_form})"))
        log(f"      URL:      {candidate.url or 'Not stated'}")
        if scores:
            line = " | ".join(
                f"{ATTRIBUTE_LABELS[attr]} {max(0, min(20, scores.get(attr, 0)))}" for attr in ATTRIBUTES
            )
            log(f"      Why:      {line}  ->  {candidate_comment(candidate)}")
        log()

    log(f"  {score_summary(result)}")
    rec_idx, rec_reason = product_recommendation(product, result)
    log(f"  Recommendation: {rec_reason}")
    note = ordering_note(result, rec_idx, product)
    if note:
        log(f"  Check before ordering: {note}")


def format_result_markdown(
    product: dict, result: Optional[SourcingResult], calls_used: int, tin: int, tout: int, cost: float
) -> str:
    lines = [f"## {product['sku']} - {product['description']}", ""]

    if product.get("error"):
        lines.append(f"**Recommendation:** {error_recommendation(product['error'])}")
    elif result is None:
        lines.append("**No structured result** (empty, refused or garbled response).")
    elif result.no_match or not result.candidates:
        reason = result.no_match_reason or "no plausible candidate found"
        lines.append(f"**No match found** - {reason}")
    else:
        lcom = product.get("lcom_price")
        lcom_text = f"${lcom:.2f}" if lcom else "n/a"
        rec_idx, rec_reason = product_recommendation(product, result)
        note = ordering_note(result, rec_idx, product)
        lines += [f"**Recommendation:** {rec_reason}", ""]
        if note:
            lines += [f"**Check before ordering:** {note}", ""]
        lines += [
            f"| Candidate | Accuracy | Unit price | Unit price confidence | vs. L-Com ({lcom_text}) | Recommended |",
            "|---|---|---|---|---|---|",
        ]
        for i, c in enumerate(result.candidates):
            accuracy, _ = candidate_score(c)
            cells = [
                f"{i + 1}. {c.manufacturer or 'Not stated'}",
                f"{accuracy}%",
                format_unit(unit_price(c)),
                confidence_label(result, i),
                format_vs_lcom(price_vs_lcom(checked_unit(result, i), lcom)),
                "YES" if i == rec_idx else "no",
            ]
            if i == rec_idx:
                cells = [f"**{x}**" for x in cells]
            lines.append("| " + " | ".join(cells) + " |")
        lines += [f"| L-Com (benchmark) | - | {lcom_text} | - | - | - |", ""]

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
                f"| Product form | {candidate.listing_form or '?'}" + ("" if candidate.same_product_form else " - **WRONG FORM, excluded**") + " |",
                *([f"| **Check before ordering** | {note} |"] if note and i - 1 == rec_idx else []),
                f"| Accuracy | {computed}% |",
                f"| Listed price | {candidate.price or 'Not stated'} |",
                f"| Qty the price covers | {candidate.quantity_covered or '?'} |",
                f"| Unit price | {format_unit(unit_price(candidate))} ({confidence_label(result, i - 1)}) |",
                f"| Unit price note | {candidate.unit_price_note or '-'} |",
                f"| vs. L-Com price | {format_vs_lcom(price_vs_lcom(checked_unit(result, i - 1), lcom))} |",
                f"| MOQ | {candidate.moq or 'Not stated'} |",
                f"| Email | {candidate.email or 'Not stated'} |",
                f"| URL | {candidate.url or 'Not stated'} |",
                f"| Spec lines seen | {candidate.spec_lines or 'none (search-result title only)'} |",
                "",
            ]
            if scores:
                lines.append("| Rubric attribute | Score (of 20) |")
                lines.append("|---|---|")
                for attr in ATTRIBUTES:
                    lines.append(f"| {ATTRIBUTE_LABELS[attr]} | {max(0, min(20, scores.get(attr, 0)))} |")
                lines += ["", candidate_comment(candidate), ""]
        lines.append(f"*{score_summary(result)}*")

    lines += ["", f"*Tool calls used: {calls_used} | Tokens: in={tin} out={tout} | Cost: ~${cost:.4f}*", ""]
    return "\n".join(lines)


def ambiguous_price_flags(rows: list) -> list:
    """Markdown bullets for every candidate whose unit price couldn't be determined or looks implausible."""
    flags = []
    for product, result in rows:
        suspect = implausible_units(result)
        for i, c in enumerate(result.candidates if result else [], start=1):
            if not c.same_product_form:
                continue  # excluded anyway; its price doesn't matter
            if i - 1 in suspect:
                reason = f"{format_unit(unit_price(c))}/unit: {IMPLAUSIBLE_NOTE} ({c.unit_price_note or 'no note'})"
            elif unit_price(c) is None:
                reason = c.unit_price_note or c.unit_price_confidence or "no unit price given"
            elif not pack_size_supported(c):
                reason = (f"quantity {c.quantity_covered} ignored - not attached to this price in the listing "
                          f"(MOQ?), priced per piece at {format_unit(unit_price(c))}")
            elif c.unit_price_confidence.startswith("promo"):
                reason = f"promo price, regular price not shown ({c.unit_price_note or 'no note'})"
            else:
                continue
            flags.append(
                f"- **{product['sku']}** candidate {i} ({c.manufacturer or 'unnamed'}): "
                f"listed \"{c.price or 'no price'}\" - {reason}"
            )
    return flags


LCOM_PRICES_CSV = os.path.join(os.path.dirname(os.path.abspath(__file__)), "lcom_prices.csv")


def load_lcom_prices(path: str = LCOM_PRICES_CSV) -> dict:
    """lcom_prices.csv (sku, pack_size, pack_price, source, date_checked) -> {SKU upper: row}."""
    try:
        with open(path, newline="", encoding="utf-8") as f:
            return {r["sku"].strip().upper(): r for r in csv.DictReader(f)}
    except FileNotFoundError:
        return {}


def apply_lcom_prices(products: list, path: str = LCOM_PRICES_CSV) -> None:
    """One source for L-Com reference prices: the CSV wins; a SKU missing from it keeps the price it
    came with (input sheet / built-in list) and gets a warning. unit price = pack_price / pack_size."""
    table = load_lcom_prices(path)
    for p in products:
        row = table.get(p["sku"].strip().upper())
        if row:
            p["lcom_price"] = round(float(row["pack_price"]) / int(row["pack_size"] or 1), 2)  # to the cent: $19.99/10 = $2.00
            p["lcom_source"], p["lcom_date"] = row["source"], row["date_checked"]
        else:
            log(f"WARNING: {p['sku']} is not in {os.path.basename(path)} - using the price it came with, unverified.")
            p["lcom_source"], p["lcom_date"] = "not in lcom_prices.csv (unverified fallback)", ""


def check_input_files(files: Optional[dict] = None) -> list:
    """Problems with the two human-maintained CSVs: missing or no data rows. A fresh checkout
    (both are .gitignored) would otherwise silently lose the L-Com prices and reviewer exclusions."""
    files = files or {
        "lcom_prices.csv": ("every L-Com price falls back to the input sheet / built-in value and is marked unverified",
                            LCOM_PRICES_CSV),
        "reviewer_exclusions.csv": ("known-bad listings will NOT be excluded", REVIEWER_EXCLUSIONS_CSV),
    }
    problems = []
    for name, (consequence, path) in files.items():
        try:
            with open(path, newline="", encoding="utf-8") as f:
                has_rows = any(any((v or "").strip() for v in row.values()) for row in csv.DictReader(f))
        except FileNotFoundError:
            problems.append(f"{name} is MISSING - {consequence}.")
            continue
        if not has_rows:
            problems.append(f"{name} is EMPTY - {consequence}.")
    return problems


def print_input_file_warnings(problems: list) -> None:
    if problems:
        bar = "!" * 72
        print("\n".join([bar, "!!! WARNING: required input file problem(s) - results will be WRONG !!!", *problems, bar]))


def price_is_unverified(product: dict) -> bool:
    source = (product.get("lcom_source") or "").lower()
    return not source or "unverified" in source or "unconfirmed" in source


def lcom_price_lines(products: list) -> list:
    """Report-header block: each SKU's L-Com price with its source and date, and the unverified ones."""
    lines = ["## L-Com reference prices", "", "| SKU | L-Com unit price | Source | Date checked |", "|---|---|---|---|"]
    for p in products:
        price = f"${p['lcom_price']:.2f}" if p.get("lcom_price") else "n/a"
        lines.append(f"| {p['sku']} | {price} | {p.get('lcom_source') or 'n/a'} | {p.get('lcom_date') or '-'} |")
    unverified = [p["sku"] for p in products if price_is_unverified(p)]
    lines += ["", "**Still on unverified L-Com prices (do not read the margin as confirmed):** "
              + (", ".join(unverified) or "none"), "", "---", ""]
    return lines


def write_report(
    sections: list, grand_in: int, grand_out: int, grand_cost: float, num_products: int, path: str,
    flags: list, products: Optional[list] = None, summary: Optional[list] = None,
) -> None:
    header = [
        "# Competitive Sourcing Research Report",
        "",
        f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M')}",
        f"Model: {MODEL} | Sites: {', '.join(SOURCING_SITES)} | Products: {num_products}",
        *(["", *summary, ""] if summary else []),
        f"Recommendation rule: >= {MIN_RECOMMEND_ACCURACY}% accuracy and a confirmed unit price "
        f">= {MIN_MARGIN_PCT}% below L-Com's price, right product form; most accurate wins, and within "
        f"{TIE_BAND_POINTS} points of it: stated price, then a named maker, then the lowest price.",
        "",
        "## Unit prices that could not be determined confidently or look implausible",
        "",
        *(flags or ["None."]),
        "",
        "---",
        "",
        *(lcom_price_lines(products) if products else []),
    ]
    footer = [
        "---",
        "",
        f"**Total tokens:** in={grand_in} out={grand_out} | **Estimated total cost:** ~${grand_cost:.4f}",
    ]
    content = "\n".join(header + sections + footer) + "\n"
    with open(path, "w", encoding="utf-8") as f:
        f.write(content)


PRODUCT_XLSX_FIELDS = ["SKU", "Product", "Keyword", "L-Com Unit Price", "L-Com Price Source", "L-Com Price Date",
                       "Recommendation"]
RECOMMENDED_XLSX_FIELDS = [
    "Recommended Manufacturer", "Recommended Email", "Recommended URL", "Recommended Unit Price",
]
ORDERING_NOTE_FIELD = "Ordering Note"  # right after the green Recommended block; amber when filled
AMBER = PatternFill("solid", fgColor="FFE699")
CANDIDATE_XLSX_FIELDS = [
    "Name", "Accuracy", "Listed Price", "Unit Price", "Unit Price Confidence", "vs. L-Com Price",
    "MOQ", "URL", "Email", "Match Tier", "Comment",
]
GREEN = PatternFill("solid", fgColor="C6EFCE")


def write_excel_report(rows: list, path: str) -> None:
    """rows: list of (product: dict, result: Optional[SourcingResult]) tuples.
    Sheet "Results": one row per product. Sheet "Comparison": candidates vs L-Com side by side.
    The green fill marks who to buy from: the Recommended columns on Results, the row on
    Comparison. Nothing is green when no candidate qualifies."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Results"
    comp = wb.create_sheet("Comparison")

    header = PRODUCT_XLSX_FIELDS + RECOMMENDED_XLSX_FIELDS + [ORDERING_NOTE_FIELD]
    for n in range(1, MAX_CANDIDATES + 1):
        header += [f"Manufacturer {n} {field}" for field in CANDIDATE_XLSX_FIELDS]
    ws.append(header)
    comp.append(["SKU", "Candidate", "Manufacturer", "Accuracy", "Unit Price", "Unit Price Confidence",
                 "vs. L-Com Price", "Recommended"])

    for product, result in rows:
        lcom = product.get("lcom_price")
        rec_idx, rec_reason = product_recommendation(product, result)
        row = [product["sku"], product.get("product_name", ""), product["description"], lcom,
               product.get("lcom_source", ""), product.get("lcom_date", ""), rec_reason]
        candidates = result.candidates if result else []
        if rec_idx is None:
            row += [""] * len(RECOMMENDED_XLSX_FIELDS)
        else:
            rec = candidates[rec_idx]
            row += [rec.manufacturer or "", rec.email or "", rec.url or "", round(unit_price(rec), 4)]
        note = ordering_note(result, rec_idx, product)
        row.append(note)
        for i in range(MAX_CANDIDATES):
            if i < len(candidates):
                c = candidates[i]
                computed, scores = candidate_score(c)
                unit = unit_price(c)
                unit_cell = round(unit, 4) if unit is not None else "UNDETERMINED"
                confidence = confidence_label(result, i)
                vs = format_vs_lcom(price_vs_lcom(checked_unit(result, i), lcom))
                row += [
                    c.manufacturer or "",
                    computed,
                    c.price or "",
                    unit_cell,
                    confidence,
                    vs,
                    c.moq or "",
                    c.url or "",
                    c.email or "",
                    match_tier(computed),
                    candidate_comment(c),
                ]
                comp.append([product["sku"], i + 1, c.manufacturer or "", computed, unit_cell,
                             confidence, vs, "YES" if i == rec_idx else "no"])
                if i == rec_idx:
                    for cell in comp[comp.max_row]:
                        cell.fill = GREEN
            else:
                row += [""] * len(CANDIDATE_XLSX_FIELDS)
        if product.get("error"):  # never green: the Recommended cell carries the error text
            comp.append([product["sku"], "ERROR", "", None, None, "", "", rec_reason])
        comp.append([product["sku"], "L-Com", "L-Com (benchmark)", None, lcom])
        comp.append([])
        ws.append(row)
        if rec_idx is not None:
            start = len(PRODUCT_XLSX_FIELDS)
            for cell in ws[ws.max_row][start : start + len(RECOMMENDED_XLSX_FIELDS)]:
                cell.fill = GREEN
        if note:
            ws[ws.max_row][len(PRODUCT_XLSX_FIELDS) + len(RECOMMENDED_XLSX_FIELDS)].fill = AMBER

    for sheet in (ws, comp):
        for cell in sheet[1]:
            cell.font = Font(bold=True)
        for col_cells in sheet.columns:
            width = max((len(str(cell.value)) for cell in col_cells if cell.value), default=10)
            sheet.column_dimensions[col_cells[0].column_letter].width = min(max(width + 2, 10), 60)

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
    lcom_col = headers.get("lcom sale price")

    products = []
    for row in ws.iter_rows(min_row=header_row + 1):
        sku = row[sku_col - 1].value
        if not sku:
            continue
        keyword = row[keyword_col - 1].value if keyword_col else None
        product_name = row[product_col - 1].value if product_col else None
        lcom_price = row[lcom_col - 1].value if lcom_col else None
        # L-Com prices packs as one SKU ("..., Package/10") - compare candidates per unit.
        pack = re.search(r"package\s*/\s*(\d+)", f"{keyword} {product_name}", re.IGNORECASE)
        if pack and isinstance(lcom_price, (int, float)):
            lcom_price /= int(pack.group(1))
        products.append({
            "sku": str(sku).strip(),
            "description": str(keyword or product_name or "").strip(),
            "product_name": str(product_name or "").strip(),
            "lcom_price": float(lcom_price) if isinstance(lcom_price, (int, float)) else None,
        })
    return products


async def connect_tools(nimble_api_key: str):
    http_client = httpx2.AsyncClient(headers={"Authorization": f"Bearer {nimble_api_key}"})
    return http_client


def read_lcom_catalog(source: Optional[str] = None) -> list:
    """Future input path: L-Com's own catalog -> product dicts shaped like PRODUCTS /
    read_products(): {"sku", "description", "product_name", "lcom_price"} (per-unit USD).

    TODO: implement. Fetch or parse L-Com's catalog (their category/listing pages, or a
    bulk export if one exists) and extract SKU, name, description and price per product.
    Watch for pack-priced SKUs ("Package/N") - divide to a per-unit price like read_products().
    Stubbed for now: no scraper exists yet (l-com.com renders search results with
    JavaScript, so a plain fetch doesn't see them) and the product list is still curated
    by hand. Returns [] so callers fall back to PRODUCTS.
    """
    return []


def pick_skus(products: list, skus: list) -> tuple:
    """(products whose SKU is in `skus`, in the order asked, first match wins; SKUs not found).
    Case-insensitive - for re-running specific products, e.g. after a transient API error."""
    by_sku = {}
    for p in products:
        by_sku.setdefault(p["sku"].strip().upper(), p)
    picked, missing = [], []
    for s in dict.fromkeys(s.strip().upper() for s in skus):
        if s in by_sku:
            picked.append(by_sku[s])
        else:
            missing.append(s)
    return picked, missing


def percentile(values: list, q: float) -> float:
    """Nearest-rank percentile; 0 for an empty list."""
    if not values:
        return 0.0
    ordered = sorted(values)
    return ordered[min(len(ordered) - 1, max(0, math.ceil(q * len(ordered)) - 1))]


def busy_time(calls: list) -> float:
    """Seconds during which at least one of these calls was in flight. Summing durations
    counted the prefetch's three parallel searches three times (2026-09-29 run showed an
    impossible "Anthropic 2%")."""
    total, end = 0.0, None
    for c in sorted(calls, key=lambda c: c["start"]):
        start, stop = c["start"], c["start"] + c["secs"]
        if end is None or start >= end:
            total += stop - start
            end = stop
        elif stop > end:
            total += stop - end
            end = stop
    return total


def peak_concurrency(calls: list) -> int:
    """Most Nimble calls in flight at the same moment."""
    events = sorted([(c["start"], 1) for c in calls] + [(c["start"] + c["secs"], -1) for c in calls],
                    key=lambda e: (e[0], e[1]))  # ends before starts at the same instant
    peak = now = 0
    for _, delta in events:
        now += delta
        peak = max(peak, now)
    return peak


def stop_reason(result: Optional[SourcingResult], calls_used: int, failed_calls: int) -> str:
    """Why a product's search ended - so a lower budget's effect can be told apart: found enough
    vs ran out of calls (and how many of those failed) vs the model gave up with budget left."""
    budget = f"{calls_used}/{MAX_TOOL_CALLS_PER_PRODUCT} calls, {failed_calls} failed"
    if result is None:
        return f"no usable result ({budget})"
    n = len(result.candidates)
    if n >= MAX_CANDIDATES:
        return f"found {MAX_CANDIDATES} candidates ({budget})"
    if calls_used >= MAX_TOOL_CALLS_PER_PRODUCT:
        return f"budget used up with {n} candidates ({budget})"
    return f"model stopped with {n} candidates, budget left ({budget})"


def timing_summary(calls: list, products: list, wall_secs: float) -> str:
    """Where the wall-clock time went: Nimble vs Anthropic (= product time minus its Nimble
    time), timeouts, and Nimble latency per tool/site for picking a timeout."""
    by_sku = {}
    for c in calls:
        by_sku.setdefault(c["sku"], []).append(c)
    nimble = sum(busy_time(group) for group in by_sku.values())  # waiting time, overlaps counted once
    product_total = sum(p["secs"] for p in products)
    anthropic = product_total - nimble
    timeouts = [c for c in calls if c["outcome"] == "timeout"]
    timeout_secs = sum(c["secs"] for c in timeouts)
    pct = lambda x: f"{100 * x / product_total:.0f}%" if product_total else "-"
    lines = [
        f"Wall clock: {wall_secs / 60:.1f} min for {len(products)} products "
        f"({wall_secs / max(1, len(products)):.0f}s per product)",
        f"  Waiting on Nimble:     {nimble / 60:6.1f} min  {pct(nimble)}  ({len(calls)} calls)",
        f"    of which timeouts:   {timeout_secs / 60:6.1f} min  {pct(timeout_secs)}  ({len(timeouts)} calls)",
        f"  Anthropic + overhead:  {anthropic / 60:6.1f} min  {pct(anthropic)}",
        f"  Peak simultaneous Nimble calls: {peak_concurrency(calls)}",
        "Nimble latency (s)      calls  median   p90    max  timeouts",
    ]
    groups = {}
    for c in calls:
        groups.setdefault(f"{c['tool'].replace('nimble_', '')} {c['site']}", []).append(c)
    for name, group in sorted(groups.items()):
        secs = [c["secs"] for c in group]
        lines.append(f"  {name:22} {len(secs):5}  {percentile(secs, .5):6.1f} {percentile(secs, .9):6.1f} {max(secs):6.1f}"
                     f"  {sum(c['outcome'] == 'timeout' for c in group):8}")
    reasons = collections.Counter(p.get("stop", "?").split(" (")[0] for p in products)
    lines.append("Why products stopped: " + "; ".join(f"{r}: {n}" for r, n in reasons.most_common()))
    ok = [c["secs"] for c in calls if c["outcome"] == "ok"]
    buckets = [(0, 15), (15, 30), (30, 45), (45, 60), (60, 90), (90, 10**9)]
    lines.append("Successful call latency: " + "  ".join(
        f"{lo}-{hi if hi < 10**9 else ''}s: {sum(lo <= x < hi for x in ok)}" for lo, hi in buckets))
    for c in timeouts:
        lines.append(f"  TIMEOUT {c['sku']} {c['tool']} {c['url']}")
    return "\n".join(lines)


def parse_args():
    parser = argparse.ArgumentParser(description="Competitive sourcing research POC")
    parser.add_argument(
        "--input", help="Path to an .xlsx with SKU/Product/Keyword columns. Defaults to the built-in products."
    )
    parser.add_argument(
        "--add-builtin", action="store_true",
        help="Also run the built-in products after the --input ones (applied after --limit).",
    )
    parser.add_argument(
        "--sku", nargs="+", metavar="SKU",
        help="Only run these SKUs (case-insensitive), looked up in --input and the built-in products. "
             "Ignores --limit. Quote SKUs with special characters, e.g. \"C&P9M\".",
    )
    parser.add_argument("--output", help="Path for the .xlsx results file. Defaults to a timestamped filename.")
    parser.add_argument("--limit", type=int, help="Only process the first N products (for a quick/cheap test run).")
    parser.add_argument(
        "--inspect-tools", action="store_true", help="Print Nimble tool schemas and exit (no Anthropic tokens spent)."
    )
    return parser.parse_args()


async def run_batch(products: list, research, report_path: str, excel_path: str, totals: dict,
                    input_path: Optional[str] = None) -> tuple:
    """Research every product (MAX_CONCURRENT_PRODUCTS at a time) and save the report and Excel after each
    one, so a crash or stop keeps everything finished so far. `research(product)` returns
    (result, tokens_in, tokens_out, calls_used). A product that raises, or comes back with no usable
    result, is recorded as errored (product["error"]) - never as a no-match. A billing, credit or
    authentication error stops the batch: products not yet started are recorded as errored without being
    run, instead of failing one by one. Returns (done, why_stopped); done[i] = (markdown, (product, result))."""
    # One slot per product in input order: (report section, excel row) once finished.
    done = [None] * len(products)
    slots = asyncio.Semaphore(MAX_CONCURRENT_PRODUCTS)
    save_lock = asyncio.Lock()
    stop = {"reason": ""}

    async def run_one(idx: int, product: dict):
        CURRENT_SKU.set(product["sku"])
        async with slots:
            log(f"Product {idx + 1} of {len(products)} ...")
            started = time.monotonic()
            try:
                if stop["reason"]:
                    raise BatchStopped(f"not run - batch stopped early ({stop['reason']})")
                result, tin, tout, calls_used = await research(product)
                if result is None:
                    product["error"] = "no structured result (empty, refused or garbled response)"
                print_product_result(product, result, calls_used)
                cost = (tin / 1_000_000 * HAIKU_INPUT_PER_MTOK) + (tout / 1_000_000 * HAIKU_OUTPUT_PER_MTOK)
                totals["in"] += tin
                totals["out"] += tout
                totals["cost"] += cost
                log(f"  Tokens:     in={tin} out={tout}  (~${cost:.4f})")
                secs = time.monotonic() - started
                nimble = busy_time([c for c in CALL_LOG if c["sku"] == product["sku"]])
                failed = sum(c["outcome"] != "ok" for c in CALL_LOG if c["sku"] == product["sku"])
                stopped = stop_reason(result, calls_used, failed)
                PRODUCT_LOG.append({"sku": product["sku"], "secs": secs, "calls": calls_used, "stop": stopped})
                log(f"  Time:       {secs:.0f}s  (Nimble {nimble:.0f}s, Anthropic+other {secs - nimble:.0f}s, "
                    f"{calls_used} tool calls)")
                log(f"  Stopped:    {stopped}")
                done[idx] = (format_result_markdown(product, result, calls_used, tin, tout, cost), (product, result))
            except Exception as exc:  # noqa: BLE001 - POC: keep going on any per-product failure
                reason = str(exc) if isinstance(exc, BatchStopped) else error_text(exc)
                product["error"] = reason
                if is_fatal_api_error(exc) and not stop["reason"]:
                    stop["reason"] = reason
                    log("!" * 72)
                    log(f"!!! BILLING / CREDIT / AUTH ERROR - stopping the batch: {reason}")
                    log("!!! Products already finished are saved; the rest will not be run.")
                    log("!" * 72)
                log(f"  Error researching this product: {reason}")
                done[idx] = (format_result_markdown(product, None, 0, 0, 0, 0.0), (product, None))
        # The lock keeps two finishing products from interleaving writes.
        async with save_lock:
            finished = [d for d in done if d]
            try:
                write_report([s for s, _ in finished], totals["in"], totals["out"], totals["cost"],
                             len(products), report_path, ambiguous_price_flags([r for _, r in finished]),
                             [p for p, _ in (d[1] for d in finished)],
                             summary_lines([d[1] for d in finished], len(products), input_path, stop["reason"]))
                write_excel_report([r for _, r in finished], excel_path)
            except OSError as exc:  # e.g. the .xlsx is open in Excel - retry on the next product
                log(f"  Could not save results yet ({exc}) - will retry after the next product.")

    await asyncio.gather(*(run_one(i, p) for i, p in enumerate(products)))
    return done, stop["reason"]


async def main_async():
    args = parse_args()
    load_dotenv()  # .env next to where you run it; real env vars still win
    print_input_file_warnings(check_input_files())

    anthropic_key = os.environ.get("ANTHROPIC_API_KEY")
    nimble_key = os.environ.get("NIMBLE_API_KEY")
    missing = [n for n, v in [("ANTHROPIC_API_KEY", anthropic_key), ("NIMBLE_API_KEY", nimble_key)] if not v]
    if missing:
        print(f"Missing required environment variable(s): {', '.join(missing)}")
        print("Add them to a .env file in this folder (or set them in the shell), e.g.:")
        print("  ANTHROPIC_API_KEY=sk-ant-...")
        print("  NIMBLE_API_KEY=...")
        sys.exit(1)

    builtin = read_lcom_catalog() or PRODUCTS
    products = read_products(args.input) if args.input else builtin
    if args.sku:
        products, missing = pick_skus(products + (builtin if args.input else []), args.sku)
        if missing:
            print(f"SKU(s) not found in {args.input or 'the built-in products'} or the built-in list: {', '.join(missing)}")
            if not products:
                sys.exit(1)
    else:
        if args.limit:
            products = products[: args.limit]
        if args.input and args.add_builtin:
            products = products + builtin
    if not products:
        print(f"No products found in {args.input} - check it has a 'SKU' header column.")
        sys.exit(1)

    apply_lcom_prices(products)
    client = AsyncAnthropic(api_key=anthropic_key, max_retries=ANTHROPIC_MAX_RETRIES)

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

                run_started = time.monotonic()
                totals = {"in": 0, "out": 0, "cost": 0.0}
                timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                report_path = f"sourcing_report_{timestamp}.md"
                excel_path = args.output or f"sourcing_results_{timestamp}.xlsx"
                print(f"Saving results after every product to {report_path} / {excel_path}")
                print(f"Running up to {MAX_CONCURRENT_PRODUCTS} products at once")

                async def research(product):
                    return await research_product(client, session, tools, product)

                done, stopped_why = await run_batch(products, research, report_path, excel_path, totals, args.input)

                print("-" * 72)
                print(timing_summary(CALL_LOG, PRODUCT_LOG, time.monotonic() - run_started))
                print(f"Anthropic retries (429 rate limit / 529 overloaded / 5xx): {dict(ANTHROPIC_RETRIES) or 'none'}")
                print(f"Nimble rate-limit retries: {NIMBLE_RATE_LIMITS['count'] or 'none'}")
                print(
                    f"Total tokens: in={totals['in']} out={totals['out']}  |  "
                    f"Estimated total cost: ~${totals['cost']:.4f}"
                )

                rows = [d[1] for d in done if d]
                print("\n".join(line.replace("**", "") for line in summary_lines(rows, len(products), args.input, stopped_why)
                                if not line.startswith("Re-run")))
                if rerun_command(rows, args.input):
                    print("Re-run the errored products (PowerShell):")
                    print("  " + rerun_command(rows, args.input))
                flags = ambiguous_price_flags([d[1] for d in done if d])
                if flags:
                    print("Unit prices that could not be determined confidently:")
                    print("\n".join(flags))
                print(f"Report written to: {os.path.abspath(report_path)}")
                print(f"Excel results written to: {os.path.abspath(excel_path)}")


def main():
    asyncio.run(main_async())


if __name__ == "__main__":
    main()
