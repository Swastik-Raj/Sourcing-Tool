"""Offline check of unit-price / L-Com margin / recommendation logic. Run: python test_pricing.py"""
import os
import tempfile

import openpyxl

from sourcing_agent import (
    AttributeBreakdown, Candidate, SourcingResult, recommend, unit_price, write_excel_report,
)

FULL = AttributeBreakdown(product_type=20, category_spec=20, shielding_material=20, mount_form=20, gender_pins=20)
NINETY = AttributeBreakdown(product_type=20, category_spec=20, shielding_material=20, mount_form=10, gender_pins=20)


def cand(name, total, qty, conf, breakdown=FULL):
    return Candidate(manufacturer=name, price_total=total, quantity_covered=qty,
                     unit_price_confidence=conf, attribute_breakdown=breakdown)


# Bundle price is divided down; ambiguous never yields a unit price.
assert unit_price(cand("A", 19.75, 10, "inferred")) == 1.975
assert unit_price(cand("A", 19.75, 10, "ambiguous")) is None
assert unit_price(cand("A", 19.75, None, "stated")) is None

# 100% match barely under L-Com loses to 90% match with a big gap.
r = SourcingResult(product="x", no_match=False, candidates=[
    cand("Close", 21.0, 1, "stated"),
    cand("Cheap", 5.0, 1, "stated", NINETY),
    cand("Mystery", 1.0, 1, "ambiguous"),
])
idx, reason = recommend(r, 22.59)
assert idx == 1, reason

# No margin anywhere -> explicit "do not source", no default pick.
r = SourcingResult(product="x", no_match=False, candidates=[cand("Close", 21.0, 1, "stated")])
idx, reason = recommend(r, 22.59)
assert idx is None and "do not source" in reason

# No L-Com price -> no recommendation.
assert recommend(r, None)[0] is None

# Excel: recommended block is green, others aren't.
r = SourcingResult(product="x", no_match=False, candidates=[cand("Close", 21.0, 1, "stated"), cand("Cheap", 5.0, 1, "stated")])
path = os.path.join(tempfile.mkdtemp(), "t.xlsx")
write_excel_report([({"sku": "S", "description": "d", "lcom_price": 22.59}, r)], path)
wb = openpyxl.load_workbook(path)
ws = wb["Results"]
names = {c.value: c.fill.fgColor.rgb for c in ws[2] if c.value in ("Close", "Cheap")}
assert names["Cheap"].endswith("C6EFCE") and not names["Close"].endswith("C6EFCE"), names
assert wb["Comparison"].cell(3, 8).value == "YES"
print("ok")
