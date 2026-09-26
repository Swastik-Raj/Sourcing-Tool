"""Offline check of unit-price / L-Com margin / recommendation logic. Run: python test_pricing.py"""
import os
import tempfile

import openpyxl

from sourcing_agent import (
    AttributeBreakdown, Candidate, SourcingResult, dedupe_manufacturers, implausible_units, is_garbled, pick_skus, recommend, strip_page_chrome, unit_price, write_excel_report,
)

FULL = AttributeBreakdown(product_type=20, category_spec=20, shielding_material=20, mount_form=20, gender_pins=20)
NINETY = AttributeBreakdown(product_type=20, category_spec=20, shielding_material=20, mount_form=10, gender_pins=20)


def cand(name, total, qty, conf, breakdown=FULL):
    return Candidate(manufacturer=name, price_total=total or 0, quantity_covered=qty or 0, unit_price_note="",
                     unit_price_confidence=conf, attribute_breakdown=breakdown, email=f"{name}@x", url=f"u/{name}")


def result(*cands):
    return SourcingResult(product="x", no_match=False, candidates=list(cands))


# Bundle price is divided down; ambiguous never yields a unit price.
assert unit_price(cand("A", 19.75, 10, "inferred")) == 1.975
assert unit_price(cand("A", 19.75, 10, "ambiguous")) is None
assert unit_price(cand("A", 19.75, None, "stated")) is None

# 100% match barely under L-Com loses to 90% match with a big gap.
idx, reason = recommend(result(cand("Close", 21.0, 1, "stated"), cand("Cheap", 4.0, 1, "stated", NINETY),
                               cand("Mystery", 1.0, 1, "ambiguous")), 22.59)
assert idx == 1, reason

# 30-79% cheaper no longer qualifies (bar is 80%).
idx, reason = recommend(result(cand("Half", 11.0, 1, "stated")), 22.59)
assert idx is None and ">= 80% cheaper" in reason and "do not source" in reason, reason

# HDFF regression: price/MOQ gave $0.0032 "inferred" next to two $1.09 "stated" -> implausible, not picked.
hdff = result(cand("A", 1.09, 1, "stated"), cand("B", 1.09, 1, "stated"), cand("FARSINCE", 1.58, 500, "inferred"))
assert implausible_units(hdff) == {2}
idx, reason = recommend(hdff, 23.79)
assert idx == 0, reason

# Implausible price can't make a product clear the bar on its own.
idx, reason = recommend(result(cand("Real", 20.0, 1, "stated"), cand("Real2", 21.0, 1, "stated"),
                               cand("Bug", 1.58, 500, "inferred")), 23.79)
assert idx is None and "implausible" in reason, reason

# ECF504-BAS regression: 2 candidates ($2.10 vs $34.72) is two real prices, not a parsing error.
assert implausible_units(result(cand("Dongguan Baimiya", 2.10, 1, "stated"), cand("Wusheng", 34.72, 1, "stated"))) == set()

# Same manufacturer twice -> keep the higher-scoring one, no backfill; placeholder names are exempt.
dupes = [cand("PremierCable", 15.5, 1, "stated", NINETY), cand("Other", 9.0, 1, "stated"),
         cand(" premiercable ", 18.66, 1, "stated"), cand("Unknown/Generic", 1.0, 1, "stated"),
         cand("Unknown/Generic", 2.0, 1, "stated")]
kept = dedupe_manufacturers(dupes)
assert [(c.manufacturer, c.price_total) for c in kept] == [
    ("Other", 9.0), (" premiercable ", 18.66), ("Unknown/Generic", 1.0), ("Unknown/Generic", 2.0)], kept

# Stated beats a (plausible) inferred price at equal accuracy, even though inferred is cheaper.
idx, reason = recommend(result(cand("Inferred", 1.5, 1, "inferred"), cand("Stated", 2.0, 1, "stated")), 23.79)
assert idx == 1, reason

# Garbled model output (real ECF504-BAS values) is detected; clean output isn't.
bad = cand("Do nggu an", 2.10, 1, "stated")
bad.url, bad.distributor_site = "ht tps ://w ww .al ib ab a.c om/ pro duc t-d eta il/ US B-a", " al ib ab a.c om"
good = cand("Dongguan", 2.10, 1, "stated")
good.url, good.distributor_site = "https://www.alibaba.com/product-detail/USB-a-Male-to-USB-B_1601687478124.html", "alibaba.com"
assert is_garbled(result(good, bad)) and not is_garbled(result(good)) and not is_garbled(None)

# No L-Com price -> no recommendation.
assert recommend(result(cand("A", 1.0, 1, "stated")), None)[0] is None

# Excel: Recommended columns filled + green; manufacturer blocks not green; blank when nothing qualifies.
path = os.path.join(tempfile.mkdtemp(), "t.xlsx")
write_excel_report([
    ({"sku": "S1", "description": "d", "lcom_price": 22.59},
     result(cand("Close", 15.0, 1, "stated"), cand("Cheap", 2.0, 1, "stated"))),
    ({"sku": "S2", "description": "d", "lcom_price": 22.59}, result(cand("Close", 21.0, 1, "stated"))),
], path)
wb = openpyxl.load_workbook(path)
ws = wb["Results"]
header = [c.value for c in ws[1]]
rec = {h: header.index(h) for h in ("Recommended Manufacturer", "Recommended Email",
                                    "Recommended URL", "Recommended Unit Price")}
row = ws[2]
assert [row[i].value for i in rec.values()] == ["Cheap", "Cheap@x", "u/Cheap", 2.0]
assert all(row[i].fill.fgColor.rgb.endswith("C6EFCE") for i in rec.values())
assert not any(c.fill.fgColor.rgb.endswith("C6EFCE") for c in row[max(rec.values()) + 1:])
assert all(ws[3][i].value is None for i in rec.values())
assert not any(c.fill.fgColor.rgb.endswith("C6EFCE") for c in ws[3])
assert wb["Comparison"].cell(3, 8).value == "YES"

# 5 candidates: all get an Excel block, and the 5th can be the recommendation.
five = result(*[cand(f"M{n}", 15.0 + n, 1, "stated") for n in range(1, 5)], cand("M5", 2.0, 1, "stated"))
assert recommend(five, 22.59)[0] == 4
write_excel_report([({"sku": "S5", "description": "d", "lcom_price": 22.59}, five)], path)
ws = openpyxl.load_workbook(path)["Results"]
header = [c.value for c in ws[1]]
assert "Manufacturer 5 Name" in header and ws[2][header.index("Manufacturer 5 Name")].value == "M5"
assert ws[2][header.index("Recommended Manufacturer")].value == "M5"
# Page chrome stripped before truncation: menus/URLs gone, listing title/price/supplier/product URL kept.
import json
menu = "".join(f"*   [Category {n} & Things](https://www.made-in-china.com/cat/{n}.html)\n" for n in range(300))
listing = ("## [LC Female-SC Female Simplex Adapter](https://fm.en.made-in-china.com/product/QUt/China-LC.html?x=1)\n"
           "**US$0.07-0.25**\n10 Pieces (MOQ)\n[Shenzhen FiberMania Technology Co., Ltd.](https://fm.en.made-in-china.com/)\n")
page = json.dumps({"content": "[Sign in](https://login.x.com/?next=y)\n" + menu + listing}) + ' {"conversation_id": "c"}'
seen = strip_page_chrome(page)
assert len(page) > 20000 and len(seen) < 800 and "login" not in seen, (len(seen), seen[:200])
assert "https://fm.en.made-in-china.com/product/QUt/China-LC.html)" in seen and "US$0.07" in seen
assert "Shenzhen FiberMania Technology Co., Ltd." in seen and "fm.en.made-in-china.com/)" not in seen
assert strip_page_chrome("plain text, no json") == "plain text, no json"

# --sku: case-insensitive, order as asked, duplicates collapsed, unknown SKUs reported.
pool = [{"sku": "FOA-020C"}, {"sku": "C&P9M"}, {"sku": "HG2409U-PRO"}]
picked, missing = pick_skus(pool, ["hg2409u-pro", "FOA-020C", "FOA-020C", "NOPE"])
assert [p["sku"] for p in picked] == ["HG2409U-PRO", "FOA-020C"] and missing == ["NOPE"]
print("ok")
