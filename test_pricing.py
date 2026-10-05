"""Offline check of unit-price / L-Com margin / recommendation logic. Run: python test_pricing.py"""
import os
import tempfile

import openpyxl

from sourcing_agent import (
    AttributeBreakdown, Candidate, SourcingResult, candidate_comment, candidate_score, dedupe_manufacturers,
    implausible_units, is_garbled, pack_size_supported, pick_skus, percentile, recommend, score_summary,
    strip_page_chrome, timing_summary, unit_price, write_excel_report,
)

FULL = AttributeBreakdown(product_type=20, category_spec=20, shielding_material=20, mount_form=20, gender_pins=20)
NINETY = AttributeBreakdown(product_type=20, category_spec=20, shielding_material=20, mount_form=10, gender_pins=20)
SEVENTY = AttributeBreakdown(product_type=20, category_spec=20, shielding_material=10, mount_form=10, gender_pins=10)


def cand(name, total, qty, conf, breakdown=FULL, title="", price="", form=True, listing_form="panel mount coupler",
         caveats="", moq=None, spec_lines=""):
    return Candidate(manufacturer=name, price_total=total or 0, quantity_covered=qty or 0, unit_price_note="",
                     unit_price_confidence=conf, attribute_breakdown=breakdown.model_copy(), email=f"{name}@x",
                     url=f"u/{name}", listing_title=title, price=price, same_product_form=form,
                     listing_form=listing_form, listing_caveats=caveats, moq=moq, spec_lines=spec_lines)


def result(*cands):
    return SourcingResult(product="x", no_match=False, candidates=list(cands))


# Bundle price is divided down only when the pack size is attached to the listing; ambiguous never priced.
assert unit_price(cand("AS95", 19.75, 10, "inferred", title="RJ45 Panel Mount Coupler(10PCS)", price="$19.75")) == 1.975
assert unit_price(cand("A", 19.75, 10, "ambiguous", title="Coupler 10PCS")) is None
assert unit_price(cand("A", 19.75, None, "stated")) is None

# MOQ is not a pack size (item 3). Suzhou Bulovb, exact listing text from the 2026-09-25 run: $6.88 per piece, not $0.0688.
bulovb = cand("Suzhou Bulovb Electronic Co., Ltd.", 6.88, 100, "inferred",
              title="USB2.0 a Male & a Female to B Female Printer Print Converter Adapter Connector USB 2.0 Port",
              price="US$0.99-6.88")
assert not pack_size_supported(bulovb) and unit_price(bulovb) == 6.88
# HDFF / FARSINCE, the original case: "$1.27-1.58, Min. order 500 pieces" -> $1.58 per piece, not $0.0032.
farsince = cand("FARSINCE", 1.58, 500, "inferred", title="FARSINCE 4K 8K HDMI Panel Coupler HDMI Female to Female",
                price="$1.27-1.58 Min. order: 500 pieces")
assert not pack_size_supported(farsince) and unit_price(farsince) == 1.58
assert not pack_size_supported(cand("X", 6.88, 100, "stated", price="US$0.99-6.88 100 Pieces (MOQ)"))
assert not pack_size_supported(cand("X", 1.5, 100, "stated", price="$1.50 (100-499 pieces)"))

# A per-piece price stated next to the pack price counts as the pack size ("$23.68 ($4.74/pc)" = 5-pack).
assert unit_price(cand("AliX", 23.68, 5, "inferred", price="$23.68 ($4.74/pc)")) == 23.68 / 5
assert unit_price(cand("AliX", 23.68, 5, "inferred", price="$23.68 ($2.00/pc)")) == 23.68  # doesn't match -> per piece

# Cable backstop (2026-09-28 run): a stated length or a plain "cable" title is excluded for coupler/adapter
# targets, whatever the model said; couplers "for ... cable" are left alone.
from sourcing_agent import apply_form_rules
coupler_target = {"sku": "ECF504-UABS", "description": "USB Adapter A-B, Shielded"}
r = result(cand("Cooyear", 0.65, 1, "stated", title="50cm USB 2.0 Type A Female to B Male Adapter Cable Fast Charging"),
           cand("Wusheng", 2.0, 1, "stated", title="USB A-B PANEL ADPT SHIELDED"),
           cand("Inline", 1.0, 1, "stated", title="Cat6 FTP Shielded Female to Female Inline Cable Coupler"))
apply_form_rules(r, coupler_target)
assert [c.same_product_form for c in r.candidates] == [False, True, True], [c.listing_form for c in r.candidates]
assert "length 50cm" in r.candidates[0].listing_form
antenna = result(cand("Pigtail", 5.0, 1, "stated", title="2.4GHz 9dBi antenna with 5m cable"))
apply_form_rules(antenna, {"sku": "HG", "description": "2.4 GHz 9dBi Omnidirectional Antenna, N-Female Connector"})
assert antenna.candidates[0].same_product_form  # not a coupler/adapter target: left to the model

# Inline couplers are the same form as a panel-mount coupler target (real titles the model excluded,
# 2026-09-28); multi-port, wrong-connector and cable listings stay excluded.
sc_target = {"sku": "ECF504-SC5E", "description": "Cat5e RJ45 Coupler Shielded (8x8) Panel Mount Style"}
r = result(cand("SC5E-inline", 1.0, 1, "stated", title="Cat6 Waterproof Shielded RJ45 Inline Coupler Female to Female Straight",
                form=False, listing_form="Inline coupler"),
           cand("SC6-inline", 1.0, 1, "stated", title="CAT6 RJ45 8p8c Network Jack in-Line Coupler Female to Female",
                form=False, listing_form="In-line coupler"),
           cand("Cable-coupler", 1.0, 1, "stated", title="Cat6 FTP Shielded Female to Female Inline Cable Coupler",
                form=False, listing_form="Inline coupler (no cable)"),
           cand("4port", 1.0, 1, "stated", title="RJ45 4-Port Inline Coupler Female", form=False, listing_form="4 port inline coupler"),
           cand("RJ11", 1.0, 1, "stated", title="RJ11 6P4C Inline Coupler", form=False, listing_form="RJ11 inline coupler"),
           cand("Patch", 1.0, 1, "stated", title="RJ45 Inline Coupler with 1m Patch Cable", form=False, listing_form="inline coupler"))
apply_form_rules(r, sc_target)
assert [c.same_product_form for c in r.candidates] == [True, True, True, False, False, False], \
    [(c.manufacturer, c.same_product_form) for c in r.candidates]
assert "mount scored separately" in r.candidates[0].listing_form

# 100% match barely under L-Com loses to 90% match with a big gap.
idx, reason = recommend(result(cand("Close", 21.0, 1, "stated"), cand("Cheap", 4.0, 1, "stated", NINETY),
                               cand("Mystery", 1.0, 1, "ambiguous")), 22.59)
assert idx == 1, reason

# Messages name the bar that failed (item 7).
idx, reason = recommend(result(cand("Half", 11.0, 1, "stated")), 22.59)
assert idx is None and "price margin only" in reason and ">= 80% cheaper" in reason and "do not source" in reason.lower(), reason
idx, reason = recommend(result(cand("Close", 2.0, 1, "stated", SEVENTY)), 22.59)
assert idx is None and "accuracy only" in reason and "70%" in reason and "does qualify" in reason, reason
idx, reason = recommend(result(cand("Bad", 21.0, 1, "stated", SEVENTY)), 22.59)
assert idx is None and "accuracy and price" in reason, reason

# Implausible guard still works on a genuinely too-low price.
hdff = result(cand("A", 1.09, 1, "stated"), cand("B", 1.10, 1, "stated"), cand("Cheapo", 0.05, 1, "stated"))
assert implausible_units(hdff) == {2}
idx, reason = recommend(hdff, 23.79)
assert idx == 0, reason
idx, reason = recommend(result(cand("Real", 20.0, 1, "stated"), cand("Real2", 21.0, 1, "stated"),
                               cand("Bug", 0.05, 1, "inferred")), 23.79)
assert idx is None and "implausible" in reason, reason
idx, reason = recommend(result(cand("Acc", 20.0, 1, "stated"), cand("Acc2", 21.0, 1, "stated"), cand("Unpriced", 0, 1, "ambiguous"),
                               cand("Bug", 0.05, 1, "inferred"), cand("Top", 0.0, 1, "ambiguous")), 23.79)
assert "but is " in reason or "but its price" in reason, reason

# Wrong product form zeroes the match and is never recommended, even when cheapest (item 1, SC5E Lanka vs Kabasi).
lanka = cand("Lanka Industrial Automation", 0.70, 1, "stated", NINETY, form=False, listing_form="patch cable with coupler end")
kabasi = cand("Xiamen Kabasi Electric", 2.33, 1, "stated", NINETY)
assert candidate_score(lanka)[0] == 0 and "Wrong product form" in candidate_comment(lanka)
idx, reason = recommend(result(lanka, kabasi), 22.59)
assert idx == 1, reason
idx, reason = recommend(result(lanka), 22.59)
assert idx is None and "product form" in reason and "patch cable" in reason, reason

# Missing scores say "scoring failed", never a margin message (item 4, HDFF).
unscored = cand("Mo-Tech", 1.5, 1, "stated")
unscored.attribute_breakdown = None
idx, reason = recommend(result(unscored), 23.79)
assert idx is None and reason.startswith("Scoring failed"), reason

# Accuracy first: a 100% inferred price beats an 80% stated one; within 5 points, stated wins.
EIGHTY = AttributeBreakdown(product_type=20, category_spec=20, shielding_material=20, mount_form=20, gender_pins=0)
NINETY_FIVE = AttributeBreakdown(product_type=20, category_spec=20, shielding_material=20, mount_form=15, gender_pins=20)
idx, _ = recommend(result(cand("Stated80", 2.0, 1, "stated", EIGHTY), cand("Inferred100", 1.5, 1, "inferred")), 23.79)
assert idx == 1
idx, reason = recommend(result(cand("Inferred100", 1.5, 1, "inferred"), cand("Stated95", 2.0, 1, "stated", NINETY_FIVE)), 23.79)
assert idx == 1 and "within 5 accuracy points" in reason, reason
# Tie-break among equally accurate qualifiers (item 8): stated > inferred > promo, then named maker, then lowest price.
idx, _ = recommend(result(cand("Promo seller", 1.09, 1, "promo (regular price unknown)"), cand("Acme", 2.0, 1, "stated")), 23.79)
assert idx == 1
idx, _ = recommend(result(cand("Unknown seller", 1.0, 1, "stated"), cand("Acme", 2.0, 1, "stated")), 23.79)
assert idx == 1
idx, _ = recommend(result(cand("Acme", 2.0, 1, "stated"), cand("Beta Co", 1.5, 1, "stated")), 23.79)
assert idx == 1

# Summary line is templated from the table's scores (item 5); comments don't invent a reason for zeros.
line = score_summary(result(kabasi, lanka))
assert "90% (20/20/20/10/20)" in line and "0% (20/20/20/10/20, wrong form (patch cable with coupler end))" in line, line
no_mount = AttributeBreakdown(product_type=20, category_spec=20, shielding_material=20, mount_form=0, gender_pins=20)
assert candidate_comment(cand("ROHO", 14.79, 1, "stated", no_mount)) == "0 pts on: Mount/form factor"

# ECF504-BAS regression: 2 candidates ($2.10 vs $34.72) is two real prices, not a parsing error.
assert implausible_units(result(cand("Dongguan Baimiya", 2.10, 1, "stated"), cand("Wusheng", 34.72, 1, "stated"))) == set()

# Same manufacturer twice -> keep the higher-scoring one, no backfill; placeholder names are exempt.
dupes = [cand("PremierCable", 15.5, 1, "stated", NINETY), cand("Other", 9.0, 1, "stated"),
         cand(" premiercable ", 18.66, 1, "stated"), cand("Unknown/Generic", 1.0, 1, "stated"),
         cand("Unknown/Generic", 2.0, 1, "stated")]
kept = dedupe_manufacturers(dupes)
assert [(c.manufacturer, c.price_total) for c in kept] == [
    ("Other", 9.0), (" premiercable ", 18.66), ("Unknown/Generic", 1.0), ("Unknown/Generic", 2.0)], kept

# Stated beats a (plausible) inferred price, even though inferred is cheaper.
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

# Fix 1: antenna band/port. Roho's exact title from the 29 Sep run (and the review's wording of it) must no
# longer be recommendable for a single-band, single-port target; single-band antennas are untouched.
from sourcing_agent import apply_spec_rules, ordering_notes, ordering_note, ORDERING_NOTE_FIELD
hg = {"sku": "HG2409U-PRO", "description": "2.4 GHz 9dBi Omnidirectional Antenna, N-Female Connector", "lcom_price": 99.00}
ROHO_95 = AttributeBreakdown(product_type=20, category_spec=20, shielding_material=20, mount_form=20, gender_pins=15)
for roho_title in ("2way 2.4GHz 5.8GHz 5-9dBi Dual Band MIMO Omni Direction N Female Connector Pole Mount Fiberglass Antenna",
                   "2way 2.4GHz/5.8GHz 5.9dBi Dual-Band MIMO Omni-Direction N-Female Connector Pole-Mount Fiberglass Antenna"):
    r = result(cand("Roho Connector", 14.79, 1, "stated", ROHO_95, title=roho_title))
    apply_spec_rules(r, hg)
    assert candidate_score(r.candidates[0])[0] == 75 and "single-band target" in candidate_comment(r.candidates[0])
    assert recommend(r, hg["lcom_price"])[0] is None
for single in ("9dBi High Gain 2.4GHz Omni Fiberglass Antenna N Female", "2400-2500MHz Omni Outdoor Antenna 9dBi N Female"):
    r = result(cand("Single", 14.79, 1, "stated", ROHO_95, title=single))
    apply_spec_rules(r, hg)
    assert candidate_score(r.candidates[0])[0] == 95, single
r = result(cand("Coupler", 1.0, 1, "stated", title="Dual Band 2.4GHz/5.8GHz thing"))
apply_spec_rules(r, {"description": "HDMI Panel Mount Adapter, Female to Female"})
assert candidate_score(r.candidates[0])[0] == 100  # non-antenna targets never touched

# Fix 2: "check before ordering" notes, from this run's three listings (29 Sep) plus a clean one.
dsub = {"sku": "C&P9M", "description": "Insertion Type D-Sub Connector, DB9 Male"}
notes = ordering_notes(cand("FF TEK", 0.25, 1, "stated", title="High Quality PCB DIP Mount D-SUB Standard Connectors Male Female VGA Dsub Connector",
                            price="US$0.03-0.25", moq="500 pieces"), dsub)
assert any("D-sub pin count" in n and "VGA" in n for n in notes) and any("500+ pieces" in n for n in notes), notes
notes = ordering_notes(cand("Centron", 6.50, 1, "stated", title="Milcom MC-6BP MC-6BR Lightning Arrestor Coaxial Surge Protector "
                            "Bulkhead N-Type Male/Female 6GHz Transmitter RF", moq="5 pieces"),
                       {"description": "Coaxial Surge Protector, 18kA, 50 ohm, N-Type F/F Bulkhead, 1 Pole"})
assert any("MC-6BP, MC-6BR" in n for n in notes), notes
notes = ordering_notes(cand("TOPNET", 0.075, 1, "stated", title="Sc LC Simplex Duplex Fiber Optic Coupler Adaptor",
                            price="US$0.062-0.075", moq="100 Piece"),
                       {"description": "LC to SC Simplex Multimode Fiber Optic Adapter"})
assert any("simplex and duplex" in n for n in notes) and not any("MOQ" in n for n in notes), notes  # MOQ 100 < 500
assert any("1,000+" in n for n in ordering_notes(cand("XTZ", 0.20, 1, "stated", price="$0.20 1,000 Pieces (MOQ)"),
                                                  {"description": "HDMI Panel Mount Adapter, Female to Female"}))
assert ordering_notes(cand("Eternalstar", 1.25, 1, "stated", title="DVI to DVI Adapter Female to Female Converter DVI-I (24+5) Female to Female",
                           price="$1.25"), {"description": "DVI 24+5 female to female coupler"}) == []
assert ordering_notes(cand("Nice U.mi", 4.24, 1, "stated", title="1-4PCS RJ45 Panel Mount Coupler Shielded D-Type RJ45 Connector CAT6 "
                           "Female To Female LAN Network Bulkhead Pass Through Socket", price="$4.24"),
                      {"description": "Cat6 RJ45 Coupler Shielded (8x8) Panel Mount Style"}) == []
assert ordering_notes(cand("X", 1.0, 1, "stated", caveats="only the MC-6BR variant is F/F"), {"description": "x"}) == \
    ["model: only the MC-6BR variant is F/F"]

# MOQ/tier remarks move out of the model caveat into the >= 500 check (exact caveats from the 29 Sep run, where
# the structured moq field said "Not stated"). 500 is flagged, 499 isn't; useful caveats are kept.
def note_for(caveat, moq="Not stated"):
    return ordering_notes(cand("X", 1.0, 1, "stated", caveats=caveat, moq=moq), {"description": "x"})
assert note_for("Minimum order quantity is 500 pieces") == ["price needs an order of 500+ pieces (MOQ)"]  # HDFF
assert note_for("MOQ 1000 pieces at this price; lower volumes (10-9999) available at $1.05-1.35") == \
    ["price needs an order of 1,000+ pieces (MOQ)"]  # TDG1026KS-C6
assert note_for("Minimum order quantity is 499 pieces") == []
assert note_for("Minimum order 100 pieces; pricing shown is for 100-499 piece quantity tier") == []  # VIC00001
assert note_for("Tiered pricing at volume; Square Flange form factor specified but adapter gender may vary by model selection") == \
    ["model: Square Flange form factor specified but adapter gender may vary by model selection"]  # FOA-020C
assert note_for("PE plastic material (not silicone); female port only", moq="1,000 pieces") == \
    ["price needs an order of 1,000+ pieces (MOQ)", "model: PE plastic material (not silicone); female port only"]  # CAPUSB-A
assert note_for("CAT6 spec listed; target is CAT5e") == ["model: CAT6 spec listed; target is CAT5e"]  # ECF504-SC5E

# ...and they land in the Excel Ordering Note column (amber) only for a recommended, flagged product.
path = os.path.join(tempfile.mkdtemp(), "notes.xlsx")
write_excel_report([
    ({"sku": "S1", "description": "Insertion Type D-Sub Connector, DB9 Male", "lcom_price": 3.98},
     result(cand("FF TEK", 0.25, 1, "stated", title="Male Female VGA Dsub Connector"))),
    ({"sku": "S2", "description": "DVI 24+5 female to female coupler", "lcom_price": 28.19},
     result(cand("Eternalstar", 1.25, 1, "stated", title="DVI Female to Female Coupler"))),
    ({"sku": "S3", "description": "d", "lcom_price": 5.0}, result(cand("Pricey", 4.0, 1, "stated", title="VGA thing"))),
], path)
ws = openpyxl.load_workbook(path)["Results"]
col = [c.value for c in ws[1]].index(ORDERING_NOTE_FIELD)
assert "VGA" in ws[2][col].value and ws[2][col].fill.fgColor.rgb.endswith("FFE699")
assert ws[3][col].value in (None, "") and ws[4][col].value in (None, "")

# UABS $0.11 "adapter" (29 Sep, 102713): really a printer data cable at a junk listing price. Three independent
# catches now: the cable phrase, the model's own low-price warning, and the price check once the $1.09 promo
# prices beside it are fixed.
from sourcing_agent import apply_price_rules, cable_evidence
uabs = {"sku": "ECF504-UABS", "description": "USB Adapter A-B, Shielded", "lcom_price": 20.19}
tongze = cand("unnamed", 0.11, 1, "stated", title="USB 2.0 High-Speed Square Port Printer Data Cable Adapter with a Male to B Male "
              "Shielded Magnetic Ring", price="$0.11", moq="5 pieces",
              caveats="Price unusually low; verify actual product form and shielding specifications on order")
cand1 = cand("unnamed", 1.63, 1, "stated", title="USB2.0 Type a to B Printer Scanner Cable High Speed Data Transfer Cord Shielded "
             "PVC Jacket Male Connector for Computer", price="$0.77-1.63")
cand4 = cand("unnamed", 1.09, 1, "inferred", SEVENTY, title="High Speed USB 2.0 Type A Female To Type B Male USB Printer Scanner "
             "Adapter Data Sync Coupler Converter Connector", price="$1.09")
cand4.unit_price_note = "Title explicitly lists female-to-male adapter; promo pricing shown ($1.09 new shopper discount from base price ~$2.31)"
cand5 = cand("unnamed", 1.09, 1, "inferred", SEVENTY, title="USB 2.0 A Male & Female to USB Type B Print Converter Adapter", price="$1.09")
cand5.unit_price_note = "Title lists both male and female variants; promo pricing ($1.09 new shopper discount)"
assert "Printer Data Cable" in cable_evidence(tongze.listing_title) and "Scanner Cable" in cable_evidence(cand1.listing_title)
assert cable_evidence("Cat6 IP67 Waterproof RJ45 Bulkhead Connector Outdoor Ethernet Cable Panel Mount") == ""  # real coupler
assert cable_evidence("SC-LC Simplex Hybrid Adapter with Flange Metal Fiber Optic Patch Cord Pigtail") == ""   # real adapter
r = result(cand1, tongze, cand4, cand5)
apply_form_rules(r, uabs)
apply_price_rules(r)
assert [c.same_product_form for c in r.candidates] == [False, False, True, True]
assert cand4.price_total == 2.31 and "regular price $2.31" in cand4.code_notes[0]
assert cand5.unit_price_confidence == "promo (regular price unknown)"
assert 1 in implausible_units(r)  # the model's own warning alone flags it
idx, reason = recommend(r, uabs["lcom_price"])
assert idx is None and "accuracy only" in reason, reason
# Without the warning or the cable rule, the price check alone now fires: reference is $1.16/$2.31, not the $1.09 promo.
plain = cand("unnamed", 0.11, 1, "stated")
others = [cand("A", 1.63, 1, "stated"), cand("B", 1.16, 1, "stated"), cand("C", 1.09, 1, "promo (regular price unknown)")]
assert 0 in implausible_units(result(plain, *others))
# The promo backstop leaves correctly used regular prices alone.
ok = cand("X", 2.28, 1, "stated"); ok.unit_price_note = "Regular price $2.28 (promo $1.09 for new shoppers)"
ok2 = cand("Y", 5.77, 1, "stated"); ok2.unit_price_note = "Single unit; promo price shown but regular price $5.77 used"
leak = cand("Z", 1.09, 1, "stated", price="$1.09"); leak.unit_price_note = "Single piece price shown at $1.09 (promotional price, regular $1.49)"
lcsp = cand("L", 22.84, 1, "stated", price="$17.13 $22.84 -25%")  # replay false positive, 102713 LCSP1050 cand 4
lcsp.unit_price_note = "Regular price after promo discount; 50ohm explicitly stated"
apply_price_rules(result(ok, ok2, leak, lcsp))
assert (ok.price_total, ok.unit_price_confidence, ok2.price_total, leak.price_total) == (2.28, "stated", 5.77, 1.49)
assert (lcsp.price_total, lcsp.unit_price_confidence, lcsp.code_notes) == (22.84, "stated", [])

# --sku: case-insensitive, order as asked, duplicates collapsed, unknown SKUs reported.
pool = [{"sku": "FOA-020C"}, {"sku": "C&P9M"}, {"sku": "HG2409U-PRO"}]
picked, missing = pick_skus(pool, ["hg2409u-pro", "FOA-020C", "FOA-020C", "NOPE"])
assert [p["sku"] for p in picked] == ["HG2409U-PRO", "FOA-020C"] and missing == ["NOPE"]
# Tool wrapper against the real MCP result type (a fake with the wrong field name once hid a
# bug that failed every call): rate limit -> backoff + retry, then the page text comes through.
import asyncio
import sourcing_agent
from mcp import types as mcp_types


class FakeSession:
    replies = [mcp_types.CallToolResult(content=[mcp_types.TextContent(type="text", text="429 Too Many Requests")], is_error=True),
               mcp_types.CallToolResult(content=[mcp_types.TextContent(type="text", text='{"content": "Widget $1.09"}')])]

    async def call_tool(self, name, arguments, read_timeout_seconds=None):
        return self.replies.pop(0)


async def no_sleep(_):
    pass


sourcing_agent.asyncio.sleep, real_sleep = no_sleep, sourcing_agent.asyncio.sleep
tool = sourcing_agent.make_bounded_tool(
    type("T", (), {"name": "nimble_extract", "description": "", "inputSchema": {"type": "object", "properties": {}}})(),
    FakeSession(), sourcing_agent.ToolBudget(3), "TEST")
assert asyncio.run(tool.call({"url": "https://www.alibaba.com/x"})) == "Widget $1.09"
assert sourcing_agent.NIMBLE_RATE_LIMITS["count"] == 1 and sourcing_agent.CALL_LOG[-1]["outcome"] == "ok"
sourcing_agent.asyncio.sleep = real_sleep

# Timing summary: Anthropic time = product time - Nimble time; timeouts listed; percentiles sane.
assert percentile([1, 2, 3, 4, 10], 0.5) == 3 and percentile([1, 2, 3, 4, 10], 0.9) == 10 and percentile([], 0.5) == 0
calls = [{"sku": "A", "tool": "nimble_extract", "site": "alibaba.com", "url": "u1", "start": 0.0, "secs": 20.0, "outcome": "ok"},
         {"sku": "A", "tool": "nimble_extract", "site": "aliexpress.com", "url": "u2", "start": 10.0, "secs": 120.0, "outcome": "timeout"}]
summary = timing_summary(calls, [{"sku": "A", "secs": 200.0, "calls": 2, "stop": "found 5 candidates (2/13 calls, 1 failed)"}], 200.0)
import re
assert re.search(r"Waiting on Nimble:\s+2\.2 min\s+65%", summary) and re.search(r"Anthropic \+ overhead:\s+1\.2 min\s+35%", summary), summary  # 10-20s overlap counted once
assert "TIMEOUT A nimble_extract u2" in summary and "15-30s: 1" in summary, summary
assert "Peak simultaneous Nimble calls: 2" in summary and "found 5 candidates: 1" in summary, summary
assert re.search(r"extract aliexpress\.com\s+1 .* 1$", summary, re.M), summary  # per-site timeout column

# Peak concurrency: back-to-back calls don't overlap; three at once do.
from sourcing_agent import peak_concurrency, stop_reason, site_search_url, prefetch_site_searches, MAX_TOOL_CALLS_PER_PRODUCT
assert peak_concurrency([{"start": 0, "secs": 5}, {"start": 5, "secs": 5}]) == 1
from sourcing_agent import busy_time
assert busy_time([{"start": 0, "secs": 10}, {"start": 0, "secs": 10}, {"start": 0, "secs": 12}]) == 12  # 3 parallel = 12s, not 32s
assert busy_time([{"start": 0, "secs": 5}, {"start": 10, "secs": 5}]) == 10
assert peak_concurrency([{"start": 0, "secs": 5}, {"start": 1, "secs": 5}, {"start": 2, "secs": 1}]) == 3

# Stop reasons tell "found enough" from "ran out of calls" from "gave up with budget left".
five = result(*[cand(f"M{n}", 1.0, 1, "stated") for n in range(5)])
assert stop_reason(five, 9, 0).startswith("found 5 candidates")
assert stop_reason(result(cand("A", 1.0, 1, "stated")), MAX_TOOL_CALLS_PER_PRODUCT, 4) == \
    f"budget used up with 1 candidates ({MAX_TOOL_CALLS_PER_PRODUCT}/{MAX_TOOL_CALLS_PER_PRODUCT} calls, 4 failed)"
assert stop_reason(result(cand("A", 1.0, 1, "stated")), 4, 0).startswith("model stopped with 1 candidates, budget left")
assert stop_reason(None, 2, 2).startswith("no usable result")

# Prefetch: the three site searches run at the same time through the budgeted wrapper (3 calls used).
assert site_search_url("made-in-china.com", "Cat6 RJ45 Coupler (8x8)") == \
    "https://www.made-in-china.com/products-search/hot-china-products/Cat6_RJ45_Coupler_8x8.html"
assert site_search_url("alibaba.com", "LC to SC Adapter") == "https://www.alibaba.com/trade/search?SearchText=LC+to+SC+Adapter"


class SlowSession:
    in_flight = peak = 0

    async def call_tool(self, name, arguments, read_timeout_seconds=None):
        SlowSession.in_flight += 1
        SlowSession.peak = max(SlowSession.peak, SlowSession.in_flight)
        await asyncio.sleep(0.05)
        SlowSession.in_flight -= 1
        return mcp_types.CallToolResult(content=[mcp_types.TextContent(type="text", text=f"page for {arguments['url']}")])


budget = sourcing_agent.ToolBudget(MAX_TOOL_CALLS_PER_PRODUCT)
extract_tool = sourcing_agent.make_bounded_tool(
    type("T", (), {"name": "nimble_extract", "description": "", "inputSchema": {"type": "object", "properties": {}}})(),
    SlowSession(), budget, "PREFETCH")
pages = asyncio.run(prefetch_site_searches([extract_tool], {"sku": "S", "description": "Cat6 RJ45 Coupler"}))
assert budget.used == 3 and SlowSession.peak == 3, (budget.used, SlowSession.peak)
assert all(f"=== {site} search results" in pages for site in sourcing_agent.SOURCING_SITES)

# --- Srijan's v1 review (2026-10-04) ---
from sourcing_agent import (apply_lcom_prices, contact_types, load_lcom_prices, lcom_price_lines, ordering_notes,
                            price_is_unverified, price_tier_note)

# 1. Contact type is a hard exclude. C&P9M's four real 102713 candidates were all solder/PCB; the target now says crimp.
cp9m = {"sku": "C&P9M", "description": "Insertion Type D-Sub Connector, DB9 Male, Crimp Contacts"}
assert contact_types(cp9m["description"]) == {"crimp"}
assert contact_types(next(p for p in sourcing_agent.PRODUCTS if p["sku"] == "C&P9M")["description"]) == {"crimp"}
solder_titles = [
    "High Quality Straight PCB Wire Mount 9 Position Solder Cup D-SUB dB9 Standard Connectors Male",
    "10PCS DB9 Female Male PCB Mount serial port Connector Solder Type D-Sub RS232 COM CONNECTORS 9pin socket 9p Adapter FOR PCB",
    "PCB Solder/Screw Gold-plated Vertical Horizontal 9 15 25 23 37 Pin Blue Black Vga D-sub Db15 Db9 Female Male Dsub Connectors",
    "5PCS DB9 DB15 DB25 37 Female Male PCB Mount serial port Connector Solder Type D-Sub RS232 CONNECTORS 9pin socket Adapter FOR PCB",
]
crimp_ok = cand("CrimpCo", 0.5, 1, "stated", title="DB9 Male D-Sub Crimp Contact Connector Plastic Shell")
both = cand("BothCo", 0.4, 1, "stated", title="DB9 Male D-Sub Connector Crimp/Solder Cup")
silent = cand("SilentCo", 0.3, 1, "stated", title="DB9 Male D-Sub Connector Gold Plated")
r = result(*[cand(f"S{i}", 0.25, 1, "stated", title=t, listing_form="PCB solder connector") for i, t in enumerate(solder_titles)],
           crimp_ok, both, silent)
apply_form_rules(r, cp9m)
assert [c.same_product_form for c in r.candidates] == [False] * 4 + [True] * 3
assert "wrong contact type" in r.candidates[0].listing_form and candidate_score(r.candidates[0])[0] == 0
idx, reason = recommend(result(*r.candidates[:4]), 3.98)
assert idx is None and "product form" in reason, reason
assert ordering_notes(crimp_ok, cp9m) == []
assert any("several contact types" in n for n in ordering_notes(both, cp9m))
assert any("contact type isn't stated" in n for n in ordering_notes(silent, cp9m))
# A target that states no contact type (any RJ45 coupler) is unaffected, and so are its notes.
rj45 = {"sku": "ECF504-SC6", "description": "Cat6 RJ45 Coupler Shielded (8x8) Panel Mount Style"}
pcb_coupler = cand("K", 2.33, 1, "stated", title="Cnlinko RJ45 Shielded Panel Mount Coupler PCB Board Connector Solder")
apply_form_rules(result(pcb_coupler), rj45)
assert pcb_coupler.same_product_form and ordering_notes(pcb_coupler, rj45) == []

# Spec text that reads like a short cable gets a note (never an exclusion) on a coupler/adapter target.
hdff_t = {"sku": "HDFF", "description": "HDMI Panel Mount Adapter, Female to Female"}
hd = cand("XTZ", 0.2, 1, "stated", title="HDMI a Female to Female with Panel Mount Adapter",
          caveats="specs: Jacket PVC, AWG 24/28/26, Length (Customized)")
assert any("short cable" in n for n in ordering_notes(hd, hdff_t))
assert not any("short cable" in n for n in ordering_notes(cand("P", 0.2, 1, "stated", title="HDMI Panel Coupler"), hdff_t))
apply_form_rules(result(hd), hdff_t)
assert hd.same_product_form

# 3. Price ranges use the high end; a discount like "-25%" is not a range; tiers are named in the note.
capusb3 = cand("Dongguan", 0.05, 1, "stated", title="Silicone USB Cap Port Cover", price="US$0.05-1.58 (1,000 Pieces MOQ)")
capusb2 = cand("Mao Jia", 0.08, 1, "stated", title="Usb Waterproof Cap", price="$0.07-0.08 (1,000-9,999 pieces)")
cappromo = cand("L", 22.84, 1, "stated", title="x", price="$17.13 $22.84 -25%")
apply_price_rules(result(capusb3, capusb2, cappromo))
assert (capusb3.price_total, capusb2.price_total, cappromo.price_total) == (1.58, 0.08, 22.84)
assert "high end" in capusb3.code_notes[0] and not cappromo.code_notes
tiered = cand("Lung Kay", 2.68, 1, "stated", title="USB adapter", price="$2.68 100-999 pieces $2.58 ≥1,000 pieces")
assert price_tier_note(tiered) == "price is for the 100-999 pieces tier (others: $2.58 at ≥1,000 pieces)", price_tier_note(tiered)
assert price_tier_note(capusb3) == ""

# 2. L-Com prices come from lcom_prices.csv; missing SKUs keep their price and are marked unverified.
csv_path = os.path.join(tempfile.mkdtemp(), "lcom_prices.csv")
with open(csv_path, "w", newline="") as f:
    f.write("sku,pack_size,pack_price,source,date_checked\n"
            'CAPUSB-A,10,19.99,"Srijan, live check",2026-10-04\nHDFF,1,31.70,"Srijan, live check",2026-10-04\n'
            "OLD,1,5.00,unverified,\n")
prods = [{"sku": "CAPUSB-A", "lcom_price": 3.998}, {"sku": "hdff", "lcom_price": 23.79},
         {"sku": "OLD", "lcom_price": 4.0}, {"sku": "MISSING", "lcom_price": 9.99}]
apply_lcom_prices(prods, csv_path)
assert prods[0]["lcom_price"] == 2.0 and prods[1]["lcom_price"] == 31.70 and prods[2]["lcom_price"] == 5.0
assert prods[3]["lcom_price"] == 9.99 and price_is_unverified(prods[3]) and price_is_unverified(prods[2])
assert not price_is_unverified(prods[0]) and prods[0]["lcom_date"] == "2026-10-04"
header = "\n".join(lcom_price_lines(prods))
assert "Still on unverified L-Com prices (do not read the margin as confirmed):** OLD, MISSING" in header
# Margin at the new CAPUSB-A reference: $0.45 is 77.5% cheaper than $2.00, so it no longer qualifies; $0.40 does.
assert recommend(result(cand("A", 0.45, 1, "stated")), 2.0)[0] is None
assert recommend(result(cand("A", 0.40, 1, "stated")), 2.0)[0] == 0
# The seeded file itself: Srijan's four values, C&P9M flagged, every built-in/input SKU present.
seeded = load_lcom_prices()
assert {s: round(float(seeded[s]["pack_price"]) / int(seeded[s]["pack_size"]), 2) for s in ("HDFF", "FOA-020C", "ECF504-SC6", "CAPUSB-A")} == \
       {"HDFF": 31.70, "FOA-020C": 72.99, "ECF504-SC6": 20.54, "CAPUSB-A": 2.0}
assert seeded["C&P9M"]["source"] == "Newark distributor price, unconfirmed"
assert all(p["sku"].upper() in seeded for p in sourcing_agent.PRODUCTS)
# The Results sheet carries each price's source and date.
path = os.path.join(tempfile.mkdtemp(), "r.xlsx")
write_excel_report([({"sku": "HDFF", "description": "d", "lcom_price": 31.70, "lcom_source": "Srijan, live check",
                      "lcom_date": "2026-10-04"}, None)], path)
ws = openpyxl.load_workbook(path).worksheets[0]
hdr = [c.value for c in ws[1]]
row = dict(zip(hdr, [c.value for c in ws[2]]))
assert row["L-Com Price Source"] == "Srijan, live check" and row["L-Com Price Date"] == "2026-10-04"
# --- Srijan's answers, round 2 (2026-10-04) ---
from sourcing_agent import apply_reviewer_exclusions, contradictions, format_result_markdown

# Reviewer exclusions: seeded for HDFF / Xiangtianzhong and VIC00001 / Xindaying; matches maker name or URL.
xtz = cand("SHENZHEN XIANGTIANZHONG TECHNOLOGY CO., LTD.", 0.2, 1, "stated", title="HDMI a Female to Female with Panel Mount Adapter")
other = cand("FARSINCE", 1.58, 1, "stated", title="HDMI Panel Coupler")
by_url = cand("", 0.2, 1, "stated", title="HDMI Panel Adapter")
by_url.url = "https://xtz-tech.en.made-in-china.com/product/hOtASZeTkIcF/xiangtianzhong-hdmi.html"
rr = result(xtz, other, by_url)
apply_reviewer_exclusions(rr, "hdff")
assert [c.same_product_form for c in rr.candidates] == [False, True, False]
assert candidate_comment(xtz).startswith("excluded by reviewer: short cable per its product page")
assert candidate_score(xtz)[0] == 0
vic = result(cand("Shenzhen Xindaying Technology Co., Ltd.", 1.6, 1, "stated"))
apply_reviewer_exclusions(vic, "VIC00001")
assert not vic.candidates[0].same_product_form
untouched = result(cand("Shenzhen Xindaying Technology Co., Ltd.", 1.6, 1, "stated"))
apply_reviewer_exclusions(untouched, "HDFF")  # exclusion is per SKU
assert untouched.candidates[0].same_product_form

# Self-contradiction: flagged with a note, not zeroed. Real texts from the 2026-10-04 page reads.
kabasi = cand("Kabasi", 2.33, 1, "stated", title="Cnlinko Yt-RJ45 Shielded Industrial Panel Mount Bulkhead Female Coupler",
              spec_lines="Compatible with CAT5e and compliant with EIA568B; Terminal Type: Cat.3-Cat.6A")
assert contradictions(kabasi) and "category" in contradictions(kabasi)[0]
lungkay = cand("LUNG KAY", 2.68, 1, "stated", title="LUNG KAY USB 2.0 Type a Female to USB Type B Male Adaptor",
               spec_lines="Connector A: Type A male; Connector B: Type B Male")
assert [("gender" in n) for n in contradictions(lungkay)] == [True]
hyconnect = cand("Hy", 0.7, 1, "stated", title="8P8C Cat 6 STP FTP RJ45 to RJ45 Inline Coupler Shielded Cat5e Cat6 Cat6A Keystone Jack",
                 spec_lines="Category: Cat5e Cat6 Cat6a; Gender: Female R45 Keystone Jack")
assert contradictions(hyconnect) == []  # a legitimate multi-category listing is untouched
assert candidate_score(kabasi)[0] == 100 and any("contradicts" in n for n in ordering_notes(kabasi, rj45))
assert "contradicts" in candidate_comment(kabasi)

# IDC is its own contact type: IDC-only vs a crimp target is excluded; "IDC ... Crimp" states both, so it's kept with a note.
idc_only = cand("I", 0.5, 1, "stated", title="DB9 Male IDC Ribbon D-Sub Connector")
idc_crimp = cand("IC", 0.5, 1, "stated", title="DSUB E09P IDC RIBBON DB9 9P Male IDC D Sub Connector Insulation Displacement Crimp Type")
apply_form_rules(result(idc_only, idc_crimp), cp9m)
assert not idc_only.same_product_form and idc_crimp.same_product_form
assert any("several contact types" in n for n in ordering_notes(idc_crimp, cp9m))

# A candidate with no URL can't be recommended; one with a URL still can.
nourl = cand("NoUrl", 0.2, 1, "stated"); nourl.url = None
withurl = cand("HasUrl", 0.3, 1, "stated")
idx, reason = recommend(result(nourl), 10.0)
assert idx is None and "no URL" in reason and "locate it manually" in reason, reason
assert recommend(result(nourl, withurl), 10.0)[0] == 1

# The report shows the spec lines read, or says none were.
md = format_result_markdown({"sku": "K", "description": "d", "lcom_price": 20.54}, result(kabasi, hyconnect), 5, 1, 1, 0.0)
assert "| Spec lines seen | Compatible with CAT5e" in md
md2 = format_result_markdown({"sku": "K", "description": "d", "lcom_price": 20.54}, result(other), 5, 1, 1, 0.0)
assert "none (search-result title only)" in md2
# A fresh checkout (the CSVs are .gitignored) must not silently skip them: missing or empty is reported loudly.
from sourcing_agent import check_input_files, print_input_file_warnings
tmp = tempfile.mkdtemp()
empty_csv = os.path.join(tmp, "empty.csv")
with open(empty_csv, "w") as f:
    f.write("sku,supplier_or_url,reason\n\n")
full_csv = os.path.join(tmp, "full.csv")
with open(full_csv, "w") as f:
    f.write("sku,supplier_or_url,reason\nHDFF,x,y\n")
probs = check_input_files({"gone.csv": ("it matters", os.path.join(tmp, "gone.csv")), "empty.csv": ("it matters", empty_csv),
                           "full.csv": ("it matters", full_csv)})
assert probs == ["gone.csv is MISSING - it matters.", "empty.csv is EMPTY - it matters."], probs
assert check_input_files() == []  # the real files in this checkout are present and non-empty
import contextlib, io
buf = io.StringIO()
with contextlib.redirect_stdout(buf):
    print_input_file_warnings(probs)
    print_input_file_warnings([])
assert buf.getvalue().count("!!! WARNING") == 1 and "gone.csv is MISSING" in buf.getvalue()

# --- A crashed product must never read as a no-match (offline: a fake `research`, no API or Nimble calls) ---
import anthropic
import httpx
from sourcing_agent import (GREEN, error_text, is_fatal_api_error, product_recommendation, ps_quote, rerun_command,
                            run_batch, run_counts)


def api_error(cls, status, message):
    resp = httpx.Response(status, request=httpx.Request("POST", "https://api.anthropic.com/v1/messages"))
    return cls(message, response=resp, body={"type": "error", "error": {"type": "x", "message": message}})


credit_error = api_error(anthropic.BadRequestError, 400, "Your credit balance is too low to access the Anthropic API.")
assert is_fatal_api_error(credit_error)
assert is_fatal_api_error(api_error(anthropic.AuthenticationError, 401, "invalid x-api-key"))
assert not is_fatal_api_error(api_error(anthropic.BadRequestError, 400, "prompt is too long"))
assert not is_fatal_api_error(RuntimeError("boom")) and not is_fatal_api_error(UnicodeEncodeError("charmap", "x", 0, 1, "bad"))
assert error_text(RuntimeError("boom\n  twice")) == "RuntimeError: boom twice" and len(error_text(RuntimeError("x" * 500))) == 140


def fake_results():
    good = result(cand("GoodCo", 0.2, 1, "stated"))
    nomatch = SourcingResult(product="x", no_match=True, no_match_reason="nothing close", candidates=[])
    return good, nomatch


def batch(products, behaviours, tmp):
    """Runs run_batch with one scripted outcome per SKU: a result, None, or an exception to raise."""
    calls = []

    async def research(product):
        calls.append(product["sku"])
        outcome = behaviours[product["sku"]]
        if isinstance(outcome, BaseException):
            raise outcome
        return outcome, 100, 10, 3

    sourcing_agent.MAX_CONCURRENT_PRODUCTS = 1  # one at a time, so which products ran is deterministic
    try:
        done, why = asyncio.run(run_batch(products, research, os.path.join(tmp, "r.md"), os.path.join(tmp, "r.xlsx"),
                                          {"in": 0, "out": 0, "cost": 0.0}, "in put.xlsx"))
    finally:
        sourcing_agent.MAX_CONCURRENT_PRODUCTS = 3
    return done, why, calls


good, nomatch = fake_results()
mk = lambda sku: {"sku": sku, "description": "d", "product_name": "p", "lcom_price": 10.0}
tmp = tempfile.mkdtemp()
done, why, calls = batch([mk("OK"), mk("C&P9M"), mk("NORES"), mk("NOMATCH")],
                         {"OK": good, "C&P9M": RuntimeError("boom"), "NORES": None, "NOMATCH": nomatch}, tmp)
rows = [d[1] for d in done]
assert calls == ["OK", "C&P9M", "NORES", "NOMATCH"] and why == ""
crash, nores, genuine = rows[1][0], rows[2][0], rows[3][0]
for crashed in (crash, nores):
    text = product_recommendation(crashed, None)[1]
    assert text.startswith("Error researching this product: ") and text.endswith(" - re-run with --sku"), text
    assert "No candidates found" not in text and product_recommendation(crashed, None)[0] is None
assert "RuntimeError: boom" in product_recommendation(crash, None)[1]
assert "no structured result" in product_recommendation(nores, None)[1]
assert product_recommendation(genuine, rows[3][1])[1] == "No candidates found - nothing to source." and "error" not in genuine
assert recommend(None, 10.0)[1] != "No candidates found - nothing to source."  # even a bare None result can't pass as a no-match
assert run_counts(rows) == (1, 1, 2)
cmd = rerun_command(rows, "in put.xlsx")
assert cmd == "python sourcing_agent.py --input 'in put.xlsx' --sku 'C&P9M' 'NORES'", cmd
assert ps_quote("it's") == "'it''s'" and rerun_command([rows[0], rows[3]]) == ""
# The saved report and Excel carry the error, never green, never "No candidates found" for it.
md = open(os.path.join(tmp, "r.md"), encoding="utf-8").read()
assert "**1 recommended** | 1 no recommendation | **2 errored**" in md and f"`{cmd}`" in md
assert "**Recommendation:** Error researching this product: RuntimeError: boom - re-run with --sku" in md
wb = openpyxl.load_workbook(os.path.join(tmp, "r.xlsx"))
res, comp = wb["Results"], wb["Comparison"]
hdr = [c.value for c in res[1]]
by_sku = {r[0].value: r for r in res.iter_rows(min_row=2)}
rec_col = hdr.index("Recommendation")
assert by_sku["C&P9M"][rec_col].value == "Error researching this product: RuntimeError: boom - re-run with --sku"
assert by_sku["NORES"][rec_col].value.startswith("Error researching this product: no structured result")
assert by_sku["NOMATCH"][rec_col].value == "No candidates found - nothing to source."
for sku in ("C&P9M", "NORES"):
    assert not by_sku[sku][hdr.index("Recommended Manufacturer")].value and not by_sku[sku][hdr.index("Recommended URL")].value
    assert not any(c.fill.fgColor.rgb == GREEN.fgColor.rgb for c in by_sku[sku])
errs = [r for r in comp.iter_rows(min_row=2) if r[1].value == "ERROR"]
assert [r[0].value for r in errs] == ["C&P9M", "NORES"] and all(r[7].value.startswith("Error researching") for r in errs)
assert not any(c.fill.fgColor.rgb == GREEN.fgColor.rgb for r in errs for c in r)
assert any(r[0].value == "OK" and r[7].value == "YES" for r in comp.iter_rows(min_row=2))

# A billing/credit/auth error stops the batch: finished products are saved, the rest are not run at all.
tmp = tempfile.mkdtemp()
done, why, calls = batch([mk("A"), mk("B"), mk("C"), mk("D")], {"A": good, "B": credit_error, "C": good, "D": good}, tmp)
assert calls == ["A", "B"], calls  # C and D were never researched
assert "credit balance is too low" in why
rows = [d[1] for d in done]
assert run_counts(rows) == (1, 0, 3)
assert rows[2][0]["error"].startswith("not run - batch stopped early (BadRequestError: Your credit balance is too low")
assert rerun_command(rows).endswith("--sku 'B' 'C' 'D'")
md = open(os.path.join(tmp, "r.md"), encoding="utf-8").read()
assert "**BATCH STOPPED EARLY:** BadRequestError: Your credit balance is too low" in md and "**1 recommended**" in md
# An ordinary crash does not stop the batch.
tmp = tempfile.mkdtemp()
done, why, calls = batch([mk("A"), mk("B"), mk("C")], {"A": RuntimeError("x"), "B": good, "C": good}, tmp)
assert calls == ["A", "B", "C"] and why == "" and run_counts([d[1] for d in done]) == (2, 0, 1)

# --- Listing titles in the Results sheet: appended after every existing column, nothing moved or renamed ---
from sourcing_agent import (CANDIDATE_XLSX_FIELDS, MAX_CANDIDATES, ORDERING_NOTE_FIELD, PRODUCT_XLSX_FIELDS,
                            RECOMMENDED_XLSX_FIELDS)
existing = PRODUCT_XLSX_FIELDS + RECOMMENDED_XLSX_FIELDS + [ORDERING_NOTE_FIELD] + [
    f"Manufacturer {n} {f}" for n in range(1, MAX_CANDIDATES + 1) for f in CANDIDATE_XLSX_FIELDS]
titles = ["Recommended Listing Title"] + [f"Manufacturer {n} Listing Title" for n in range(1, MAX_CANDIDATES + 1)]
pa = cand("Alpha", 0.2, 1, "stated", title="Alpha's own Cat6 Coupler Title")        # the pick: cheap and accurate
pb = cand("Beta", 15.0, 1, "stated", title="Beta Listing, Panel Mount")             # not the pick: too expensive
pc = cand("Gamma", 0.3, 1, "stated", title=None)                                    # a candidate with no title at all
prod = {"sku": "TITLED", "description": "d", "product_name": "p", "lcom_price": 10.0}
path = os.path.join(tempfile.mkdtemp(), "titles.xlsx")
write_excel_report([(prod, result(pb, pa, pc)),
                    ({**prod, "sku": "NOREC"}, result(pb)),                          # no recommendation: titles still listed
                    ({**prod, "sku": "ERR", "error": "RuntimeError: boom"}, None)], path)
ws = openpyxl.load_workbook(path).worksheets[0]
hdr = [c.value for c in ws[1]]
assert hdr[:len(existing)] == existing and hdr[len(existing):] == titles, hdr   # existing columns untouched, titles last
rows_by_sku = {r[0].value: dict(zip(hdr, [c.value for c in r])) for r in ws.iter_rows(min_row=2)}
t = rows_by_sku["TITLED"]
assert t["Recommended Manufacturer"] == "Alpha" and t["Recommended Listing Title"] == "Alpha's own Cat6 Coupler Title"
assert [t[f"Manufacturer {n} Listing Title"] for n in (1, 2, 3, 4, 5)] == ["Beta Listing, Panel Mount",
                                                                         "Alpha's own Cat6 Coupler Title", None, None, None]
assert rows_by_sku["NOREC"]["Recommended Listing Title"] is None and rows_by_sku["NOREC"]["Manufacturer 1 Listing Title"] == "Beta Listing, Panel Mount"
assert rows_by_sku["ERR"]["Recommended Listing Title"] is None and rows_by_sku["ERR"]["Recommendation"].startswith("Error researching")
print("ok")
