"""Offline tests for content_agent.py with a stubbed model: no network, no key, no Walmart. Run: python test_content_agent.py"""
import contextlib
import copy
import csv
import io
import json
import os
import sys
import tempfile
from decimal import Decimal

import openpyxl

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
sys.stdout.reconfigure(encoding="utf-8", errors="replace")
import content_agent as ca   # noqa: E402  (module is swapped by the mutation harness)
import test_order_sheet as T   # noqa: E402  (reuses its small results/approved fixtures)

BRAND = "Acme"
FACT_ROWS = [("F1", "product_type", "Cat6 RJ45 inline coupler", "page_read"), ("F2", "connector_a", "RJ45 female", "page_read"),
             ("F3", "connector_b", "RJ45 female", "page_read"), ("F4", "shielding", "shielded", "sample_inspected"),
             ("F5", "category_rating", "Cat6", "page_read"), ("F6", "pack_contents", "pack of 2", "sample_inspected"),
             ("F7", "mounting", "panel mount", "listing_title"), ("F8", "certifications", "RoHS", "unverified"),
             ("F9", "material", "brass", "unverified")]
SENTENCES = [("This coupler joins two network cables so they work as one longer run.", ["F1"]),
             ("Plug a cable into each end and the connection is made without any tools.", ["F1", "F2"]),
             ("Both ends are female so each end takes a plug from a patch cable.", ["F2", "F3"]),
             ("The body is shielded which helps keep stray noise away from the signal.", ["F4"]),
             ("It is rated for Cat6 networks and suits everyday office and home wiring.", ["F5"]),
             ("Each pack holds a pair so you can finish two links in one order.", ["F6"]),
             ("It sits in the middle of a cable run and keeps the path tidy and short.", ["F1"]),
             ("Use it to extend a run when one cable is not long enough for the job.", ["F1"]),
             ("The simple design has no moving parts to wear out or to adjust over time.", ["F1"]),
             ("Both connectors line up so cables seat firmly and stay put once pushed in.", ["F2", "F3"]),
             ("It suits desks, wall plates, cabinets and other places where cabling meets.", ["F1"]),
             ("Check the plug on each cable before you buy so that it fits this coupler.", ["F2", "F3"]),
             ("A shielded body is useful where cables run close to other equipment.", ["F4"]),
             ("It is a plain, practical part for people who build and repair networks.", ["F1"]),
             ("Keep a spare pair in the drawer for the day a run turns out too short.", ["F6"])]
TITLE = "Acme Cat6 RJ45 Shielded Inline Coupler, 2 Pack"
BULLETS = [("Cat6 rated for fast network links", ["F5"]), ("Shielded body helps reduce interference", ["F4"]),
           ("RJ45 female ports on both ends", ["F2", "F3"])]


def good_output(**over):
    out = {"title": TITLE, "short_description": "A shielded Cat6 RJ45 inline coupler that joins two cables. Pack of 2.",
           "long_description": " ".join(s for s, _ in SENTENCES), "key_features": [b for b, _ in BULLETS],
           "claims": [{"text": TITLE, "fact_ids": ["F1", "F5"]}]
           + [{"text": s, "fact_ids": ids} for s, ids in SENTENCES] + [{"text": b, "fact_ids": ids} for b, ids in BULLETS]}
    out.update(over)
    return out


def with_text(**over):
    """good_output with its text changed AND a claim added for the new text, so only the targeted rule can fire."""
    out = good_output()
    for k, v in over.items():
        out[k] = v
    return out


class Stub:
    """A scripted model: each call returns the next scripted reply; records every message list it was sent."""
    def __init__(self, *replies):
        self.replies, self.sent = list(replies), []

    def __call__(self, system, messages, max_tokens):
        self.sent.append(copy.deepcopy(messages))
        r = self.replies.pop(0) if len(self.replies) > 1 else self.replies[0]      # the last reply repeats, like a stubborn model
        if isinstance(r, Exception):
            raise r
        return (r if isinstance(r, str) else json.dumps(r)), 1000, 500


def write_facts(path, rows, sku="AAA", extra=()):
    with open(path, "w", newline="", encoding="utf-8") as f:
        w = csv.writer(f)
        w.writerow(ca.FACT_COLUMNS)
        for fid, field, value, status in list(rows) + list(extra):
            w.writerow([sku, fid, field, value, status, "test"])


class Env:
    def __init__(self, facts=FACT_ROWS, extra=(), brand=BRAND):
        self.box = T.Box(T.basic_rows(), [{}])
        self.facts = os.path.join(self.box.dir, "product_facts.csv")
        write_facts(self.facts, facts, extra=extra)
        self.old = os.environ.get("BRAND_NAME")
        if brand:
            os.environ["BRAND_NAME"] = brand
        else:
            os.environ.pop("BRAND_NAME", None)

    def close(self):
        if self.old is None:
            os.environ.pop("BRAND_NAME", None)
        else:
            os.environ["BRAND_NAME"] = self.old

    def generate(self, *replies, allow=False, dry=False, stub=None):
        stub = stub or Stub(*replies)
        with contextlib.redirect_stdout(io.StringIO()) as out:
            res = ca.generate(self.box.approved, self.box.results, self.facts, None, allow, dry, stub,
                              os.path.join(self.box.dir, "out"), T.NOW, self.box.excl)
        res["stub"], res["stdout"] = stub, out.getvalue()
        return res


def run(*replies, **kw):
    env = Env(**{k: kw.pop(k) for k in ("facts", "extra", "brand") if k in kw})
    try:
        return env.generate(*replies, **kw), env
    finally:
        env.close()


def reasons(item, rule=None):
    return [(r, f, w) for r, f, t, w in item["failures"] if rule in (None, r)]


def has(item, rule, word):
    return any(rule == r and word in w for r, _, _, w in [(a, b, c, d) for a, b, c, d in item["failures"]])


def blocked_with(out, rule, word):
    it = out["items"][0]
    assert it["status"] == "BLOCKED" and has(it, rule, word), it["failures"]


# ---------- happy path ----------

def test_a_good_output_passes_every_validator_and_builds_the_workbook():
    out, env = run(good_output())
    it = out["items"][0]
    assert it["status"] == "READY FOR REVIEW" and it["failures"] == [], it["failures"]
    assert out["calls"] == 1 and out["cost"] == Decimal("0.0035")                            # 1000 in + 500 out at $1/$5 per Mtok
    wb = openpyxl.load_workbook(out["path"])
    assert wb.sheetnames == ["Listings", "Facts used", "Review notes", "To do before upload", "Check data"]
    ws = wb["Listings"]
    assert "Draft content for human review. Not published." in ws["A1"].value
    head = [c.value for c in ws[4]]
    assert head[:6] == ["SKU", "Status", "Brand", "Title", "Short description", "Long description"]
    row = {h: ws.cell(5, i + 1).value for i, h in enumerate(head)}
    assert row["Status"] == "READY FOR REVIEW" and row["Brand"] == BRAND and row["Title"] == TITLE
    assert row["Selling price"] == "PRICE NEEDED" and row["UPC"] == "UPC NEEDED (GS1)"
    assert row["Title length"] == "=LEN(D5)" and row["Long description words"].startswith("=IF(TRIM(F5)")
    assert row["Key feature 1 length"] == '=IF(G5="","",LEN(G5))' and row["Key feature 4 length"] == '=IF(J5="","",LEN(J5))'
    assert len(ws.conditional_formatting) >= 4
    assert os.path.exists(out["md"]) and "Not published" in open(out["md"], encoding="utf-8").read()
    facts_sheet = [r for r in wb["Facts used"].iter_rows(min_row=2, values_only=True)]
    assert {r[1] for r in facts_sheet} == {"F1", "F2", "F3", "F4", "F5", "F6"}                # unverified + listing_title kept out
    assert any(r[1] == "F1" and TITLE in r[6] for r in facts_sheet)
    todo = " ".join(str(r[1]) for r in wb["To do before upload"].iter_rows(min_row=2, values_only=True))
    for word in ("photos", "item-spec", "UPC", "price", "Brand registration", "reads every word"):
        assert word in todo, word


def test_the_model_is_shown_only_usable_facts_and_nothing_about_cost_or_sellers():
    out, env = run(good_output())
    sent = out["stub"].sent[0][0]["content"]
    for secret in ("Seller A", "Seller B", "alibaba", "aliexpress", "L-Com", "l-com", "0.90", "ECF", "brass", "RoHS", "panel mount", "margin"):
        assert secret not in sent, secret
    assert "RJ45 female" in sent and '"fact_id": "F4"' in sent and BRAND in sent
    assert "F8" not in sent and "F9" not in sent and "F7" not in sent                          # unverified / listing_title


# ---------- 1. limits ----------

def test_title_over_the_limit_fails():
    long_title = "Acme Cat6 RJ45 Shielded Inline Coupler " + "Cat6 " * 30
    out, _ = run(good_output(title=long_title))
    assert has(out["items"][0], "limits", "over the 150 limit")


def test_all_caps_word_and_special_characters_in_the_title_fail():
    out, _ = run(good_output(title="Acme Cat6 RJ45 COUPLER Shielded Inline, 2 Pack"))
    assert has(out["items"][0], "limits", "ALL CAPS") and not has(out["items"][0], "limits", "RJ45")
    out, _ = run(good_output(title="Acme Cat6 RJ45 Shielded Inline Coupler! 2 Pack"))
    assert has(out["items"][0], "limits", "special characters")
    assert not has(run(good_output())[0]["items"][0], "limits", "ALL CAPS")


def test_subjective_title_words_fail():
    out, _ = run(good_output(title="Acme Premium Cat6 RJ45 Shielded Inline Coupler"))
    assert has(out["items"][0], "limits", "subjective")


def test_html_and_markdown_fail_plain_text_rule():
    out, _ = run(good_output(long_description=good_output()["long_description"] + " <b>Great</b>"))
    assert has(out["items"][0], "limits", "no HTML")
    out, _ = run(good_output(short_description="**Shielded** Cat6 RJ45 coupler"))
    assert has(out["items"][0], "limits", "markdown")


def test_description_under_150_words_fails():
    short = " ".join(s for s, _ in SENTENCES[:5])
    out, _ = run(good_output(long_description=short))
    assert has(out["items"][0], "limits", "under the 150 minimum")


def test_bullet_over_80_characters_and_wrong_count_fail():
    big = "Cat6 rated for fast network links " * 3
    out, _ = run(good_output(key_features=[big] + [b for b, _ in BULLETS[1:]]))
    assert has(out["items"][0], "limits", "over the 80 limit")
    out, _ = run(good_output(key_features=[b for b, _ in BULLETS[:2]]))
    assert has(out["items"][0], "limits", "need 3-10")


def test_short_description_and_long_description_caps():
    out, _ = run(good_output(short_description="Shielded Cat6 RJ45 coupler. " * 30))
    assert has(out["items"][0], "limits", "over the 500 limit")
    out, _ = run(good_output(long_description=good_output()["long_description"] * 6))
    assert has(out["items"][0], "limits", "over the 4000 limit")


# ---------- 2. evidence ----------

def test_claim_without_fact_ids_and_unknown_fact_id_fail():
    g = good_output()
    g["claims"][1] = {"text": g["claims"][1]["text"], "fact_ids": []}
    out, _ = run(g)
    assert has(out["items"][0], "evidence", "no fact_ids")
    g = good_output()
    g["claims"][1]["fact_ids"] = ["F99"]
    out, _ = run(g)
    assert has(out["items"][0], "evidence", "F99 does not exist")


def test_unverified_fact_is_never_usable_and_listing_title_needs_the_flag():
    g = good_output()
    g["claims"][2]["fact_ids"] = ["F8"]
    out, _ = run(g)
    assert has(out["items"][0], "evidence", "unverified fact is never usable")
    g = good_output()
    g["claims"][2]["fact_ids"] = ["F7"]
    out, _ = run(g)
    assert has(out["items"][0], "evidence", "needs --allow-listing-facts")
    out, env = run(g, allow=True)                                                              # with the flag it passes...
    assert out["items"][0]["status"] == "READY FOR REVIEW"
    wb = openpyxl.load_workbook(out["path"])                                                   # ...and is marked in the review sheet
    marked = [r for r in wb["Facts used"].iter_rows(min_row=2, values_only=True) if r[1] == "F7"]
    assert marked and "LISTING-TITLE FACT" in marked[0][7]


def test_every_sentence_bullet_and_title_needs_a_claim():
    g = good_output(long_description=good_output()["long_description"] + " This part also fits every cable in the world.")
    out, _ = run(g)
    assert has(out["items"][0], "evidence", "not covered by any claim")
    g = good_output()
    g["claims"] = [c for c in g["claims"] if c["text"] != TITLE]
    out, _ = run(g)
    assert any(f == "title" for _, f, _ in reasons(out["items"][0], "evidence"))


# ---------- 3. unsupported spec tokens ----------

def test_invented_cat6a_fails_but_spelling_variants_of_cat6_pass():
    bad = good_output(title="Acme Cat6a RJ45 Shielded Inline Coupler, 2 Pack")
    bad["claims"][0]["text"] = bad["title"]
    out, _ = run(bad)
    assert has(out["items"][0], "spec_token", "'cat6a'")
    for variant in ("Cat.6", "Cat 6", "CAT-6", "cat6"):
        g = good_output(title=f"Acme {variant} RJ45 Shielded Inline Coupler, 2 Pack")
        g["claims"][0]["text"] = g["title"]
        g["title"] = g["title"]
        out, _ = run(g)
        assert not [x for x in reasons(out["items"][0], "spec_token")], (variant, out["items"][0]["failures"])
    out, _ = run(good_output(), facts=[("F1", "product_type", "Cat.6 RJ-45 inline coupler", "page_read")] + [r for r in FACT_ROWS[1:] if r[0] not in ("F5",)])
    assert not reasons(out["items"][0], "spec_token"), out["items"][0]["failures"]          # facts spelled 'Cat.6 RJ-45', text says 'Cat6 RJ45'


def test_invented_certification_ip_rating_unit_and_standards_fail():
    for word, text in (("cert:ul", "It is UL listed."), ("ip67", "It is rated IP67."), ("10ft", "It is 10 ft long."),
                       ("4k", "It carries 4K video."), ("hdmi2.0", "It works with HDMI 2.0."), ("gold", "The contacts are gold."),
                       ("waterproof", "It is waterproof."), ("qty:5", "Sold as a pack of 5."), ("qty:10", "Includes 10 pcs.")):
        g = good_output(short_description=text)
        out, _ = run(g)
        assert has(out["items"][0], "spec_token", f"'{word}'"), (word, out["items"][0]["failures"])


def test_connector_and_standard_names_must_be_in_the_facts():
    for word, text in (("dvi", "Fits DVI plugs."), ("usb", "A USB adapter."), ("bnc", "BNC style."), ("sma", "SMA parts."),
                       ("ntype", "N-Type parts."), ("db9", "A DB-9 plug."), ("8p8c", "An 8P8C jack."), ("hdmi", "HDMI ready.")):
        out, _ = run(good_output(short_description=text))
        assert has(out["items"][0], "spec_token", f"'{word}'"), (word, out["items"][0]["failures"])


def test_lc_sc_st_fc_count_only_when_standalone_and_uppercase():
    for text in ("Fits LC ports.", "Fits SC ports.", "An ST plug.", "An FC plug."):
        out, _ = run(good_output(short_description=text))
        assert reasons(out["items"][0], "spec_token"), text
    for text in ("The lc of it.", "It is sc6 sized.", "Cast stock fcs."):                          # lower case or inside a word: not a token
        out, _ = run(good_output(short_description=text))
        assert not reasons(out["items"][0], "spec_token"), text
    out, _ = run(good_output(short_description="Fits LC ports."), facts=FACT_ROWS + [("F20", "connector_a", "LC duplex", "page_read")])
    assert not reasons(out["items"][0], "spec_token")


def test_tokens_match_whole_tokens_in_the_facts_ignoring_case():
    out, _ = run(good_output(short_description="A SHIELDED inline coupler with RJ-45 ports, female."))
    assert not reasons(out["items"][0], "spec_token"), out["items"][0]["failures"]
    out, _ = run(good_output(short_description="An unshielded coupler."))                          # 'shielded' in facts is not 'unshielded'
    assert has(out["items"][0], "spec_token", "'unshielded'")
    out, _ = run(good_output(short_description="A male plug."))                                    # 'female' in facts is not 'male'
    assert has(out["items"][0], "spec_token", "'male'")
    out, _ = run(good_output(short_description="A small, golden finish for the usual busy office, in the form of a resale tool."))
    assert not reasons(out["items"][0], "spec_token"), out["items"][0]["failures"]                # sma/gold/usb/male inside other words


# ---------- 4. forbidden terms ----------

def test_lcom_and_its_variants_are_forbidden():
    for text in ("Same as L-Com.", "Like an L Com part.", "A LCOM part.", "An l-com part.", "From L-COM."):
        out, _ = run(good_output(short_description=text))
        assert has(out["items"][0], "forbidden", "L-Com"), text


def test_sku_supplier_platform_and_competitor_names_are_forbidden():
    for text, label in (("Part AAA fits.", "SKU"), ("Part A-A-A fits.", "SKU"), ("Made by Seller C.", "supplier"),
                        ("Found on Alibaba.", "platform"), ("Seen on AliExpress.", "platform"), ("From Made-in-China.", "platform"),
                        ("Like Monoprice.", "competitor"), ("Better than Cables To Go.", "competitor"), ("A StarTech style.", "competitor")):
        out, _ = run(good_output(short_description=text))
        assert has(out["items"][0], "forbidden", label), (text, out["items"][0]["failures"])


def test_promotional_phrases_are_forbidden():
    for text in ("Free shipping today.", "The best coupler.", "A cheap coupler.", "Guaranteed to fit.", "Genuine part.", "The original design.",
                 "Lifetime use.", "OEM quality.", "Top rated.", "Number #1 choice."):
        out, _ = run(good_output(short_description=text))
        assert has(out["items"][0], "forbidden", "forbidden"), text
    out, _ = run(good_output(short_description="Originally made for offices."))                   # 'original' inside another word is fine
    assert not reasons(out["items"][0], "forbidden")


# ---------- 5. brand ----------

def test_brand_unset_refuses_to_generate_but_dry_run_and_check_work():
    env = Env(brand=None)
    try:
        try:
            env.generate(good_output())
        except ca.ContentError as e:
            assert "BRAND_NAME" in str(e)
        else:
            raise AssertionError("expected a refusal")
        dry = env.generate(dry=True)
        assert ca.BRAND_PLACEHOLDER in dry["stdout"] and dry["stub"].sent == []
    finally:
        env.close()
    out, env2 = run(good_output())
    os.environ.pop("BRAND_NAME", None)
    try:
        res = ca.check_workbook(out["path"])
        assert any("BRAND_NAME is not set" in w for w in res["warnings"]) and res["results"]["AAA"] == []
    finally:
        os.environ["BRAND_NAME"] = BRAND


def test_brand_twice_in_title_or_equal_to_a_competitor_fails():
    out, _ = run(good_output(title="Acme Cat6 RJ45 Shielded Acme Inline Coupler"))
    assert has(out["items"][0], "brand", "more than once")
    out, _ = run(good_output(), brand="StarTech")
    assert has(out["items"][0], "brand", "forbidden name")
    out, _ = run(good_output(), brand="L-Com")
    assert has(out["items"][0], "brand", "forbidden name")
    long_brand = "Acme" * 20
    env = Env(brand=long_brand)
    try:
        env.generate(good_output())
    except ca.ContentError as e:
        assert "limit" in str(e)
    else:
        raise AssertionError("expected a refusal")
    finally:
        env.close()


# ---------- 6. price and cost ----------

def test_dollar_amounts_and_cost_words_fail():
    for text in ("Only $5 each.", "Priced at 5 USD.", "A few dollars.", "Costs cents."):
        out, _ = run(good_output(short_description=text))
        assert has(out["items"][0], "price", "never states money"), text


def test_price_and_upc_placeholders_and_values_come_only_from_the_facts_file():
    out, _ = run(good_output())
    ws = openpyxl.load_workbook(out["path"])["Listings"]
    assert (ws.cell(5, ca.C_PRICE).value, ws.cell(5, ca.C_UPC).value) == ("PRICE NEEDED", "UPC NEEDED (GS1)")
    out, _ = run(good_output(), extra=[("F30", "selling_price", "24.99", "page_read"), ("F31", "upc", "036000291452", "page_read")])
    ws = openpyxl.load_workbook(out["path"])["Listings"]
    assert ws.cell(5, ca.C_PRICE).value == 24.99 and ws.cell(5, ca.C_UPC).value == "036000291452" and out["items"][0]["status"] == "READY FOR REVIEW"
    assert "24.99" not in out["stub"].sent[0][0]["content"] and "036000291452" not in out["stub"].sent[0][0]["content"]


# ---------- 7. UPC ----------

def test_upc_check_digit_is_validated():
    assert ca.upc_ok("036000291452") and not ca.upc_ok("036000291453") and not ca.upc_ok("12345") and not ca.upc_ok("03600029145a")
    out, _ = run(good_output(), extra=[("F31", "upc", "036000291453", "page_read")])
    assert out["items"][0]["status"] == "BLOCKED" and has(out["items"][0], "upc", "check digit")
    out, _ = run(good_output(), extra=[("F30", "selling_price", "free", "page_read")])
    assert has(out["items"][0], "price", "positive amount")


# ---------- retries, blocking, billing ----------

def test_invalid_json_then_a_valid_retry_succeeds_and_is_costed_twice():
    out, _ = run("```json\n{}\n```", good_output())
    assert out["items"][0]["status"] == "READY FOR REVIEW" and out["calls"] == 2
    assert "not the required JSON" in out["stub"].sent[1][-1]["content"]


def test_invalid_json_twice_is_blocked():
    out, _ = run("not json", '{"title": "x"}')
    it = out["items"][0]
    assert it["status"] == "BLOCKED" and has(it, "json", "not valid JSON twice") and out["calls"] == 2


def test_a_repair_that_fixes_the_problem_passes_and_gets_the_failures_listed():
    bad = good_output(short_description="Only $5 each.")
    out, _ = run(bad, good_output())
    assert out["items"][0]["status"] == "READY FOR REVIEW" and out["calls"] == 2
    assert "[price]" in out["stub"].sent[1][-1]["content"]


def test_a_repair_that_still_fails_is_blocked_with_the_failing_text_kept_and_no_third_call():
    bad = good_output(short_description="Only $5 each.")
    out, _ = run(bad, bad)
    it = out["items"][0]
    assert it["status"] == "BLOCKED" and has(it, "price", "never states money") and out["calls"] == 2
    assert it["listing"]["short_description"] == "Only $5 each."                                   # we never edit the model's text
    ws = openpyxl.load_workbook(out["path"])["Listings"]
    assert ws.cell(5, 2).value == "BLOCKED" and ws.cell(5, 5).value == "Only $5 each."


def test_a_skus_with_no_usable_facts_is_blocked_without_calling_the_model():
    out, _ = run(facts=[("F1", "product_type", "Cat6 coupler", "unverified")])
    assert out["calls"] == 0 and out["items"][0]["status"] == "BLOCKED" and has(out["items"][0], "evidence", "no usable facts")


class FakeApiError(Exception):
    def __init__(self, status_code, message=""):
        super().__init__(message)
        self.status_code, self.message = status_code, message


def test_billing_errors_stop_the_batch_and_save_what_finished():
    assert ca.is_billing_error(FakeApiError(402)) and ca.is_billing_error(FakeApiError(401)) and ca.is_billing_error(FakeApiError(403))
    assert ca.is_billing_error(FakeApiError(400, "Your credit balance is too low")) and not ca.is_billing_error(FakeApiError(400, "bad field"))
    assert not ca.is_billing_error(FakeApiError(500, "credit"))
    env = Env()
    box = T.Box(T.basic_rows() + [dict(sku="BBB", keyword="K2", rec=("S", T.URL_B), cands=[T.cand("S", T.URL_B, "$1", 1, "10")])],
                [{}, {"sku": "BBB", "listing_url": T.URL_B}])
    write_facts(env.facts, FACT_ROWS)
    with open(env.facts, "a", newline="", encoding="utf-8") as f:
        csv.writer(f).writerows([["BBB", r[0], r[1], r[2], r[3], "t"] for r in FACT_ROWS])
    try:
        with contextlib.redirect_stdout(io.StringIO()):
            out = ca.generate(box.approved, box.results, env.facts, None, False, False, Stub(good_output(), FakeApiError(402, "credit balance")),
                              os.path.join(box.dir, "out"), T.NOW, box.excl)
        st = {it["sku"]: it["status"] for it in out["items"]}
        assert st == {"AAA": "READY FOR REVIEW", "BBB": "NOT RUN"} and os.path.exists(out["path"])
    finally:
        env.close()


def test_cost_is_logged_per_call_and_totalled_and_the_key_is_never_printed():
    os.environ["ANTHROPIC_API_KEY"] = "sk-ant-SECRET-123"
    try:
        out, _ = run(good_output())
        assert "sk-ant-SECRET-123" not in out["stdout"] and "[cost] AAA" in out["stdout"] and "[cost] total $0.00350 for 1 call(s)" in out["stdout"]
        assert ca.call_cost(1_000_000, 1_000_000) == Decimal("6.00")
    finally:
        os.environ.pop("ANTHROPIC_API_KEY", None)
    try:
        ca.live_model()
    except ca.ContentError as e:
        assert "ANTHROPIC_API_KEY" in str(e)
    else:
        raise AssertionError("expected a refusal")


# ---------- dry run, facts template, check ----------

def test_dry_run_prints_the_prompt_and_calls_nothing_and_writes_nothing():
    out, env = run(dry=True)
    assert out["stub"].sent == [] and "usable_facts" in out["stdout"] and "RJ45 female" in out["stdout"]
    assert not os.path.exists(os.path.join(env.box.dir, "out"))


def test_facts_template_writes_a_starter_and_never_overwrites():
    box = T.Box(T.basic_rows(), [{}])
    path = os.path.join(box.dir, "facts.csv")
    got, n, notes = ca.facts_template(box.approved, box.results, path, T.NOW, box.excl)
    rows = list(csv.DictReader(open(path, encoding="utf-8-sig")))
    assert n == 1 and {r["sku"] for r in rows} == {"AAA"}
    by_field = {}
    for r in rows:
        by_field.setdefault(r["field"], []).append(r)
    assert by_field["product_type"][0]["status"] == "unverified" and by_field["product_type"][0]["value"] == "Test cable"
    assert all(r["status"] == "unverified" for r in by_field["spec_line"])                          # the Ordering Note lines
    assert all(r["value"] == "" and r["status"] == "" for f in ("connector_a", "selling_price", "upc") for r in by_field[f])
    assert {f for f in ca.TEMPLATE_FIELDS} <= set(by_field)
    before = open(path, "rb").read()
    try:
        ca.facts_template(box.approved, box.results, path, T.NOW, box.excl)
    except ca.ContentError as e:
        assert "already exists" in str(e)
    else:
        raise AssertionError("expected a refusal")
    assert open(path, "rb").read() == before
    filled = ca.read_facts(path)
    assert filled and all(r["value"] and r["status"] in ca.ALL_STATUSES for r in filled)               # blank template rows are skipped on read


def test_the_facts_file_is_validated():
    box = T.Box(T.basic_rows(), [{}])
    p = os.path.join(box.dir, "f.csv")
    write_facts(p, [("F1", "x", "v", "maybe")])
    for bad, word in ((lambda: ca.read_facts(p), "status"), (lambda: ca.read_facts(os.path.join(box.dir, "nope.csv")), "not found")):
        try:
            bad()
        except ca.ContentError as e:
            assert word in str(e)
        else:
            raise AssertionError(word)
    write_facts(p, [("F1", "x", "v", "page_read"), ("F1", "y", "w", "page_read")])
    try:
        ca.read_facts(p)
    except ca.ContentError as e:
        assert "duplicate" in str(e)
    else:
        raise AssertionError("expected duplicate refusal")


def edit(path, **cells):
    wb = openpyxl.load_workbook(path)
    ws = wb["Listings"]
    head = {c.value: c.column for c in ws[4]}
    for name, value in cells.items():
        ws.cell(5, head[name.replace("_", " ")], value)
    wb.save(path)


def test_check_passes_a_clean_file_and_catches_human_edits():
    out, env = run(good_output())
    assert ca.check_workbook(out["path"])["results"] == {"AAA": []}
    cases = [(dict(Title="Acme Cat6a RJ45 Shielded Inline Coupler, 2 Pack"), "spec_token", "cat6a"),
             (dict(Short_description="Cheapest coupler anywhere."), "forbidden", "promotional"),
             (dict(Short_description="Same as L-Com part."), "forbidden", "L-Com"),
             (dict(Short_description="Only $4."), "price", "money"),
             (dict(Long_description="Too short now."), "limits", "150"),
             (dict(Key_feature_1="x" * 90), "limits", "80"),
             (dict(Title="Acme Cat6 RJ45 Shielded Inline Coupler <b>"), "limits", "HTML"),
             (dict(Brand="Other"), "brand", "differs"),
             (dict(UPC="036000291453"), "upc", "check digit"),
             (dict(Selling_price="lots"), "price", "positive")]
    for cells, rule, word in cases:
        import shutil
        copy_path = out["path"] + f".{rule}.xlsx"
        shutil.copy(out["path"], copy_path)
        edit(copy_path, **cells)
        res = ca.check_workbook(copy_path)["results"]["AAA"]
        assert any(r == rule and word.lower() in w.lower() for r, _, _, w in res), (cells, res)


def test_check_cli_exits_nonzero_on_failures_and_prints_them():
    out, env = run(good_output())
    edit(out["path"], Short_description="Free shipping here.")
    buf, code = io.StringIO(), 0
    with contextlib.redirect_stdout(buf):
        try:
            ca.main(["check", out["path"]])
        except SystemExit as e:
            code = e.code
    assert code == 1 and "AAA: 1 problem(s)" in buf.getvalue() and "forbidden" in buf.getvalue()


def test_check_refuses_a_file_that_generate_did_not_make():
    p = os.path.join(tempfile.mkdtemp(), "x.xlsx")
    openpyxl.Workbook().save(p)
    try:
        ca.check_workbook(p)
    except ca.ContentError as e:
        assert "not made by this tool" in str(e)
    else:
        raise AssertionError("expected a refusal")


def test_expired_excluded_and_unmatched_products_are_skipped_not_written():
    env = Env()
    try:
        box = T.Box(T.basic_rows(), [{"approved_at": "2026-09-20T10:00:00Z", "expires_at": "2026-09-27T10:00:00Z"}])
        try:
            ca.generate(box.approved, box.results, env.facts, None, False, True, None, None, T.NOW, box.excl)
        except ca.ContentError as e:
            assert "expired" in str(e)
        else:
            raise AssertionError("expected a refusal")
        box = T.Box(T.basic_rows(), [{"listing_url": "https://www.alibaba.com/product-detail/Other_9.html"}])
        try:
            ca.generate(box.approved, box.results, env.facts, None, False, True, None, None, T.NOW, box.excl)
        except ca.ContentError as e:
            assert "not found in the results file" in str(e)
        else:
            raise AssertionError("expected a refusal")
        box = T.Box(T.basic_rows(), [{}], excl=[["AAA", "Seller A", "wrong part"]])
        try:
            ca.generate(box.approved, box.results, env.facts, None, False, True, None, None, T.NOW, box.excl)
        except ca.ContentError as e:
            assert "excluded by the reviewer" in str(e)
        else:
            raise AssertionError("expected a refusal")
    finally:
        env.close()


def test_the_script_never_touches_walmart_the_inputs_or_other_agents():
    box = T.Box(T.basic_rows(), [{}])
    write_facts(os.path.join(box.dir, "product_facts.csv"), FACT_ROWS)
    before = {p: open(p, "rb").read() for p in (box.approved, box.results, box.state, box.excl)}
    os.environ["BRAND_NAME"] = BRAND
    try:
        with contextlib.redirect_stdout(io.StringIO()):
            ca.generate(box.approved, box.results, os.path.join(box.dir, "product_facts.csv"), None, False, False, Stub(good_output()),
                        os.path.join(box.dir, "out"), T.NOW, box.excl)
    finally:
        os.environ.pop("BRAND_NAME", None)
    assert before == {p: open(p, "rb").read() for p in before}
    src = open(os.path.join(HERE, "content_agent.py"), encoding="utf-8").read()
    for name in ("requests", "httpx", "urllib", "smtplib", "socket", "subprocess", "walmart.com", "marketplace.walmartapis"):
        assert name not in src.replace("Walmart Marketplace", "").replace("walmart_listings", ""), name
    assert "float(" in src and "ANTHROPIC_API_KEY" in src and "print(key" not in src


def test_the_real_approved_orders_file_if_there_is_one():
    path = os.path.join(HERE, "approved_orders.csv")
    if not os.path.exists(path):
        print("  (no real approved_orders.csv here; covered by the sandbox dry run)")
        return
    results = sorted(p for p in os.listdir(os.path.join(HERE, "Excel Output Sheets")) if p.startswith("sourcing_results_"))
    assert results and ca.read_approved(path)[0]
    tmp = tempfile.mkdtemp()
    got, n, _ = ca.facts_template(path, os.path.join(HERE, "Excel Output Sheets", results[-1]), os.path.join(tmp, "facts.csv"))
    assert n >= 1


def test_a_run_whose_worst_case_cost_is_over_the_limit_warns_before_calling():
    old = ca.COST_WARN
    ca.COST_WARN = Decimal("0.0001")
    try:
        out, _ = run(good_output())
    finally:
        ca.COST_WARN = old
    assert out["stdout"].startswith("WARNING: worst-case cost") and out["calls"] == 1
    assert "WARNING: worst-case" not in run(good_output())[0]["stdout"]


def test_hdmi_version_in_the_facts_supports_plain_hdmi_but_not_the_reverse():
    facts = [("F1", "product_type", "HDMI 2.0 cable", "page_read")]
    only = good_output(short_description="An HDMI cable.")
    out, _ = run(only, facts=facts)
    assert not [x for x in reasons(out["items"][0], "spec_token") if "hdmi" in x[2]]
    out, _ = run(good_output(short_description="An HDMI 2.0 cable."), facts=[("F1", "product_type", "HDMI cable", "page_read")])
    assert has(out["items"][0], "spec_token", "hdmi2.0")


def test_supplier_names_do_not_block_everyday_words():
    assert ca.supplier_terms("Unknown (AliExpress listing)") == [] and ca.supplier_terms("Unknown") == []
    cable = {t.lower() for t in ca.supplier_terms("PCM Cable Co., Ltd")}
    assert "cable" not in cable and "pcm cable" in cable
    assert "farsince" in {t.lower() for t in ca.supplier_terms("Ningbo Farsince Electronic Co., Ltd.")}
    assert "kronz" in {t.lower() for t in ca.supplier_terms("KRONZ (Guangzhou) Electronics Co., Ltd")}
    forb = ca.forbidden_list([], ["Premier Cable Co., Limited"])
    assert not [1 for _, rx in forb if rx.search("A shielded network cable coupler with a premier finish")]
    assert [1 for _, rx in forb if rx.search("Made by Premier Cable Co.")]


def test_check_reports_a_blocked_empty_row_plainly():
    out, _ = run(facts=[("F1", "product_type", "Cat6 coupler", "unverified")])
    res = ca.check_workbook(out["path"])["results"]["AAA"]
    assert [r for r, _, _, w in res] == ["run"] and "no content to check" in res[0][3]


def test_conditional_format_fill_sets_both_fg_and_bg_colour_in_the_styles_xml():
    import re as _re
    import zipfile
    out, _ = run(good_output())
    styles = zipfile.ZipFile(out["path"]).read("xl/styles.xml").decode("utf-8")
    dxfs = _re.search(r"<dxfs.*?</dxfs>", styles, _re.S)[0]
    assert dxfs.count('<fgColor rgb="FFFFC7CE"') >= 1 and dxfs.count('<bgColor rgb="FFFFC7CE"') >= 1, dxfs
    assert dxfs.count('<fgColor rgb="FFFFC7CE"') == dxfs.count('<bgColor rgb="FFFFC7CE"')
    assert ca.RED_FILL.fgColor.rgb == ca.RED_FILL.bgColor.rgb == "FFFFC7CE"


def review_warnings(out):
    ws = openpyxl.load_workbook(out["path"])["Review notes"]
    return [r for r in ws.iter_rows(min_row=2, values_only=True) if r[1] == "WARNING"]


def test_title_starting_with_the_brand_gives_no_warning_ignoring_case():
    out, _ = run(good_output())
    assert out["items"][0]["warnings"] == [] and review_warnings(out) == []
    g = good_output(title=TITLE.replace("Acme", "ACME"))
    g["claims"][0]["text"] = g["title"]
    out, _ = run(g)
    assert out["items"][0]["warnings"] == []


def test_title_without_the_brand_or_with_another_brand_warns_but_never_blocks():
    for title in ("Cat6 RJ45 Shielded Inline Coupler, 2 Pack Acme", "Zenith Cat6 RJ45 Shielded Inline Coupler, 2 Pack",
                  "Acmeco Cat6 RJ45 Shielded Inline Coupler, 2 Pack"):
        g = good_output(title=title)
        g["claims"][0]["text"] = title
        out, _ = run(g)
        it = out["items"][0]
        assert it["status"] == "READY FOR REVIEW" and it["failures"] == [], (title, it["failures"])      # a warning, not a block
        rows = review_warnings(out)
        assert len(rows) == 1 and rows[0][0] == "AAA" and "does not start with the brand 'Acme'" in rows[0][4], rows


def test_title_brand_check_is_skipped_when_brand_is_unset():
    assert ca.title_brand_warning("Zenith coupler", None) == [] and ca.title_brand_warning("Zenith coupler", "") == []
    assert ca.title_brand_warning("", "Acme") == []
    os.environ.pop("BRAND_NAME", None)
    out, _ = run(good_output(), brand=None, dry=True)                           # dry run works without a brand and has no warnings
    assert out["stub"].sent == []


if __name__ == "__main__":
    for name, fn in sorted(globals().items()):
        if name.startswith("test_"):
            fn()
    print("content agent tests passed")
