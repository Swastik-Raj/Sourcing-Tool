"""Offline checks for email_agent.py - no real email, no Anthropic or Nimble call. Run: python test_email_agent.py

Summarization is stubbed (keyword rules), so these tests prove the plumbing - which SKU and listing a reply is
filed under, what the report shows, that UTF-8 survives - not the quality of a real LLM summary."""
import contextlib
import csv
import glob
import io
import json
import os
import re
import shutil
import subprocess
import sys
import tempfile

import openpyxl
from dotenv import dotenv_values

import email_agent as ea

HERE = os.path.dirname(os.path.abspath(__file__))
FIXTURES = os.path.join(HERE, "test_fixtures", "email_replies")
HEADER = ["SKU", "Product", "Keyword", "L-Com Unit Price", "L-Com Price Source", "L-Com Price Date", "Recommendation",
          "Recommended Manufacturer", "Recommended Email", "Recommended URL", "Recommended Unit Price", "Ordering Note",
          "Manufacturer 1 URL", "Manufacturer 1 MOQ", "Manufacturer 2 URL", "Manufacturer 2 MOQ"]
TITLE_HEADER = HEADER + ["Manufacturer 1 Listing Title"]
NO_REC = "No candidate qualifies - accuracy only"


def row(sku, product="Couplers", keyword=None, rec="Buy from Candidate 1", maker="Acme", email="", url=None, price=2.5,
        note="", moq="100", title=None):
    """A Results row; the recommended listing is Manufacturer 1, so its MOQ is `moq`."""
    url = url if url is not None else f"http://listing/{sku}"
    r = [sku, product, keyword or f"{sku} keyword", 20, "unverified", "", rec, maker, email, url, price, note,
         url, moq, "http://other", "999"]
    return r + [title] if title is not None else r


def make_xlsx(path, rows, header=HEADER):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Results"
    ws.append(header)
    for r in rows:
        ws.append(r + [None] * (len(header) - len(r)))
    wb.save(path)
    wb.close()


def key(sku, url=None):
    return ea.listing_key(sku, url if url is not None else f"http://listing/{sku}", "")


def drafts_from(rows, header=HEADER, exclusions=()):
    with tempfile.TemporaryDirectory() as d:
        path = os.path.join(d, "r.xlsx")
        make_xlsx(path, rows, header)
        cands, skipped = ea.load_results(path, list(exclusions))
        return [ea.make_draft(c) for c in cands], skipped


def cand(sku, email="", maker="Acme", note="", url=None):
    url = url if url is not None else f"u/{sku}"
    return {"sku": sku, "product": "p", "description": "d", "manufacturer": maker, "email": email, "url": url,
            "unit_price": None, "moq": "", "title": "", "note": note, "key": ea.listing_key(sku, url, maker)}


class FakeSMTP:
    """Stands in for smtplib.SMTP: records messages, connects to nothing."""
    sent = []
    created = 0

    def __init__(self, host, port):
        FakeSMTP.created += 1

    def __enter__(self):
        return self

    def __exit__(self, *a):
        return False

    def starttls(self):
        pass

    def login(self, user, password):
        pass

    def send_message(self, msg):
        FakeSMTP.sent.append(msg)


def write_exclusions(path, rows):
    with open(path, "w", newline="", encoding="utf-8") as f:
        w = csv.writer(f)
        w.writerow(["sku", "supplier_or_url", "reason"])
        w.writerows(rows)


def write_near_misses(path, rows):
    with open(path, "w", newline="", encoding="utf-8") as f:
        w = csv.writer(f)
        w.writerow(["sku", "supplier_or_url", "concern", "question"])
        w.writerows(rows)


@contextlib.contextmanager
def sandbox(exclusions=(("OTHER", "nobody", "placeholder row"),), near_misses=()):
    """Temp dir for state and reports, a real-looking SHIPPING_ADDRESS, an exclusions file, no .env, no API key, SMTP
    faked. Nothing leaves the process."""
    old = (ea.HERE, ea.STATE_PATH, ea.REVIEWER_EXCLUSIONS_CSV, ea.smtplib.SMTP, ea.summarize_reply, dict(os.environ), ea.NEAR_MISSES_CSV)
    tmp = tempfile.mkdtemp()
    ea.HERE, ea.STATE_PATH = tmp, os.path.join(tmp, "email_state.json")
    ea.REVIEWER_EXCLUSIONS_CSV = os.path.join(tmp, "reviewer_exclusions.csv")
    write_exclusions(ea.REVIEWER_EXCLUSIONS_CSV, exclusions)
    ea.NEAR_MISSES_CSV = os.path.join(tmp, "near_misses.csv")
    if near_misses is not None:                       # None = the file does not exist
        write_near_misses(ea.NEAR_MISSES_CSV, near_misses)
    ea.smtplib.SMTP = FakeSMTP
    os.environ.update(SMTP_HOST="smtp.invalid", SMTP_USER="u@z.example", SMTP_PASSWORD="x",
                      SHIPPING_ADDRESS="Zync Technologies\\n100 Test Way\\nPlano, TX 75024")
    os.environ.pop("ANTHROPIC_API_KEY", None)
    FakeSMTP.sent, FakeSMTP.created = [], 0

    def no_llm(text):
        raise AssertionError("the real summarizer was called")

    ea.summarize_reply = no_llm
    try:
        yield tmp
    finally:
        ea.HERE, ea.STATE_PATH, ea.REVIEWER_EXCLUSIONS_CSV, ea.smtplib.SMTP, ea.summarize_reply = old[:5]
        ea.NEAR_MISSES_CSV = old[6]
        os.environ.clear()
        os.environ.update(old[5])


def run_cli(argv, stdin_bytes=b""):
    """Runs main() in-process. stdin is a cp1252 wrapper, as a Windows console pipe is by default."""
    out = io.TextIOWrapper(io.BytesIO(), encoding="utf-8", newline="")
    old = (sys.argv, sys.stdout, sys.stdin)
    sys.argv, sys.stdout = ["email_agent.py"] + argv, out
    sys.stdin = io.TextIOWrapper(io.BytesIO(stdin_bytes), encoding="cp1252", newline="")
    try:
        ea.main()
    finally:
        sys.argv, sys.stdout, sys.stdin = old
    out.flush()
    return out.buffer.getvalue().decode("utf-8")


def cli_exit(argv):
    """Runs main() expecting it to stop with an error message; returns that message."""
    try:
        run_cli(argv)
    except SystemExit as e:
        return str(e.code)
    raise AssertionError("expected the command to stop with an error")


def state_file_exists():
    return os.path.exists(ea.STATE_PATH)


def read_state():
    return json.load(open(ea.STATE_PATH, encoding="utf-8"))


# ---------- drafting, template, fallback, skipping ----------

def test_load_and_drafts():
    rows = [row("A1", email="sales@acme.com", moq="100", url="http://b", note="model: confirm shielded"),
            row("B2", maker="", url="http://only-url"),
            row("C3", rec=NO_REC, maker=None, url=""),   # no recommendation
            row("D4", rec="No candidates found - nothing to source.", maker=None, url=""),
            row("E5", rec="Error researching this product: RuntimeError: boom - re-run with --sku", maker="Acme", url="http://e")]
    drafts, skipped = drafts_from(rows)
    assert [d["sku"] for d in drafts] == ["A1", "B2"]  # no-recommendation and error rows get no draft
    assert [(s["sku"], s["kind"]) for s in skipped] == [("C3", "none"), ("D4", "none"), ("E5", "error")]
    a, b = drafts
    assert a["moq"] == "100" and "Hello Acme team," in a["body"] and "your listing (http://b)" in a["body"]
    assert "Plano, TX" in a["body"] and "branding/engraving" in a["body"] and "Zync Technologies" in a["body"]
    assert "A1" in a["subject"]
    assert a["label"] == "" and b["label"] == "NO EMAIL - submit via the inquiry form at http://only-url"
    assert "NO EMAIL" not in b["body"] and "inquiry form" not in b["body"]  # the label never rides inside the pasted text
    assert b["body"].startswith("Hello,")  # no maker name: no "Hello  team"
    assert "{" not in a["body"] + a["subject"] + b["body"]


def test_drafts_drop_the_price_line():
    d = drafts_from([row("A1", email="a@x.com", price=0.226, moq="500")])[0][0]
    assert not re.search(r"listed around|/unit|\$\d|MOQ", d["body"]), d["body"]


def test_listing_title_or_just_the_url():
    rows = [row("A1", email="a@x.com", keyword="OUR TARGET DESCRIPTION", title="Seller's own Cat6 Coupler Title"),
            row("B2", email="b@x.com", keyword="OUR TARGET DESCRIPTION")]          # this results file has no title for B2
    drafts, _ = drafts_from(rows, TITLE_HEADER)
    a, b = drafts
    assert 'your listing "Seller\'s own Cat6 Coupler Title" (http://listing/A1)' in a["body"]
    assert "your listing (http://listing/B2)" in b["body"]
    assert "OUR TARGET DESCRIPTION" not in a["body"] + b["body"]                    # never quoted as "your listing for ..."
    noisy = drafts_from([row("C3", maker="Acme", url="", email="")])[0][0]
    assert "your products" in noisy["body"]
    # real results files: runs before the title columns existed refer to the URL only; newer ones use the title
    for path in sorted(glob.glob(os.path.join(HERE, "Excel Output Sheets", "sourcing_results_*.xlsx")), key=os.path.getmtime)[-1:]:
        for c in ea.load_results(path, [])[0]:
            body = ea.make_draft(c)["body"]
            expected = f'your listing "{c["title"]}" ({c["url"]})' if c["title"] else f"your listing ({c['url']})"
            assert expected in body and "listing for" not in body


def test_titles_come_from_the_search_agents_own_sheet():
    """End to end with the real writer: sourcing_agent.write_excel_report -> email_agent.load_results -> draft."""
    import sourcing_agent as sa
    full = sa.AttributeBreakdown(product_type=20, category_spec=20, shielding_material=20, mount_form=20, gender_pins=20)

    def sc(name, title, total, url):
        return sa.Candidate(manufacturer=name, price_total=total, quantity_covered=1, unit_price_note="",
                            unit_price_confidence="stated", same_product_form=True, listing_form="coupler",
                            listing_caveats="", attribute_breakdown=full, listing_title=title, price=f"${total}", url=url)

    prod = lambda sku: {"sku": sku, "description": "OUR TARGET DESCRIPTION", "product_name": "Couplers", "lcom_price": 10.0}
    # T1: the cheaper second candidate wins, so the title must come from the recommended one, not Manufacturer 1
    t1 = sa.SourcingResult(product="x", no_match=False, candidates=[
        sc("Alpha", "Alpha's Cat6 Coupler", 0.5, "http://shop/alpha"), sc("Beta", "Beta Cat6 Panel Mount Coupler, Shielded", 0.3, "http://shop/beta")])
    t2 = sa.SourcingResult(product="x", no_match=False, candidates=[sc("Gamma", None, 0.4, "http://shop/gamma")])   # no title
    with tempfile.TemporaryDirectory() as tmp:
        path = os.path.join(tmp, "real_writer.xlsx")
        sa.write_excel_report([(prod("T1"), t1), (prod("T2"), t2)], path)
        header = [c.value for c in openpyxl.load_workbook(path).worksheets[0][1]]
        assert header[-6:] == ["Recommended Listing Title"] + [f"Manufacturer {n} Listing Title" for n in range(1, 6)]
        cands, skipped = ea.load_results(path, [])
        by_sku = {c["sku"]: c for c in cands}
        assert skipped == [] and by_sku["T1"]["manufacturer"] == "Beta"
        assert by_sku["T1"]["title"] == "Beta Cat6 Panel Mount Coupler, Shielded"
        d1, d2 = ea.make_draft(by_sku["T1"]), ea.make_draft(by_sku["T2"])
        assert 'your listing "Beta Cat6 Panel Mount Coupler, Shielded" (http://shop/beta)' in d1["body"]
        assert "your listing (http://shop/gamma)" in d2["body"] and "OUR TARGET DESCRIPTION" not in d1["body"] + d2["body"]
        assert "Alpha" not in d1["body"]                                       # not the other candidate's title
        # the writer fills the Recommended column itself, from the recommended candidate (not Manufacturer 1's)
        rec_cell = {r[0].value: r[header.index("Recommended Listing Title")].value
                    for r in openpyxl.load_workbook(path).worksheets[0].iter_rows(min_row=2)}
        assert rec_cell == {"T1": "Beta Cat6 Panel Mount Coupler, Shielded", "T2": None}, rec_cell
        # each header works on its own: with only "Recommended Listing Title" filled, and with only the Manufacturer N one
        def only(columns_to_keep, name):
            wb = openpyxl.load_workbook(path)
            ws = wb.worksheets[0]
            for i, h in enumerate(header, start=1):
                if h.endswith("Listing Title") and h not in columns_to_keep:
                    for r in range(2, ws.max_row + 1):
                        ws.cell(r, i).value = None
            out = os.path.join(tmp, name)
            wb.save(out)
            wb.close()
            return {c["sku"]: c["title"] for c in ea.load_results(out, [])[0]}["T1"]

        want = "Beta Cat6 Panel Mount Coupler, Shielded"
        assert only({"Recommended Listing Title"}, "only_recommended.xlsx") == want
        assert only({"Manufacturer 2 Listing Title"}, "only_manufacturer_n.xlsx") == want
        assert only(set(), "no_titles.xlsx") == ""


# ---------- near-miss drafts: opt-in, labelled for us, clean for the seller ----------

NM_HEADER = HEADER + ["Manufacturer 1 Name", "Manufacturer 1 Email", "Manufacturer 2 Name", "Manufacturer 2 Email"]
NM_ROWS = [  # (sku, seller URL in candidate block 2, concern, the one question for the seller)
    ("NM1", "http://listing/nm1-near", "Surge rating (18 kA) not stated on the listing", "Could you confirm the surge current rating (in kA) of this arrester?"),
    ("NM2", "http://listing/nm2-near", "Cat6a not stated; shielding unconfirmed", "Could you confirm the category rating of this coupler and whether it is shielded?"),
    ("NM3", "http://listing/nm3-near", "Price is a $2-$16 range; the check used $16", "Your listing shows a range of prices. Could you tell us the unit price for our sample quantity?"),
]


def nm_results(path, extra_rows=()):
    """Two recommended rows (REC1, REC2) plus NM1-NM3: no recommended seller, but a near-miss candidate in block 2."""
    rows = [row("REC1", maker="Acme", email="sales@acme.com") + ["Acme", "sales@acme.com", "", ""],
            row("REC2", maker="Beta Ltd") + ["Beta Ltd", "", "", ""]]
    for sku, url, _, _ in NM_ROWS:
        r = row(sku, rec=NO_REC, maker=None, url="")
        r[HEADER.index("Manufacturer 2 URL")], r[HEADER.index("Manufacturer 2 MOQ")] = url, "50"
        rows.append(r + ["Other Co", "", f"{sku} Maker Ltd", ""])
    make_xlsx(path, rows + list(extra_rows), NM_HEADER)


def default_output(path):
    return run_cli(["drafts", path])


def test_near_miss_is_opt_in_and_default_output_is_unchanged():
    with sandbox(near_misses=NM_ROWS) as tmp:
        path = os.path.join(tmp, "r.xlsx")
        nm_results(path)
        with_file = default_output(path)                                  # near_misses.csv exists, flag not given
        os.remove(ea.NEAR_MISSES_CSV)
        without_file = default_output(path)
        assert with_file == without_file                                  # the list changes nothing unless asked for
        assert "NEAR-MISS" not in with_file and "NEEDS HUMAN CHECK" not in with_file and "--near-miss" not in with_file
        assert with_file.count("=" * 70) == 2 and "2 draft(s); 1 with an email address. Nothing sent." in with_file
        assert with_file.rstrip().endswith("Nothing sent.") and "NM1" not in with_file
        # the default layout of a draft header is pinned exactly (this is what `drafts` printed before the flag existed)
        first = with_file.split("=" * 70 + "\n", 1)[1]
        assert first.startswith("REC1 | Acme | To: sales@acme.com\nSubject: Sample order inquiry - REC1 keyword (REC1)\n\nHello Acme team,\n")
        second = with_file.split("=" * 70 + "\n")[2]
        assert second.startswith("REC2 | Beta Ltd | To: NO EMAIL - inquiry form\nNO EMAIL - submit via the inquiry form at http://listing/REC2\nSubject: ")
        # the same three drafts through the old code path (make_draft) render exactly as the printed text says
        drafts = [ea.make_draft(c) for c in ea.load_candidates(path)]
        assert "\n".join(ea.format_draft(d) for d in drafts) in with_file
        # the real results file (gitignored, so skipped when absent): the default stays 5 drafts with the real near_misses.csv present
        real = sorted(glob.glob(os.path.join(HERE, "Excel Output Sheets", "sourcing_results_*.xlsx")), key=os.path.getmtime)[-1:]
        if real and os.path.exists(os.path.join(HERE, "near_misses.csv")):
            ea.REVIEWER_EXCLUSIONS_CSV = os.path.join(HERE, "reviewer_exclusions.csv")
            out = run_cli(["drafts", real[0]])
            n_drafts = int(re.search(r"^(\d+) draft\(s\)", out, re.M)[1])
            assert "NEAR-MISS" not in out and "NEEDS HUMAN CHECK" not in out and out.count("=" * 70) == n_drafts


def test_near_miss_flag_adds_labelled_drafts_after_the_normal_ones():
    with sandbox(near_misses=NM_ROWS) as tmp:
        path = os.path.join(tmp, "r.xlsx")
        nm_results(path)
        out = run_cli(["drafts", path, "--near-miss"])
        normal, _, extra = out.partition("NEAR-MISS DRAFTS")
        assert "2 draft(s); 1 with an email address. Nothing sent." in normal and normal.count("=" * 70) == 2 and normal.count("#" * 70) == 1
        assert "the search agent did NOT recommend these" in extra and "Nothing sent." in extra.rstrip().splitlines()[-1]
        assert extra.rstrip().endswith("3 near-miss draft(s), 0 with an email address. Nothing sent.")
        for sku, url, concern, question in NM_ROWS:
            block = extra.split(f"{sku} | ")[1].split("=" * 70)[0]
            assert f"NEEDS HUMAN CHECK: {concern}\n" in block                 # the concern, in the header
            assert f"NO EMAIL - submit via the inquiry form at {url}\n" in block and f"4. {question}" in block
            assert block.index("NEEDS HUMAN CHECK") < block.index("Subject:") < block.index("Hello")
        assert "NM1 Maker Ltd" in extra and "Other Co" not in extra                # the block the CSV row names (block 2), not block 1
        assert not state_file_exists() and FakeSMTP.created == 0                   # read only: nothing stored, nothing sent


def test_near_miss_label_stays_out_of_the_seller_text_and_asks_one_question():
    with sandbox(near_misses=NM_ROWS) as tmp:
        path = os.path.join(tmp, "r.xlsx")
        nm_results(path)
        recommended = [ea.make_draft(c) for c in ea.load_candidates(path)]
        near, notes = ea.near_miss_drafts(path, recommended, ea.load_exclusions(), ea.load_near_misses())
        assert notes == [] and [d["sku"] for d in near] == ["NM1", "NM2", "NM3"]
        for d, (sku, url, concern, question) in zip(near, NM_ROWS):
            text = d["subject"] + "\n" + d["body"]
            assert d["check"] == concern and d["questions"] == [question]
            assert "NEEDS HUMAN CHECK" not in text and concern not in text and "near-miss" not in text.lower() and "human" not in text.lower()
            items = re.findall(r"^(\d+)\. (.*)$", d["body"], re.M)
            assert [n for n, _ in items] == ["1", "2", "3", "4"] and items[3][1] == question      # one added question, nothing else
            # reuses the existing generator: the body is exactly make_draft's body for the same candidate, plus that one question
            plain = ea.make_draft({k: v for k, v in d.items() if k not in ("check", "questions", "subject", "body", "label")})
            assert d["body"] == plain["body"].replace("3. The estimated lead time for samples?",
                                                      f"3. The estimated lead time for samples?\n4. {question}")
            assert d["subject"] == plain["subject"] and d["label"] == plain["label"] and d["key"] == plain["key"]
            assert d["body"].startswith(f"Hello {d['manufacturer']} team,") and "100 Test Way" in d["body"]


def test_near_miss_drafts_pass_the_same_leak_scan():
    with sandbox(near_misses=NM_ROWS) as tmp:
        path = os.path.join(tmp, "r.xlsx")
        nm_results(path)
        near, _ = ea.near_miss_drafts(path, [], ea.load_exclusions(), ea.load_near_misses())
        for d in near:
            found = INTERNAL.search(d["subject"] + "\n" + d["body"])
            assert not found, (d["sku"], found[0])
    # the real near_misses.csv and the real latest results file (both kept out of git, so skipped when absent)
    real_csv = os.path.join(HERE, "near_misses.csv")
    files = sorted(glob.glob(os.path.join(HERE, "Excel Output Sheets", "sourcing_results_*.xlsx")), key=os.path.getmtime)
    if os.path.exists(real_csv):
        rows = ea.load_near_misses(real_csv)
        assert len(rows) >= 1
        for r in rows:
            assert not INTERNAL.search(r["question"]), (r["sku"], r["question"])      # the wording itself is clean
            assert r["question"].endswith("?") and r["concern"].strip()
        if files:
            drafts, _ = ea.load_results(files[-1], ea.load_exclusions(os.path.join(HERE, "reviewer_exclusions.csv")))
            near, notes = ea.near_miss_drafts(files[-1], [ea.make_draft(c) for c in drafts], ea.load_exclusions(os.path.join(HERE, "reviewer_exclusions.csv")), rows)
            assert len(near) + len(notes) == len(rows)
            for d in near:
                assert not INTERNAL.search(d["subject"] + "\n" + d["body"]), (d["sku"], INTERNAL.search(d["body"])[0])
                assert "NEEDS HUMAN CHECK" not in d["body"] and d["check"] not in d["body"]


def test_near_miss_respects_reviewer_exclusions_duplicates_and_missing_rows():
    with sandbox(near_misses=NM_ROWS + [("NOSUCH", "http://x", "c", "q?"), ("NM1", "http://not-a-candidate", "c", "q?"),
                                        ("REC2", "http://listing/REC2", "already recommended", "Could you confirm something?")],
                 exclusions=[("NM2", "http://listing/nm2-near", "page unavailable")]) as tmp:
        path = os.path.join(tmp, "r.xlsx")
        nm_results(path)
        out = run_cli(["drafts", path, "--near-miss"])
        extra = out.partition("NEAR-MISS DRAFTS")[2]
        assert "NM1 | " in extra and "NM3 | " in extra and "NM2 | " not in extra          # NM2 is reviewer-excluded
        assert "NOTE: near-miss NM2 (NM2 Maker Ltd): excluded by reviewer (page unavailable) - skipped" in extra
        assert "NOTE: near-miss NOSUCH: not found in the results file (no such SKU) - skipped" in extra
        assert "NOTE: near-miss NM1: not found in the results file (seller or URL not among its candidates) - skipped" in extra
        assert "NOTE: near-miss REC2 (Beta Ltd): already a recommended draft - not drafted twice" in extra
        assert extra.count("NEEDS HUMAN CHECK") == 2 and extra.rstrip().endswith("2 near-miss draft(s), 0 with an email address. Nothing sent.")
    # a seller name in the CSV (instead of a URL) selects the block too
    with sandbox(near_misses=[("NM1", "nm1 maker", "name match", "Could you confirm the rating?")]) as tmp:
        path = os.path.join(tmp, "r.xlsx")
        nm_results(path)
        near, notes = ea.near_miss_drafts(path, [], [], ea.load_near_misses())
        assert [d["manufacturer"] for d in near] == ["NM1 Maker Ltd"] and notes == []
    # no near_misses.csv at all: a loud warning, the normal drafts are untouched, nothing crashes
    with sandbox(near_misses=None) as tmp:
        path = os.path.join(tmp, "r.xlsx")
        nm_results(path)
        out = run_cli(["drafts", path, "--near-miss"])
        assert "!!! WARNING: near_misses.csv is missing - no near-miss drafts were added." in out and "NEEDS HUMAN CHECK" not in out
        assert "2 draft(s); 1 with an email address. Nothing sent." in out


def test_near_miss_is_drafts_only_and_read_only():
    with sandbox(near_misses=NM_ROWS) as tmp:
        path = os.path.join(tmp, "r.xlsx")
        nm_results(path)
        before = open(ea.NEAR_MISSES_CSV, "rb").read()
        run_cli(["drafts", path, "--near-miss"])
        assert open(ea.NEAR_MISSES_CSV, "rb").read() == before and not state_file_exists()
        assert FakeSMTP.created == 0
        for cmd in ("send", "report"):                       # the flag does not exist there: they can never include near-miss drafts
            assert cli_exit([cmd, path, "--near-miss"]) == "2"
        assert not state_file_exists() and FakeSMTP.created == 0
        report = run_cli(["report", path])
        md = open(re.search(r"Wrote (.+\.md)", report)[1], encoding="utf-8").read()
        assert "NEEDS HUMAN CHECK" not in md and "Other Co" not in md and "NM1 Maker" not in md       # no near-miss seller
        assert re.search(r"\| NM1 \|.*not contacted", md)                                              # the SKU is still "not contacted"


def test_config_values_are_used():
    with sandbox():
        os.environ["SENDER_NAME"] = "Test Sender"
        d = drafts_from([row("A1", email="a@x.com")])[0][0]
    assert "100 Test Way\nPlano, TX 75024" in d["body"] and d["body"].count("Test Sender") == 2


# ---------- SHIPPING_ADDRESS: warn on the default, refuse to send ----------

def test_shipping_address_default_warns_and_refuses_real_sends():
    with sandbox() as tmp:
        path = os.path.join(tmp, "r.xlsx")
        make_xlsx(path, [row("A1", email="a@x.com")])
        del os.environ["SHIPPING_ADDRESS"]
        assert ea.shipping_address() == (ea.DEFAULT_SHIPPING, True)
        out = run_cli(["drafts", path])                                   # drafts are still generated, with a loud warning
        assert "!!! WARNING: SHIPPING_ADDRESS is not set" in out and "Plano, TX" in out and "Hello Acme team," in out
        msg = cli_exit(["send", path])                                    # a real send is refused before any prompt
        assert msg.startswith("Refusing to send: SHIPPING_ADDRESS is not set") and FakeSMTP.created == 0
        for same_as_default in ("Zync Technologies, Plano, TX", "  zync technologies\\nplano,  tx  ", ""):
            os.environ["SHIPPING_ADDRESS"] = same_as_default
            assert ea.shipping_address()[1], same_as_default
        sent = []
        try:
            ea.confirm_and_send([ea.make_draft(cand("A1", "a@x.com"))], {}, send=sent.append, ask=lambda _: "y")
            raise AssertionError("sent with the default address")
        except ea.SendRefused:
            assert sent == []
        n = ea.confirm_and_send([ea.make_draft(cand("A1", "a@x.com"))], {}, send=sent.append, ask=lambda _: "y",
                                test_recipient="me@test.example")          # a [TEST] send is safe and still allowed
        assert n == 1
        os.environ["SHIPPING_ADDRESS"] = "Zync Technologies\\n100 Test Way\\nPlano, TX 75024"
        assert ea.shipping_address() == ("Zync Technologies\n100 Test Way\nPlano, TX 75024", False)
        assert "SHIPPING_ADDRESS" not in run_cli(["drafts", path]).split("Subject")[0]   # no warning once it is set


def test_shipping_address_forms_in_dotenv():
    """The three ways to write a multi-line SHIPPING_ADDRESS in .env all produce the same address."""
    want = "Zync Technologies\n100 Test Way\nPlano, TX 75024"
    forms = {"double quotes with \\n": 'SHIPPING_ADDRESS="Zync Technologies\\n100 Test Way\\nPlano, TX 75024"\n',
             "real line breaks inside quotes": 'SHIPPING_ADDRESS="Zync Technologies\n100 Test Way\nPlano, TX 75024"\n',
             "single quotes with literal \\n": "SHIPPING_ADDRESS='Zync Technologies\\n100 Test Way\\nPlano, TX 75024'\n"}
    with tempfile.TemporaryDirectory() as tmp:
        for label, text in forms.items():
            path = os.path.join(tmp, ".env")
            open(path, "w", encoding="utf-8").write(text)
            old = os.environ.get("SHIPPING_ADDRESS")
            os.environ["SHIPPING_ADDRESS"] = dotenv_values(path)["SHIPPING_ADDRESS"]
            try:
                assert ea.shipping_address() == (want, False), label
            finally:
                os.environ.pop("SHIPPING_ADDRESS") if old is None else os.environ.update(SHIPPING_ADDRESS=old)


# ---------- nothing internal may reach a seller ----------

INTERNAL = re.compile(r"l-?com|margin|accuracy|\d\s?%|promo|implausible|model:|unverified|ordering note|check before ordering|"
                      r"target is|unknown|unnamed|rubric|tie-break|reviewer|pick the right option|confirm it is the target",
                      re.IGNORECASE)
REAL_NOTES = [  # copied from the 2026-10-04 run (161054), plus one of every other note the search agent can write
    "model: Inline coupler form; waterproof rated but target does not specify this",
    "price needs an order of 10,000+ pieces (MOQ); price is for the 10,000-49,999 pieces tier (others: $0.50 at ≥50,000 pieces)",
    "price needs an order of 500+ pieces (MOQ); model: 90-degree angled variant; slightly different form factor",
    "listing offers male and female variants - pick the right option when ordering; listing mentions several contact types "
    "(IDC/crimp; target is crimp) - check before ordering; model: Pack includes both male and female connectors",
    "title names a different D-sub pin count (\"VGA\") - confirm it is the target's variant",
    "one listing covers several models (MC-6BP, MC-6BR) - confirm which one ships",
    "promo price - regular price not shown on the listing",
    "listing's quantity 10 wasn't tied to this price - priced per piece",
    "contact type isn't stated on the listing (target is crimp) - confirm before ordering",
    "listing contradicts itself on category: \"compatible with CAT5e\" but also lists up to Cat.6 - confirm the rating with the seller",
    "listing contradicts itself on gender: Type A is both male and female - confirm with the seller",
    "listing text mentions \"PVC jacket\" - may be a short cable rather than a plain adapter, verify the product page",
    "model: unit price 95% below L-Com, margin 91%, implausible and promo-looking; accuracy 80%",
]


def test_no_internal_text_in_any_draft():
    makers = ["Unknown", "Unknown brand (AliExpress listing)", "E-HONG", None, "Acme Co.", "X"] * 3
    emails = ["a@x.com", "", ""] * 5
    rows = [row(f"S{i}", maker=makers[i], email=emails[i], note=n, price=0.226) for i, n in enumerate(REAL_NOTES)]
    drafts, _ = drafts_from(rows)
    assert len(drafts) == len(REAL_NOTES)
    for d in drafts:
        text = f"{d['subject']}\n{d['body']}"
        assert not INTERNAL.search(text), (d["sku"], INTERNAL.search(text)[0], text)
    # placeholder makers get a plain greeting, never "Hello Unknown team"
    assert drafts[0]["body"].startswith("Hello,") and drafts[1]["body"].startswith("Hello,")
    assert drafts[2]["body"].startswith("Hello E-HONG team,") and drafts[3]["body"].startswith("Hello,")
    # ...and the newest real results file (kept out of git, so skipped when absent) must scan clean too
    files = glob.glob(os.path.join(HERE, "Excel Output Sheets", "sourcing_results_*.xlsx"))
    for path in sorted(files, key=os.path.getmtime)[-1:]:
        for c in ea.load_results(path, [])[0]:
            d = ea.make_draft(c)
            found = INTERNAL.search(d["subject"] + d["body"])
            assert not found, (path, d["sku"], found[0])


def test_notes_become_polite_questions():
    qs = ea.confirmation_questions(REAL_NOTES[3])
    assert len(qs) == 2 and all(q.endswith("?") for q in qs), qs
    assert any("male and female options" in q for q in qs) and any("more than one contact type" in q for q in qs)
    assert ea.confirmation_questions("model: only the model's free text") == []   # free text is never forwarded
    assert ea.confirmation_questions("something nobody wrote a question for") == []
    assert ea.confirmation_questions("") == [] and ea.confirmation_questions(None) == []
    ours = [n for n in REAL_NOTES if "model:" not in n]       # in real notes the model's caveat is always last
    assert len(ea.confirmation_questions("; ".join(ours))) == ea.MAX_QUESTIONS and len(ours) > ea.MAX_QUESTIONS
    d = drafts_from([row("A1", email="a@x.com", note=REAL_NOTES[4] + "; " + REAL_NOTES[5])])[0][0]
    assert "\n4. Your listing title mentions \"VGA\"" in d["body"]
    assert "\n5. Your listing covers several models (MC-6BP, MC-6BR)" in d["body"]


# ---------- reviewer exclusions: skip rows, say how many ----------

def test_reviewer_exclusions_skip_rows_and_say_so():
    excl = [("C&P9M", "https://www.aliexpress.us/item/2251832668010347.html", "IDC ribbon, mixed pack"),
            ("HDFF", "xiangtianzhong", "short cable")]
    rows = [row("C&P9M", maker="Unknown", url="https://www.aliexpress.us/item/2251832668010347.html"),
            row("HDFF", maker="SHENZHEN XIANGTIANZHONG TECHNOLOGY CO., LTD.", url="http://xtz/p"),
            row("C&P9M2", maker="Other", url="https://www.aliexpress.us/item/2251832668010347.html"),   # same URL, other SKU
            row("OK1", maker="Fine Co", email="a@x.com"),
            row("NR", rec=NO_REC, maker=None, url="")]
    with sandbox(exclusions=excl) as tmp:
        before = open(ea.REVIEWER_EXCLUSIONS_CSV, "rb").read()
        path = os.path.join(tmp, "r.xlsx")
        make_xlsx(path, rows)
        out = run_cli(["drafts", path])
        assert "Skipped 3 row(s): 1 no recommended seller, 2 excluded by reviewer_exclusions.csv." in out
        assert "2 draft(s)" in out and "C&P9M |" not in out and "HDFF |" not in out and "C&P9M2 |" in out
        _, skipped = ea.load_results(path)
        assert [s["reason"] for s in skipped if s["kind"] == "reviewer"] == ["excluded by reviewer: IDC ribbon, mixed pack",
                                                                            "excluded by reviewer: short cable"]
        assert open(ea.REVIEWER_EXCLUSIONS_CSV, "rb").read() == before          # read-only
        # a missing file skips nothing and says so loudly
        os.remove(ea.REVIEWER_EXCLUSIONS_CSV)
        out = run_cli(["drafts", path])
        assert "reviewer_exclusions.csv is missing or empty - known-bad sellers are NOT being skipped" in out and "C&P9M |" in out
    real = os.path.join(HERE, "reviewer_exclusions.csv")
    if os.path.exists(real):                                                    # the real file excludes this run's C&P9M pick
        assert ea.excluded_reason("C&P9M", "", "https://www.aliexpress.us/item/2251832668010347.html", ea.load_exclusions(real))


# ---------- same recipient: warn at the prompt; one email per SKU stays ----------

def test_duplicate_recipient_warning_and_one_email_per_sku():
    drafts = [ea.make_draft(cand("A1", "sales@acme.com")), ea.make_draft(cand("A2", "Sales@Acme.com")),
              ea.make_draft(cand("B1", "b@beta.com")), ea.make_draft(cand("F1", ""))]
    with sandbox():
        buf, sent = io.StringIO(), []
        with contextlib.redirect_stdout(buf):
            n = ea.confirm_and_send(drafts, {}, send=sent.append, ask=lambda _: "y")
        out = buf.getvalue()
        assert n == 3 and [d["sku"] for d in sent] == ["A1", "A2", "B1"]       # still one email per SKU
        assert "WARNING: 2 separate emails (A1, A2) are addressed to the same recipient sales@acme.com" in out
        assert "NOTE: same address as A2" in out and "NOTE: same address as A1" in out and out.count("NOTE: same address") == 2
        buf = io.StringIO()
        with contextlib.redirect_stdout(buf):
            ea.confirm_and_send(drafts[:1] + drafts[2:], {}, yes=True, send=sent.append, ask=lambda _: "no")
        assert "WARNING" not in buf.getvalue()                                  # no duplicates, no warning
        buf = io.StringIO()
        with contextlib.redirect_stdout(buf):
            ea.confirm_and_send(drafts, {}, yes=True, send=sent.append, ask=lambda _: "no")
        assert "WARNING: 2 separate emails" in buf.getvalue()                   # also before the batch "yes" prompt
        buf = io.StringIO()
        with contextlib.redirect_stdout(buf):
            ea.confirm_and_send(drafts, {}, yes=True, send=sent.append, ask=lambda _: "yes", test_recipient="me@test.example")
        assert "WARNING: 2 separate emails" not in buf.getvalue()               # a test send to one address is expected
    plain, _ = drafts_from([row("A3", maker="Acme", email=""), row("A4", maker="Acme", email="")])
    assert plain[0]["url"] != plain[1]["url"] and all(d["label"] for d in plain)   # forms stay per listing


# ---------- confirmation gate ----------

def test_gate_nothing_sends_without_explicit_yes():
    drafts = [ea.make_draft(cand("A", "a@x.com")), ea.make_draft(cand("B", "b@x.com")), ea.make_draft(cand("C", ""))]

    def attempt(answers, **kw):
        sent, state = [], {}
        it = iter(answers)
        n = ea.confirm_and_send(drafts, state, send=sent.append, ask=lambda _: next(it), **kw)
        return n, sent, state

    with sandbox():
        for answers in (["", ""], ["n", "N"], ["no", "nope"], [" ", "\t"], ["yes", "yes"], ["maybe", "ok"], ["q"]):
            assert attempt(answers) == (0, [], {}), answers   # Enter, n, typos, 'yes' at a per-email prompt, q: nothing sent
        n, sent, state = attempt(["y", "n"])
        assert n == 1 and [d["sku"] for d in sent] == ["A"] and set(state) == {drafts[0]["key"]}
        for answers in ([""], ["y"], ["YES PLEASE"], ["ye"], ["no"]):      # batch mode needs exactly the word "yes"
            n, sent, _ = attempt(answers, yes=True)
            assert (n, sent) == (0, []), answers
        assert attempt(["yes"], yes=True)[0] == 2                          # C has no email: never sent

        def eof(_):
            raise EOFError()

        def must_not_send(d):
            raise AssertionError("sent")

        try:  # EOF on stdin (a closed pipe) must not send either
            ea.confirm_and_send(drafts, {}, send=must_not_send, ask=eof)
            raise AssertionError("EOF swallowed")
        except EOFError:
            pass


def test_smtp_is_reachable_only_through_the_gate():
    src = open(os.path.join(HERE, "email_agent.py"), encoding="utf-8").read()
    assert src.count("send(d)") == 1                       # one call site, inside confirm_and_send, after the prompt
    assert src.count("smtp_send") == 2                     # its definition and confirm_and_send's default `send=`
    with sandbox() as tmp:
        path = os.path.join(tmp, "r.xlsx")
        make_xlsx(path, [row("A1", email="a@x.com"), row("B2", email="")])
        ea.summarize_reply = lambda text: {"status": "needs_info", "summary": "stub"}
        pricing = os.path.join(FIXTURES, "pricing.txt")
        for argv in (["drafts", path], ["mark", "B2", "--results", path], ["report", path],
                     ["reply", "A1", "--results", path, "--file", pricing]):
            run_cli(argv)
        assert FakeSMTP.created == 0                       # drafts, mark, report and reply never open an SMTP connection
    # At the real prompt: a copy of the agent runs in a temp dir pointed at a closed local port, so even a bug
    # could only fail to connect - it could not reach anyone.
    with tempfile.TemporaryDirectory() as tmp:
        for f in ("email_agent.py", "email_template.txt"):
            shutil.copy(os.path.join(HERE, f), tmp)
        make_xlsx(os.path.join(tmp, "r.xlsx"), [row("A1", email="a@x.com"), row("B2", email="b@x.com")])
        env = {k: v for k, v in os.environ.items() if k not in ("ANTHROPIC_API_KEY", "SHIPPING_ADDRESS")}
        env.update(SMTP_HOST="127.0.0.1", SMTP_PORT="1", SMTP_USER="u@z.example", SMTP_PASSWORD="x")

        def send(stdin, *flags, shipping=True):
            e = dict(env, SHIPPING_ADDRESS="Zync Technologies\\n100 Test Way\\nPlano, TX 75024") if shipping else env
            return subprocess.run([sys.executable, "email_agent.py", "send", "r.xlsx", *flags], cwd=tmp, input=stdin,
                                  text=True, capture_output=True, env=e, timeout=60)

        for stdin in ("\n\n", "n\nn\n", "x\nq\n"):
            p = send(stdin)
            assert "Sent 0 email(s)" in p.stdout and "refused" not in p.stderr.lower(), (stdin, p.stdout, p.stderr)
        p = send("")                                       # stdin closed: crashes at the prompt, sends nothing
        assert p.returncode != 0 and "EOFError" in p.stderr and "refused" not in p.stderr.lower()
        p = send("no\n", "--yes")
        assert "Sent 0 email(s)" in p.stdout and "refused" not in p.stderr.lower()
        p = send("y\ny\n", shipping=False)                 # default shipping address: refused before any prompt
        assert p.returncode != 0 and "Refusing to send: SHIPPING_ADDRESS" in p.stderr and "Sent" not in p.stdout
        assert "refused" not in p.stderr.lower().replace("refusing", "")
        assert not os.path.exists(os.path.join(tmp, "email_state.json"))     # nothing was sent, nothing stored


# ---------- --test-recipient: a safe first send ----------

def test_test_recipient_redirects_everything():
    with sandbox() as tmp:
        path = os.path.join(tmp, "r.xlsx")
        make_xlsx(path, [row("A1", email="sales@acme.com"), row("B2", email="info@beta.com"),
                         row("C3", email="", url="http://form")])
        drafts = [ea.make_draft(c) for c in ea.load_candidates(path)]
        sent, state = [], {}
        n = ea.confirm_and_send(drafts, state, yes=True, send=sent.append, ask=lambda _: "yes", test_recipient="me@test.example")
        assert n == 3 and state == {}                                  # test sends are never recorded as sent
        assert {d["email"] for d in sent} == {"me@test.example"}
        assert all(d["subject"].startswith("[TEST] ") and d["body"].startswith("[TEST - this would have gone to: ") for d in sent)
        assert "sales@acme.com" in sent[0]["body"] and "inquiry form http://form" in sent[2]["body"]
        # Through the CLI with a fake SMTP and an auto-"y": every message is addressed to the test address only.
        real = ea.confirm_and_send
        ea.confirm_and_send = lambda drafts, state, yes=False, test_recipient="": real(
            drafts, state, yes, ask=lambda _: "y", test_recipient=test_recipient)
        try:
            out = run_cli(["send", path, "--test-recipient", "me@test.example"])
        finally:
            ea.confirm_and_send = real
        assert "Sent 3 email(s)" in out and FakeSMTP.created == 3
        assert [m["To"] for m in FakeSMTP.sent] == ["me@test.example"] * 3
        assert all(m["Subject"].startswith("[TEST] ") for m in FakeSMTP.sent)
        everything = "".join(m.as_string() for m in FakeSMTP.sent)
        assert "To: sales@acme.com" not in everything and "To: info@beta.com" not in everything
        assert not state_file_exists()                                 # a test send stores nothing
        for bad in ("sales@acme.com", "not-an-address"):              # a manufacturer's address, or junk, is refused
            assert "test address" in cli_exit(["send", path, "--test-recipient", bad])
        assert FakeSMTP.created == 3


# ---------- replies (stubbed LLM), matching and the report ----------

def stub_summarize(text):
    low = text.lower()
    if "unsubscribe" in low or "auto-reply" in low or "automated message" in low:
        return {"status": "auto_reply_or_spam", "summary": "auto-reply or spam, no information"}
    if "discontinued" in low or "do not provide" in low:
        return {"status": "dead_end", "summary": "declined"}
    if "usd" in low or "price" in low or "报价" in text:
        return {"status": "pricing_provided", "summary": "gave sample pricing: " + text.strip().splitlines()[0][:40]}
    return {"status": "needs_info", "summary": "asks us for more information"}


def read_fixture(name):
    return open(os.path.join(FIXTURES, name), encoding="utf-8").read()


def report_rows(results_path):
    out = run_cli(["report", results_path])
    md = open(re.search(r"Wrote (.+\.md)", out)[1], encoding="utf-8").read()
    rows = {}
    for line in md.splitlines():
        if line.startswith("| ") and "---" not in line:
            cells = [c.strip() for c in line.strip("|").split(" | ")]
            rows[cells[0]] = cells
    return md, rows


# report columns: 0 SKU, 1 Product, 2 Manufacturer, 3 Contact, 4 Contact method, 5 Sent, 6 Reply received, 7 Reply status, 8 Summary

def test_replies_matching_and_report():
    with sandbox() as tmp:
        path = os.path.join(tmp, "r.xlsx")
        make_xlsx(path, [row("ECF504-SC6", maker="Acme Co.", email="sales@acme.com"),
                         row("TDG1026KS-C6", maker="Acme Co.", email="sales@acme.com"),     # same maker, second SKU
                         row("HDFF", maker="Beta Ltd", email=""),
                         row("VIC00001", maker="Gamma", email=""),
                         row("FOA-020C", maker="Delta", email="d@delta.com"),
                         row("LCSP1050", maker="Echo", email=""),
                         row("NOREPLY", maker="Zeta", email="z@zeta.com"),
                         row("SPAMMED", maker="Eta", email="e@eta.com")])
        ea.summarize_reply = stub_summarize
        fixture_for = {"ECF504-SC6": "pricing.txt", "HDFF": "needs_info.txt", "VIC00001": "declined.txt",
                       "FOA-020C": "autoreply.txt", "LCSP1050": "chinese.txt", "SPAMMED": "spam.txt"}
        for sku, name in fixture_for.items():
            assert f"{sku} (" in run_cli(["reply", sku, "--results", path, "--file", os.path.join(FIXTURES, name)])
        run_cli(["mark", "HDFF", "--results", path])                    # contacted by hand through the inquiry form
        md, rows = report_rows(path)
        assert rows["HDFF"][4] == "inquiry form submitted" and rows["HDFF"][6:8] == ["yes", "needs_info"]
        assert rows["ECF504-SC6"][6:8] == ["yes", "pricing_provided"] and rows["ECF504-SC6"][2] == "Acme Co."
        assert rows["VIC00001"][4:8] == ["inquiry form not yet submitted", "no (inquiry form)", "yes", "dead_end"]
        assert rows["FOA-020C"][4:8] == ["email drafted, not sent", "no (email)", "yes", "auto_reply_or_spam"]
        assert rows["SPAMMED"][7] == "auto_reply_or_spam" and rows["VIC00001"][7] == "dead_end"   # declined is not spam
        assert rows["NOREPLY"][6:8] == ["no", ""] and rows["TDG1026KS-C6"][6:8] == ["no", ""]
        assert rows["LCSP1050"][7] == "pricing_provided" and "您好" in rows["LCSP1050"][8]   # Chinese reply, summary kept
        # One reply covering two products from the same maker: it is filed once per SKU it is pasted against.
        two = os.path.join(FIXTURES, "two_products.txt")
        run_cli(["reply", "ECF504-SC6", "--results", path, "--file", two])
        run_cli(["reply", "TDG1026KS-C6", "--results", path, "--file", two])
        _, rows = report_rows(path)
        assert rows["TDG1026KS-C6"][2] == "Acme Co." and rows["TDG1026KS-C6"][6:8] == ["yes", "pricing_provided"]
        # Non-ASCII survives on disk as UTF-8 (the state file, the markdown report, the Excel report).
        state = read_state()
        assert state[key("LCSP1050")]["reply"] == read_fixture("chinese.txt") and "您好" in md
        xlsx = max(glob.glob(os.path.join(tmp, "Reports", "*.xlsx")), key=os.path.getmtime)
        wb = openpyxl.load_workbook(xlsx)
        values = {r[0].value: [c.value for c in r] for r in wb.active.iter_rows(min_row=2)}
        assert [c.value for c in wb.active[1]] == ea.REPORT_HEADER and "您好" in values["LCSP1050"][8]
        wb.close()


def test_stdin_is_read_as_utf8():
    chinese = read_fixture("chinese.txt")
    seen = []
    with sandbox() as tmp:
        path = os.path.join(tmp, "r.xlsx")
        make_xlsx(path, [row("A1", email="a@x.com")])
        ea.summarize_reply = lambda text: seen.append(text) or {"status": "pricing_provided", "summary": "ok"}
        run_cli(["reply", "A1", "--results", path], stdin_bytes=chinese.encode("utf-8"))   # a cp1252 pipe, as on Windows
        assert seen == [chinese], ascii(seen)
        assert read_state()[key("A1")]["reply"] == chinese


def test_reply_and_mark_reject_unknown_skus_and_store_nothing():
    with sandbox() as tmp:
        path = os.path.join(tmp, "r.xlsx")
        make_xlsx(path, [row("A1", email="a@x.com"), row("NR", rec=NO_REC, maker=None, url="")])
        pricing = os.path.join(FIXTURES, "pricing.txt")
        msg = cli_exit(["reply", "ECF504-SC66", "--results", path, "--file", pricing])    # typo; the real LLM must not be called
        assert "ECF504-SC66: not a SKU in" in msg and "A1" in msg and msg.endswith("Nothing stored.")
        assert "not a SKU in" in cli_exit(["mark", "NOPE", "--results", path])
        msg = cli_exit(["reply", "NR", "--results", path, "--file", pricing])             # in the file, but no seller
        assert "no recommended seller" in msg and msg.endswith("Nothing stored.")
        assert not state_file_exists()
        ea.summarize_reply = stub_summarize
        run_cli(["reply", "a1", "--results", path, "--file", pricing])                    # case-insensitive
        assert list(read_state()) == [key("A1")]


def test_second_reply_is_appended_not_replaced():
    inputs = []

    def recording_stub(text):
        inputs.append(text)
        return stub_summarize(text)

    with sandbox() as tmp:
        path = os.path.join(tmp, "r.xlsx")
        make_xlsx(path, [row("A1", email="a@x.com")])
        ea.summarize_reply = recording_stub
        run_cli(["reply", "A1", "--results", path, "--file", os.path.join(FIXTURES, "needs_info.txt")])
        out = run_cli(["reply", "A1", "--results", path, "--file", os.path.join(FIXTURES, "pricing.txt")])
        entry = read_state()[key("A1")]
        assert len(entry["replies"]) == 2
        assert read_fixture("needs_info.txt") in entry["reply"] and read_fixture("pricing.txt") in entry["reply"]
        assert "--- reply 1" in entry["reply"] and "--- reply 2" in entry["reply"]
        assert entry["reply"].index("drawing") < entry["reply"].index("USD 3.20")       # earlier text first
        assert inputs[1] == entry["reply"] and "pricing_provided" in out and entry["status"] == "pricing_provided"
        run_cli(["reply", "A1", "--results", path, "--file", os.path.join(FIXTURES, "pricing.txt")])   # same paste again
        assert len(read_state()[key("A1")]["replies"]) == 2
        before = json.dumps(read_state(), sort_keys=True)

        def failing(text):
            raise RuntimeError("summarizer down")

        ea.summarize_reply = failing                                                     # a failed summary stores nothing
        try:
            run_cli(["reply", "A1", "--results", path, "--file", os.path.join(FIXTURES, "declined.txt")])
        except RuntimeError:
            pass
        assert json.dumps(read_state(), sort_keys=True) == before


def test_state_is_keyed_by_sku_and_listing():
    assert ea.listing_key("a1", "https://Shop.example/item/1/?spm=x#top", "Acme") == "A1|shop.example/item/1"
    assert ea.listing_key("A1", "https://shop.example/item/1", "Other Name") == ea.listing_key("A1", "https://shop.example/item/1/", "Acme")
    assert ea.listing_key("A1", "", "Acme Co.") == "A1|acme co." and ea.listing_key("A1", "", "Other") != ea.listing_key("A1", "", "Acme Co.")
    with sandbox() as tmp:
        old, new = os.path.join(tmp, "old.xlsx"), os.path.join(tmp, "new.xlsx")
        make_xlsx(old, [row("A1", maker="Seller X", email="x@x.com", url="http://shop/x"), row("B2", maker="Seller B", email="", url="http://shop/b")])
        make_xlsx(new, [row("A1", maker="Seller Y", email="y@y.com", url="http://shop/y"), row("B2", maker="Seller B", email="", url="http://shop/b")])
        ea.summarize_reply = stub_summarize
        run_cli(["reply", "A1", "--results", old, "--file", os.path.join(FIXTURES, "pricing.txt")])
        run_cli(["mark", "B2", "--results", old])
        d_old = [ea.make_draft(c) for c in ea.load_candidates(old)]
        ea.save_state({**read_state(), key("A1", "http://shop/x"): {**read_state()[key("A1", "http://shop/x")], "sent": "2026-10-05 09:00"}})
        _, rows = report_rows(old)
        assert rows["A1"][4] == "email sent" and rows["A1"][6] == "yes"
        # A re-run now recommends Seller Y for A1: the old seller's reply and sent mark must not show under Y.
        md, rows = report_rows(new)
        assert rows["A1"][2] == "Seller Y" and rows["A1"][4] == "email drafted, not sent" and rows["A1"][6:8] == ["no", ""]
        assert rows["B2"][4] == "inquiry form submitted"                      # B2's seller did not change: still shown
        assert "1 earlier record(s) belong to a seller that is no longer the recommended one" in md and "A1 (Seller X)" in md
        # ...and the new seller still gets its email: Seller X's sent mark does not count as sent.
        d_new = [ea.make_draft(c) for c in ea.load_candidates(new)]
        sent = []
        ea.confirm_and_send(d_new, read_state(), yes=True, send=sent.append, ask=lambda _: "yes")
        assert [d["sku"] for d in sent] == ["A1"] and [d["email"] for d in sent] == ["y@y.com"]
        assert ea.confirm_and_send(d_old, read_state(), yes=True, send=sent.append, ask=lambda _: "yes") == 0   # X is already done


def test_reply_statuses_include_auto_reply_or_spam():
    assert "auto_reply_or_spam" in ea.REPLY_STATUSES and "dead_end" in ea.REPLY_STATUSES
    assert ea.parse_summary('{"status": "auto_reply_or_spam", "summary": "out of office"}')["status"] == "auto_reply_or_spam"
    assert ea.parse_summary('{"status": "dead_end", "summary": "declined"}')["status"] == "dead_end"
    try:
        ea.parse_summary('{"status": "spam", "summary": ""}')
        raise AssertionError("an unknown status was accepted")
    except ValueError:
        pass


def test_contact_method_column():
    drafts = [ea.make_draft(cand("E1", "a@x.com")), ea.make_draft(cand("E2", "b@x.com")),
              ea.make_draft(cand("F1", "")), ea.make_draft(cand("F2", ""))]
    state = {drafts[0]["key"]: {"sent": "2026-10-05 09:00"}, drafts[2]["key"]: {"sent": "2026-10-05 09:01"}}
    skipped = [{"sku": "N1", "product": "p", "manufacturer": "", "reason": "no recommended seller", "kind": "none"},
               {"sku": "N2", "product": "p", "manufacturer": "", "reason": "error researching this product", "kind": "error"}]
    rows = ea.build_report(drafts, state, skipped)
    assert "Contact method" in ea.REPORT_HEADER and ea.REPORT_HEADER.index("Contact method") == 4
    assert [r[4] for r in rows] == ["email sent", "email drafted, not sent", "inquiry form submitted",
                                    "inquiry form not yet submitted", "not contacted", "not contacted"]
    assert rows[4][8] == "no recommended seller" and rows[5][2] == "" and rows[4][6] == "no"
    with sandbox() as tmp:
        path = os.path.join(tmp, "r.xlsx")
        make_xlsx(path, [row("A1", email="a@x.com"), row("NR", rec=NO_REC, maker=None, url=""),
                         row("ER", rec="Error researching this product: boom - re-run with --sku", maker="Acme")])
        _, by_sku = report_rows(path)
        assert [by_sku[s][4] for s in ("A1", "NR", "ER")] == ["email drafted, not sent", "not contacted", "not contacted"]


def test_drafts_never_touch_the_state_file():
    with sandbox() as tmp:
        path = os.path.join(tmp, "r.xlsx")
        make_xlsx(path, [row("A1", email="a@x.com"), row("B2", email="")])
        run_cli(["drafts", path])
        assert not state_file_exists()                                    # drafts creates nothing
        run_cli(["report", path])
        assert not state_file_exists()                                    # neither does the report
        open(ea.STATE_PATH, "w", encoding="utf-8").write('{"keep": {"sent": "x"}}')
        before = open(ea.STATE_PATH, "rb").read()
        run_cli(["drafts", path])
        assert open(ea.STATE_PATH, "rb").read() == before                 # an existing state file is left byte-for-byte


# ---------- structured quote fields: the model extracts (stubbed here), the code verifies ----------

def run_quote(fixture, raw, state=None, draft=None):
    """A stub LLM returns `raw` for the fixture reply; record_reply verifies it against the reply text."""
    state = {} if state is None else state
    draft = draft or ea.make_draft(cand("A1", "a@x.com"))
    full = {"status": "pricing_provided", "summary": "stub summary", **raw}
    result = ea.record_reply(state, draft, read_fixture(fixture), summarize=lambda text: full)
    return result, state[draft["key"]]


QUOTE_CASES = [
    ("a per-unit sample price", "quote_per_unit.txt",
     {"sample_available": "yes", "sample_quantity": 5, "sample_unit_price": "3.20", "price_basis": "per_unit",
      "lead_time_days": 5, "currency": "USD",
      "evidence": {"sample_quantity": "5 pcs are available as samples", "sample_unit_price": "Sample price is USD 3.20 per piece",
                   "lead_time_days": "Lead time for samples is 5 days"}},
     {"sample_available": "yes", "sample_quantity": 5, "sample_unit_price": "3.20", "sample_total_price": None, "price_basis": "per_unit",
      "lead_time_days": 5, "currency": "USD", "bulk_tiers": None, "moq": None}, []),
    ("a total sample fee", "quote_total_fee.txt",
     {"sample_available": "yes", "sample_quantity": 10, "sample_total_price": "45.00", "price_basis": "total_for_quantity",
      "lead_time_days": 7, "currency": "USD",
      "evidence": {"sample_quantity": "in total for 10 pcs", "sample_total_price": "The sample fee is $45.00 in total",
                   "lead_time_days": "Sample lead time is 7 days"}},
     {"sample_quantity": 10, "sample_unit_price": None, "sample_total_price": "45.00", "price_basis": "total_for_quantity",
      "lead_time_days": 7, "currency": "USD"}, []),
    ("bulk tiers", "quote_bulk_tiers.txt",
     {"sample_quantity": 5, "sample_unit_price": "3.20", "price_basis": "per_unit", "moq": 100, "production_lead_time_days": 20,
      "shipping_terms": "FOB", "currency": "USD", "quote_valid_until": "2026-11-15",
      "bulk_tiers": [{"min_qty": 100, "unit_price": "1.80", "evidence": "100-499 pcs USD 1.80/pc"},
                     {"min_qty": 500, "unit_price": "1.50", "evidence": "500-999 pcs USD 1.50/pc"},
                     {"min_qty": 1000, "unit_price": "1.20", "evidence": "1000+ pcs USD 1.20/pc"}],
      "evidence": {"sample_quantity": "Samples: 5 pcs", "sample_unit_price": "5 pcs at USD 3.20 each", "moq": "MOQ is 100 pcs",
                   "production_lead_time_days": "Production lead time is 20 days"}},
     {"moq": 100, "production_lead_time_days": 20, "shipping_terms": "FOB", "quote_valid_until": "2026-11-15", "currency": "USD",
      "bulk_tiers": [{"min_qty": 100, "unit_price": "1.80"}, {"min_qty": 500, "unit_price": "1.50"}, {"min_qty": 1000, "unit_price": "1.20"}]}, []),
    ("a branding fee", "quote_branding_fee.txt",
     {"branding_possible": "yes", "branding_fee": "30.00", "branding_min_qty": 50, "sample_quantity": 10, "sample_unit_price": "2.50",
      "price_basis": "per_unit", "currency": "USD",
      "evidence": {"branding_fee": "The engraving fee is USD 30.00 one-time setup", "branding_min_qty": "minimum 50 pcs for engraving",
                   "sample_quantity": "for 10 pcs", "sample_unit_price": "Sample price is USD 2.50 per piece"}},
     {"branding_possible": "yes", "branding_fee": "30.00", "branding_min_qty": 50, "sample_unit_price": "2.50", "currency": "USD"}, []),
    ("a quote in RMB", "quote_rmb.txt",
     {"sample_quantity": 10, "sample_unit_price": "22.50", "price_basis": "per_unit", "shipping_cost": "80", "shipping_terms": "EXW",
      "currency": "RMB",
      "evidence": {"sample_quantity": "for 10 pcs", "sample_unit_price": "Sample price is RMB 22.50 per piece",
                   "shipping_cost": "Shipping cost is RMB 80"}},
     {"sample_unit_price": "22.50", "shipping_cost": "80", "shipping_terms": "EXW", "currency": "CNY"}, []),
    ("an ambiguous price", "quote_ambiguous.txt",
     {"sample_quantity": 10, "sample_unit_price": "15", "price_basis": "unclear",
      "evidence": {"sample_quantity": "for 10 pcs sample", "sample_unit_price": "The price is 15 for 10 pcs sample"}},
     {"sample_quantity": 10, "sample_unit_price": "15", "price_basis": "unclear", "currency": "unknown"}, []),
    ("a reply in Chinese", "quote_chinese.txt",
     {"sample_available": "yes", "sample_quantity": 5, "sample_unit_price": "3.2", "price_basis": "per_unit", "lead_time_days": 5,
      "branding_possible": "yes", "branding_fee": "50", "currency": "USD",
      "evidence": {"sample_quantity": "样品数量5件", "sample_unit_price": "样品价格每件3.2美元", "lead_time_days": "样品交期5天",
                   "branding_fee": "工程费50美元"}},
     # 美元 is a USD marker (matched before the 元 inside it), so the Chinese quote is USD and nothing is mixed
     {"sample_quantity": 5, "sample_unit_price": "3.2", "lead_time_days": 5, "branding_fee": "50", "branding_possible": "yes",
      "currency": "USD"}, []),
    ("a reply with questions for us", "quote_needs_info.txt",
     {"status": "needs_info", "needs_from_us": ["drawing", "planned order quantity", " preferred shielding material "]},
     {"needs_from_us": ["drawing", "planned order quantity", "preferred shielding material"], "sample_unit_price": None,
      "currency": None, "price_basis": None, "bulk_tiers": None, "sample_quantity": None}, []),
    ("a stub number that is not in the reply", "quote_per_unit.txt",
     {"sample_quantity": 5, "sample_unit_price": "2.50", "price_basis": "per_unit", "currency": "USD",
      "evidence": {"sample_quantity": "5 pcs are available as samples", "sample_unit_price": "USD 2.50 per piece"}},
     {"sample_quantity": 5, "sample_unit_price": None, "price_basis": None, "currency": None},
     ["sample_unit_price: evidence quote not found in the reply text"]),
]


def test_quote_fields_through_the_stub_llm():
    for name, fixture, raw, want, warnings in QUOTE_CASES:
        result, entry = run_quote(fixture, raw)
        for field, value in want.items():
            got = entry[field]
            if field == "bulk_tiers" and got:
                got = [{k: t[k] for k in ("min_qty", "unit_price")} for t in got]
            assert got == value, (name, field, got, value)
        assert entry["quote_warnings"] == warnings, (name, entry["quote_warnings"])
        assert entry["status"] == raw.get("status", "pricing_provided") and entry["summary"] == "stub summary"   # names unchanged
        assert set(ea.QUOTE_KEYS) <= set(entry)                                      # every field is present, null when unstated
        for money in ea.QUOTE_MONEY_FIELDS:                                           # money is a string that parses as a Decimal
            assert entry[money] is None or (isinstance(entry[money], str) and ea.Decimal(entry[money]) >= 0), (name, money)
        proven = [f for f in ea.QUOTE_EVIDENCE_FIELDS if entry[f] is not None]
        assert sorted(entry["evidence"]) == sorted(proven), (name, entry["evidence"])  # evidence for every non-null field only
        for field, quote in entry["evidence"].items():
            assert ea._norm(quote) in ea._norm(read_fixture(fixture)), (name, field)
    # persisted under SKU + listing, as a plain JSON-safe entry
    state = {}
    draft = ea.make_draft(cand("A1", "a@x.com", url="http://shop/a1"))
    run_quote("quote_bulk_tiers.txt", QUOTE_CASES[2][2], state, draft)
    assert list(state) == ["A1|shop/a1"] and json.loads(json.dumps(state))["A1|shop/a1"]["bulk_tiers"][1]["min_qty"] == 500


VERIFY_TABLE = [  # (what the stub returned, the field, expected value, warning that must appear (or None))
    ({"sample_unit_price": "3.50", "evidence": {"sample_unit_price": "Sample price is USD 3.20 per piece"}}, "sample_unit_price", None,
     "sample_unit_price: the number 3.50 does not appear in its evidence quote"),
    ({"sample_unit_price": "3.20"}, "sample_unit_price", None, "sample_unit_price: no evidence quote given"),
    ({"sample_unit_price": "3.20", "evidence": {"sample_unit_price": "  "}}, "sample_unit_price", None, "no evidence quote given"),
    ({"sample_unit_price": "3.20", "evidence": {"sample_unit_price": "sample   PRICE is usd 3.20\nper piece"}}, "sample_unit_price", "3.20", None),
    ({"sample_unit_price": 3.2, "evidence": {"sample_unit_price": "Sample price is USD 3.20 per piece"}}, "sample_unit_price", "3.2", None),
    ({"sample_unit_price": "$3.20", "evidence": {"sample_unit_price": "Sample price is USD 3.20 per piece"}}, "sample_unit_price", "3.20", None),
    ({"sample_unit_price": "abc", "evidence": {"sample_unit_price": "USD 3.20"}}, "sample_unit_price", None, "'abc' is not an amount"),
    ({"sample_unit_price": "-3.20", "evidence": {"sample_unit_price": "USD 3.20"}}, "sample_unit_price", None, "is not an amount"),
    ({"sample_quantity": 0, "evidence": {"sample_quantity": "5 pcs"}}, "sample_quantity", None, "is not a positive whole number"),
    ({"sample_quantity": 2.5, "evidence": {"sample_quantity": "5 pcs"}}, "sample_quantity", None, "is not a positive whole number"),
    ({"sample_quantity": "ten", "evidence": {"sample_quantity": "5 pcs"}}, "sample_quantity", None, "is not a positive whole number"),
    ({"sample_quantity": True, "evidence": {"sample_quantity": "5 pcs"}}, "sample_quantity", None, "is not a positive whole number"),
    ({"sample_quantity": 5.0, "evidence": {"sample_quantity": "5 pcs are available as samples"}}, "sample_quantity", 5, None),
    ({"sample_quantity": "5", "evidence": {"sample_quantity": "5 pcs are available as samples"}}, "sample_quantity", 5, None),
    ({"sample_quantity": 50, "evidence": {"sample_quantity": "5 pcs are available as samples"}}, "sample_quantity", None, "the number 50 does not appear"),
    ({"sample_quantity": 5, "evidence": {"sample_quantity": "five pieces"}}, "sample_quantity", None, "evidence quote not found"),
    ({"lead_time_days": 5, "evidence": {"lead_time_days": "Lead time for samples is 5 days"}}, "lead_time_days", 5, None),
    ({"lead_time_days": 9, "evidence": {"lead_time_days": "Lead time for samples is 5 days"}}, "lead_time_days", None, "the number 9 does not appear"),
    ({"shipping_terms": "DDP"}, "shipping_terms", None, "shipping_terms: not found in the reply text"),
    ({"sample_available": "maybe"}, "sample_available", "unknown", "'maybe' is not yes/no/unknown"),
    ({"sample_available": None}, "sample_available", "unknown", None),
    ({"sample_unit_price": "3.20", "price_basis": "weekly", "evidence": {"sample_unit_price": "Sample price is USD 3.20 per piece"}},
     "price_basis", "unclear", "price_basis: 'weekly' is not"),
    ({"sample_unit_price": "3.20", "currency": "EUR", "evidence": {"sample_unit_price": "Sample price is USD 3.20 per piece"}},
     "currency", "USD", "currency: the model said 'EUR' but the evidence shows USD"),
    ({"sample_unit_price": "3.20", "sample_quantity": 5, "evidence": {"sample_unit_price": "Sample price is USD 3.20 per piece",
      "sample_quantity": "5 pcs are available"}, "currency": "$"}, "currency", "USD", None),
    ({"bulk_tiers": [{"min_qty": 100, "unit_price": "1.80"}]}, "bulk_tiers", None, "bulk_tiers[0].min_qty: no evidence quote given"),
    ({"bulk_tiers": [{"min_qty": 200, "unit_price": "1.80", "evidence": "100-499 pcs USD 1.80/pc"}]}, "bulk_tiers", None,
     "bulk_tiers[0].min_qty: the number 200 does not appear"),
    ({"bulk_tiers": [{"min_qty": 100, "unit_price": "9.99", "evidence": "100-499 pcs USD 1.80/pc"}]}, "bulk_tiers", None,
     "bulk_tiers[0].unit_price: the number 9.99 does not appear"),
    ({"bulk_tiers": ["100 pcs"]}, "bulk_tiers", None, "bulk_tiers[0]: needs a positive whole min_qty"),
]


def test_numbers_are_verified_in_code():
    reply_fixture = "quote_per_unit.txt"
    for raw, field, want, warning in VERIFY_TABLE:
        reply = read_fixture("quote_bulk_tiers.txt") if "bulk_tiers" in raw else read_fixture(reply_fixture)
        got = ea.verify_quote({"status": "pricing_provided", **raw}, reply)
        value = got[field]
        assert value == want, (raw, field, value, want)
        if warning:
            assert any(warning in w for w in got["quote_warnings"]), (raw, got["quote_warnings"])
        else:
            assert got["quote_warnings"] == [], (raw, got["quote_warnings"])
    # a nulled field leaves no evidence behind; the others keep theirs
    got = ea.verify_quote({"sample_quantity": 5, "sample_unit_price": "9.99",
                           "evidence": {"sample_quantity": "5 pcs are available as samples", "sample_unit_price": "USD 9.99"}},
                          read_fixture(reply_fixture))
    assert got["sample_unit_price"] is None and list(got["evidence"]) == ["sample_quantity"]
    # mixed currencies in the evidence never pick one
    mixed = ea.verify_quote({"sample_unit_price": "3.20", "shipping_cost": "80", "evidence": {
        "sample_unit_price": "USD 3.20", "shipping_cost": "RMB 80"}}, "Sample USD 3.20 each. Shipping RMB 80.")
    assert mixed["currency"] == "unknown" and any("mixes currencies" in w for w in mixed["quote_warnings"])


def test_currency_detection_follows_the_stated_markers():
    table = [("USD 3.20", "USD"), ("US$3.20", "USD"), ("$3.20", "USD"), ("usd 3.20", "USD"), ("RMB 22", "CNY"), ("CNY 22", "CNY"),
             ("¥22", "CNY"), ("EUR 5", "EUR"), ("€5", "EUR"), ("22 yuan", "unknown"), ("22 dollars", "unknown")]
    for evidence, want in table:
        number = re.search(r"\d+(?:\.\d+)?", evidence)[0]
        got = ea.verify_quote({"sample_unit_price": number, "evidence": {"sample_unit_price": evidence}}, f"Our price: {evidence} each.")
        assert got["sample_unit_price"] == number and got["currency"] == want, (evidence, got["currency"], got["quote_warnings"])


CHINESE_MARKERS = [  # (evidence, expected currency)
    ("3.2美元", "USD"), ("3.2 美元", "USD"), ("美元3.2", "USD"), ("美元 3.2", "USD"),            # 美元 is USD, never CNY
    ("3.2美金", "USD"), ("3.2 美金", "USD"), ("美金3.2", "USD"), ("美金 3.2", "USD"),
    ("每件3.2美元", "USD"), ("报价3.2美金/件", "USD"), ("USD 3.2", "USD"), ("3.2$", "USD"),
    ("22人民币", "CNY"), ("人民币22", "CNY"), ("人民币 22", "CNY"),
    ("22元", "CNY"), ("22 元", "CNY"), ("每件22元", "CNY"), ("22.5元/件", "CNY"), ("RMB 22", "CNY"), ("¥22", "CNY"), ("￥22", "CNY"),
    ("每单元22", "unknown"), ("22个单元", "unknown"), ("22元件", "unknown"), ("元22", "unknown"), ("22 dollars", "unknown"),   # not markers
]


def test_chinese_currency_markers():
    for evidence, want in CHINESE_MARKERS:
        number = re.search(r"\d+(?:\.\d+)?", evidence)[0]
        got = ea.verify_quote({"sample_unit_price": number, "evidence": {"sample_unit_price": evidence}}, f"您好 {evidence} 谢谢")
        assert got["sample_unit_price"] == number and got["currency"] == want, (evidence, got["currency"], got["quote_warnings"])
        assert got["quote_warnings"] == [], (evidence, got["quote_warnings"])
        assert ea._currency_codes(evidence) == ({want} if want != "unknown" else set()), (evidence, ea._currency_codes(evidence))
    # a USD marker is blanked before CNY is tried: 美元 / 美金 never also produce CNY, and a real 元 next to one still does
    assert ea._currency_codes("3.2美元") == {"USD"} and ea._currency_codes("3.2美金") == {"USD"}
    assert ea._currency_codes("3.2美元，约23元") == {"USD", "CNY"} and ea._currency_codes("23元(约3.2美元)") == {"USD", "CNY"}
    # The ordering itself, independent of the digit rule: even with a loose bare-元 CNY pattern, USD is matched first and
    # its marker blanked, so 美元 / 美金 are never read as CNY.
    original = ea.CURRENCY_PATTERNS
    ea.CURRENCY_PATTERNS = (original[0], ("CNY", re.compile(r"元")), original[2])
    try:
        assert ea._currency_codes("3.2美元") == {"USD"} and ea._currency_codes("美金3.2") == {"USD"}
        assert ea._currency_codes("约23元") == {"CNY"} and ea._currency_codes("3.2美元，约23元") == {"USD", "CNY"}
    finally:
        ea.CURRENCY_PATTERNS = original
    # what the model says it quoted in is understood too (a bare 元 included) and compared with the evidence
    for said, code in [("美元", "USD"), ("美金", "USD"), ("人民币", "CNY"), ("元", "CNY"), ("RMB", "CNY"), ("￥", "CNY"), ("$", "USD"), ("€", "EUR")]:
        assert ea._currency_code(said) == code, said
    ok = ea.verify_quote({"sample_unit_price": "3.2", "currency": "美元", "evidence": {"sample_unit_price": "3.2美元"}}, "单价3.2美元")
    assert ok["currency"] == "USD" and ok["quote_warnings"] == []
    wrong = ea.verify_quote({"sample_unit_price": "3.2", "currency": "人民币", "evidence": {"sample_unit_price": "3.2美元"}}, "单价3.2美元")
    assert wrong["currency"] == "USD" and wrong["quote_warnings"] == ["currency: the model said '人民币' but the evidence shows USD"]


def test_a_reply_that_mixes_chinese_currencies_is_unknown():
    reply = "样品价格每件3.2美元，运费80元。"
    mixed = ea.verify_quote({"sample_unit_price": "3.2", "shipping_cost": "80", "currency": "USD",
                             "evidence": {"sample_unit_price": "样品价格每件3.2美元", "shipping_cost": "运费80元"}}, reply)
    assert (mixed["sample_unit_price"], mixed["shipping_cost"]) == ("3.2", "80")           # the numbers are proven...
    assert mixed["currency"] == "unknown"                                                    # ...but the currency never picks a side
    assert "currency: the evidence mixes currencies" in mixed["quote_warnings"]
    assert any(w.startswith("currency: the model said 'USD'") for w in mixed["quote_warnings"])
    # both markers inside ONE evidence quote mix as well, in either order
    for evidence in ("每件3.2美元（约23元）", "约23元（每件3.2美元）", "3.2美金 / 23人民币"):
        number = re.search(r"\d+(?:\.\d+)?", evidence)[0]
        got = ea.verify_quote({"sample_unit_price": number, "evidence": {"sample_unit_price": evidence}}, evidence)
        assert got["currency"] == "unknown" and "currency: the evidence mixes currencies" in got["quote_warnings"], evidence
    # the same marker in two fields is not a mix
    same = ea.verify_quote({"sample_unit_price": "3.2", "shipping_cost": "35", "evidence": {
        "sample_unit_price": "每件3.2美元", "shipping_cost": "运费35美金"}}, "每件3.2美元，运费35美金")
    assert same["currency"] == "USD" and same["quote_warnings"] == []


def test_reply_cli_records_and_prints_the_verified_quote():
    with sandbox() as tmp:
        path = os.path.join(tmp, "r.xlsx")
        make_xlsx(path, [row("A1", email="a@x.com", url="http://shop/a1")])
        raw = {"status": "pricing_provided", "summary": "gave a quote", **QUOTE_CASES[0][2],
               "bulk_tiers": [{"min_qty": 100, "unit_price": "9.99", "evidence": "invented tier text"}]}
        ea.summarize_reply = lambda text: raw
        out = run_cli(["reply", "A1", "--results", path, "--file", os.path.join(FIXTURES, "quote_per_unit.txt")])
        assert "A1 (Acme): pricing_provided" in out and "sample_quantity=5" in out and "sample_unit_price=3.20" in out
        assert "WARNING (field set to null): bulk_tiers[0].min_qty: evidence quote not found in the reply text" in out
        entry = read_state()["A1|shop/a1"]
        assert entry["sample_quantity"] == 5 and entry["bulk_tiers"] is None and entry["status"] == "pricing_provided"
        # a second reply is appended and the whole thread is re-verified against the combined text
        raw2 = {"status": "pricing_provided", "summary": "more", "moq": 100, "evidence": {"moq": "MOQ is 100 pcs"}}
        ea.summarize_reply = lambda text: raw2 if "--- reply 2" in text else raw
        run_cli(["reply", "A1", "--results", path, "--file", os.path.join(FIXTURES, "quote_bulk_tiers.txt")])
        entry = read_state()["A1|shop/a1"]
        assert entry["moq"] == 100 and entry["sample_quantity"] is None and len(entry["replies"]) == 2
        # a status outside the existing names is refused and nothing is stored
        before = json.dumps(read_state(), sort_keys=True)
        ea.summarize_reply = lambda text: {"status": "great_news", "summary": "x"}
        try:
            run_cli(["reply", "A1", "--results", path, "--file", os.path.join(FIXTURES, "quote_rmb.txt")])
            raise AssertionError("accepted an unknown status")
        except ValueError as e:
            assert "bad status" in str(e)
        assert json.dumps(read_state(), sort_keys=True) == before


def test_parse_summary_and_report_rows():
    ok = ea.parse_summary('Sure: {"status": "needs_info", "summary": "Wants quantity."} done')
    assert ok == {"status": "needs_info", "summary": "Wants quantity."}
    d = ea.make_draft(cand("B", "", maker=""))
    r = ea.build_report([d], {d["key"]: {"sent": "t", "reply": "x", **ok}})[0]
    assert r[5] == "yes (inquiry form, t)" and r[6] == "yes" and r[7] == "needs_info" and r[4] == "inquiry form submitted"


if __name__ == "__main__":
    for name, fn in sorted(globals().items()):
        if name.startswith("test_"):
            fn()
    print("email agent tests passed")
