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


@contextlib.contextmanager
def sandbox(exclusions=(("OTHER", "nobody", "placeholder row"),)):
    """Temp dir for state and reports, a real-looking SHIPPING_ADDRESS, an exclusions file, no .env, no API key, SMTP
    faked. Nothing leaves the process."""
    old = (ea.HERE, ea.STATE_PATH, ea.REVIEWER_EXCLUSIONS_CSV, ea.smtplib.SMTP, ea.summarize_reply, dict(os.environ))
    tmp = tempfile.mkdtemp()
    ea.HERE, ea.STATE_PATH = tmp, os.path.join(tmp, "email_state.json")
    ea.REVIEWER_EXCLUSIONS_CSV = os.path.join(tmp, "reviewer_exclusions.csv")
    write_exclusions(ea.REVIEWER_EXCLUSIONS_CSV, exclusions)
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
