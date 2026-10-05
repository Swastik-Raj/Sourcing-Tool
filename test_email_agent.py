"""Offline checks for email_agent.py. Run: python test_email_agent.py"""
import os
import tempfile

import openpyxl

import email_agent as ea

HEADER = ["SKU", "Product", "Keyword", "L-Com Unit Price", "Recommendation", "Recommended Manufacturer",
          "Recommended Email", "Recommended URL", "Recommended Unit Price", "Ordering Note",
          "Manufacturer 1 URL", "Manufacturer 1 MOQ", "Manufacturer 2 URL", "Manufacturer 2 MOQ"]


def make_xlsx(path):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Results"
    ws.append(HEADER)
    ws.append(["A1", "Couplers", "Cat6 Coupler", 20, "", "Acme", "sales@acme.com", "http://b", 2.5, "confirm shielded",
               "http://a", "100", "http://b", "500"])
    ws.append(["B2", "Cables", "Cat5e Cable", 9, "", None, None, "http://only-url", 1.0, None, "", "", "", ""])
    ws.append(["C3", "Nothing", "No rec", 5, "no qualifier", None, None, None, None, None, "", "", "", ""])
    # A crashed product is skipped even if a stale recommendation were somehow filled in.
    ws.append(["E5", "Crashed", "Err", 5, "Error researching this product: RuntimeError: boom - re-run with --sku",
               "Acme", None, "http://e", 1.0, None, "", "", "", ""])
    wb.save(path)


def test_load_and_drafts():
    with tempfile.TemporaryDirectory() as d:
        path = os.path.join(d, "r.xlsx")
        make_xlsx(path)
        cands = ea.load_candidates(path)
    assert [c["sku"] for c in cands] == ["A1", "B2"]  # C3 has no recommendation
    assert cands[0]["moq"] == "500"  # matched by URL to Manufacturer 2
    drafts = [ea.make_draft(c) for c in cands]
    assert "Acme" in drafts[0]["body"] and "MOQ 500" in drafts[0]["body"] and "confirm shielded" in drafts[0]["body"]
    assert "Plano, TX" in drafts[0]["body"] and "No email available" not in drafts[0]["body"]
    assert "No email available" in drafts[1]["body"] and "http://only-url" in drafts[1]["body"]
    assert "{" not in drafts[0]["body"] + drafts[1]["body"] + drafts[0]["subject"]  # no unfilled placeholders


def test_send_requires_confirmation():
    ea.STATE_PATH = os.path.join(tempfile.mkdtemp(), "state.json")
    drafts = [ea.make_draft({"sku": s, "product": "p", "description": "d", "manufacturer": "m", "email": e,
                             "url": "u", "unit_price": None, "moq": "", "note": ""})
              for s, e in [("A", "a@x.com"), ("B", ""), ("C", "c@x.com")]]
    sent = []
    state = {}
    # per-email: yes to A, no to C; B has no email so is never offered
    answers = iter(["y", "n"])
    n = ea.confirm_and_send(drafts, state, send=sent.append, ask=lambda _: next(answers))
    assert n == 1 and [d["sku"] for d in sent] == ["A"] and "sent" in state["A"]
    # batch: anything but 'yes' sends nothing; already-sent A is skipped
    assert ea.confirm_and_send(drafts, state, yes=True, send=sent.append, ask=lambda _: "no") == 0
    assert ea.confirm_and_send(drafts, state, yes=True, send=sent.append, ask=lambda _: "yes") == 1
    assert [d["sku"] for d in sent] == ["A", "C"]


def test_parse_summary_and_report():
    ok = ea.parse_summary('Sure: {"status": "needs_info", "summary": "Wants quantity."} done')
    assert ok == {"status": "needs_info", "summary": "Wants quantity."}
    try:
        ea.parse_summary('{"status": "maybe", "summary": ""}')
        raise AssertionError("bad status accepted")
    except ValueError:
        pass
    d = ea.make_draft({"sku": "B", "product": "p", "description": "d", "manufacturer": "", "email": "",
                       "url": "u", "unit_price": None, "moq": "", "note": ""})
    row = ea.build_report([d], {"B": {"sent": "t", "reply": "x", **ok}})[0]
    assert row[4] == "yes (inquiry form, t)" and row[5] == "yes" and row[6] == "needs_info"


if __name__ == "__main__":
    test_load_and_drafts()
    test_send_requires_confirmation()
    test_parse_summary_and_report()
    print("email agent tests passed")
