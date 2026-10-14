"""Stub agent used by the tests and by nothing else: no network, no real work. The 'search' mode imitates
sourcing_agent.py for the UI: the same progress, cost and stop lines (copied from that file's print and log calls), a
small results workbook at --output saved after every product, and a report in the working folder. What it does is chosen
by the UI_STUB_SEARCH environment variable (ok, nokey, billing, slow, stubborn), which only this file reads."""
import argparse
import os
import subprocess
import sys
import time
from datetime import datetime
from pathlib import Path

p = argparse.ArgumentParser()
p.add_argument("mode", choices=["ok", "fail", "slow", "secret", "hostile", "term", "search"])
p.add_argument("--sku", nargs="+")
p.add_argument("--results")
p.add_argument("--input")
p.add_argument("--limit", type=int)
p.add_argument("--output")
p.add_argument("--add-builtin", action="store_true")
p.add_argument("--changes-state", action="store_true")
a = p.parse_args()


def search() -> None:
    how = os.environ.get("UI_STUB_SEARCH", "ok")
    if how == "nokey":
        print("Missing required environment variable(s): ANTHROPIC_API_KEY, NIMBLE_API_KEY")
        print("Add them to a .env file in this folder (or set them in the shell), e.g.:")
        print("  ANTHROPIC_API_KEY=sk-ant-...")
        print("  NIMBLE_API_KEY=...")
        sys.exit(1)
    sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
    import openpyxl
    from ui import search as ui_search
    from ui.config import Settings
    s = Settings(project=Path.cwd(), data=Path.cwd() / "ui_data")
    builtin = ui_search.builtin_products(s)
    rows = ui_search.input_products(s, a.input.replace("\\", "/")) if a.input else []
    if a.sku:                                       # like the agent: looked up in the input rows, then the built-in list
        products = [next(x for x in rows + builtin if x[0] == k) for k in a.sku]
    elif a.input:
        products = rows[:a.limit] + (builtin if a.add_builtin else [])
    else:
        products = builtin
    if how in ("slow", "stubborn"):
        if how == "stubborn":                       # ignore polite stop requests, so only a forced tree kill ends it
            import signal
            for name in ("SIGTERM", "SIGINT", "SIGBREAK"):
                if hasattr(signal, name):
                    signal.signal(getattr(signal, name), signal.SIG_IGN)
        child = subprocess.Popen([sys.executable, "-c", "import time; time.sleep(60)"])
        print(f"child pid {child.pid}", flush=True)
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    report = f"sourcing_report_{stamp}.md"
    print("Competitive Sourcing Research POC")
    print(f"Model: stub  |  Products: {len(products)}  |  Sites: aliexpress.com, alibaba.com, made-in-china.com")
    print("Tool call budget: 13 per product")
    print(f"Saving results after every product to {report} / {a.output}")
    print("Running up to 3 products at once", flush=True)

    head = ["SKU", "Product", "Keyword", "L-Com Unit Price", "L-Com Price Source", "L-Com Price Date", "Recommendation",
            "Recommended Manufacturer", "Recommended Email", "Recommended URL", "Recommended Unit Price", "Ordering Note",
            "Manufacturer 1 Name", "Manufacturer 1 Accuracy", "Manufacturer 1 URL", "Recommended Listing Title"]
    rows, total_in, total_out, total_cost, stopped = [], 0, 0, 0.0, ""

    def save() -> None:
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Results"
        ws.append(head)
        for r in rows:
            ws.append(r)
        wb.save(a.output)
        Path(report).write_text(f"# Competitive Sourcing Research Report\n\nProducts: {len(rows)} of {len(products)}\n", encoding="utf-8")

    for i, (sku, desc) in enumerate(products):
        tag = f"[{sku}] "
        print(f"{tag}Product {i + 1} of {len(products)} ...", flush=True)
        if how in ("slow", "stubborn") and i:
            time.sleep(30)
        if how == "billing" and i >= 1:
            reason = stopped and f"not run - batch stopped early ({stopped})"
            if not stopped:
                stopped = reason = "BadRequestError: Your credit balance is too low to access the Anthropic API."
                print(tag + "!" * 72)
                print(f"{tag}!!! BILLING / CREDIT / AUTH ERROR - stopping the batch: {reason}")
                print(f"{tag}!!! Products already finished are saved; the rest will not be run.")
                print(tag + "!" * 72)
            print(f"{tag}  Error researching this product: {reason}", flush=True)
            rows.append([sku, "", desc, 10.0, "L-Com list", "2026-01-01", f"Error researching this product: {reason} - re-run with --sku"] + [""] * 9)
        else:
            tin, tout = 12000 + i, 900
            cost = tin / 1_000_000 * 1.00 + tout / 1_000_000 * 5.00
            total_in, total_out, total_cost = total_in + tin, total_out + tout, total_cost + cost
            print(tag + "-" * 72)
            print(f"{tag}{sku} - \"{desc}\"  (5 tool calls used)")
            if i % 2 == 0:
                url = f"https://www.alibaba.com/product/{i}"
                rec = f"Buy from Candidate 1 (Stub Maker {i}): 90% accuracy at $1.25/unit (stated), 87% cheaper than L-Com."
                rows.append([sku, "", desc, 10.0, "L-Com list", "2026-01-01", rec, f"Stub Maker {i}", "", url, 1.25,
                             "Check the pack size" if i == 0 else "", f"Stub Maker {i}", 90, url, "Stub listing"])
            else:
                rec = "No candidate qualifies - price margin only (L-Com unit price $10.00). Do not source from any of these."
                rows.append([sku, "", desc, 10.0, "L-Com list", "2026-01-01", rec] + [""] * 9)
            print(f"{tag}  Recommendation: {rec}")
            print(f"{tag}  Tokens:     in={tin} out={tout}  (~${cost:.4f})")
            print(f"{tag}  Time:       41s  (Nimble 30s, Anthropic+other 11s, 5 tool calls)")
            print(f"{tag}  Stopped:    model finished", flush=True)
        save()

    print("-" * 72)
    print(f"Total tokens: in={total_in} out={total_out}  |  Estimated total cost: ~${total_cost:.4f}")
    if stopped:
        print(f"BATCH STOPPED EARLY: {stopped}. Finished products were saved; the rest were not run.")
    print(f"Report written to: {os.path.abspath(report)}")
    print(f"Excel results written to: {os.path.abspath(a.output)}", flush=True)


sku_text = " ".join(a.sku) if a.sku else None
if a.mode == "search":
    search()
    sys.exit(0)
print(f"stub start mode={a.mode} sku={sku_text}", flush=True)
if a.mode == "ok":
    print("stub done", flush=True)
elif a.mode == "fail":
    print("stub failing on purpose", file=sys.stderr, flush=True)
    sys.exit(3)
elif a.mode == "secret":
    print("ANTHROPIC_API_KEY=sk-ant-abcdefgh12345678 contact boss@example.com phone +1 214 555 0100", flush=True)
elif a.mode == "hostile":
    print("<script>alert('x')</script> <img src=x onerror=alert(1)>", flush=True)
elif a.mode == "term":
    print("L-Com price 12.00 | LCOM SKU AB-1 | lcom_prices.csv | l com | LCom unit price | workflow lcom-sourcing_results_20260101_100000", flush=True)
elif a.mode == "slow":
    child = subprocess.Popen([sys.executable, "-c", "import time; time.sleep(60)"])   # a grandchild, to prove tree kill
    print(f"child pid {child.pid}", flush=True)
    for i in range(60):
        print(f"tick {i}", flush=True)
        time.sleep(1)
