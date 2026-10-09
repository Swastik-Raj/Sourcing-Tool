"""Stub agent used by the tests and by nothing else: no network, no files, no real work."""
import argparse
import subprocess
import sys
import time

p = argparse.ArgumentParser()
p.add_argument("mode", choices=["ok", "fail", "slow", "secret", "hostile", "term"])
p.add_argument("--sku")
p.add_argument("--results")
p.add_argument("--changes-state", action="store_true")
a = p.parse_args()

print(f"stub start mode={a.mode} sku={a.sku}", flush=True)
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
