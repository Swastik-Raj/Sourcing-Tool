"""Start the local web UI:  python run_ui.py   (UI_HOST / UI_PORT override 127.0.0.1:8000)"""
import os
import sys

from ui.config import is_loopback, load_settings


def main() -> None:
    s = load_settings()
    if not is_loopback(s.host):
        if os.environ.get("UI_ALLOW_NON_LOOPBACK") != "1":
            sys.exit(f"Refusing to start on {s.host}: the UI has no login, so it only runs on this computer "
                     "(127.0.0.1). Set UI_ALLOW_NON_LOOPBACK=1 to override.")
        print("*" * 70 + f"\nWARNING: listening on {s.host} with NO LOGIN. Anyone who can reach this port can run the agents.\n" + "*" * 70)
    import uvicorn
    from ui.app import create_app
    print(f"L-Com agents UI: http://{s.host}:{s.port}/   (Ctrl+C to stop)")
    uvicorn.run(create_app(s), host=s.host, port=s.port, log_level="warning")


if __name__ == "__main__":
    main()
