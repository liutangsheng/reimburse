from __future__ import annotations

import sys

import main as cli_main
import web_ui


def run() -> None:
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    if hasattr(sys.stderr, "reconfigure"):
        sys.stderr.reconfigure(encoding="utf-8", errors="replace")
    if len(sys.argv) <= 1:
        web_ui.main()
    else:
        cli_main.main()


if __name__ == "__main__":
    run()
