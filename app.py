from __future__ import annotations

import sys

import main as cli_main
import web_ui


def run() -> None:
    if len(sys.argv) <= 1:
        web_ui.main()
    else:
        cli_main.main()


if __name__ == "__main__":
    run()
