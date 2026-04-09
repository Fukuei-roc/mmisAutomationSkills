from __future__ import annotations

import argparse
import json
import subprocess
import sys
from pathlib import Path


SCRIPT_PATH = Path(__file__).resolve().parent / "playwright_linked_fault_notices_query.py"


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="Run MMIS 1A linked fault notices query")
    parser.add_argument("--work-order-no", required=True, help="Work order number, e.g. 115-1A-23391")
    return parser


def main() -> int:
    args = build_parser().parse_args()
    command = [sys.executable, str(SCRIPT_PATH), "--work-order-no", args.work_order_no]
    env = dict(**__import__("os").environ)
    env["PYTHONIOENCODING"] = "utf-8"
    completed = subprocess.run(command, capture_output=True, text=True, encoding="utf-8", errors="replace", env=env)

    if completed.stdout:
        print(completed.stdout.strip())
    if completed.stderr:
        print(completed.stderr.strip(), file=sys.stderr)
    return completed.returncode


if __name__ == "__main__":
    sys.exit(main())
