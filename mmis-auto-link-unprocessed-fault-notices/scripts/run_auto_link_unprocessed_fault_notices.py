from __future__ import annotations

import argparse
import json
import os
import subprocess
import sys
from pathlib import Path


SCRIPT_PATH = Path(__file__).resolve().parent / "auto_link_unprocessed_fault_notices.py"


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="Run MMIS batch auto-link for unprocessed fault notices")
    parser.add_argument("--file", help="Optional Excel file path override")
    parser.add_argument("--skip-filled", action="store_true", help="Skip rows whose column I already has data")
    return parser


def main() -> int:
    args = build_parser().parse_args()
    command = [sys.executable, str(SCRIPT_PATH)]
    if args.file:
        command.extend(["--file", args.file])
    if args.skip_filled:
        command.append("--skip-filled")

    env = dict(os.environ)
    env["PYTHONIOENCODING"] = "utf-8"
    completed = subprocess.run(command, capture_output=True, text=True, encoding="utf-8", errors="replace", env=env)
    if completed.stdout:
        print(completed.stdout.strip())
    if completed.stderr:
        print(completed.stderr.strip(), file=sys.stderr)
    return completed.returncode


if __name__ == "__main__":
    raise SystemExit(main())
