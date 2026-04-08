from __future__ import annotations

import argparse
import os
import subprocess
import sys
from pathlib import Path


MMIS_CLIENT = Path(__file__).resolve().with_name("mmisClient.py")


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser()
    parser.add_argument("--format-excel", action="store_true")
    return parser


def main() -> int:
    args = build_parser().parse_args()
    command = [
        sys.executable,
        str(MMIS_CLIENT),
        "download-unprocessed-fault-reports",
    ]
    if args.format_excel:
        command.append("--format-excel")
    env = os.environ.copy()
    env["PYTHONIOENCODING"] = "utf-8"
    process = subprocess.run(command, check=False, env=env)
    return process.returncode


if __name__ == "__main__":
    raise SystemExit(main())
