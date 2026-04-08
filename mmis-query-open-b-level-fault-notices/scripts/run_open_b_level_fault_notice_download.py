from __future__ import annotations

import argparse
import os
import subprocess
import sys
from pathlib import Path


MMIS_CLIENT = Path(
    r"C:\Users\NMMIS\.codex\skills\mmis-query-unprocessed-fault-notices\scripts\mmisClient.py"
)


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser()
    parser.add_argument("--level", default="B")
    parser.add_argument("--depot", default="新竹機務段")
    return parser


def main() -> int:
    args = build_parser().parse_args()
    env = os.environ.copy()
    env["PYTHONIOENCODING"] = "utf-8"
    process = subprocess.run(
        [
            sys.executable,
            str(MMIS_CLIENT),
            "download-open-b-level-fault-reports",
            "--level",
            args.level,
            "--depot",
            args.depot,
        ],
        check=False,
        env=env,
    )
    return process.returncode


if __name__ == "__main__":
    raise SystemExit(main())
