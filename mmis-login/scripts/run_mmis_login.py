from __future__ import annotations

import os
import subprocess
import sys
from pathlib import Path


MMIS_CLIENT = Path(
    r"C:\Users\NMMIS\.codex\skills\mmis-query-unprocessed-fault-notices\scripts\mmisClient.py"
)


def main() -> int:
    env = os.environ.copy()
    env["PYTHONIOENCODING"] = "utf-8"
    process = subprocess.run(
        [sys.executable, str(MMIS_CLIENT), "login"],
        check=False,
        env=env,
    )
    return process.returncode


if __name__ == "__main__":
    raise SystemExit(main())
