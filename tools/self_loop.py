import subprocess  # nosec B404
import sys
import time
import datetime
import os
import shutil
from pathlib import Path

INTERVAL_MIN = int(os.getenv("IMPROVE_INTERVAL_MIN", "60"))
BACKUP_DIR = "backups"
os.makedirs(BACKUP_DIR, exist_ok=True)

REPO = Path(__file__).resolve().parents[1]
IMPROVER = REPO / "tools" / "auto_improver.py"


def make_backup() -> None:
    ts = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    dst = os.path.join(BACKUP_DIR, f"backup_{ts}")
    os.makedirs(dst, exist_ok=True)
    for fn in ("agent.py", "cloud.py", "tools/auto_improver.py"):
        if os.path.exists(fn):
            shutil.copy(fn, os.path.join(dst, fn.replace("/", "_")))
    print(f"[BACKUP] {dst}")


def loop_once() -> None:
    # Static, trusted argv; explicit interpreter and cwd.
    # nosec B603: argv is constant, no untrusted input; shell=False.
    rc = subprocess.call([sys.executable, str(IMPROVER), "--apply"], cwd=str(REPO))  # nosec B603
    if rc == 0:
        print("[LOOP] ✅ Applied successfully.")
    else:
        print("[LOOP] ❌ Tests failed. Skipping apply.")

    print(f"[LOOP] sleep {INTERVAL_MIN}m")
    time.sleep(INTERVAL_MIN * 60)


if __name__ == "__main__":
    while True:
        make_backup()
        loop_once()
