#!/usr/bin/env python3
"""
Afton Dashboard - daily auto-update.
Rebuilds data/dashboard.json from live sources and pushes it to GitHub Pages.

Run by Windows Task Scheduler each morning. Requires GITHUB_TOKEN in the
environment (already set on this machine) OR a git credential helper.
"""
import os, subprocess, logging
from datetime import datetime
from pathlib import Path

import build_metrics  # same folder

SCRIPT_DIR = Path(__file__).parent
LOG_DIR = SCRIPT_DIR / "logs"
LOG_DIR.mkdir(exist_ok=True)
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s %(levelname)s %(message)s",
    handlers=[logging.FileHandler(LOG_DIR / "update_dashboard.log"), logging.StreamHandler()],
)
log = logging.getLogger(__name__)

REPO  = "github.com/mikeyozery-beep/afton-dashboard.git"
TOKEN = os.getenv("GITHUB_TOKEN")


def git(*args, check=True):
    return subprocess.run(["git", *args], cwd=SCRIPT_DIR, check=check,
                          capture_output=True, text=True, timeout=60)


def main():
    log.info("=" * 60)
    log.info("Rebuilding dashboard.json from live sources...")
    out = build_metrics.main()
    if out.get("errors"):
        log.warning("Extractor warnings: %s", out["errors"])

    # stage the data + any code changes
    git("add", "data/dashboard.json")
    status = git("status", "--short").stdout.strip()
    if "data/dashboard.json" not in status:
        log.info("No change in dashboard.json - nothing to publish.")
        return

    msg = f"Auto-update dashboard data - {datetime.now():%Y-%m-%d %H:%M}"
    git("commit", "-m", msg)

    # push (use token if available, else rely on configured credentials)
    if TOKEN:
        url = f"https://{TOKEN}@{REPO}"
        git("push", url, "HEAD:main")
    else:
        git("push", "origin", "main")
    log.info("Published: https://mikeyozery-beep.github.io/afton-dashboard/dashboard_live.html")


if __name__ == "__main__":
    main()
