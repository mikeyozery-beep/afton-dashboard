#!/usr/bin/env python3
"""
Afton Dashboard - Extract Metrics and Push to GitHub
Runs the extraction and commits/pushes metrics.json to GitHub
"""

import os
import json
import subprocess
import logging
from pathlib import Path
from datetime import datetime

# Setup logging
SCRIPT_DIR = Path(__file__).parent
LOG_DIR = SCRIPT_DIR / "logs"
LOG_DIR.mkdir(exist_ok=True)

logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler(LOG_DIR / f"push_github_{datetime.now().strftime('%Y%m%d_%H%M%S')}.log"),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger(__name__)

METRICS_FILE = SCRIPT_DIR / "data" / "metrics.json"
GITHUB_TOKEN = os.getenv("GITHUB_TOKEN")
GITHUB_REPO = "https://github.com/mikeyozery-beep/afton-dashboard.git"

def run_extraction():
    """Run the existing extraction script"""
    logger.info("=" * 70)
    logger.info("Running metrics extraction...")
    logger.info("=" * 70)

    try:
        # Import and run the extraction logic
        import extract_metrics_enhanced
        # The script logs its own output
        logger.info("[OK] Extraction completed")
        return True
    except Exception as e:
        logger.error(f"Extraction failed: {e}")
        return False

def push_to_github():
    """Commit and push metrics.json to GitHub"""
    logger.info("=" * 70)
    logger.info("Pushing to GitHub...")
    logger.info("=" * 70)

    if not GITHUB_TOKEN:
        logger.warning("GITHUB_TOKEN environment variable not set - skipping GitHub push")
        logger.info("To enable GitHub push, set GITHUB_TOKEN environment variable")
        return False

    if not METRICS_FILE.exists():
        logger.error(f"metrics.json not found at {METRICS_FILE}")
        return False

    try:
        # Change to script directory
        os.chdir(SCRIPT_DIR)

        # Configure git remote with token
        remote_url = GITHUB_REPO.replace("https://", f"https://{GITHUB_TOKEN}@")

        logger.info("Configuring git remote...")
        subprocess.run(
            ["git", "remote", "set-url", "origin", remote_url],
            check=True,
            capture_output=True,
            timeout=10
        )

        # Stage metrics.json
        logger.info("Staging metrics.json...")
        subprocess.run(
            ["git", "add", "data/metrics.json"],
            check=True,
            capture_output=True,
            timeout=10
        )

        # Create commit message with timestamp
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        commit_msg = f"Update metrics - {timestamp}"

        logger.info(f"Committing: {commit_msg}")
        subprocess.run(
            ["git", "commit", "-m", commit_msg],
            check=True,
            capture_output=True,
            timeout=10
        )

        # Push to main branch
        logger.info("Pushing to GitHub main branch...")
        result = subprocess.run(
            ["git", "push", "origin", "main"],
            check=True,
            capture_output=True,
            timeout=30,
            text=True
        )

        logger.info("[OK] Successfully pushed to GitHub")
        logger.info(f"Repository: {GITHUB_REPO}")
        return True

    except subprocess.CalledProcessError as e:
        logger.error(f"Git command failed: {e}")
        if e.stderr:
            logger.error(f"Error details: {e.stderr}")
        return False
    except Exception as e:
        logger.error(f"Error pushing to GitHub: {e}")
        return False

def main():
    logger.info("=" * 70)
    logger.info("AFTON DASHBOARD - EXTRACT & PUSH WORKFLOW")
    logger.info("=" * 70)

    # Step 1: Extract metrics
    if not run_extraction():
        logger.error("Extraction failed - aborting push")
        return

    # Step 2: Push to GitHub
    if push_to_github():
        logger.info("=" * 70)
        logger.info("[SUCCESS] Metrics extracted and pushed to GitHub")
        logger.info("=" * 70)
    else:
        logger.warning("Extraction succeeded but GitHub push failed")
        logger.info("Metrics are saved locally but not synced to GitHub")

if __name__ == '__main__':
    main()
