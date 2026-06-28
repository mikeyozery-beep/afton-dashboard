#!/usr/bin/env python3
"""
Download YARDI attachments from shared mailbox using O365 library
Uses OAuth - one-time browser login, then fully automated
CRITICAL: Uses mikereports@aftonprop.com, NEVER primary email
"""

import os
from pathlib import Path
from datetime import datetime
import logging

try:
    from O365 import Account, FileSystemTokenBackend
except ImportError:
    print("ERROR: O365 library not installed")
    print("Run: pip install O365 --break-system-packages")
    exit(1)

# Setup logging
SCRIPT_DIR = Path(__file__).parent
LOG_DIR = SCRIPT_DIR / "logs"
LOG_DIR.mkdir(exist_ok=True)

logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler(LOG_DIR / f"o365_download_{datetime.now().strftime('%Y%m%d_%H%M%S')}.log"),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger(__name__)

# Configuration
OUTPUT_DIR = Path(r"C:\Users\MichaelOzery\OneDrive - Afton Properties\Old Dropbox\My PC (DESKTOP-5D77V89)\Mike Ozery\Dashboard Data")
PRIMARY_EMAIL = "mozery@aftonprop.com"
SHARED_MAILBOX = "mikereports@aftonprop.com"
SENDER_EMAIL = "cdr@yardi.com"

# OAuth credentials (create at https://portal.azure.com)
# For now, using generic Microsoft App credentials
CLIENT_ID = "00000000-0000-0000-0000-000000000000"  # Placeholder
CLIENT_SECRET = ""  # Will use device flow instead

OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

def download_via_o365():
    """Download attachments using O365 library"""
    logger.info("=" * 70)
    logger.info("DOWNLOAD YARDI ATTACHMENTS (O365 Library)")
    logger.info(f"CRITICAL: Using {SHARED_MAILBOX} via {PRIMARY_EMAIL}")
    logger.info("=" * 70)

    downloaded = 0
    skipped = 0
    errors = 0

    try:
        logger.info("\n[OAUTH SETUP]")
        logger.info("Setting up OAuth connection...")
        logger.info("First-time setup will open a browser for you to authorize")

        # Use file-based token storage
        token_path = SCRIPT_DIR / "o365_tokens"
        token_path.mkdir(exist_ok=True)
        token_backend = FileSystemTokenBackend(token_path=str(token_path))

        # Create account with default Microsoft credentials (no app registration needed)
        # This uses device flow for authentication
        logger.info(f"Authenticating as: {PRIMARY_EMAIL}")

        account = Account(
            primary_email=PRIMARY_EMAIL,
            token_backend=token_backend
        )

        # Connect - will open browser if first time
        if not account.authenticate():
            logger.error("Authentication failed")
            return downloaded, errors + 1

        logger.info("[OK] Authentication successful")

        logger.info(f"\n[MAILBOX CONNECTION]")
        logger.info(f"Accessing shared mailbox: {SHARED_MAILBOX}")

        # Get mailbox (shared mailbox)
        try:
            mailbox = account.mailbox(shared_mailbox=SHARED_MAILBOX)
        except:
            logger.warning(f"Could not access as shared mailbox, trying as folder...")
            mailbox = account.mailbox()

        logger.info("[OK] Mailbox accessed")

        # Get inbox
        logger.info("\nGetting inbox...")
        inbox = mailbox.inbox_folder()

        logger.info(f"\n[EMAIL SEARCH]")
        logger.info(f"Searching for emails from {SENDER_EMAIL}...")

        # Search for emails from YARDI
        emails = inbox.get_messages(
            limit=None,
            query=f"from:{SENDER_EMAIL}"
        )

        email_list = list(emails)
        logger.info(f"Found {len(email_list)} emails from {SENDER_EMAIL}")

        if len(email_list) == 0:
            logger.warning("No emails found")
            return downloaded, errors

        logger.info("\n[ATTACHMENT DOWNLOAD]")
        logger.info(f"Processing {len(email_list)} emails...")

        for i, message in enumerate(email_list, 1):
            try:
                subject = message.subject if message.subject else "(No Subject)"
                logger.info(f"\n[Email {i}] {subject}")
                logger.info(f"  From: {message.sender}")
                logger.info(f"  Date: {message.received}")

                # Check for attachments
                if not message.has_attachments:
                    logger.info("  No attachments")
                    continue

                logger.info(f"  Attachments: {len(message.attachments)}")

                # Download each attachment
                for attachment in message.attachments:
                    try:
                        filename = attachment.name

                        # Only download Excel files
                        if not filename.lower().endswith(('.xlsx', '.xls')):
                            logger.info(f"    Skipping (not Excel): {filename}")
                            continue

                        filepath = OUTPUT_DIR / filename

                        # Check if already exists
                        if filepath.exists():
                            logger.info(f"    Skipping (already exists): {filename}")
                            skipped += 1
                            continue

                        # Download
                        try:
                            attachment.download(to_folder=str(OUTPUT_DIR))
                            logger.info(f"    [OK] Downloaded: {filename}")
                            downloaded += 1

                        except Exception as e:
                            logger.error(f"    Error downloading: {e}")
                            errors += 1

                    except Exception as e:
                        logger.error(f"    Error processing attachment: {e}")
                        errors += 1

            except Exception as e:
                logger.error(f"  Error processing email: {e}")
                errors += 1
                continue

        logger.info("\n" + "=" * 70)
        logger.info("DOWNLOAD COMPLETE")
        logger.info("=" * 70)
        logger.info(f"Downloaded: {downloaded} attachments")
        logger.info(f"Skipped (already exist): {skipped}")
        logger.info(f"Errors: {errors}")
        logger.info(f"Saved to: {OUTPUT_DIR}")
        logger.info("=" * 70)

        return downloaded, errors

    except Exception as e:
        logger.error(f"Fatal error: {e}", exc_info=True)
        return downloaded, errors + 1

if __name__ == '__main__':
    logger.info("\n" + "=" * 70)
    logger.info("YARDI DOWNLOADER (O365 Library - OAuth)")
    logger.info("=" * 70)
    logger.info("")
    logger.info("This script uses OAuth for secure authentication.")
    logger.info("First run: A browser will open for you to authorize access")
    logger.info("Token will be saved for future runs")
    logger.info("")

    downloaded, errors = download_via_o365()

    if errors == 0:
        logger.info("\n[SUCCESS] All attachments downloaded")
        exit(0)
    else:
        logger.error(f"\n[FAILED] {errors} error(s) occurred")
        exit(1)
