#!/usr/bin/env python3
"""
Download YARDI attachments via IMAP
Direct connection to mikereports@aftonprop.com mailbox
CRITICAL: Always uses mikereports@aftonprop.com, NEVER primary email
"""

import os
import imaplib
import email
import logging
from pathlib import Path
from datetime import datetime
import getpass

# Setup logging
SCRIPT_DIR = Path(__file__).parent
LOG_DIR = SCRIPT_DIR / "logs"
LOG_DIR.mkdir(exist_ok=True)

logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler(LOG_DIR / f"imap_download_{datetime.now().strftime('%Y%m%d_%H%M%S')}.log"),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger(__name__)

# Configuration
OUTPUT_DIR = Path(r"C:\Users\MichaelOzery\OneDrive - Afton Properties\Old Dropbox\My PC (DESKTOP-5D77V89)\Mike Ozery\Dashboard Data")
TARGET_EMAIL = "mikereports@aftonprop.com"  # CRITICAL: NEVER primary email
SENDER_EMAIL = "cdr@yardi.com"
IMAP_SERVER = "outlook.office365.com"
IMAP_PORT = 993

OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

def download_via_imap(email_address, password):
    """Download attachments via IMAP"""
    logger.info("=" * 70)
    logger.info("DOWNLOAD YARDI ATTACHMENTS VIA IMAP")
    logger.info(f"CRITICAL: Using {TARGET_EMAIL} ONLY")
    logger.info("=" * 70)

    downloaded = 0
    skipped = 0
    errors = 0

    try:
        logger.info("\n[IMAP CONNECTION]")
        logger.info(f"Connecting to {IMAP_SERVER}:{IMAP_PORT}...")

        # Connect to IMAP server
        imap = imaplib.IMAP4_SSL(IMAP_SERVER, IMAP_PORT)
        logger.info("  SSL connection established")

        # Login
        logger.info(f"Logging in as {email_address}...")
        try:
            imap.login(email_address, password)
            logger.info("  [OK] Authentication successful")
        except imaplib.IMAP4.error as e:
            logger.error(f"  [ERROR] Authentication failed: {e}")
            logger.error("  Check email address and password/app password")
            return downloaded, errors + 1

        # Select Inbox
        logger.info("\nSelecting Inbox...")
        status, mailbox_list = imap.select('INBOX')
        if status != 'OK':
            logger.error("  [ERROR] Could not select Inbox")
            return downloaded, errors + 1

        message_count = int(mailbox_list[0])
        logger.info(f"  [OK] Inbox has {message_count} messages")

        # Search for emails from YARDI
        logger.info(f"\nSearching for emails from {SENDER_EMAIL}...")
        status, message_ids = imap.search(None, f'FROM "{SENDER_EMAIL}"')

        if status != 'OK':
            logger.error("  [ERROR] Search failed")
            return downloaded, errors + 1

        message_ids = message_ids[0].split()
        logger.info(f"  Found {len(message_ids)} emails from {SENDER_EMAIL}")

        if len(message_ids) == 0:
            logger.warning("  No emails found - check sender email address")
            return downloaded, errors

        # Process each email
        logger.info("\n[EMAIL PROCESSING]")

        for i, msg_id in enumerate(message_ids, 1):
            try:
                # Fetch email
                status, msg_data = imap.fetch(msg_id, '(RFC822)')
                if status != 'OK':
                    logger.error(f"  [{i}] Could not fetch message")
                    errors += 1
                    continue

                # Parse email
                msg = email.message_from_bytes(msg_data[0][1])

                email_from = msg.get('From', 'Unknown')
                email_subject = msg.get('Subject', 'No Subject')
                email_date = msg.get('Date', 'Unknown')

                logger.info(f"\n[Email {i}] {email_subject}")
                logger.info(f"  From: {email_from}")
                logger.info(f"  Date: {email_date}")

                # Check for attachments
                has_attachments = False
                for part in msg.walk():
                    if part.get_content_disposition() == 'attachment':
                        has_attachments = True
                        break

                if not has_attachments:
                    logger.info("  No attachments")
                    continue

                logger.info("  Attachments found:")

                # Download attachments
                for part in msg.walk():
                    if part.get_content_disposition() != 'attachment':
                        continue

                    filename = part.get_filename()

                    # Skip if no filename
                    if not filename:
                        logger.info("    Skipping (no filename)")
                        continue

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

                    # Save attachment
                    try:
                        payload = part.get_payload(decode=True)
                        with open(filepath, 'wb') as f:
                            f.write(payload)
                        logger.info(f"    [OK] Downloaded: {filename}")
                        downloaded += 1

                    except Exception as e:
                        logger.error(f"    Error saving attachment: {e}")
                        errors += 1

            except Exception as e:
                logger.error(f"  [{i}] Error processing email: {e}")
                errors += 1
                continue

        # Close connection
        logger.info("\n[CLEANUP]")
        imap.close()
        imap.logout()
        logger.info("  [OK] Disconnected from IMAP server")

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
    logger.info("=" * 70)
    logger.info("YARDI IMAP DOWNLOADER")
    logger.info("=" * 70)

    print("\n" + "=" * 70)
    print("YARDI REPORT DOWNLOADER (IMAP)")
    print("=" * 70)
    print(f"\nTarget mailbox: {TARGET_EMAIL}")
    print("\nNOTE: If you have 2FA enabled on your account:")
    print("  1. Create an app password in Office 365")
    print("  2. Use the app password (not your regular password)")
    print("\nHow to create app password:")
    print("  https://support.microsoft.com/en-us/account-billing/using-app-passwords-with-apps-that-dont-support-multi-factor-authentication")
    print("\n" + "=" * 70)

    email_addr = input(f"\nEmail address [{TARGET_EMAIL}]: ").strip()
    if not email_addr:
        email_addr = TARGET_EMAIL

    password = getpass.getpass(f"Password (or app password): ")

    if not password:
        logger.error("Password required")
        exit(1)

    logger.info(f"\nStarting download for {email_addr}...")
    logger.info("")

    downloaded, errors = download_via_imap(email_addr, password)

    if errors == 0:
        logger.info("\n[SUCCESS] All attachments downloaded")
        exit(0)
    else:
        logger.error(f"\n[FAILED] {errors} error(s) occurred")
        exit(1)
