#!/usr/bin/env python3
"""
Download YARDI attachments from SHARED MAILBOX via IMAP
Authenticates with primary account, accesses shared mailbox
CRITICAL: Always uses mikereports@aftonprop.com shared mailbox
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
        logging.FileHandler(LOG_DIR / f"imap_shared_{datetime.now().strftime('%Y%m%d_%H%M%S')}.log"),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger(__name__)

# Configuration
OUTPUT_DIR = Path(r"C:\Users\MichaelOzery\OneDrive - Afton Properties\Old Dropbox\My PC (DESKTOP-5D77V89)\Mike Ozery\Dashboard Data")
PRIMARY_EMAIL = "mozery@aftonprop.com"  # Your primary account
SHARED_MAILBOX = "mikereports@aftonprop.com"  # Shared mailbox to access
SENDER_EMAIL = "cdr@yardi.com"
IMAP_SERVER = "outlook.office365.com"
IMAP_PORT = 993

OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

def list_mailboxes(imap):
    """List all available mailboxes/folders"""
    logger.info("\nAvailable mailboxes:")
    status, mailboxes = imap.list()
    for mailbox in mailboxes:
        logger.info(f"  {mailbox.decode()}")

def download_via_shared_mailbox(password):
    """Download attachments from shared mailbox via IMAP"""
    logger.info("=" * 70)
    logger.info("DOWNLOAD YARDI ATTACHMENTS FROM SHARED MAILBOX")
    logger.info(f"CRITICAL: Accessing {SHARED_MAILBOX} via {PRIMARY_EMAIL}")
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

        # Login with primary account
        logger.info(f"Logging in as {PRIMARY_EMAIL}...")
        try:
            imap.login(PRIMARY_EMAIL, password)
            logger.info("  [OK] Authentication successful")
        except imaplib.IMAP4.error as e:
            logger.error(f"  [ERROR] Authentication failed: {e}")
            logger.error("  Check email address and password/app password")
            return downloaded, errors + 1

        # List mailboxes (for debugging)
        logger.info("\n[MAILBOX DISCOVERY]")
        list_mailboxes(imap)

        # Try to access shared mailbox in different ways
        logger.info(f"\nTrying to access shared mailbox: {SHARED_MAILBOX}...")

        # Try different folder name formats
        folder_attempts = [
            SHARED_MAILBOX,
            f'"{SHARED_MAILBOX}"',
            f'[Gmail]/{SHARED_MAILBOX}',
            'INBOX',  # Fallback to primary inbox first to test
        ]

        inbox = None
        for folder_name in folder_attempts:
            try:
                status, mailbox_list = imap.select(folder_name)
                if status == 'OK':
                    logger.info(f"  [OK] Connected to: {folder_name}")
                    inbox = folder_name
                    break
                else:
                    logger.debug(f"  Could not select {folder_name}: {status}")
            except Exception as e:
                logger.debug(f"  Error trying {folder_name}: {e}")
                continue

        if not inbox:
            logger.error(f"  [ERROR] Could not access {SHARED_MAILBOX}")
            logger.error("  Make sure:")
            logger.error("    1. The shared mailbox is configured in your Office 365 account")
            logger.error("    2. You have permission to access it")
            logger.error("    3. IMAP is enabled for shared mailboxes")

            logger.info("\n[TROUBLESHOOTING]")
            logger.info("Try adding the shared mailbox to Outlook Web Access:")
            logger.info("  1. Go to https://outlook.office365.com")
            logger.info("  2. Settings > Shared mailboxes")
            logger.info("  3. Add: mikereports@aftonprop.com")
            logger.info("  4. Retry this script")

            return downloaded, errors + 1

        # Get message count
        message_count = int(mailbox_list[0])
        logger.info(f"  Mailbox has {message_count} messages")

        # Search for emails from YARDI
        logger.info(f"\nSearching for emails from {SENDER_EMAIL}...")
        status, message_ids = imap.search(None, f'FROM "{SENDER_EMAIL}"')

        if status != 'OK':
            logger.error("  [ERROR] Search failed")
            return downloaded, errors + 1

        message_ids = message_ids[0].split()
        logger.info(f"  Found {len(message_ids)} emails from {SENDER_EMAIL}")

        if len(message_ids) == 0:
            logger.warning("  No emails found from YARDI")
            return downloaded, errors

        # Process each email
        logger.info("\n[EMAIL PROCESSING]")
        logger.info(f"Processing {len(message_ids)} emails...")

        for i, msg_id in enumerate(message_ids, 1):
            try:
                # Fetch email
                status, msg_data = imap.fetch(msg_id, '(RFC822)')
                if status != 'OK':
                    logger.error(f"  [{i}/{len(message_ids)}] Could not fetch message")
                    errors += 1
                    continue

                # Parse email
                msg = email.message_from_bytes(msg_data[0][1])

                email_from = msg.get('From', 'Unknown')
                email_subject = msg.get('Subject', 'No Subject')
                email_date = msg.get('Date', 'Unknown')

                logger.info(f"\n[Email {i}/{len(message_ids)}] {email_subject[:60]}")
                logger.info(f"  From: {email_from}")

                # Check for attachments
                has_attachments = False
                for part in msg.walk():
                    if part.get_content_disposition() == 'attachment':
                        has_attachments = True
                        break

                if not has_attachments:
                    logger.info("  No attachments")
                    continue

                # Download attachments
                for part in msg.walk():
                    if part.get_content_disposition() != 'attachment':
                        continue

                    filename = part.get_filename()

                    # Skip if no filename
                    if not filename:
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
                        logger.error(f"    Error saving: {e}")
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
    import sys

    print("\n" + "=" * 70)
    print("YARDI REPORT DOWNLOADER (SHARED MAILBOX)")
    print("=" * 70)
    print(f"\nPrimary account: {PRIMARY_EMAIL}")
    print(f"Shared mailbox:  {SHARED_MAILBOX}")
    print("\nNOTE: If you have 2FA enabled on your account:")
    print("  1. Create an app password in Office 365")
    print("  2. Use the app password (not your regular password)")
    print("\nHow to create app password:")
    print("  https://support.microsoft.com/en-us/account-billing/using-app-passwords-with-apps-that-dont-support-multi-factor-authentication")
    print("\n" + "=" * 70)

    # Check if password provided as command line argument
    if len(sys.argv) > 1:
        password = sys.argv[1]
        print(f"\nUsing password from command line")
    else:
        # Try getpass, fallback to regular input if it fails
        try:
            password = getpass.getpass(f"\nPassword (or app password) for {PRIMARY_EMAIL}: ")
        except:
            print("\n(Password input issue detected, using text input instead)")
            password = input(f"Password (or app password) for {PRIMARY_EMAIL}: ")

    if not password:
        print("ERROR: Password required")
        exit(1)

    logger.info(f"\nStarting download from {SHARED_MAILBOX}...")
    logger.info("")

    downloaded, errors = download_via_shared_mailbox(password)

    if errors == 0:
        logger.info("\n[SUCCESS] All attachments downloaded")
        exit(0)
    else:
        logger.error(f"\n[FAILED] {errors} error(s) occurred")
        exit(1)
