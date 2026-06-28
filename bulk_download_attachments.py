#!/usr/bin/env python3
"""
Bulk Download YARDI Report Attachments from Outlook
Downloads ALL Excel attachments from cdr@yardi.com emails

CRITICAL: Always uses mikereports@aftonprop.com mailbox, NEVER primary email
"""

import os
import time
import logging
from pathlib import Path
from datetime import datetime

try:
    import win32com.client
    import win32com.pywintypes as pywintypes
except ImportError:
    print("ERROR: pywin32 not installed")
    print("Run: pip install pywin32 --break-system-packages")
    exit(1)

# Setup logging
SCRIPT_DIR = Path(__file__).parent
LOG_DIR = SCRIPT_DIR / "logs"
LOG_DIR.mkdir(exist_ok=True)

logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler(LOG_DIR / f"bulk_download_{datetime.now().strftime('%Y%m%d_%H%M%S')}.log"),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger(__name__)

# Configuration
OUTPUT_DIR = Path(r"C:\Users\MichaelOzery\OneDrive - Afton Properties\Old Dropbox\My PC (DESKTOP-5D77V89)\Mike Ozery\Dashboard Data")
SENDER_EMAIL = "cdr@yardi.com"
TARGET_MAILBOX = "mikereports@aftonprop.com"  # CRITICAL: Never use default email

OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

def get_mikereports_inbox():
    """Get inbox for mikereports@aftonprop.com - NEVER primary email"""
    logger.info("\n[MAILBOX CONNECTION]")
    logger.info(f"Target mailbox: {TARGET_MAILBOX}")
    logger.info("Connecting to Outlook...")

    outlook = None
    mapi = None
    inbox = None

    try:
        # Connect to Outlook
        try:
            outlook = win32com.client.GetObject(None, "Outlook.Application")
            logger.info("  Connected to running Outlook instance")
        except:
            logger.info("  No running Outlook found, creating new instance...")
            outlook = win32com.client.Dispatch("Outlook.Application")

        time.sleep(0.5)  # Let COM stabilize

        # Get MAPI namespace
        mapi = outlook.GetNamespace("MAPI")
        logger.info("  Got MAPI namespace")

        time.sleep(0.5)  # Let COM stabilize

        # Search for mikereports mailbox
        logger.info(f"  Searching for mailbox: {TARGET_MAILBOX}...")

        stores_list = []
        for store in mapi.Stores:
            try:
                store_name = store.DisplayName
                stores_list.append(store_name)
                logger.info(f"    Found store: {store_name}")
            except:
                continue

        logger.info(f"  Total mailboxes found: {len(stores_list)}")

        # Now try to get the correct inbox
        for store in mapi.Stores:
            try:
                root_folder = store.GetRootFolder()
                store_name = root_folder.Name

                logger.info(f"  Checking: {store_name}")

                # Check if this is the mikereports mailbox
                if TARGET_MAILBOX.lower() in store_name.lower() or \
                   "mikereports" in store_name.lower() or \
                   "aftonprop" in store_name.lower():

                    logger.info(f"    [MATCH] This is the target mailbox!")

                    # Get Inbox folder
                    inbox = root_folder.Folders("Inbox")
                    logger.info(f"[✓] Successfully connected to: {store_name} > Inbox")

                    return inbox, outlook, mapi

            except Exception as e:
                logger.debug(f"    Error checking store: {e}")
                continue

        # If we got here, mikereports mailbox wasn't found
        logger.error(f"[✗] Could not find mailbox for: {TARGET_MAILBOX}")
        logger.error("    Available mailboxes:")
        for store_name in stores_list:
            logger.error(f"      - {store_name}")

        return None, outlook, mapi

    except Exception as e:
        logger.error(f"Fatal error connecting to Outlook: {e}", exc_info=True)
        return None, outlook, mapi

def bulk_download():
    """Download all attachments from YARDI emails"""
    logger.info("=" * 70)
    logger.info("BULK DOWNLOAD YARDI ATTACHMENTS")
    logger.info("CRITICAL: Using {0} mailbox ONLY".format(TARGET_MAILBOX))
    logger.info("=" * 70)

    outlook = None
    mapi = None

    try:
        # Get inbox
        inbox, outlook, mapi = get_mikereports_inbox()

        if not inbox:
            logger.error("ERROR: Could not connect to {0}".format(TARGET_MAILBOX))
            logger.error("Make sure this mailbox is configured in Outlook")
            return 0, 1

        # Process emails
        logger.info(f"\nSearching for emails from {SENDER_EMAIL}...")
        logger.info("Sorting emails by received time...")

        items = inbox.Items
        items.Sort("[ReceivedTime]", False)  # Ascending: oldest first

        logger.info(f"Total emails in inbox: {items.Count}")

        downloaded = 0
        skipped = 0
        errors = 0
        processed = 0

        logger.info("Processing emails...")

        for i, item in enumerate(items):
            try:
                # Check sender
                try:
                    sender = item.SenderEmailAddress
                except:
                    continue

                if SENDER_EMAIL.lower() not in sender.lower():
                    continue

                # Check attachments
                if item.Attachments.Count == 0:
                    continue

                processed += 1

                try:
                    email_date = item.ReceivedTime
                except:
                    continue

                logger.info(f"\n[Email {processed}] {email_date.strftime('%Y-%m-%d %H:%M:%S')} - {item.Subject}")
                logger.info(f"    Attachments: {item.Attachments.Count}")

                # Download each attachment
                for attachment in item.Attachments:
                    try:
                        filename = attachment.Filename

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
                            attachment.SaveAsFile(str(filepath))
                            logger.info(f"    ✓ Downloaded: {filename}")
                            downloaded += 1
                        except Exception as e:
                            logger.error(f"    Error downloading {filename}: {e}")
                            errors += 1

                    except Exception as e:
                        logger.error(f"    Error processing attachment: {e}")
                        errors += 1

            except Exception as e:
                logger.debug(f"  Error processing email {i}: {e}")
                errors += 1
                continue

        logger.info("\n" + "=" * 70)
        logger.info(f"DOWNLOAD COMPLETE")
        logger.info("=" * 70)
        logger.info(f"Downloaded: {downloaded} attachments")
        logger.info(f"Skipped (already exist): {skipped}")
        logger.info(f"Errors: {errors}")
        logger.info(f"Processed: {processed} YARDI emails")
        logger.info(f"Saved to: {OUTPUT_DIR}")
        logger.info("=" * 70)

        return downloaded, errors

    except Exception as e:
        logger.error(f"Fatal error: {e}", exc_info=True)
        return 0, 1

    finally:
        # Cleanup COM objects
        try:
            if inbox:
                del inbox
            if mapi:
                del mapi
            if outlook:
                del outlook
        except:
            pass

if __name__ == '__main__':
    downloaded, errors = bulk_download()
    exit(0 if errors == 0 else 1)
