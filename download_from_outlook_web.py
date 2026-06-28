#!/usr/bin/env python3
"""
Download YARDI attachments from Outlook Web Access (OWA)
Uses browser automation via Selenium to access mikereports@aftonprop.com
CRITICAL: Always uses mikereports@aftonprop.com, NEVER primary email
"""

import os
import time
import logging
from pathlib import Path
from datetime import datetime

try:
    from selenium import webdriver
    from selenium.webdriver.common.by import By
    from selenium.webdriver.support.ui import WebDriverWait
    from selenium.webdriver.support import expected_conditions as EC
    from selenium.webdriver.chrome.options import Options as ChromeOptions
except ImportError:
    print("ERROR: selenium not installed")
    print("Run: pip install selenium --break-system-packages")
    exit(1)

# Setup logging
SCRIPT_DIR = Path(__file__).parent
LOG_DIR = SCRIPT_DIR / "logs"
LOG_DIR.mkdir(exist_ok=True)

logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler(LOG_DIR / f"owa_download_{datetime.now().strftime('%Y%m%d_%H%M%S')}.log"),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger(__name__)

# Configuration
OUTPUT_DIR = Path(r"C:\Users\MichaelOzery\OneDrive - Afton Properties\Old Dropbox\My PC (DESKTOP-5D77V89)\Mike Ozery\Dashboard Data")
TARGET_EMAIL = "mikereports@aftonprop.com"  # CRITICAL: NEVER primary email
SENDER_EMAIL = "cdr@yardi.com"
OWA_URL = "https://outlook.office365.com"

OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

def download_from_owa():
    """Download attachments from Outlook Web Access"""
    logger.info("=" * 70)
    logger.info("DOWNLOAD YARDI ATTACHMENTS FROM OUTLOOK WEB")
    logger.info(f"CRITICAL: Using {TARGET_EMAIL} ONLY")
    logger.info("=" * 70)

    driver = None
    downloaded = 0
    skipped = 0
    errors = 0

    try:
        logger.info("\n[BROWSER SETUP]")
        logger.info(f"Opening Outlook Web at {OWA_URL}...")

        # Create Chrome options
        chrome_options = ChromeOptions()
        chrome_options.add_argument("--start-maximized")
        # Do NOT use headless mode - need to see login/MFA prompts

        # Create driver
        driver = webdriver.Chrome(options=chrome_options)
        logger.info("  Browser created")

        # Navigate to OWA
        driver.get(OWA_URL)
        logger.info("  Navigated to Outlook Web")

        # Wait for login
        logger.info("\n[LOGIN REQUIRED]")
        logger.info("Please log in with: mikereports@aftonprop.com")
        logger.info("Complete any MFA prompts if required...")
        logger.info("Script will continue automatically once logged in...")

        # Wait for inbox to load (indicates successful login)
        wait = WebDriverWait(driver, 300)  # 5 minute timeout for login

        try:
            wait.until(EC.presence_of_all_elements_located((By.XPATH, "//span[contains(@class, 'lvItemsContainerClass')]")))
            logger.info("[OK] Login successful - inbox loaded")
        except:
            logger.error("Login timeout or inbox not found")
            logger.error("Make sure you're logging in with: mikereports@aftonprop.com")
            return downloaded, errors + 1

        # Wait a moment for page to stabilize
        time.sleep(2)

        logger.info("\n[EMAIL PROCESSING]")
        logger.info(f"Looking for emails from {SENDER_EMAIL}...")

        # Get all email rows
        email_elements = driver.find_elements(By.XPATH, "//div[@role='option']")
        logger.info(f"Found {len(email_elements)} emails in inbox")

        for i, email_elem in enumerate(email_elements):
            try:
                # Get email text
                email_text = email_elem.text

                # Check if from YARDI
                if SENDER_EMAIL.lower() not in email_text.lower():
                    continue

                logger.info(f"\n[Email {i+1}] {email_text[:80]}...")

                # Click email to open
                email_elem.click()
                time.sleep(1)

                # Wait for attachments section to load
                try:
                    attachment_section = WebDriverWait(driver, 10).until(
                        EC.presence_of_element_located((By.XPATH, "//div[contains(@class, 'attachmentContainer')]"))
                    )
                    logger.info("    Has attachments")

                    # Get all attachment links
                    attachment_links = driver.find_elements(By.XPATH, "//a[contains(@aria-label, 'Download')]")

                    for attachment in attachment_links:
                        try:
                            filename = attachment.get_attribute("title")

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

                            # Download - this will trigger browser download
                            logger.info(f"    Downloading: {filename}")
                            attachment.click()
                            time.sleep(2)  # Wait for download

                            # Check if file appeared in Downloads folder
                            downloads_folder = Path.home() / "Downloads" / filename
                            if downloads_folder.exists():
                                # Move to Dashboard Data folder
                                import shutil
                                shutil.move(str(downloads_folder), str(filepath))
                                logger.info(f"    [OK] Downloaded: {filename}")
                                downloaded += 1
                            else:
                                logger.warning(f"    Download did not appear in Downloads folder")
                                errors += 1

                        except Exception as e:
                            logger.error(f"    Error processing attachment: {e}")
                            errors += 1

                except:
                    logger.info("    No attachments or could not find attachment section")
                    continue

                # Go back to inbox
                back_button = driver.find_element(By.XPATH, "//button[contains(@aria-label, 'Back')]")
                back_button.click()
                time.sleep(1)

            except Exception as e:
                logger.error(f"  Error processing email {i}: {e}")
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

    finally:
        if driver:
            logger.info("\nClosing browser...")
            driver.quit()

if __name__ == '__main__':
    logger.info("Installation check:")
    logger.info("  pip install selenium --break-system-packages")
    logger.info("  Download ChromeDriver: https://chromedriver.chromium.org/")
    logger.info("  Place chromedriver.exe in this directory")
    logger.info("")

    downloaded, errors = download_from_owa()
    exit(0 if errors == 0 else 1)
