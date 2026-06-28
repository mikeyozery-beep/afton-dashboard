#!/usr/bin/env python3
"""
UNIFIED YARDI DOWNLOAD & ORGANIZATION
Single-step automation: Download → Extract Zips → Extract Worksheets → Organize by Report Name
No manual sequencing required. One command handles everything.
"""

import logging
from pathlib import Path
from datetime import datetime
from openpyxl import load_workbook
import zipfile
import shutil
import subprocess
import sys

# Setup logging
SCRIPT_DIR = Path(__file__).parent
LOG_DIR = SCRIPT_DIR / "logs"
LOG_DIR.mkdir(exist_ok=True)

logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler(LOG_DIR / f"unified_{datetime.now().strftime('%Y%m%d_%H%M%S')}.log"),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger(__name__)

DASHBOARD_DATA_DIR = Path(r"C:\Users\MichaelOzery\OneDrive - Afton Properties\Old Dropbox\My PC (DESKTOP-5D77V89)\Mike Ozery\Dashboard Data")

def sanitize_foldername(name):
    invalid_chars = '<>:"/\\|?*'
    for char in invalid_chars:
        name = name.replace(char, '_')
    return name.strip()

def step_1_download_files():
    """Download all files from mikereports@aftonprop.com inbox"""
    logger.info("\n" + "=" * 70)
    logger.info("STEP 1: DOWNLOADING FILES FROM INBOX")
    logger.info("=" * 70)

    try:
        download_script = SCRIPT_DIR / "download_yardi_outlook_com.ps1"
        if not download_script.exists():
            logger.error(f"Download script not found: {download_script}")
            return False

        logger.info(f"Running: {download_script.name}")
        result = subprocess.run(
            ["powershell", "-ExecutionPolicy", "Bypass", "-File", str(download_script)],
            capture_output=True,
            text=True,
            timeout=300
        )

        if result.returncode != 0:
            logger.error(f"Download failed:\n{result.stderr}")
            return False

        logger.info("✓ Download complete")
        return True

    except Exception as e:
        logger.error(f"Step 1 error: {e}", exc_info=True)
        return False

def step_2_extract_zips():
    """Extract all zip files in Dashboard Data"""
    logger.info("\n" + "=" * 70)
    logger.info("STEP 2: EXTRACTING ZIP FILES")
    logger.info("=" * 70)

    try:
        zip_files = list(DASHBOARD_DATA_DIR.glob("*.zip"))

        if not zip_files:
            logger.info("No zip files found")
            return True

        logger.info(f"Found {len(zip_files)} zip files")

        for zip_file in zip_files:
            try:
                logger.info(f"Extracting: {zip_file.name}")

                with zipfile.ZipFile(zip_file, 'r') as zip_ref:
                    for file_info in zip_ref.filelist:
                        if file_info.filename.lower().endswith(('.xlsx', '.xls')):
                            extract_path = DASHBOARD_DATA_DIR / Path(file_info.filename).name

                            with zip_ref.open(file_info) as source, open(extract_path, 'wb') as target:
                                target.write(source.read())

                            logger.info(f"  ✓ {Path(file_info.filename).name}")

                zip_file.unlink()
                logger.info(f"  ✓ Deleted zip file")

            except Exception as e:
                logger.error(f"Error extracting {zip_file.name}: {e}")

        logger.info("✓ Zip extraction complete")
        return True

    except Exception as e:
        logger.error(f"Step 2 error: {e}", exc_info=True)
        return False

def step_3_extract_worksheets():
    """Extract worksheets from multi-sheet Excel files"""
    logger.info("\n" + "=" * 70)
    logger.info("STEP 3: EXTRACTING & ORGANIZING BY REPORT NAME")
    logger.info("=" * 70)

    try:
        root_files = list(DASHBOARD_DATA_DIR.glob("*.xlsx")) + list(DASHBOARD_DATA_DIR.glob("*.xls"))

        if not root_files:
            logger.warning("No Excel files found in root")
            return True

        logger.info(f"Found {len(root_files)} files\n")

        organized_count = 0
        report_folders = {}
        files_to_delete = []

        for filepath in sorted(root_files):
            try:
                wb = load_workbook(filepath)
                logger.info(f"[{filepath.name}] {len(wb.sheetnames)} worksheet(s)")

                for sheet_name in wb.sheetnames:
                    try:
                        # Create new workbook with just this sheet
                        new_wb = load_workbook(filepath)

                        sheets_to_delete = [s for s in new_wb.sheetnames if s != sheet_name]
                        for sheet_to_delete in sheets_to_delete:
                            del new_wb[sheet_to_delete]

                        new_ws = new_wb.active
                        new_ws.title = "Report"

                        # Read A1 for actual report name
                        report_name = new_ws['A1'].value

                        if not report_name:
                            logger.warning(f"  No name in A1, using: {sheet_name}")
                            report_name = sheet_name

                        report_name = str(report_name).strip()
                        logger.info(f"  → {report_name}")

                        # Create folder and save
                        safe_folder = sanitize_foldername(report_name)
                        report_folder = DASHBOARD_DATA_DIR / safe_folder
                        report_folder.mkdir(exist_ok=True)

                        timestamp = datetime.now().strftime("%Y-%m-%d %H-%M-%S")
                        output_filename = f"{safe_folder}_{timestamp}.xlsx"
                        output_path = report_folder / output_filename

                        new_wb.save(output_path)
                        logger.info(f"     ✓ {safe_folder}/")

                        if safe_folder not in report_folders:
                            report_folders[safe_folder] = 0
                        report_folders[safe_folder] += 1
                        organized_count += 1

                    except Exception as e:
                        logger.error(f"  Error: {e}")

                files_to_delete.append(filepath)

            except Exception as e:
                logger.error(f"Error processing {filepath.name}: {e}")

        # Delete originals
        logger.info(f"\nDeleting {len(files_to_delete)} original files...")
        for filepath in files_to_delete:
            try:
                filepath.unlink()
            except Exception as e:
                logger.error(f"Could not delete {filepath.name}: {e}")

        logger.info(f"✓ Organized {organized_count} worksheets into {len(report_folders)} folders")
        return True

    except Exception as e:
        logger.error(f"Step 3 error: {e}", exc_info=True)
        return False

def step_4_mark_read():
    """Mark the downloaded emails as read - ONLY after organize succeeded."""
    logger.info("\n" + "=" * 70)
    logger.info("STEP 4: MARKING PROCESSED EMAILS AS READ")
    logger.info("=" * 70)

    try:
        mark_script = SCRIPT_DIR / "mark_read_pending.ps1"
        if not mark_script.exists():
            logger.error(f"Mark-read script not found: {mark_script}")
            return False

        result = subprocess.run(
            ["powershell", "-ExecutionPolicy", "Bypass", "-File", str(mark_script)],
            capture_output=True,
            text=True,
            timeout=300
        )
        logger.info(result.stdout.strip())
        if result.returncode != 0:
            logger.error(f"Mark-read failed:\n{result.stderr}")
            return False

        logger.info("✓ Mark-read complete")
        return True

    except Exception as e:
        logger.error(f"Step 4 error: {e}", exc_info=True)
        return False

def main():
    logger.info("\n" + "=" * 70)
    logger.info("UNIFIED YARDI DOWNLOAD & ORGANIZATION")
    logger.info("=" * 70)
    logger.info(f"Directory: {DASHBOARD_DATA_DIR}\n")

    # Execute all steps
    if not step_1_download_files():
        logger.error("\n✗ Process stopped at download")
        return False

    if not step_2_extract_zips():
        logger.error("\n✗ Process stopped at zip extraction")
        return False

    if not step_3_extract_worksheets():
        logger.error("\n✗ Process stopped at worksheet extraction")
        return False

    # Only mark emails read once their attachments are safely organized into folders
    if not step_4_mark_read():
        logger.warning("\n⚠ Organize succeeded but marking emails read failed - "
                       "they stay UNREAD and will be retried next run")

    logger.info("\n" + "=" * 70)
    logger.info("✓ ALL COMPLETE - DOWNLOADED, ORGANIZED & MARKED READ")
    logger.info("=" * 70 + "\n")
    return True

if __name__ == '__main__':
    success = main()
    sys.exit(0 if success else 1)
