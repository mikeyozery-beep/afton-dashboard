#!/usr/bin/env python3
"""
Standalone organize step (no download).
Splits multi-sheet workbooks and files each report into a folder named by its
A1 title. Delegates to the shared, optimized organizer in
unified_download_organize.py so there is a single source of truth.
"""

import sys
from unified_download_organize import organize_data_dir, DASHBOARD_DATA_DIR, logger

if __name__ == '__main__':
    logger.info("=" * 70)
    logger.info("ORGANIZE REPORTS (standalone)")
    logger.info("=" * 70)
    count, folders = organize_data_dir(DASHBOARD_DATA_DIR)
    sys.exit(0)
