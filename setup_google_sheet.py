#!/usr/bin/env python3
"""
Setup Google Sheet tabs and formatting for Afton Dashboard
"""

import json
import logging
from pathlib import Path
from datetime import datetime

try:
    from google.auth.transport.requests import Request
    from google.oauth2.credentials import Credentials
    from google_auth_oauthlib.flow import InstalledAppFlow
    from googleapiclient.discovery import build
except ImportError as e:
    print(f"ERROR: Missing Google libraries - {e}")
    print("Run: pip install google-auth-oauthlib google-auth-httplib2 google-api-python-client --break-system-packages")
    exit(1)

# Setup logging
SCRIPT_DIR = Path(__file__).parent
LOG_DIR = SCRIPT_DIR / "logs"
LOG_DIR.mkdir(exist_ok=True)

logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler(LOG_DIR / f"setup_sheet_{datetime.now().strftime('%Y%m%d_%H%M%S')}.log"),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger(__name__)

# Google Sheets API Configuration
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
SHEET_ID = "1PAnz3SC7mKRiRdRsmeUIp-zP0gsGQDAl_RR4pmHG6YE"
CREDENTIALS_FILE = SCRIPT_DIR / "credentials.json"

def get_sheets_service():
    """Authenticate and get Google Sheets service"""
    creds = None
    token_file = SCRIPT_DIR / "token.json"

    if token_file.exists():
        creds = Credentials.from_authorized_user_file(token_file, SCOPES)

    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            if not CREDENTIALS_FILE.exists():
                raise FileNotFoundError(f"credentials.json not found at {CREDENTIALS_FILE}")

            # Load credentials.json with UTF-8 BOM handling
            try:
                flow = InstalledAppFlow.from_client_secrets_file(CREDENTIALS_FILE, SCOPES)
            except json.JSONDecodeError:
                # If standard loading fails due to BOM, load manually
                with open(CREDENTIALS_FILE, 'r', encoding='utf-8-sig') as f:
                    client_config = json.load(f)
                flow = InstalledAppFlow.from_client_config(client_config, SCOPES)

            # Desktop app: run_local_server opens browser automatically
            creds = flow.run_local_server(port=0, open_browser=True)

        with open(token_file, 'w') as token:
            token.write(creds.to_json())

    return build('sheets', 'v4', credentials=creds)

def create_sheet_tabs(service):
    """Create tabs if they don't exist"""
    logger.info("Setting up sheet tabs...")

    tab_names = ["METRICS", "MARKETING", "HIGHLIGHTS", "BOTTOM_PERFORMERS", "CHART_DATA", "PRIORITIES", "METADATA"]

    # Get existing sheets
    sheet_metadata = service.spreadsheets().get(spreadsheetId=SHEET_ID).execute()
    existing_tabs = {sheet['properties']['title'] for sheet in sheet_metadata['sheets']}
    logger.info(f"Existing tabs: {existing_tabs}")

    requests = []

    # Remove unwanted default sheets
    for sheet in sheet_metadata['sheets']:
        if sheet['properties']['title'] not in tab_names:
            logger.info(f"Removing tab: {sheet['properties']['title']}")
            requests.append({
                'deleteSheet': {
                    'sheetId': sheet['properties']['sheetId']
                }
            })

    # Create missing tabs
    for tab_name in tab_names:
        if tab_name not in existing_tabs:
            logger.info(f"Creating tab: {tab_name}")
            requests.append({
                'addSheet': {
                    'properties': {
                        'title': tab_name,
                        'gridProperties': {
                            'rowCount': 100,
                            'columnCount': 10
                        }
                    }
                }
            })

    if requests:
        service.spreadsheets().batchUpdate(
            spreadsheetId=SHEET_ID,
            body={'requests': requests}
        ).execute()
        logger.info(f"[OK] Created/Updated {len(requests)} tabs")

def setup_metrics_tab(service):
    """Setup METRICS tab headers"""
    logger.info("Setting up METRICS tab...")

    headers = [["Metric Name", "Current Value", "Change", "Detail / Notes"]]

    service.spreadsheets().values().update(
        spreadsheetId=SHEET_ID,
        range="METRICS!A1:D1",
        valueInputOption="RAW",
        body={"values": headers}
    ).execute()

    logger.info("[OK] METRICS tab headers set")

def setup_marketing_tab(service):
    """Setup MARKETING tab headers"""
    logger.info("Setting up MARKETING tab...")

    headers = [["Metric", "Value", "Trend"]]

    service.spreadsheets().values().update(
        spreadsheetId=SHEET_ID,
        range="MARKETING!A1:C1",
        valueInputOption="RAW",
        body={"values": headers}
    ).execute()

    logger.info("[OK] MARKETING tab headers set")

def setup_highlights_tab(service):
    """Setup HIGHLIGHTS tab headers"""
    logger.info("Setting up HIGHLIGHTS tab...")

    headers = [["KEY HIGHLIGHTS", "AREAS OF FOCUS"]]

    service.spreadsheets().values().update(
        spreadsheetId=SHEET_ID,
        range="HIGHLIGHTS!A1:B1",
        valueInputOption="RAW",
        body={"values": headers}
    ).execute()

    logger.info("[OK] HIGHLIGHTS tab headers set")

def setup_bottom_performers_tab(service):
    """Setup BOTTOM_PERFORMERS tab headers"""
    logger.info("Setting up BOTTOM_PERFORMERS tab...")

    headers = [["Community", "Occupancy", "Leased Occ.", "% Collected", "Rent Growth (Ann.)", "NOI Var. to Budget"]]

    service.spreadsheets().values().update(
        spreadsheetId=SHEET_ID,
        range="BOTTOM_PERFORMERS!A1:F1",
        valueInputOption="RAW",
        body={"values": headers}
    ).execute()

    logger.info("[OK] BOTTOM_PERFORMERS tab headers set")

def setup_chart_data_tab(service):
    """Setup CHART_DATA tab"""
    logger.info("Setting up CHART_DATA tab...")

    data = [
        ["NOI Performance"],
        ["Category", "Percentage"],
        ["Above Budget", 53],
        ["On Line", 27],
        ["Below Budget", 20],
        [],
        ["Net Income Excl. Interest"],
        ["Category", "Percentage"],
        ["Above Target", 58],
        ["At Target", 24],
        ["Below Target", 18],
    ]

    service.spreadsheets().values().update(
        spreadsheetId=SHEET_ID,
        range="CHART_DATA!A1:B11",
        valueInputOption="RAW",
        body={"values": data}
    ).execute()

    logger.info("[OK] CHART_DATA tab set")

def setup_priorities_tab(service):
    """Setup PRIORITIES tab headers"""
    logger.info("Setting up PRIORITIES tab...")

    headers = [["#", "Priority Item"]]

    service.spreadsheets().values().update(
        spreadsheetId=SHEET_ID,
        range="PRIORITIES!A1:B1",
        valueInputOption="RAW",
        body={"values": headers}
    ).execute()

    logger.info("[OK] PRIORITIES tab headers set")

def setup_metadata_tab(service):
    """Setup METADATA tab"""
    logger.info("Setting up METADATA tab...")

    data = [
        ["Field", "Value"],
        ["Last Updated", "—"],
        ["Report Date", "—"],
        ["Data Source", "Excel Extract (Python)"],
        ["Next Update", "Daily 9:00 AM"],
    ]

    service.spreadsheets().values().update(
        spreadsheetId=SHEET_ID,
        range="METADATA!A1:B5",
        valueInputOption="RAW",
        body={"values": data}
    ).execute()

    logger.info("[OK] METADATA tab set")

def main():
    try:
        logger.info("=" * 70)
        logger.info("Setting up Google Sheet for Afton Dashboard")
        logger.info("=" * 70)

        # Authenticate
        logger.info("Authenticating with Google Sheets API...")
        service = get_sheets_service()
        logger.info("[OK] Connected to Google Sheets")

        # Create tabs
        create_sheet_tabs(service)

        # Setup each tab
        setup_metrics_tab(service)
        setup_marketing_tab(service)
        setup_highlights_tab(service)
        setup_bottom_performers_tab(service)
        setup_chart_data_tab(service)
        setup_priorities_tab(service)
        setup_metadata_tab(service)

        logger.info("=" * 70)
        logger.info("[OK] Google Sheet setup complete!")
        logger.info(f"Sheet URL: https://docs.google.com/spreadsheets/d/{SHEET_ID}/edit")
        logger.info("=" * 70)

    except Exception as e:
        logger.error(f"Fatal error: {e}", exc_info=True)

if __name__ == '__main__':
    main()
