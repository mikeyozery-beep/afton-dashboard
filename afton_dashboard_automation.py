import os
import json
import pickle
import logging
from pathlib import Path
from datetime import datetime

try:
    from google.auth.transport.requests import Request
    from google.oauth2.credentials import Credentials
    from google_auth_oauthlib.flow import InstalledAppFlow
    from googleapiclient.discovery import build
    from googleapiclient.errors import HttpError
except ImportError:
    print("ERROR: Missing dependencies. Run: pip install -r requirements.txt")
    exit(1)

SCRIPT_DIR = Path(__file__).parent
TOKENS_DIR = SCRIPT_DIR / "tokens"
LOG_DIR = SCRIPT_DIR / "logs"
DATA_DIR = SCRIPT_DIR / "data"
TOKENS_DIR.mkdir(exist_ok=True)
LOG_DIR.mkdir(exist_ok=True)
DATA_DIR.mkdir(exist_ok=True)

logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler(LOG_DIR / f"dashboard_sync_{datetime.now().strftime('%Y%m%d_%H%M%S')}.log"),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger(__name__)

GOOGLE_SCOPES = ['https://www.googleapis.com/auth/spreadsheets.readonly']
SHEET_ID = '1PAnz3SC7mKRiRdRsmeUIp-zP0gsGQDAl_RR4pmHG6YE'
METRICS_FILE = DATA_DIR / 'metrics.json'

def get_google_credentials():
    token_file = TOKENS_DIR / 'google_token.pickle'

    if token_file.exists():
        with open(token_file, 'rb') as token:
            creds = pickle.load(token)
        if creds and creds.expired and creds.refresh_token:
            try:
                creds.refresh(Request())
                with open(token_file, 'wb') as token:
                    pickle.dump(creds, token)
            except Exception as e:
                logger.warning(f"Token refresh failed: {e}. Reauthorizing...")
                os.remove(token_file)
                return get_google_credentials()
        return creds

    logger.info("\n" + "="*70)
    logger.info("AUTHORIZATION REQUIRED")
    logger.info("="*70)

    flow = InstalledAppFlow.from_client_secrets_file('credentials.json', GOOGLE_SCOPES)

    # Try to open browser automatically, but provide manual URL if it fails
    try:
        creds = flow.run_local_server(port=8080, open_browser=True)
    except Exception as e:
        logger.warning(f"Auto-browser failed ({e}). Using manual authorization...")
        logger.info("\n1. Copy this authorization URL:")
        auth_url, _ = flow.authorization_url(prompt='consent')
        logger.info(f"\n   {auth_url}\n")
        logger.info("2. Paste it in your web browser and authorize")
        logger.info("3. Copy the authorization code from the browser")
        logger.info("4. Paste it here and press Enter\n")

        auth_code = input("Enter authorization code: ").strip()

        if not auth_code:
            raise Exception("No authorization code provided")

        creds = flow.fetch_token(code=auth_code)

    with open(token_file, 'wb') as token:
        pickle.dump(creds, token)

    logger.info("✓ Authorization successful!")
    logger.info("="*70 + "\n")
    return creds

def fetch_and_cache_metrics():
    """Fetch metrics from Google Sheet and save to local JSON"""
    try:
        creds = get_google_credentials()
        service = build('sheets', 'v4', credentials=creds)

        logger.info(f"Fetching metrics from sheet {SHEET_ID}...")

        result = service.spreadsheets().values().get(
            spreadsheetId=SHEET_ID,
            range="'Current Metrics'!A1:D100"
        ).execute()

        values = result.get('values', [])
        logger.info(f"Retrieved {len(values)} rows from sheet")

        if not values or len(values) < 2:
            logger.warning("No data found in sheet")
            return None

        # Parse metrics (skip header row)
        metrics = []
        for row in values[1:]:
            if len(row) >= 2:
                metric = {
                    'name': row[0] if len(row) > 0 else '',
                    'value': row[1] if len(row) > 1 else '',
                    'updated': row[2] if len(row) > 2 else '',
                    'source': row[3] if len(row) > 3 else ''
                }
                metrics.append(metric)

        # Save to local JSON file
        data = {
            'timestamp': datetime.now().isoformat(),
            'metrics': metrics,
            'count': len(metrics),
            'success': True
        }

        with open(METRICS_FILE, 'w') as f:
            json.dump(data, f, indent=2)

        logger.info(f"Cached {len(metrics)} metrics to {METRICS_FILE}")
        return data

    except HttpError as error:
        logger.error(f"Google Sheets API error: {error}")
        return None
    except Exception as e:
        logger.error(f"Error fetching metrics: {e}", exc_info=True)
        return None

def get_cached_metrics():
    """Get metrics from local cache"""
    if METRICS_FILE.exists():
        try:
            with open(METRICS_FILE, 'r') as f:
                data = json.load(f)
                logger.info(f"Using cached metrics ({len(data.get('metrics', []))} items)")
                return data
        except Exception as e:
            logger.error(f"Error reading cache: {e}")
    return None

def main():
    try:
        logger.info("=" * 60)
        logger.info("Starting Afton Dashboard Automation")
        logger.info("=" * 60)

        # Try to fetch fresh metrics from Google Sheet
        data = fetch_and_cache_metrics()

        # If fetch failed, use cached version
        if not data:
            logger.info("Fetching fresh data failed, using cache...")
            data = get_cached_metrics()

        if data:
            logger.info(f"✓ Dashboard has {data.get('count', 0)} metrics")
            logger.info(f"✓ Data saved to {METRICS_FILE}")
            logger.info(f"✓ Dashboard can access at: http://localhost:8000/data/metrics.json")
        else:
            logger.warning("No metrics available (fresh or cached)")

        logger.info("=" * 60)
        logger.info("Dashboard automation completed!")
        logger.info("=" * 60)

    except Exception as e:
        logger.error(f"Dashboard automation failed: {e}", exc_info=True)

if __name__ == '__main__':
    main()
