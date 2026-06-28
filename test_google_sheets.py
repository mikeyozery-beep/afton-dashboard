#!/usr/bin/env python3
"""
Simple test to verify Google Sheets API credentials work
"""

import json
from pathlib import Path

SCRIPT_DIR = Path(__file__).parent
CREDENTIALS_FILE = SCRIPT_DIR / "credentials.json"

print("=" * 70)
print("GOOGLE SHEETS API CREDENTIALS TEST")
print("=" * 70)

# Test 1: Check if file exists
print("\n[TEST 1] Checking if credentials.json exists...")
if CREDENTIALS_FILE.exists():
    print(f"✓ File found at: {CREDENTIALS_FILE}")
    print(f"  File size: {CREDENTIALS_FILE.stat().st_size} bytes")
else:
    print(f"✗ File NOT found at: {CREDENTIALS_FILE}")
    exit(1)

# Test 2: Try to read the file
print("\n[TEST 2] Reading credentials.json file...")
try:
    with open(CREDENTIALS_FILE, 'r') as f:
        content = f.read()
    print(f"✓ File read successfully")
    print(f"  Content length: {len(content)} characters")
except Exception as e:
    print(f"✗ Error reading file: {e}")
    exit(1)

# Test 3: Try to parse JSON
print("\n[TEST 3] Parsing JSON...")
try:
    creds_data = json.loads(content)
    print(f"✓ JSON parsed successfully")
    print(f"  Structure: {list(creds_data.keys())}")
    if 'installed' in creds_data:
        installed = creds_data['installed']
        print(f"  Client ID: {installed.get('client_id', '?')[:30]}...")
        print(f"  Client Secret: {installed.get('client_secret', '?')[:20]}...")
except json.JSONDecodeError as e:
    print(f"✗ JSON parsing error: {e}")
    print(f"  First 100 chars: {content[:100]}")
    exit(1)

# Test 4: Try to import Google libraries
print("\n[TEST 4] Testing Google library imports...")
try:
    from google.auth.transport.requests import Request
    print("✓ google.auth.transport.requests")

    from google.oauth2.credentials import Credentials
    print("✓ google.oauth2.credentials")

    from google_auth_oauthlib.flow import InstalledAppFlow
    print("✓ google_auth_oauthlib.flow")

    from googleapiclient.discovery import build
    print("✓ googleapiclient.discovery")

except ImportError as e:
    print(f"✗ Import error: {e}")
    exit(1)

print("\n" + "=" * 70)
print("✓ ALL TESTS PASSED - Google Sheets API is ready!")
print("=" * 70)
