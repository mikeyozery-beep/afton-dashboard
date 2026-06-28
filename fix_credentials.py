#!/usr/bin/env python3
"""
Fix credentials.json - Add all necessary OAuth redirect URIs
"""

import json
from pathlib import Path

SCRIPT_DIR = Path(__file__).parent
CREDENTIALS_FILE = SCRIPT_DIR / "credentials.json"

print("=" * 70)
print("FIXING CREDENTIALS.JSON FOR GOOGLE OAUTH")
print("=" * 70)

# Step 1: Read with UTF-8 BOM handling
print(f"\n[STEP 1] Reading credentials.json...")
try:
    with open(CREDENTIALS_FILE, 'r', encoding='utf-8-sig') as f:
        creds_data = json.load(f)
    print("✓ File read successfully")
except Exception as e:
    print(f"✗ Error reading file: {e}")
    exit(1)

# Step 2: Check structure
print(f"\n[STEP 2] Checking credentials structure...")
if 'installed' not in creds_data:
    print("✗ No 'installed' section found")
    exit(1)

installed = creds_data['installed']
print(f"✓ Client ID: {installed.get('client_id', '?')[:30]}...")

# Step 3: Update redirect URIs with all possible values
print(f"\n[STEP 3] Adding comprehensive redirect URIs...")
new_redirect_uris = [
    "http://localhost:8080/",
    "http://localhost:8000/",
    "http://localhost/",
    "http://127.0.0.1:8080/",
    "http://127.0.0.1:8000/",
    "http://127.0.0.1/",
    "urn:ietf:wg:oauth:2.0:oob",
]

installed['redirect_uris'] = new_redirect_uris
print(f"✓ Added {len(new_redirect_uris)} redirect URIs:")
for uri in new_redirect_uris:
    print(f"  - {uri}")

# Step 4: Write back WITHOUT BOM (UTF-8 standard)
print(f"\n[STEP 4] Saving credentials.json...")
try:
    with open(CREDENTIALS_FILE, 'w', encoding='utf-8') as f:
        json.dump(creds_data, f, indent=2)
    print("✓ Saved successfully (UTF-8 clean)")
except Exception as e:
    print(f"✗ Error writing file: {e}")
    exit(1)

# Step 5: Verify
print(f"\n[STEP 5] Verifying...")
try:
    with open(CREDENTIALS_FILE, 'r', encoding='utf-8-sig') as f:
        verify_data = json.load(f)
    verify_uris = verify_data['installed'].get('redirect_uris', [])
    print(f"✓ Verified: {len(verify_uris)} redirect URIs")
except Exception as e:
    print(f"✗ Verification failed: {e}")
    exit(1)

print("\n" + "=" * 70)
print("✓ CREDENTIALS.JSON FIXED!")
print("=" * 70)
print("\n** IMPORTANT **")
print("You also need to add these redirect URIs to Google Cloud Console:")
print("1. Go to https://console.cloud.google.com")
print("2. Select your 'afton-properties-dashboard' project")
print("3. Go to APIs & Services → Credentials")
print("4. Click on the OAuth 2.0 Client ID (Desktop app)")
print("5. Add these URIs to 'Authorized redirect URIs':")
for uri in new_redirect_uris:
    print(f"   {uri}")
print("6. Click Save")
print("\nThen try again:")
print("  python setup_google_sheet.py")
