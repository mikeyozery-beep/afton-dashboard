#!/usr/bin/env python3
"""
Live staffing-vacancy count from the SharePoint 'Open Positions' list via
Microsoft Graph (DELEGATED auth, device-code sign-in -> cached refresh token).

One-time setup (interactive, do this once):
    python graph_staffing.py login
        -> opens a microsoft.com/devicelogin code; sign in as a user who can
           see Maggie's 'Open Positions' list. Token is cached locally.

After that, the daily task calls open_positions_count() silently using the
cached refresh token (auto-renews with use; ~90-day idle lifetime).

No app registration needed: uses Microsoft's public "Graph CLI" client id.
"""
import sys, json
from pathlib import Path

import msal
import requests

# ---- config -----------------------------------------------------------------
TENANT    = "3f73f992-5f90-448f-a139-a3ceb1ad94e5"          # aftonprop.com
CLIENT_ID = "14d82eec-204b-4c2f-b7e8-296a70dab67e"          # MS Graph Command Line Tools (public client)
AUTHORITY = f"https://login.microsoftonline.com/{TENANT}"
SCOPES    = ["Sites.Read.All"]

SITE_HOST = "aftonpropeties-my.sharepoint.com"
SITE_PATH = "/personal/maggie_aftonprop_com"
LIST_NAME = "Open Positions"
STATUS_FIELD_HINT = "status"   # matched case-insensitively against the list's field internal names
OPEN_VALUE = "open"

GRAPH = "https://graph.microsoft.com/v1.0"
CACHE = Path(__file__).parent / "graph_token_cache.json"


# ---- auth -------------------------------------------------------------------
def _app():
    cache = msal.SerializableTokenCache()
    if CACHE.exists():
        cache.deserialize(CACHE.read_text())
    app = msal.PublicClientApplication(CLIENT_ID, authority=AUTHORITY, token_cache=cache)
    return app, cache


def _save(cache):
    if cache.has_state_changed:
        CACHE.write_text(cache.serialize())


def get_token_silent():
    app, cache = _app()
    accounts = app.get_accounts()
    if not accounts:
        return None
    res = app.acquire_token_silent(SCOPES, account=accounts[0])
    _save(cache)
    return res.get("access_token") if res else None


def interactive_login():
    app, cache = _app()
    flow = app.initiate_device_flow(scopes=SCOPES)
    if "user_code" not in flow:
        raise RuntimeError("Device flow failed: " + json.dumps(flow, indent=2))
    print(flow["message"], flush=True)          # "go to microsoft.com/devicelogin and enter CODE"
    res = app.acquire_token_by_device_flow(flow)  # blocks until the user completes sign-in
    _save(cache)
    if "access_token" not in res:
        raise RuntimeError("Sign-in failed: " + res.get("error_description", str(res)))
    print("\n[OK] Signed in. Token cached at", CACHE, flush=True)
    return res["access_token"]


# ---- graph reads ------------------------------------------------------------
def _get(url, token):
    r = requests.get(url, headers={"Authorization": f"Bearer {token}"}, timeout=30)
    r.raise_for_status()
    return r.json()


def open_positions_count(token=None):
    """Return {'value': <#Open>, 'total': <#rows>} from the live SharePoint list."""
    token = token or get_token_silent()
    if not token:
        raise RuntimeError("No cached token. Run once:  python graph_staffing.py login")

    site = _get(f"{GRAPH}/sites/{SITE_HOST}:{SITE_PATH}", token)
    site_id = site["id"]

    lists = _get(f"{GRAPH}/sites/{site_id}/lists?$top=200", token)
    lst = next((l for l in lists.get("value", [])
                if LIST_NAME.lower() in (str(l.get("displayName", "")).lower(),
                                         str(l.get("name", "")).lower())), None)
    if not lst:
        names = [l.get("displayName") for l in lists.get("value", [])]
        raise RuntimeError(f"List '{LIST_NAME}' not found. Available: {names}")
    list_id = lst["id"]

    count = total = 0
    status_key = None
    url = f"{GRAPH}/sites/{site_id}/lists/{list_id}/items?$expand=fields&$top=200"
    while url:
        data = _get(url, token)
        for item in data.get("value", []):
            total += 1
            fields = item.get("fields", {})
            if status_key is None:
                status_key = next((k for k in fields if k.lower() == STATUS_FIELD_HINT), None)
            val = str(fields.get(status_key, "")).strip().lower() if status_key else ""
            if val == OPEN_VALUE:
                count += 1
        url = data.get("@odata.nextLink")
    return {"value": count, "total": total, "source": "SharePoint Open Positions (Graph, live)"}


if __name__ == "__main__":
    if len(sys.argv) > 1 and sys.argv[1] == "login":
        interactive_login()
    print(json.dumps(open_positions_count(), indent=2))
