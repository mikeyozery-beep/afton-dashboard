#!/usr/bin/env python3
"""
Live download of Kylie's "Afton Revenue Management Dashboard.xlsx" (the Tradeouts /
Renewal Rates source) from SharePoint via Microsoft Graph (DELEGATED auth, device-code
sign-in -> cached refresh token). Self-contained: no dependency on OneDrive sync.

One-time setup (interactive, do this once):
    python graph_revmgmt.py login
        -> opens a microsoft.com/devicelogin code; sign in as a user who can see the
           file (it's shared with mozery@aftonprop.com). Token is cached locally.

After that, build_metrics calls fetch() silently using the cached refresh token
(auto-renews with use). On any failure it returns None and the build falls back to a
local copy. Shares the token cache with graph_staffing.py.

No app registration needed: uses Microsoft's public "Graph CLI" client id.
"""
import sys, json, base64
from pathlib import Path

import msal
import requests

# ---- config -----------------------------------------------------------------
TENANT    = "3f73f992-5f90-448f-a139-a3ceb1ad94e5"          # aftonprop.com
CLIENT_ID = "14d82eec-204b-4c2f-b7e8-296a70dab67e"          # MS Graph Command Line Tools (public client)
AUTHORITY = f"https://login.microsoftonline.com/{TENANT}"
SCOPES    = ["Files.Read.All"]

GRAPH = "https://graph.microsoft.com/v1.0"
CACHE = Path(__file__).parent / "graph_token_cache.json"    # shared with graph_staffing.py

# The shared workbook. driveId/itemId are stable; the web URL is the resilient fallback
# (re-resolves driveId/itemId if the file is ever re-shared or moved).
DRIVE_ID = "b!QhcsImigQkOlvG0Y3W1AD2uRkbLjMlZAr3YDMzMPGVR5Sn0T4HeLTo7KldtVG3RL"
ITEM_ID  = "014BTSNLQSL3ALMCHYMBELM65OKML74Z53"
SHARE_URL = ("https://aftonpropeties-my.sharepoint.com/personal/kylie_aftonprop_com/"
             "Documents/Revenue Management/Analyses/Afton Revenue Management Dashboard.xlsx")

CACHE_DIR  = Path(__file__).parent / ".cache"
CACHE_FILE = CACHE_DIR / "Afton Revenue Management Dashboard.xlsx"


# ---- auth -------------------------------------------------------------------
def _app():
    cache = msal.SerializableTokenCache()
    if CACHE.exists():
        cache.deserialize(CACHE.read_text())
    return msal.PublicClientApplication(CLIENT_ID, authority=AUTHORITY, token_cache=cache), cache


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


def interactive_login(poll_minutes=12):
    import time
    app, cache = _app()
    flow = app.initiate_device_flow(scopes=SCOPES)
    if "user_code" not in flow:
        raise RuntimeError("Device flow failed: " + json.dumps(flow, indent=2))
    print(flow["message"], flush=True)          # "go to microsoft.com/devicelogin and enter CODE"
    # Robust poll: drive our own deadline so a system clock change can't make msal think the
    # code already expired (which makes acquire_token_by_device_flow return after one poll).
    deadline = time.monotonic() + poll_minutes * 60     # monotonic: immune to wall-clock jumps
    interval = int(flow.get("interval", 5))
    res = None
    while time.monotonic() < deadline:
        flow["expires_at"] = time.time() + 600          # keep msal's own exit_condition satisfied
        res = app.acquire_token_by_device_flow(flow, exit_condition=lambda f: True)  # one poll
        if "access_token" in res:
            break
        err = res.get("error")
        if err not in ("authorization_pending", "slow_down"):
            raise RuntimeError("Sign-in failed: " + res.get("error_description", str(res)))
        if err == "slow_down":
            interval += 5
        time.sleep(interval)
    _save(cache)
    if not res or "access_token" not in res:
        raise RuntimeError("Sign-in timed out after %d min (code not entered in time)." % poll_minutes)
    print("\n[OK] Signed in. Token cached at", CACHE, flush=True)
    return res["access_token"]


# ---- download ---------------------------------------------------------------
def _encode_share_url(url):
    b64 = base64.urlsafe_b64encode(url.encode()).decode().rstrip("=")
    return "u!" + b64


def _content_url(token):
    """Resolve the download endpoint, preferring stable driveId/itemId; on 404 fall back
    to re-resolving via the sharing URL."""
    direct = f"{GRAPH}/drives/{DRIVE_ID}/items/{ITEM_ID}/content"
    r = requests.head(direct, headers={"Authorization": f"Bearer {token}"},
                      allow_redirects=False, timeout=30)
    if r.status_code in (200, 302):
        return direct
    item = requests.get(f"{GRAPH}/shares/{_encode_share_url(SHARE_URL)}/driveItem",
                        headers={"Authorization": f"Bearer {token}"}, timeout=30)
    item.raise_for_status()
    j = item.json()
    return j.get("@microsoft.graph.downloadUrl") or \
        f"{GRAPH}/drives/{j['parentReference']['driveId']}/items/{j['id']}/content"


def fetch(dest=CACHE_FILE, token=None):
    """Download the workbook to `dest`. Returns the Path on success, None on any failure
    (so the caller can fall back to a local copy)."""
    try:
        token = token or get_token_silent()
        if not token:
            return None
        url = _content_url(token)
        r = requests.get(url, headers={"Authorization": f"Bearer {token}"}, timeout=120)
        r.raise_for_status()
        dest.parent.mkdir(parents=True, exist_ok=True)
        dest.write_bytes(r.content)
        return dest
    except Exception as e:
        print(f"graph_revmgmt fetch failed: {e}", flush=True)
        return None


if __name__ == "__main__":
    if len(sys.argv) > 1 and sys.argv[1] == "login":
        interactive_login()
    p = fetch()
    print(json.dumps({"downloaded": str(p) if p else None,
                      "bytes": p.stat().st_size if p else 0}, indent=2))
