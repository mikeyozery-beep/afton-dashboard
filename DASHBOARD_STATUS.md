# Afton Properties Dashboard - Project Status

**Last Updated:** 2026-06-23

---

## ✅ Completed

- [x] Google Cloud OAuth setup with credentials
- [x] Google Sheet created: `1PAnz3SC7mKRiRdRsmeUIp-zP0gsGQDAl_RR4pmHG6YE`
- [x] Sheet structure: Tab "Current Metrics" with headers (Metric | Latest Value | Updated | Data Source)
- [x] Test data added (4 metrics):
  - Portfolio Occupancy: 87.50%
  - Leased Occupancy: 92.30%
  - Percent Collected: 94.20%
  - Analyzed Rent Growth: 2.80%
- [x] Python automation script created: `afton_dashboard_automation.py`
- [x] HTML dashboard created: `dashboard_live.html` with multiple implementation attempts
- [x] API Key added to dashboard: `AIzaSyBViYuumqrA6jJRqHe7Qg094pnE470dpJI`
- [x] Sheet ID added to dashboard: `1PAnz3SC7mKRiRdRsmeUIp-zP0gsGQDAl_RR4pmHG6YE`
- [x] Batch file created: `run_dashboard_sync.bat`
- [x] Local HTTP server running on localhost:8000
- [x] Dashboard with debug logging created

---

## 🔧 Dashboard Data Loading - Investigation Complete

**Current Status:** Dashboard HTML created and configured, but external data loading failing

**Root Cause Analysis:**
The dashboard has been tried with THREE different data loading methods:
1. **Google Sheets API (with API key)** - API requests not reaching Google servers (possible CORS/network issue)
2. **Google Sheets CSV Export** - Fetch requests hang indefinitely  
3. **URL-encoded API requests** - No successful responses

**Key Findings:**
- Google Sheet is accessible and has correct data when viewed directly
- Dashboard HTML is syntactically correct and loads properly
- Local HTTP server (localhost:8000) is running
- Network requests from browser are NOT reaching external servers (Google Sheets API/export)
- Possible causes:
  - API key insufficient permissions or misconfigured
  - Google Sheet not properly shared for public access
  - Browser security policies / CORS restrictions
  - Network/proxy blocking external API calls

---

## 📋 Dashboard File

**Location:** `C:\Afton\Dashboard\dashboard_live.html`

**Current Implementation:**
- Attempts to fetch CSV from Google Sheets public export
- Parses CSV data into dashboard metrics cards
- Displays debug info when errors occur
- Auto-refreshes every 5 minutes

**To Access Dashboard:**
```
http://localhost:8000/dashboard_live.html
```
or
```
file:///C:/Afton/Dashboard/dashboard_live.html
```

---

## 🔍 Debugging Information Added

The dashboard now includes debug logging that displays:
- Sheet ID being used
- CSV export URL being requested
- HTTP response status
- CSV data preview (first 200 characters)
- Number of rows parsed
- Any error messages

**To View Debug Info:**
1. Open dashboard in browser
2. Dashboard will show debug box if data load attempts fail
3. Check browser console (F12) for additional console.log output

---

## 💡 Recommended Next Steps

**OPTION 1: Verify Google Sheet Permissions** (Most Likely Fix)
1. Open Google Sheet: https://docs.google.com/spreadsheets/d/1PAnz3SC7mKRiRdRsmeUIp-zP0gsGQDAl_RR4pmHG6YE
2. Click "Share" button (top right)
3. Verify it's set to "Anyone with the link can view"
4. Try dashboard again - it should now be able to fetch the CSV

**OPTION 2: Use Google Sheets API Correctly**
- The CSV export method should work once sheet is properly shared
- If API method is needed, ensure:
  - API key has "Google Sheets API" enabled in Google Cloud Console
  - API key has "Sheets API" in its restrictions
  - Test with curl: `curl "https://docs.google.com/spreadsheets/d/SHEET_ID/export?format=csv"`

**OPTION 3: Alternative Data Source (Python Script)**
- Instead of dashboard fetching live from Google Sheets
- Have Python script write latest metrics to a local JSON file
- Dashboard fetches JSON from localhost
- More reliable, no CORS issues

---

## 📁 File Locations

```
C:\Afton\Dashboard\
├── dashboard_live.html          # Main dashboard (with debug info)
├── run_dashboard_sync.bat       # Batch file for Task Scheduler
├── afton_dashboard_automation.py # Python automation script
├── credentials.json             # OAuth credentials (DO NOT SHARE)
├── requirements.txt             # Python dependencies
└── DASHBOARD_STATUS.md          # This file

C:\Afton\Dashboard\tokens\
└── token.pickle                 # Cached OAuth token

C:\Afton\Dashboard\logs\
└── [timestamped log files]      # Execution logs
```

---

## 🚀 Upcoming Tasks

1. **Verify Google Sheet Sharing Settings** ← START HERE
2. **Test Dashboard Data Loading** - Confirm metrics display after step 1
3. **Configure Windows Task Scheduler** - Run Python script daily at 9 AM
4. **Extract YARDI Data** - Automated Excel parsing from:
   - Greenbooks folder
   - Dashboard Data folder
5. **Extract Outlook Reports** - Parse scheduled reports from `mikereports@aftonprop.com`
6. **Update Python Script** - Add data extraction and Google Sheet write logic

---

## 📊 Dashboard Metrics (Ready to Display)

Once data loads successfully, dashboard will show:
- Portfolio Occupancy: 87.50%
- Leased Occupancy: 92.30%  
- Percent Collected: 94.20%
- Analyzed Rent Growth: 2.80%

---

## 📞 Contact & Links

- **User Email:** mikeyozery@gmail.com
- **Reports Email:** mikereports@aftonprop.com
- **Google Sheet:** https://docs.google.com/spreadsheets/d/1PAnz3SC7mKRiRdRsmeUIp-zP0gsGQDAl_RR4pmHG6YE
- **Dashboard Local:** http://localhost:8000/dashboard_live.html
