# Afton Dashboard - Setup Instructions

## Architecture Overview

```
Excel Files (Local)
    ↓ (Python extracts)
Google Sheet (Single source of truth)
    ↓ (Python syncs)
Local JSON Cache (metrics.json)
    ↓ (Dashboard fetches)
HTML Dashboard (Live display)
    ↓ (Runs every 30 seconds)
Browser Display
```

---

## Step 1: Verify Google Sheets API Setup

### Check if credentials exist:
```
ls C:\Afton\Dashboard\credentials.json
```

If NOT found, you need to create OAuth credentials:

1. Go to **Google Cloud Console** (console.cloud.google.com)
2. Create a new project: **"Afton Dashboard"**
3. Enable **Google Sheets API**
4. Create **OAuth 2.0 Desktop Application** credentials
5. Download as JSON and save to: `C:\Afton\Dashboard\credentials.json`

---

## Step 2: Setup Google Sheet Structure

This creates all 7 tabs and headers in your existing Google Sheet.

**Run this ONCE:**
```powershell
cd C:\Afton\Dashboard
python setup_google_sheet.py
```

**Output:**
```
[OK] Google Sheet setup complete!
Sheet URL: https://docs.google.com/spreadsheets/d/1PAnz3SC7mKRiRdRsmeUIp-zP0gsGQDAl_RR4pmHG6YE/edit
```

---

## Step 3: Test Data Extraction

Extract metrics from Excel → Google Sheet + Local JSON

**Run this to test:**
```powershell
cd C:\Afton\Dashboard
python extract_to_sheets.py
```

**What it does:**
1. Reads from Excel files in Greenbooks folder
2. Extracts 8 metrics (Portfolio Occupancy, etc.)
3. Writes to Google Sheet (METRICS tab)
4. Saves to local JSON cache (`data/metrics.json`)

**Check the results:**
- Open your Google Sheet and check the **METRICS** tab
- Check local file: `C:\Afton\Dashboard\data\metrics.json`

---

## Step 4: View the Dashboard

Once extraction is successful, view the dashboard:

1. **Start HTTP server** (if not running):
```powershell
cd C:\Afton\Dashboard
python -m http.server 8000
```

2. **Open in browser:**
```
http://localhost:8000/dashboard_live.html
```

The dashboard will display the extracted metrics and auto-refresh every 30 seconds.

---

## Step 5: Schedule Daily Extraction (Windows Task Scheduler)

Set up automatic extraction at 9 AM daily.

### Option A: Using PowerShell Script

**Create `C:\Afton\Dashboard\run_extraction.ps1`:**
```powershell
cd C:\Afton\Dashboard
python extract_to_sheets.py
```

**Then schedule it:**
1. Open **Task Scheduler**
2. Create Basic Task: "Afton Dashboard Daily Extract"
3. Trigger: Daily at 9:00 AM
4. Action: Run program
   - Program: `powershell.exe`
   - Arguments: `-ExecutionPolicy Bypass -File "C:\Afton\Dashboard\run_extraction.ps1"`

### Option B: Using Batch File

**Create `C:\Afton\Dashboard\run_extraction.bat`:**
```batch
@echo off
cd C:\Afton\Dashboard
python extract_to_sheets.py
pause
```

**Then schedule it in Task Scheduler with this as the program**

---

## Step 6: Verify End-to-End

**Daily checklist:**
1. ✓ Python script runs at 9 AM
2. ✓ Google Sheet updates with latest metrics
3. ✓ Dashboard displays fresh data
4. ✓ No errors in logs (`C:\Afton\Dashboard\logs\`)

---

## Troubleshooting

### "credentials.json not found"
→ Follow Step 1 above to create Google OAuth credentials

### "Could not authenticate with Google"
→ Delete `token.json` and re-run the script to re-authenticate

### "Excel file not found"
→ Check file path in extract_to_sheets.py matches your Greenbooks folder

### "Dashboard shows old data"
→ Manually run: `python extract_to_sheets.py`
→ Refresh browser (Ctrl+F5 to clear cache)

---

## Data Flow Summary

| Step | Action | File | Command |
|------|--------|------|---------|
| 1 | Setup Google Sheet | setup_google_sheet.py | `python setup_google_sheet.py` |
| 2 | Extract & Upload | extract_to_sheets.py | `python extract_to_sheets.py` |
| 3 | View Dashboard | dashboard_live.html | http://localhost:8000/dashboard_live.html |
| 4 | Schedule Daily | Task Scheduler | 9:00 AM daily |

---

## Files Created

```
C:\Afton\Dashboard\
├── setup_google_sheet.py          [One-time setup script]
├── extract_to_sheets.py           [Daily extraction script]
├── dashboard_live.html            [Dashboard display]
├── credentials.json               [Google OAuth (keep secret!)]
├── token.json                     [Generated after first auth]
├── data/
│   └── metrics.json               [Local cache for dashboard]
└── logs/
    └── setup_sheet_*.log          [Setup logs]
    └── sheets_extract_*.log       [Extraction logs]
```

---

## Next: Make It Completely Automated

Once this is working, we can:
1. Set up Google Drive auto-sync for Excel files
2. Add email alerts if extraction fails
3. Add more metrics (NOI, CapEx, Staffing)
4. Deploy dashboard to public cloud (GitHub Pages)
