# 🚀 Afton Dashboard Setup Guide

## Quick Start (5 minutes)

### Step 1: Install Python Dependencies
Open **Command Prompt** or **PowerShell** in `C:\Afton\Dashboard` and run:

```bash
pip install google-auth-oauthlib google-auth-httplib2 google-api-python-client
```

### Step 2: Run the Automation Script
Still in `C:\Afton\Dashboard`, run:

```bash
python afton_dashboard_automation.py
```

**First time only - Authorization:**

The script will try to open your browser automatically. If it does:
- Click **"Allow"** when Google asks for permission
- Close the browser and the script will continue

**If the browser doesn't open (manual method):**
1. The script will print a long **authorization URL**
2. Copy the full URL from the console
3. Paste it into your web browser address bar
4. Click **"Allow"** when Google asks for permission
5. Google will show an authorization code
6. **Copy that code** and paste it back into the command prompt
7. Press Enter and the script will continue

### Step 3: Open the Dashboard
Visit: **http://localhost:8000/dashboard_live.html**

You should now see your 4 metrics displayed! ✓

---

## 🔧 How It Works

**Old Approach (Broken):** 
- Dashboard → tries to fetch from Google Sheets API → blocked by CORS

**New Approach (Fixed):**
```
Google Sheet ← Python script reads every 30 seconds
       ↓
   metrics.json (local cache)
       ↓
Dashboard reads from localhost (no CORS issues)
```

---

## 📋 What Each File Does

| File | Purpose |
|------|---------|
| `afton_dashboard_automation.py` | Fetches metrics from Google Sheet every time it runs |
| `dashboard_live.html` | Displays metrics from local data |
| `data/metrics.json` | Cache file (created automatically) |
| `run_dashboard_sync.bat` | Batch file to run Python script |

---

## ⏰ Setup Scheduled Task (Optional)

To have metrics update automatically every 30 minutes:

1. Open **Task Scheduler** (search in Windows)
2. Click **Create Basic Task**
3. Name: `Afton Dashboard Sync`
4. Trigger: **Daily** at any time
5. Action: **Start a program**
   - Program: `C:\Windows\System32\cmd.exe`
   - Arguments: `/c python C:\Afton\Dashboard\afton_dashboard_automation.py`
   - Start in: `C:\Afton\Dashboard`
6. Click **Finish**

---

## ✅ Verification Checklist

- [ ] Python dependencies installed (`pip install ...` ran successfully)
- [ ] Ran `python afton_dashboard_automation.py` once
- [ ] No errors in the log file (`logs/` folder)
- [ ] `data/metrics.json` file was created
- [ ] Dashboard shows 4 metrics at http://localhost:8000/dashboard_live.html
- [ ] Dashboard auto-refreshes every 30 seconds

---

## 🐛 Troubleshooting

**Dashboard shows "Metrics unavailable"**
- Run: `python afton_dashboard_automation.py`
- Check that `data/metrics.json` was created

**Authorization browser didn't open**
- Run: `python afton_dashboard_automation.py`
- A browser WILL open for authorization

**"Could not find a version" error**
- Make sure you're running: `pip install google-auth-oauthlib google-auth-httplib2 google-api-python-client`
- Not `pip install google-sheets-api`

**Metrics show "No cached metrics"**
- Your Google Sheet may not be accessible
- Check: https://docs.google.com/spreadsheets/d/1PAnz3SC7mKRiRdRsmeUIp-zP0gsGQDAl_RR4pmHG6YE
- Make sure it has data in "Current Metrics" tab

---

## 📊 Next Steps (After Verification)

Once the dashboard is showing metrics:

1. **Extract from YARDI files** - Update Python script to read from:
   - `C:\Users\MichaelOzery\OneDrive - Afton Properties\Old Dropbox\My PC\Mike Ozery\Greenbooks`
   - `C:\Users\MichaelOzery\OneDrive - Afton Properties\Old Dropbox\My PC\Mike Ozery\Dashboard Data`

2. **Parse Outlook emails** - Add extraction from `mikereports@aftonprop.com` scheduled reports

3. **Automate the sync** - Set up Task Scheduler to run every morning at 9 AM

---

## 🎯 Dashboard Features

- ✅ **Live metrics display** - Shows latest data from Google Sheet
- ✅ **Auto-refresh** - Updates every 30 seconds
- ✅ **Clean UI** - Afton Properties branding (green #1D9E75)
- ✅ **Status messages** - Clear error reporting
- ✅ **Local caching** - Works even if internet is temporarily down

---

## 📞 Need Help?

Check the logs: `C:\Afton\Dashboard\logs\dashboard_sync_*.log`

Each time you run the script, a new log file is created with detailed information.
