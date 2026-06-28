# How to Build the Google Sheet Structure

## Step 1: Create New Google Sheet

1. Go to **Google Sheets** (sheets.google.com)
2. Click **"+ New"** → **"Blank spreadsheet"**
3. Rename to: **"Afton Properties - Dashboard Data"**
4. Share with anyone who needs access (File → Share)
5. Note the **Sheet ID** (from URL: `https://docs.google.com/spreadsheets/d/{SHEET_ID}/edit`)

---

## Step 2: Create & Format Tabs

### **Delete the default "Sheet1" and create 7 new tabs:**

1. Right-click "Sheet1" → Rename to **"METRICS"**
2. Right-click sheet tab → "Insert 1 right" → Create **"MARKETING"**
3. Repeat for: **"HIGHLIGHTS"**, **"BOTTOM_PERFORMERS"**, **"CHART_DATA"**, **"PRIORITIES"**, **"METADATA"**

**Tab Order (Left to Right):**
```
METRICS → MARKETING → HIGHLIGHTS → BOTTOM_PERFORMERS → CHART_DATA → PRIORITIES → METADATA
```

---

## Step 3: Format Each Tab

### **TAB 1: METRICS** (Primary KPIs)

**Column Headers (Row 1):**
- A1: `Metric Name`
- B1: `Current Value`
- C1: `Change`
- D1: `Detail / Notes`

**Formatting:**
- Header row: Dark green background (#1D4D39), white text, bold
- Freeze row 1: View → Freeze → 1 row
- Column widths: A=220px, B=120px, C=100px, D=300px

**Data (Rows 2-9):**
```
Row 2:  Portfolio Occupancy           | 94.8%  | ▲ +50 bps | Prior Week: 94.3% | Budget: 95.0%
Row 3:  Leased Occupancy              | 96.1%  | ▲ +40 bps | Prior Week: 95.7% | Budget: 96.0%
Row 4:  % Collected                   | 99.1%  | ▲ +30 bps | As of: 06/15/26 | Budget: 99.0%
Row 5:  Annualized Rent Growth        | 4.7%   | ▲ +40 bps | Prior YTD: 4.3% | Budget: 4.0%
Row 6:  Trend Occupancy 30-Day        | 95.5%  | ▲ +40 bps | Budget: 95.0%
Row 7:  # Staffing Vacancies          | 12     | ▼ -2 Open | Prior Month: 14
Row 8:  NOI Var. to Budget (Thru June)| +2.8%  | —         | Prior YTD: +1.9% | Budget: 0.0%
Row 9:  CapEx vs Budget               | +2.8%  | 🔴 RED    | YTD: $1.9M Actual vs $1.8M Budget
```

**Cell Formatting:**
- Column B (Current Value): Green color (#1D9E75), bold
- Column C: Green for ▲, Red for ▼ and RED FLAG
- Column D: Light gray text, italics

---

### **TAB 2: MARKETING**

**Column Headers (Row 1):**
- A1: `Metric`
- B1: `Value`
- C1: `Trend`

**Data (Rows 2-6):**
```
Row 2:  Traffic              | 12.4K | +8% vs Prior Month to Date
Row 3:  Leads                | 342   | +12% vs Prior Month to Date
Row 4:  Applications         | 28    | +5% vs Prior Month to Date
Row 5:  Tours over Leads     | 3.8%  | Prior Month to Date: 4.0%
Row 6:  Conversions          | 8.2%  | Prior Month to Date: 8.1%
```

**Formatting:**
- Same header styling as METRICS
- Column widths: A=150px, B=100px, C=250px

---

### **TAB 3: HIGHLIGHTS**

**Left Column (Column A): Key Highlights**
- A1: Header = "KEY HIGHLIGHTS" (green background, white text)
- A2-A7: Bullet points (6 items)

**Right Column (Column B): Areas of Focus**
- B1: Header = "AREAS OF FOCUS" (green background, white text)
- B2-B4: Warning boxes (3 items)

**Formatting:**
- Column A: White background, left-aligned
- Column B: Alternating rows with light orange background (#FFF4E6)
- Both columns: Center vertically with row height 40px

---

### **TAB 4: BOTTOM_PERFORMERS**

**Column Headers (Row 1):**
- A1: `Community`
- B1: `Occupancy`
- C1: `Leased Occ.`
- D1: `% Collected`
- E1: `Rent Growth (Ann.)`
- F1: `NOI Var. to Budget`

**Data (Rows 2-4):**
```
Row 2:  Sienna Heights    | 90.4% ▼ -150 bps | 92.1% ▼ -100 bps | 96.7% ▼ -150 bps | 1.6% ▼ -120 bps | -3.5% ▼ -330 bps
Row 3:  Cassia Apartments | 88.6% ▼ -120 bps | 90.3% ▼ -80 bps  | 96.2% ▼ -140 bps | 0.9% ▼ -110 bps | -2.9% ▼ -280 bps
Row 4:  Hills of Corona   | 89.1% ▼ -100 bps | 89.6% ▼ -70 bps  | 97.1% ▼ -120 bps | 1.2% ▼ -100 bps | -2.3% ▼ -220 bps
```

**Formatting:**
- All negative indicators in red (#D9534F)
- Numbers: center-aligned
- Column widths: A=180px, B-F=140px each

---

### **TAB 5: CHART_DATA**

**Section 1: NOI Performance (Rows 1-3)**
```
Row 1: Header = "NOI Performance"
Row 2: Category | Percentage
Row 3: Above Budget | 53
Row 4: On Line | 27
Row 5: Below Budget | 20
```

**Section 2: Net Income (Rows 7-11)**
```
Row 7: Header = "Net Income Excl. Interest"
Row 8: Category | Percentage
Row 9: Above Target | 58
Row 10: At Target | 24
Row 11: Below Target | 18
```

**Formatting:**
- Headers: Bold, 13px
- Data values: Green (#1D9E75) if positive, red if negative
- Column widths: A=200px, B=100px

---

### **TAB 6: PRIORITIES**

**Column Headers (Row 1):**
- A1: `#`
- B1: `Priority Item`

**Data (Rows 2-5):**
```
Row 2:  1 | Fill 4 open maintenance positions across portfolio
Row 3:  2 | Increase occupancy at Lancaster Village from 91% to 93%
Row 4:  3 | Push June renewals to achieve 5% annualized rent growth
Row 5:  4 | Complete roof replacement at Property X on schedule/budget
```

**Formatting:**
- Column A: Centered, bold green background (#1D9E75), white text, 20px width
- Column B: Left-aligned, wrap text
- Row height: 40px for readability

---

### **TAB 7: METADATA**

**Column Headers (Row 1):**
- A1: `Field`
- B1: `Value`

**Data (Rows 2-5):**
```
Row 2:  Last Updated      | [Auto-updated by script]
Row 3:  Report Date       | June 15, 2026
Row 4:  Data Source       | Excel Extract (Python)
Row 5:  Next Update       | Daily 9:00 AM
```

**Formatting:**
- Column A: Light gray background (#f5f5f5), bold
- Column B: White background
- Column widths: A=150px, B=300px

---

## Step 4: Apply Conditional Formatting

### **In METRICS tab:**
- Column B (Current Value): Apply green text color to all percentages
- Column C (Change): Red text for "▼" items, orange for "🔴 RED"

### **In BOTTOM_PERFORMERS tab:**
- All cells with "▼": Apply red text color (#D9534F)
- All percentage values: Center alignment

---

## Step 5: Share & Document

**Share with:**
- Your team (edit access)
- Dashboard viewers (read-only link)

**Add to Description:**
```
This is the source of truth for Afton Properties Dashboard metrics.
Updated daily at 9 AM from Excel files via Python automation script.
Do NOT edit manually - all data comes from automated extraction.
```

---

## Next: Connecting to Dashboard

**Once the Google Sheet is built, we will:**
1. Set up Google Sheets API authentication
2. Configure Python script to PUSH data from Excel → Google Sheet
3. Configure dashboard to PULL data from Google Sheet → Display
4. Schedule daily updates via Windows Task Scheduler

**The data flow will be:**
```
Excel Files (Greenbooks folder)
         ↓ (Python script extracts)
Google Sheet (Single source of truth)
         ↓ (Dashboard fetches)
HTML Dashboard (Live display)
```
