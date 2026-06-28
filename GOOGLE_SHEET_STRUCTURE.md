# Afton Properties Dashboard - Google Sheet Input Structure

## Overview
This Google Sheet will be the **single source of truth** for all dashboard data. Each section maps directly to the dashboard display.

---

## Sheet Tabs Structure

### **TAB 1: METRICS (Primary KPIs)**
Location: `Sheet1` → Columns A-D

```
| Metric Name                      | Current Value | Change | Detail                              |
|----------------------------------|---------------|--------|-------------------------------------|
| Portfolio Occupancy              | 94.8%         | +50 bps| Prior Week: 94.3%, Budget: 95.0%   |
| Leased Occupancy                 | 96.1%         | +40 bps| Prior Week: 95.7%, Budget: 96.0%   |
| % Collected                      | 99.1%         | +30 bps| As of: 06/15/26, Budget: 99.0%     |
| Annualized Rent Growth           | 4.7%          | +40 bps| Prior YTD: 4.3%, Budget: 4.0%      |
| Trend Occupancy 30-Day           | 95.5%         | +40 bps| Budget: 95.0%                      |
| # Staffing Vacancies             | 12            | -2 Open| Prior Month: 14                    |
| NOI Var. to Budget (Thru June)   | +2.8%         | —      | Prior YTD: +1.9%, Budget: 0.0%     |
| CapEx vs Budget                  | +2.8%         | RED    | YTD: $1.9M Actual vs $1.8M Budget  |
```

---

### **TAB 2: MARKETING**
Location: `Marketing` → Columns A-C

```
| Metric              | Value | Detail                      |
|---------------------|-------|------------------------------|
| Traffic             | 12.4K | +8% vs Prior Month to Date   |
| Leads               | 342   | +12% vs Prior Month to Date  |
| Applications        | 28    | +5% vs Prior Month to Date   |
| Tours over Leads    | 3.8%  | Prior Month to Date: 4%      |
| Conversions         | 8.2%  | Prior Month to Date: 8.1%    |
```

---

### **TAB 3: HIGHLIGHTS & FOCUS**
Location: `Highlights` → Columns A-B

#### **Key Highlights** (Column A)
```
1. Occupancy increased 50 bps to 94.8%
2. Leased occupancy remains strong at 96.1%
3. Collections exceeded budget at 99.1%
4. NOI outperformed budget by 2.8%
5. Vacancies improved from 14 to 12
6. Trade outs and renewals on track
```

#### **Areas of Focus** (Column B)
```
1. Lancaster Village drives 75% of vacancies
2. Maintenance technician roles remain challenging in 2 markets
3. Renewal conversion rate down to 58%
```

---

### **TAB 4: BOTTOM PERFORMERS**
Location: `BottomPerformers` → Columns A-F

```
| Community         | Occupancy          | Leased Occ.        | % Collected        | Rent Growth        | NOI Var. to Budget |
|-------------------|--------------------|--------------------|--------------------|--------------------|-------------------|
| Sienna Heights    | 90.4% ▼ -150 bps   | 92.1% ▼ -100 bps   | 96.7% ▼ -150 bps   | 1.6% ▼ -120 bps    | -3.5% ▼ -330 bps |
| Cassia Apartments | 88.6% ▼ -120 bps   | 90.3% ▼ -80 bps    | 96.2% ▼ -140 bps   | 0.9% ▼ -110 bps    | -2.9% ▼ -280 bps |
| Hills of Corona   | 89.1% ▼ -100 bps   | 89.6% ▼ -70 bps    | 97.1% ▼ -120 bps   | 1.2% ▼ -100 bps    | -2.3% ▼ -220 bps |
```

---

### **TAB 5: CHART DATA**
Location: `ChartData` → Columns A-B

#### **NOI Performance** (Rows 1-3)
```
| Category        | Percentage |
|-----------------|------------|
| Above Budget    | 53         |
| On Line         | 27         |
| Below Budget    | 20         |
```

#### **Net Income** (Rows 5-7)
```
| Category        | Percentage |
|-----------------|------------|
| Above Target    | 58         |
| At Target       | 24         |
| Below Target    | 18         |
```

---

### **TAB 6: PRIORITY ITEMS**
Location: `Priorities` → Columns A-B

```
| # | Priority Item                                              |
|---|-------------------------------------------------------------|
| 1 | Fill 4 open maintenance positions across portfolio         |
| 2 | Increase occupancy at Lancaster Village from 91% to 93%    |
| 3 | Push June renewals to achieve 5% annualized rent growth    |
| 4 | Complete roof replacement at Property X on schedule/budget |
```

---

### **TAB 7: METADATA**
Location: `Metadata` → Columns A-B

```
| Field            | Value                 |
|------------------|-----------------------|
| Last Updated     | [TIMESTAMP]          |
| Report Date      | June 15, 2026        |
| Data Source      | Excel Extract        |
| Next Update      | [SCHEDULED TIME]     |
```

---

## Data Flow Architecture

```
Excel Files (Local)
    ↓
Python Extract Script (extract_metrics_enhanced.py)
    ↓
Google Sheet (This structure)
    ↓
JSON Cache (metrics.json)
    ↓
Dashboard (dashboard_live.html)
```

---

## Implementation Notes

1. **No formulas needed** - All data is input from Excel via Python script
2. **Color coding**: Use conditional formatting for red/green values
3. **Date field**: Always shows last extraction timestamp
4. **Refresh cadence**: Daily at 9 AM via Windows Task Scheduler
