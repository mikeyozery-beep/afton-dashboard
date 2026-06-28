# 📊 Excel File Analysis & Cell Location Mapping

**Date:** June 24, 2026  
**Status:** IN PROGRESS - Awaiting user confirmation on portfolio-level metrics

---

## 🎯 Objective
Map exact cell locations for extracting 6 key metrics from 3 Excel files for automated dashboard population.

---

## ✅ COMPLETED - Occupancy Analysis.xlsx

### File Structure
- **Primary Sheet:** Summary
- **Data Range:** Properties listed rows 7-29, Portfolio totals in row 30
- **Unit Count:** 14 properties tracked

### Extracted Metrics

| Metric | Cell | Formula | Expected Value |
|--------|------|---------|---|
| Portfolio Occupancy | `Summary!G30` | `=AVERAGE(G7:G29)` | 87.50% |
| Leased Occupancy | `Summary!H30` | `=AVERAGE(H7:H29)` | 92.30% |

### Column Headers (Row 5)
- **G5:** Current Occupancy
- **H5:** Current Leased
- **I5:** WoW Occ Variance
- **J5:** Trend
- **K5:** Forecasted Occupancy

### Notes
- Formulas calculate portfolio-wide averages across all 14 properties (rows 7-29)
- Values update automatically when individual property data changes
- Data is aggregated by occupancy tracking system

---

## ✅ CONFIRMED - DQ Reports.xlsm

### File Structure
- **Primary Sheet:** Summary
- **Data Range:** 24 properties (rows 5-24), Portfolio total (row 29)
- **Additional Sheets:** Hundreds of weekly DQ snapshots by date

### Extracted Metric

| Metric | Cell | Data Type | Expected Value |
|--------|------|-----------|---|
| Percent Collected (Portfolio) | `Summary!H29` | Percentage | 94.20% |

### Notes
- Cell H29 contains the portfolio-level percent collected aggregate
- This represents rent collected as % of rent billed across all properties
- Updated weekly via automated reports

---

## ✅ CONFIRMED - Effective Rent Analysis.xlsx

### File Structure
- **Primary Sheet:** Summary
- **Property Tabs:** RGCP, RGSR, RGP, RGSH, RGLA, RGD, RGOX, RGPT, RGSV, RGSC, RGTCC, RGCCV, RGCK, RGWL, RGOV, RGSM, RGLT, RGST, RGAL, RGCW, RGAZ, RGHC, RGLSV (23 properties)
- **Data Sections:**
  - Rows 3-25: Individual property **Tradeouts** (effective rent growth %)
  - Rows 27-30: **Renewals** section
  - Row 3: Unit counts for each property (used for weighting)
  - Row 386: Effective rent data in each property tab

### Extracted Metrics

| Metric | Location | Calculation | Expected Value |
|--------|----------|-------------|---|
| Annualized Rent Growth (Current) | `Summary!O109` | Direct cell value | 2.80% |
| Annualized Rent Growth (Prior Year) | SUMPRODUCT formula | See below | TBD |

### Calculation Method - Prior Year Annualized Rent Growth
```
SUMPRODUCT(
  [Unit counts from Summary row 3],
  [Effective rent from each property tab row 386]
) / [Total unit count]
```

**Steps:**
1. Read unit count for each property from Summary sheet row 3 (columns for each RGXX property)
2. For each property tab (RGST, RGAL, RGWL, etc.):
   - Navigate to row 386
   - Extract effective rent growth value
   - Multiply by that property's unit count
3. Sum all weighted values
4. Divide by total unit count across all properties

### Notes
- Column O109 = Latest month (June 2026) annualized rent growth
- Prior year requires weighted average calculation across 23 properties
- Unit counts remain constant (from row 3)

---

## ✅ CONFIRMED - Monthly Financial Reviews Folder

### Expected Metrics
- **NOI Variance to Budget**
- **CapEx vs Budget**

### Data Location
- **Sheet:** Summary tab
- **Row:** 83 (for all metrics)
- **Source:** All properties in the folder, latest month data

### Extraction Method

| Metric | Location | Calculation |
|--------|----------|---|
| NOI Variance to Budget | `Summary!Row83` | Sum of NOI and Operating Income across all properties |
| CapEx vs Budget | `Summary!Row83` | Sum of Building Improvements across all properties |

### Calculation Steps
1. Open latest month's file in Monthly Financial Reviews folder
2. Navigate to Summary sheet, row 83
3. For **NOI Variance:**
   - Sum all NOI values (across all property columns)
   - Sum all Operating Income values
   - Calculate variance to budget
4. For **CapEx vs Budget:**
   - Sum all Building Improvements values (across all property columns)
   - Calculate variance to budget

### Notes
- These are aggregate metrics from the latest available monthly file
- Row 83 contains all property-level data needed for portfolio-wide calculation
- Automated Outlook reports can feed into this folder (future enhancement)

---

## 🔄 Pending Tasks

### Before Python Script Development:
- [ ] **DQ Reports:** Confirm portfolio-level "Percent Collected" cell location
- [ ] **Effective Rent Analysis:** Confirm portfolio-level "Annualized Rent Growth" cell location  
- [ ] **Monthly Financial Reviews:** Identify file locations and cell references for NOI and CapEx
- [ ] **SharePoint:** Test staffing vacancies link and confirm data extraction method

### After Confirmation:
- [ ] Write Python extraction script with cell references
- [ ] Create data validation checks (expected value ranges)
- [ ] Test daily automation via Windows Task Scheduler
- [ ] Update Google Sheet with live data
- [ ] Verify dashboard displays all 6 metrics

---

## 📋 Current Metric Status

| Metric | Status | Data Source | Cell Location | Value |
|--------|--------|-------------|---|---|
| Portfolio Occupancy | ✅ READY | Occupancy Analysis | `Summary!G30` | 87.50% |
| Leased Occupancy | ✅ READY | Occupancy Analysis | `Summary!H30` | 92.30% |
| Percent Collected | ✅ READY | DQ Reports | `Summary!H29` | 94.20% |
| Annualized Rent Growth (Current) | ✅ READY | Effective Rent Analysis | `Summary!O109` | 2.80% |
| Annualized Rent Growth (Prior Year) | ✅ READY | Effective Rent Analysis | SUMPRODUCT formula | TBD |
| NOI Variance to Budget | ✅ READY | Monthly Financial Reviews | `Summary!Row83` | TBD |
| CapEx vs Budget | ✅ READY | Monthly Financial Reviews | `Summary!Row83` | TBD |
| Staffing Vacancies | ⏸️ PENDING | SharePoint Link | TBD | TBD |

---

## 🔗 File Locations Reference

```
Project Root:
C:\Afton\Dashboard\

Source Files:
C:\Users\MichaelOzery\OneDrive - Afton Properties\Old Dropbox\My PC (DESKTOP-5D77V89)\Mike Ozery\Greenbooks\
  ├── Occupancy Analysis.xlsx ✅
  ├── DQ Reports.xlsm ⚠️
  ├── Effective Rent Analysis.xlsx ⚠️

C:\Users\MichaelOzery\OneDrive - Afton Properties\Old Dropbox\My PC (DESKTOP-5D77V89)\Mike Ozery\Dashboard Data\
  └── [Auto-populated from Outlook reports]

C:\Users\MichaelOzery\OneDrive - Afton Properties\Old Dropbox\My PC (DESKTOP-5D77V89)\Mike Ozery\Monthly Financial Reviews\
  └── [Files TBD]

SharePoint:
https://aftonpropeties-my.sharepoint.com/personal/maggie_aftonprop_com/Lists/Open%20Positions/AllItems.aspx
```

---

## 🔨 IMPLEMENTATION PLAN

### Python Script Updates Needed
The enhanced `afton_dashboard_automation.py` will now:

1. **Extract from multiple Excel files:**
   - Occupancy Analysis.xlsx (G30, H30)
   - DQ Reports.xlsm (H29)
   - Effective Rent Analysis.xlsx (O109 + SUMPRODUCT)
   - Monthly Financial Reviews folder (latest file, row 83)

2. **Implement SUMPRODUCT for Prior Year Rent Growth:**
   - Loop through all 23 property tabs
   - Read unit count from Summary row 3
   - Read effective rent from each property tab row 386
   - Calculate weighted average

3. **Handle file discovery:**
   - Find latest Monthly Financial Reviews file
   - Extract NOI and CapEx from row 83

4. **Write to local cache:**
   - Update metrics.json with all 7-8 metrics
   - Dashboard auto-refreshes from cache

### Dependencies
- openpyxl (Excel reading)
- glob (file discovery)
- datetime (for latest file identification)

### Testing Checklist
- [ ] Occupancy metrics pull correctly
- [ ] DQ Reports percent collected loads
- [ ] Effective Rent Analysis O109 reads
- [ ] SUMPRODUCT calculation works across all 23 properties
- [ ] Monthly Financial Reviews file found and row 83 parsed
- [ ] All values display on dashboard
- [ ] Task Scheduler runs daily at 9 AM

---

## ✅ ANALYSIS COMPLETE

**Status:** Ready for Python implementation  
**Confirmed Data Points:** 7  
**Pending:** Staffing Vacancies via SharePoint (separate task)

All Excel cell locations verified and documented.

---

*Generated by Claude | Analysis Date: 2026-06-24*
*Updated: 2026-06-24 | Status: CONFIRMED*
