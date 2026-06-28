# Afton Dashboard - Complete Metrics Mapping

## DATA SOURCE STATUS
- ✅ Currently Extracted from Excel
- ❌ Not Yet Extracted
- 🔄 Needs to be Added/Updated

---

## TOP METRICS (8 Cards in 2 Rows)

### Row 1: Core Occupancy & Collections

**1. PORTFOLIO OCCUPANCY**
- Value: 94.8%
- Change: +50 bps (▲)
- Prior Week: 94.3%
- Budget: 95.0%
- **Source:** ✅ Occupancy Analysis.xlsx, Cell G30
- **Status:** Currently extracted via extract_metrics_enhanced.py

**2. LEASED OCCUPANCY**
- Value: 96.1%
- Change: +40 bps (▲)
- Prior Week: 95.7%
- Budget: 96.0%
- **Source:** ✅ Occupancy Analysis.xlsx, Cell H30
- **Status:** Currently extracted via extract_metrics_enhanced.py

**3. % COLLECTED**
- Value: 99.1%
- Change: +30 bps (▲)
- As of Date: 06/15/26
- Prior Month to Date: 98.8%
- Budget: 99.0%
- **Source:** ✅ DQ Reports.xlsm, Cell H29
- **Status:** Currently extracted via extract_metrics_enhanced.py

**4. ANNUALIZED RENT GROWTH**
- Value: 4.7%
- Change: +40 bps (▲)
- Prior YTD: 4.3%
- Budget: 4.0%
- **Source:** ✅ Effective Rent Analysis.xlsx, Cell O109 (current month) + SUMPRODUCT calculation for prior year
- **Status:** Currently extracted via extract_metrics_enhanced.py

### Row 2: Operational & Financial

**5. TREND OCCUPANCY 30-DAY**
- Value: 95.5%
- Change: +40 bps (▲)
- Budget: 95.0%
- **Source:** ❌ NOT YET EXTRACTED
- **Status:** Needs calculation/extraction from Excel (likely Occupancy Analysis or property-level data)
- **Action Required:** Add extraction logic to calculate 30-day moving average occupancy

**6. # STAFFING VACANCIES**
- Value: 12
- Change: -2 Open (▼)
- Prior Month: 14
- **Source:** ❌ NOT YET EXTRACTED
- **Status:** Need to identify source - likely SharePoint HR data or internal staffing tracker
- **Action Required:** Identify exact data source and create extraction script

**7. NOI VAR. TO BUDGET (THRU JUNE)**
- Value: +2.8%
- Prior YTD: +1.9%
- Budget: 0.0%
- **Source:** ❌ NOT YET EXTRACTED
- **Status:** Need to extract from Monthly Financial Reviews folder
- **Action Required:** Add extraction from financial statements in "5. Monthly Financial Reviews"

**8. CAPEX VS BUDGET**
- Value: +2.8% (shown as dollar amount in cell)
- Amount Over: $0.3M
- YTD Actual: $1.9M
- YTD Budget: $1.8M
- **Source:** ❌ NOT YET EXTRACTED
- **Status:** Need to extract from Monthly Financial Reviews folder
- **Action Required:** Add extraction from capital expenditure tracking

---

## MARKETING PERFORMANCE (5 Cards)

**9. TRAFFIC**
- Value: 12.4K
- Change: +8% vs Prior Month to Date
- **Source:** ❌ NOT YET EXTRACTED
- **Status:** Need to identify source (likely website analytics or CRM)
- **Action Required:** Determine data source and create extraction

**10. LEADS**
- Value: 342
- Change: +12% vs Prior Month to Date
- **Source:** ❌ NOT YET EXTRACTED
- **Status:** Need to identify source (likely CRM or marketing platform)
- **Action Required:** Determine data source and create extraction

**11. APPLICATIONS**
- Value: 28
- Change: +5% vs Prior Month to Date
- **Source:** ❌ NOT YET EXTRACTED
- **Status:** Need to identify source (likely CRM or applicant tracking)
- **Action Required:** Determine data source and create extraction

**12. TOURS OVER LEADS**
- Value: 3.8%
- Prior Month to Date: 4%
- **Source:** ❌ NOT YET EXTRACTED
- **Status:** Likely calculated from Leads and Tours data
- **Action Required:** Calculate ratio from lead and tour data

**13. CONVERSIONS**
- Value: 8.2% (Apps/Tours)
- Prior Month to Date: 8.1%
- **Source:** ❌ NOT YET EXTRACTED
- **Status:** Likely calculated as Applications/Tours ratio
- **Action Required:** Calculate from application and tour data

---

## KEY HIGHLIGHTS (Static Text - 5 Items)

**14. Highlight 1**
- Text: "Occupancy increased 50 bps to 94.8%"
- **Source:** 🔄 MANUAL/DYNAMIC - Could be auto-generated from metric #1
- **Status:** Currently hardcoded in HTML
- **Action Required:** Update to reference Portfolio Occupancy metric dynamically

**15. Highlight 2**
- Text: "Leased occupancy remains strong at 96.1%"
- **Source:** 🔄 MANUAL/DYNAMIC - Could reference metric #2
- **Status:** Currently hardcoded
- **Action Required:** Update to reference Leased Occupancy dynamically

**16. Highlight 3**
- Text: "Collections exceeded budget at 99.1%"
- **Source:** 🔄 MANUAL/DYNAMIC - Could reference metric #3
- **Status:** Currently hardcoded
- **Action Required:** Update to reference % Collected dynamically

**17. Highlight 4**
- Text: "NOI outperformed budget by 2.8%"
- **Source:** 🔄 MANUAL/DYNAMIC - Could reference metric #7
- **Status:** Currently hardcoded
- **Action Required:** Update to reference NOI Var dynamically

**18. Highlight 5**
- Text: "Vacancies improved from 14 to 12"
- **Source:** 🔄 MANUAL/DYNAMIC - Could reference metric #6
- **Status:** Currently hardcoded
- **Action Required:** Update to reference Staffing Vacancies dynamically

---

## AREAS OF FOCUS (Static Text - 3 Items)

**19. Focus Area 1**
- Text: "Lancaster Village drives 75% of vacancies"
- **Source:** ❌ NOT YET EXTRACTED
- **Status:** Hardcoded - needs data source
- **Action Required:** Identify where vacancy breakdown by property comes from

**20. Focus Area 2**
- Text: "Maintenance technician roles remain challenging in 2 markets"
- **Source:** ❌ NOT YET EXTRACTED
- **Status:** Hardcoded - qualitative insight
- **Action Required:** Determine if this should be automated or remain manual

**21. Focus Area 3**
- Text: "Renewal conversion rate down to 58%"
- **Source:** ❌ NOT YET EXTRACTED
- **Status:** Hardcoded - needs extraction
- **Action Required:** Add renewal conversion rate extraction

---

## BOTTOM PERFORMING COMMUNITIES TABLE (3 Properties x 6 Metrics)

### Properties:
**22. Sienna Heights** (Row 1)
- Occupancy: 90.4% ▼ -150 bps
- Leased Occ.: 92.1% ▼ -100 bps
- % Collected: 96.7% ▼ -150 bps
- Rent Growth (Ann.): 1.6% ▼ -120 bps
- NOI Var. to Budget: -3.5% ▼ -330 bps
- **Source:** ❌ NOT YET EXTRACTED - Property-level detail from Excel

**23. Cassia Apartments** (Row 2)
- Occupancy: 88.6% ▼ -120 bps
- Leased Occ.: 90.3% ▼ -80 bps
- % Collected: 96.2% ▼ -140 bps
- Rent Growth (Ann.): 0.9% ▼ -110 bps
- NOI Var. to Budget: -2.9% ▼ -280 bps
- **Source:** ❌ NOT YET EXTRACTED - Property-level detail from Excel

**24. Hills of Corona** (Row 3)
- Occupancy: 89.1% ▼ -100 bps
- Leased Occ.: 89.6% ▼ -70 bps
- % Collected: 97.1% ▼ -120 bps
- Rent Growth (Ann.): 1.2% ▼ -100 bps
- NOI Var. to Budget: -2.3% ▼ -220 bps
- **Source:** ❌ NOT YET EXTRACTED - Property-level detail from Excel

**Action Required:** Extract bottom 3 performing properties and their metrics from Effective Rent Analysis or property-level sheets

---

## PERFORMANCE VS BUDGET (NOI) - PIE CHART

**25. Above Budget (> +1%)**
- Percentage: 53%
- **Source:** ❌ NOT YET EXTRACTED
- **Status:** Chart data hardcoded

**26. On Line (-1% to +1%)**
- Percentage: 27%
- **Source:** ❌ NOT YET EXTRACTED
- **Status:** Chart data hardcoded

**27. Below Budget (< -1%)**
- Percentage: 20%
- **Source:** ❌ NOT YET EXTRACTED
- **Status:** Chart data hardcoded

**Action Required:** Extract NOI variance distribution across properties from financial data

---

## NET INCOME EXCL. INTEREST EXPENSE - PIE CHART

**28. Above Target (> +2%)**
- Percentage: 58%
- **Source:** ❌ NOT YET EXTRACTED
- **Status:** Chart data hardcoded

**29. At Target (-2% to +2%)**
- Percentage: 24%
- **Source:** ❌ NOT YET EXTRACTED
- **Status:** Chart data hardcoded

**30. Below Target (< -2%)**
- Percentage: 18%
- **Source:** ❌ NOT YET EXTRACTED
- **Status:** Chart data hardcoded

**Action Required:** Extract Net Income variance distribution across properties

---

## WEEKLY MEETING TAKEAWAYS (4 Priorities)

**31. Priority 1**
- Text: "Fill 4 open maintenance positions across portfolio"
- **Source:** ❌ NOT YET EXTRACTED
- **Status:** Hardcoded - manual entry
- **Action Required:** Determine if this should be automated or remain manual

**32. Priority 2**
- Text: "Increase occupancy at Lancaster Village from 91% to 93%"
- **Source:** ❌ NOT YET EXTRACTED
- **Status:** Hardcoded - manual entry
- **Action Required:** Determine if this should be automated or remain manual

**33. Priority 3**
- Text: "Push June renewals to achieve 5% annualized rent growth"
- **Source:** ❌ NOT YET EXTRACTED
- **Status:** Hardcoded - manual entry
- **Action Required:** Determine if this should be automated or remain manual

**34. Priority 4**
- Text: "Complete roof replacement at Property X on schedule and budget"
- **Source:** ❌ NOT YET EXTRACTED
- **Status:** Hardcoded - manual entry
- **Action Required:** Determine if this should be automated or remain manual

---

## EXTRACTION PRIORITY SUMMARY

### Phase 1 - Quick Wins (Already Working)
- ✅ Portfolio Occupancy
- ✅ Leased Occupancy
- ✅ % Collected
- ✅ Annualized Rent Growth

### Phase 2 - Essential Financials (Next Priority)
- ❌ NOI Var. to Budget (from Monthly Financial Reviews)
- ❌ CapEx vs Budget (from Monthly Financial Reviews)
- ❌ Trend Occupancy 30-Day (calculate from Occupancy Analysis)
- ❌ Bottom 3 Properties detail (property-level extraction)
- ❌ Chart distributions (NOI & Net Income variance)

### Phase 3 - Operational Metrics
- ❌ Staffing Vacancies (from HR/SharePoint)
- ❌ Renewal Conversion Rate (from CRM/applicant tracking)

### Phase 4 - Marketing Metrics
- ❌ Traffic, Leads, Applications, Tours (from CRM/analytics platform)

### Phase 5 - Manual/Strategic Items
- 🔄 Key Highlights (can auto-generate from metrics)
- 🔄 Areas of Focus (strategic - manual entry recommended)
- 🔄 Weekly Meeting Takeaways (strategic - manual entry recommended)

---

## NEXT STEPS

1. ✅ Set up Windows Task Scheduler (this session)
2. ❌ Create extraction for Phase 2 metrics
3. ❌ Identify data sources for Phase 3 & 4
4. ❌ Update HTML dashboard to populate metrics dynamically
5. ❌ Create manual data entry process for strategic items
