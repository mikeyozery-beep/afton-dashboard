# YARDI Scheduler Reports Analysis
## Data Available from mikereports@aftonprop.com

**Source:** Email inbox export from YARDI automation system  
**Sender:** cdr@yardi.com (YARDI Scheduler)  
**Frequency:** Daily automated reports  
**Format:** Excel (.xlsx) and ZIP archives

---

## CORE OPERATIONAL REPORTS

### 1. **Unit Occupancy Reports**
- **Report Name:** `Unit Occupancy_6413722(Property: .ap-gs)`
- **File:** `Unit_Occupancy_*.xlsx`
- **Data Contains:** Unit-level occupancy status by property
- **Dashboard Relevance:** ✅ **HIGH** - Directly feeds Portfolio & Leased Occupancy metrics
- **Extract Priority:** **PHASE 2** - Enhance current extraction logic

### 2. **Rent Roll Reports**
- **Report Names:**
  - `ResAnalytics_Rent_Roll(Property: .ap-gs)`
  - `rs_sql_am_reports_combined_rent_rolls(Property: .ap-gs)`
  - `rs_sql_rent_roll_muvan(Property: .ap-gs)`
- **Files:** `ResAnalytics_Rent_Roll_*.xlsx`, `rs_sql_am_reports_combined_rent_rolls_*.xlsx`
- **Data Contains:** Detailed lease information, unit status, rent amounts by property
- **Dashboard Relevance:** ✅ **HIGH** - Source for bottom-performing communities table
- **Extract Priority:** **PHASE 2** - Extract bottom 3 properties by occupancy

### 3. **Box Score Summary**
- **Report Name:** `ResAnalytics_Box_Score_Summary(Property: .ap-gs)`
- **File:** `ResAnalytics_Box_Score_Summary_*.xlsx`
- **Data Contains:** High-level property performance metrics (occupancy, rent, collections, financials)
- **Dashboard Relevance:** ✅ **CRITICAL** - Likely contains most metrics needed
- **Extract Priority:** **PHASE 2** - Extract NOI variance, collections, occupancy by property

---

## FINANCIAL REPORTS

### 4. **Income Statement**
- **Report Name:** `Income Statement(Property: .ap-gs Books: Accrual)`
- **File:** `rs_sql_income_statement_all_muvan_*.xlsx`
- **Data Contains:** P&L data - revenue, expenses, NOI by property
- **Dashboard Relevance:** ✅ **HIGH** - Source for NOI Variance & CapEx metrics
- **Extract Priority:** **PHASE 2** - Extract NOI variance to budget, CapEx vs budget

### 5. **Aged Receivables**
- **Report Name:** `ResARAnalytics Financial Aged Receivable(Property: .ap-gs)`
- **File:** `ResAnalytic_Aged_Receivables_*.xlsx`
- **Data Contains:** Outstanding receivables by aging bucket
- **Dashboard Relevance:** ⚠️ **MEDIUM** - Could inform collections % calculations
- **Extract Priority:** **PHASE 3** - Lower priority, nice-to-have

### 6. **Tenant Receivables**
- **Report Names:**
  - `rs_sql_tr_payables(Property: .ap-gs)`
  - `rs_sql_tr_payments_muvan(Property: .ap-gs)`
  - `rs_sql_tr_receipts(Property: .ap-gs)`
  - `rs_sql_tr_charges(Property: .ap-gs)`
- **Files:** `rs_sql_tr_*.xlsx`
- **Data Contains:** Tenant billing, payments received, outstanding charges
- **Dashboard Relevance:** ✅ **HIGH** - Source for % Collected metric
- **Extract Priority:** **PHASE 2** - Extract collections percentage

---

## RESIDENT/TENANT MANAGEMENT REPORTS

### 7. **Resident Activity (MTD)**
- **Report Name:** `rs_sql_resident_activity_mtd_muvan(Property: .ap-gs)`
- **File:** `rs_sql_resident_activity_mtd_muvan_*.xlsx`
- **Data Contains:** Month-to-date resident turnover, move-ins, move-outs
- **Dashboard Relevance:** ⚠️ **MEDIUM** - Supports occupancy trend analysis
- **Extract Priority:** **PHASE 3** - Secondary metric support

### 8. **Lease Renewals**
- **Report Name:** `rs_sql_op_lease_renewals(Property: .ap-gs)`
- **File:** Typically in ZIP bundles
- **Data Contains:** Renewal lease data, renewal rates by property
- **Dashboard Relevance:** ✅ **HIGH** - Source for renewal conversion rate (Areas of Focus)
- **Extract Priority:** **PHASE 2** - Extract renewal conversion rates

### 9. **Resident Lease Expirations**
- **Report Name:** `ResAnalytics_Resident_Lease_Expirations(Property: .ap-gs)`
- **File:** Typically in ZIP bundles
- **Data Contains:** Upcoming lease expirations by date
- **Dashboard Relevance:** ⚠️ **MEDIUM** - Predictive occupancy data
- **Extract Priority:** **PHASE 3** - Future improvements

### 10. **Rent Ready & Evictions**
- **Report Names:**
  - `rs_sql_rent_ready_am_reports(Property: .ap-gs)`
  - `rs_sql_evictions_am_reports(Property: .ap-gs)`
- **Files:** `rs_sql_rent_ready_am_reports_*.xlsx`, `rs_sql_evictions_am_reports_*.xlsx`
- **Data Contains:** Unit status ready for rent, eviction activity
- **Dashboard Relevance:** ⚠️ **MEDIUM** - Operational tracking
- **Extract Priority:** **PHASE 4** - Lower priority

---

## FINANCIAL ANALYSIS REPORTS

### 11. **Gross Potential Rent**
- **Report Name:** `ResAnalytics_Gross_Potential_Rent(Property: .ap-gs)`
- **File:** Typically in ZIP bundles
- **Data Contains:** Maximum potential rental income by unit/property
- **Dashboard Relevance:** ⚠️ **MEDIUM** - Useful for budget comparisons
- **Extract Priority:** **PHASE 3** - Secondary metric

### 12. **Work Order / Maintenance Reports**
- **Report Name:** `rs_sql_wo_am_reports(Property: .ap-gs)`
- **File:** `rs_sql_wo_am_reports_*.xlsx`
- **Data Contains:** Work orders, maintenance status by property
- **Dashboard Relevance:** 🟡 **LOW** - Not currently needed for dashboard
- **Extract Priority:** **PHASE 5** - Future consideration

### 13. **Unit Directory**
- **Report Name:** `custom_unit_directory(Property: .ap-gs)`
- **File:** Typically in ZIP bundles
- **Data Contains:** Unit-level details, features, configuration
- **Dashboard Relevance:** 🟡 **LOW** - Reference data only
- **Extract Priority:** **PHASE 5** - Reference only

---

## CONSOLIDATED SCHEDULER PACKAGES

### 14. **Scheduler_Reports.xlsx / .zip**
- **Content:** Bundle of multiple reports above
- **File:** `Scheduler_Reports.xlsx` or `Scheduler_Reports.zip`
- **Dashboard Relevance:** ✅ **HIGH** - May contain all needed data in one file
- **Extract Priority:** **PHASE 2** - Investigate structure first

---

## EXTRACTION PRIORITY ROADMAP

### 🔴 PHASE 2 (IMMEDIATE - Week 1-2)
Priority metrics that will significantly improve dashboard completeness:

1. **Box Score Summary** → Extract:
   - NOI Variance to Budget
   - Collections % by property
   - Occupancy by property
   - Bottom 3 performers

2. **Income Statement** → Extract:
   - NOI Actual vs Budget (all properties)
   - CapEx Actual vs Budget
   - Financial variance data

3. **Lease Renewals** → Extract:
   - Renewal conversion rate
   - Renewal count by property

4. **Rent Roll** → Extract:
   - Bottom 3 performing properties
   - Property-level metrics (occupancy, rent, collections)

### 🟡 PHASE 3 (Secondary - Week 3-4)
Supporting metrics:

5. **Aged Receivables** → Extract collection trends
6. **Gross Potential Rent** → Budget comparison data
7. **Resident Lease Expirations** → Predictive occupancy

### 🟢 PHASE 4-5 (Future - Nice-to-Have)
Optional enhancements:

- Work Order metrics → Maintenance tracking
- Resident Activity → Turnover analysis
- Unit Directory → Property reference

---

## TECHNICAL NOTES

**File Format:** Most reports are Excel (.xlsx) files with property-level detail  
**Property Naming:** Reports use property identifiers (e.g., ".ap-gs") in names  
**Frequency:** Daily delivery via automated YARDI scheduler  
**Access Method:** Email attachments to mikereports@aftonprop.com  
**Data Freshness:** Daily updates, typically delivered overnight

---

## NEXT STEPS

1. ✅ Identify report files in email inbox
2. ⏳ Download and examine Box Score Summary structure
3. ⏳ Create extraction logic for PHASE 2 reports
4. ⏳ Map Excel sheets to dashboard metrics
5. ⏳ Integrate into Python extraction script
6. ⏳ Test daily automation

---

## DASHBOARD METRICS COVERAGE

**Currently Working (4/34):**
- Portfolio Occupancy ✅
- Leased Occupancy ✅
- % Collected ✅
- Annualized Rent Growth ✅

**Can Extract from YARDI Reports (20/34):**
- NOI Variance to Budget ✅
- CapEx vs Budget ✅
- Trend Occupancy 30-Day ✅
- Bottom 3 Properties ✅
- Chart distributions ✅
- Renewal Conversion Rate ✅
- Additional metrics ✅

**Remaining (10/34):**
- Staffing Vacancies (HR/SharePoint)
- Marketing metrics (CRM/Analytics)
- Strategic/manual items (Key Highlights, Priorities)
