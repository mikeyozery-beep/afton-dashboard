#!/usr/bin/env python3
"""
Afton Properties Executive Dashboard - consolidated metrics builder.

Reads every live source (Occupancy, Collections, Effective Rent, Box Score,
Monthly Financial Reviews, Budgets, Open Positions, Weekly Meeting Notes) and
writes ONE rich JSON the dashboard consumes: data/dashboard.json

Run daily; financial sections are month-stamped (currently April 2026, the
latest completed Financial Review).
"""
import json, shutil, tempfile, time, traceback
from pathlib import Path
from datetime import datetime

import openpyxl
from openpyxl.utils import get_column_letter

# ----------------------------------------------------------------------------
# Paths
# ----------------------------------------------------------------------------
MO   = Path(r"C:\Users\MichaelOzery\OneDrive - Afton Properties\Old Dropbox\My PC (DESKTOP-5D77V89)\Mike Ozery")
GB   = MO / "Greenbooks"
DD   = MO / "Dashboard Data"
BOX  = DD / "BoxScore Summary"
FIN  = MO / "5. Monthly Financial Reviews"
BUDG = MO / "2026 Budgets" / "2026 Budgets Final - Broken Links"
IDX  = MO / "Portfolio Index.xlsx"

OCC_FILE  = GB / "Occupancy Analysis.xlsx"
DQ_FILE   = GB / "DQ Reports.xlsm"
RENT_FILE = GB / "Effective Rent Analysis.xlsx"
OPEN_POS  = DD / "Open Positions.xlsx"
NOTES_DIR = DD / "Weekly Meeting Notes"

SCRIPT_DIR = Path(__file__).parent
OUT_FILE   = SCRIPT_DIR / "data" / "dashboard.json"

# Which Financial Review month folder to use (latest completed). Auto-detect newest "NN. Month YYYY".
FIN_PRIOR_MONTH_STAFFING = 12   # prior-month staffing vacancy count for the delta (stored, per spec)

# ----------------------------------------------------------------------------
# Helpers
# ----------------------------------------------------------------------------
def load(path):
    """Open a workbook read-only, working around Windows/OneDrive file locks."""
    try:
        return openpyxl.load_workbook(path, read_only=True, data_only=True)
    except PermissionError:
        tmp = Path(tempfile.gettempdir()) / ("aftonlock_" + path.name)
        last = None
        for attempt in range(4):
            try:
                shutil.copy2(path, tmp)
                return openpyxl.load_workbook(tmp, read_only=True, data_only=True)
            except PermissionError as e:
                last = e
                time.sleep(0.5 * (attempt + 1))
        raise last

def num(v):
    try:
        if v is None or v == "":
            return None
        return float(v)
    except (TypeError, ValueError):
        return None

def code_from(s):
    """Normalize a property code: 'ap-rgsh' / 'RGSH Final' -> 'RGSH'."""
    if not s:
        return None
    s = str(s).strip().upper()
    if s.startswith("AP-"):
        s = s[3:]
    return s.split()[0].split("/")[0].strip()

# ----------------------------------------------------------------------------
# Section extractors  (each returns plain dict; failures are isolated)
# ----------------------------------------------------------------------------
def get_occupancy():
    """Occupancy Analysis -> Summary sheet. Portfolio + per-property occ/leased."""
    wb = load(OCC_FILE); ws = wb["Summary"]
    as_of = None
    f1 = ws["F1"].value
    if isinstance(f1, str) and "as of" in f1.lower():
        as_of = f1.split("as of")[-1].strip()
    per_property = {}
    names = {}
    for r in range(7, 30):  # rows 7-29 are properties; row 30 = Total/Average
        name = ws.cell(r, 6).value           # F
        code = code_from(ws.cell(r, 4).value)  # D
        if not name or str(name).startswith("Total"):
            continue
        per_property[code] = {
            "name": str(name).strip(),
            "occupancy": num(ws.cell(r, 7).value),   # G
            "leased":    num(ws.cell(r, 8).value),   # H
        }
        names[code] = str(name).strip()
    res = {
        "portfolio_occupancy_current":    num(ws["G30"].value),
        "portfolio_occupancy_prior_week": num(ws["S30"].value),
        "leased_occupancy_current":       num(ws["H30"].value),
        "leased_occupancy_prior_week":    num(ws["T30"].value),
        "trend_occupancy_30d":            num(ws["K30"].value),  # K = Forecasted/Trend Occupancy (level)
        "as_of": as_of,
        "_per_property": per_property,
        "_names": names,
    }
    wb.close()
    return res

def get_collections():
    """DQ Reports.xlsm -> Summary sheet. Portfolio H29 + per-property H."""
    wb = load(DQ_FILE); ws = wb["Summary"]
    as_of = None
    e1 = ws["E1"].value
    if isinstance(e1, str) and "as of" in e1.lower():
        as_of = e1.split("as of")[-1].strip()
    per_property = {}
    for r in range(6, 29):  # rows 6-28 props; row 29 = Total/Average
        name = ws.cell(r, 5).value            # E
        code = code_from(ws.cell(r, 3).value)   # C
        if not name or str(name).startswith("Total"):
            continue
        per_property[code] = num(ws.cell(r, 8).value)  # H = % rent collected
    res = {
        "percent_collected_current": num(ws["H29"].value),
        "as_of": as_of,
        "_per_property": per_property,
    }
    wb.close()
    return res

def get_rent_growth():
    """Effective Rent Analysis -> Summary. Row 109 = YoY %, O=current, N=prior."""
    wb = load(RENT_FILE); ws = wb["Summary"]
    res = {
        "rent_growth_current": num(ws["O109"].value),
        "rent_growth_prior":   num(ws["N109"].value),
    }
    wb.close()
    return res

def get_box_score():
    """Latest Box Score Summary -> Report sheet, 'Conversion Ratios' block."""
    files = sorted(BOX.glob("*.xlsx"), key=lambda p: p.stat().st_mtime)
    if not files:
        return {"error": "no box score files"}
    wb = load(files[-1]); ws = wb["Report"]
    period = None
    a3 = ws["A3"].value
    if isinstance(a3, str):
        period = a3.replace("Date =", "").strip()

    # Conversion Ratios block: header row 73, total row 101, data rows 74-100.
    def col_total(col):
        v = ws.cell(101, col).value
        n = num(v)
        if n is not None:
            return n
        return sum(num(ws.cell(r, col).value) or 0 for r in range(74, 101))

    # Columns: C=Calls D=Walk-in E=Email F=Other G=SMS H=Web I=Chat
    #          J=Unq First Contact  K=Show(tours)  L=Applied(apps)
    contact_cols = [3, 4, 5, 6, 7, 8, 9]
    traffic = int(sum(col_total(c) for c in contact_cols))
    tours   = int(col_total(11))   # Show
    apps    = int(col_total(12))   # Applied
    res = {
        "traffic": traffic,
        "tours": tours,
        "applications": apps,
        "tours_over_traffic": (tours / traffic) if traffic else None,
        "conversions_apps_over_tours": (apps / tours) if tours else None,
        "period": period,
        "source_file": files[-1].name,
    }
    wb.close()
    return res

def latest_fin_folder():
    cands = []
    for p in FIN.iterdir():
        if p.is_dir() and p.name[:2].isdigit() and "Financial Reviews" in p.name:
            cands.append(p)
    if not cands:
        return None
    # name like "04. April 2026 Financial Reviews" -> sort by leading number within newest year
    def key(p):
        return p.stat().st_mtime
    return sorted(cands, key=key)[-1]

def get_financials():
    """Aggregate the 23 'RGxx Final.xlsx' Financial Snapshots (YTD vs budget)."""
    folder = latest_fin_folder()
    if folder is None:
        return {"error": "no financial review folder"}
    month_label = folder.name.split(" Financial Reviews")[0].split(". ")[-1]

    noi_ytd = noi_bud = capex_act = capex_bud = 0.0
    noi_var_by_code, ni_var_by_code = {}, {}
    noi_buckets = {"above": 0, "online": 0, "below": 0}
    ni_buckets  = {"above": 0, "at": 0, "below": 0}
    n = 0
    for fp in sorted(folder.glob("RG* Final.xlsx")):
        code = code_from(fp.name)
        try:
            wb = load(fp); ws = wb["Financial Snapshot"]
            # row 83 NOI, row 104 Net Income, row 88 Unit Impr, row 94 Building Impr
            nj, nk, nl = num(ws.cell(83, 10).value), num(ws.cell(83, 11).value), num(ws.cell(83, 12).value)
            il = num(ws.cell(104, 12).value)
            u_j, u_k = num(ws.cell(88, 10).value), num(ws.cell(88, 11).value)
            b_j, b_k = num(ws.cell(94, 10).value), num(ws.cell(94, 11).value)
            wb.close()
        except Exception:
            continue
        n += 1
        if nj is not None: noi_ytd += nj
        if nk is not None: noi_bud += nk
        for val, acc in ((u_j, "ca"), (b_j, "ca")):
            pass
        capex_act += (u_j or 0) + (b_j or 0)
        capex_bud += (u_k or 0) + (b_k or 0)
        if nl is not None:
            noi_var_by_code[code] = nl
            if   nl >  0.01: noi_buckets["above"]  += 1
            elif nl < -0.01: noi_buckets["below"]  += 1
            else:            noi_buckets["online"] += 1
        if il is not None:
            ni_var_by_code[code] = il
            if   il >  0.02: ni_buckets["above"] += 1
            elif il < -0.02: ni_buckets["below"] += 1
            else:            ni_buckets["at"]    += 1

    def pct(d):
        tot = sum(d.values()) or 1
        return {k: round(100 * v / tot) for k, v in d.items()}

    return {
        "month_label": month_label,
        "properties_counted": n,
        "noi_variance_to_budget_ytd": (noi_ytd / noi_bud - 1) if noi_bud else None,
        "capex_actual_ytd": capex_act,
        "capex_budget_ytd": capex_bud,
        "capex_vs_budget_ytd": (capex_act / capex_bud - 1) if capex_bud else None,
        "noi_distribution": pct(noi_buckets),
        "net_income_distribution": pct(ni_buckets),
        "_noi_var_by_code": noi_var_by_code,
    }

def get_units():
    """Unit counts per property code from DQ Reports 'Index' sheet (A=code, B=#units)."""
    units = {}
    try:
        wb = load(DQ_FILE); ws = wb["Index"]
        for r in range(2, ws.max_row + 1):
            code = ws.cell(r, 1).value
            u = num(ws.cell(r, 2).value)
            if not code:
                continue
            if u is not None:
                units[code_from(code)] = int(u)
        wb.close()
    except Exception:
        pass
    return units

def get_budget_targets(month_num, units):
    """Weighted-average Occupancy & Collections budget targets from per-property
    budget files. Assumptions sheet: row 5=Occupancy, row 12=Collections %.
    Month columns start at C(=Jan); column = month_num + 2."""
    col = month_num + 2  # C=3 for January
    acc = {"occupancy": [0.0, 0.0], "collected": [0.0, 0.0]}  # [weighted_sum, units]
    for fp in BUDG.glob("RG*.xlsx"):
        code = code_from(fp.stem)
        w = units.get(code, 0)
        if w == 0:
            continue
        try:
            wb = load(fp)
            ws = wb["Assumptions"] if "Assumptions" in wb.sheetnames else wb.active
            occ = num(ws.cell(5, col).value)
            coll = num(ws.cell(12, col).value)
            wb.close()
        except Exception:
            continue
        if occ is not None:
            occ = occ / 100.0 if occ > 1.5 else occ
            acc["occupancy"][0] += occ * w; acc["occupancy"][1] += w
        if coll is not None:
            acc["collected"][0] += coll * w; acc["collected"][1] += w

    def wavg(key):
        s, u = acc[key]
        return (s / u) if u else None
    return {
        "occupancy_budget": wavg("occupancy"),
        "collected_budget": wavg("collected"),
    }

def get_rent_growth_budget(units):
    """Annualized budgeted rent growth: per property, take Total Potential Rent
    (GL 4212-0000) from the Budget tab of each financial-review file for the most
    recent month vs the prior month, annualize the MoM change, then weight by units."""
    folder = latest_fin_folder()
    if folder is None:
        return None
    wsum = wunits = 0.0
    for fp in sorted(folder.glob("RG* Final.xlsx")):
        code = code_from(fp.name)
        w = units.get(code, 0)
        if w == 0:
            continue
        try:
            wb = load(fp)
            if "Budget" not in wb.sheetnames:
                wb.close(); continue
            ws = wb["Budget"]
            # locate the "Total" column in header row 5 -> recent = Total-1, prior = Total-2
            total_col = None
            for c in range(3, ws.max_column + 1):
                if str(ws.cell(5, c).value).strip().lower() == "total":
                    total_col = c; break
            # locate Total Potential Rent row (GL 4212-0000)
            pot_row = None
            for r in range(6, min(ws.max_row, 80) + 1):
                if str(ws.cell(r, 1).value).strip() == "4212-0000":
                    pot_row = r; break
            wb_recent = wb_prior = None
            if total_col and pot_row:
                wb_recent = num(ws.cell(pot_row, total_col - 1).value)
                wb_prior  = num(ws.cell(pot_row, total_col - 2).value)
            wb.close()
        except Exception:
            continue
        if wb_recent and wb_prior and wb_prior > 0:
            annual = (wb_recent / wb_prior) ** 12 - 1   # compound-annualized MoM change
            wsum += annual * w; wunits += w
    return (wsum / wunits) if wunits else None

def get_staffing():
    """Open Positions.xlsx (SharePoint export) -> count Status == Open."""
    if not OPEN_POS.exists():
        return {"value": None, "status": "FILE_NOT_FOUND"}
    wb = load(OPEN_POS); ws = wb.active
    # find Status column from header row
    status_col = None
    for c in range(1, ws.max_column + 1):
        if str(ws.cell(1, c).value).strip().lower() == "status":
            status_col = c; break
    open_count = 0
    if status_col:
        for r in range(2, ws.max_row + 1):
            v = ws.cell(r, status_col).value
            if isinstance(v, str) and v.strip().lower() == "open":
                open_count += 1
    wb.close()
    return {"value": open_count, "prior_month": FIN_PRIOR_MONTH_STAFFING,
            "source": "Open Positions.xlsx (SharePoint export)"}

def get_takeaways():
    """Latest .docx in Weekly Meeting Notes -> list-paragraph bullets."""
    try:
        import docx
    except ImportError:
        return {"items": [], "error": "python-docx not installed"}
    docs = sorted(NOTES_DIR.glob("*.docx"), key=lambda p: p.stat().st_mtime)
    docs = [d for d in docs if not d.name.startswith("~$")]
    if not docs:
        return {"items": [], "error": "no notes found"}
    d = docx.Document(str(docs[-1]))
    bullets = []
    for p in d.paragraphs:
        t = p.text.strip()
        if not t:
            continue
        if "List" in (p.style.name or "") or t.startswith(("•", "-", "*")):
            bullets.append(t.lstrip("•-* ").strip())
    return {"items": bullets, "source_file": docs[-1].name}

# ----------------------------------------------------------------------------
# Analysis: bottom performers + focus areas (anomalies)
# ----------------------------------------------------------------------------
def build_analysis(occ, coll, fin):
    names = occ["_names"]
    occ_pp  = occ["_per_property"]
    coll_pp = coll["_per_property"]
    noi_pp  = fin.get("_noi_var_by_code", {})

    rows = []
    for code, name in names.items():
        o = occ_pp.get(code, {})
        rows.append({
            "code": code, "name": name,
            "occupancy": o.get("occupancy"),
            "leased": o.get("leased"),
            "collected": coll_pp.get(code),
            "noi_variance": noi_pp.get(code),
        })

    # composite score across available metrics (lower = worse). Normalize each metric 0..1.
    metrics = ["occupancy", "leased", "collected", "noi_variance"]
    ranges = {}
    for m in metrics:
        vals = [r[m] for r in rows if r[m] is not None]
        ranges[m] = (min(vals), max(vals)) if vals else (0, 1)
    for r in rows:
        parts = []
        for m in metrics:
            lo, hi = ranges[m]
            if r[m] is None or hi == lo:
                continue
            parts.append((r[m] - lo) / (hi - lo))
        r["score"] = sum(parts) / len(parts) if parts else 1.0

    bottom = sorted(rows, key=lambda r: r["score"])[:3]
    bottom_out = [{
        "name": r["name"],
        "occupancy": r["occupancy"], "leased": r["leased"],
        "collected": r["collected"], "noi_variance": r["noi_variance"],
        "rent_growth": None,  # not available per-property in a single cell
    } for r in bottom]

    # Focus areas = occupancy anomalies (lowest occupancy properties)
    occ_sorted = sorted([r for r in rows if r["occupancy"] is not None],
                        key=lambda r: r["occupancy"])[:3]
    sev = ["High", "Medium", "Medium"]
    focus = [{
        "rank": i + 1,
        "title": "Occupancy Concern",
        "summary": f"{r['name']} occupancy at {r['occupancy']*100:.1f}%",
        "property": r["name"],
        "severity": sev[i] if i < len(sev) else "Medium",
    } for i, r in enumerate(occ_sorted)]

    return bottom_out, focus

# ----------------------------------------------------------------------------
# Main
# ----------------------------------------------------------------------------
def safe(fn, label, errors):
    try:
        return fn()
    except Exception as e:
        errors.append(f"{label}: {e}")
        traceback.print_exc()
        return {}

def main():
    errors = []
    now = datetime.now()

    occ  = safe(get_occupancy,    "occupancy",   errors)
    coll = safe(get_collections,  "collections", errors)
    rent = safe(get_rent_growth,  "rent_growth", errors)
    box  = safe(get_box_score,    "box_score",   errors)
    fin  = safe(get_financials,   "financials",  errors)
    units = safe(get_units,       "units",       errors) or {}
    bud  = safe(lambda: get_budget_targets(now.month, units), "budgets", errors)
    rent_budget = safe(lambda: get_rent_growth_budget(units), "rent_budget", errors)
    staff= safe(get_staffing,     "staffing",    errors)
    take = safe(get_takeaways,    "takeaways",   errors)

    bottom, focus = [], []
    if occ and coll:
        try:
            bottom, focus = build_analysis(occ, coll, fin or {})
        except Exception as e:
            errors.append(f"analysis: {e}"); traceback.print_exc()

    out = {
        "generated_at": now.isoformat(),
        "as_of": {
            "occupancy":   occ.get("as_of"),
            "collections": coll.get("as_of"),
            "financials":  fin.get("month_label"),
            "boxscore":    box.get("period"),
        },
        "kpis": {
            "portfolio_occupancy": {
                "value": occ.get("portfolio_occupancy_current"),
                "prior_week": occ.get("portfolio_occupancy_prior_week"),
                "budget": bud.get("occupancy_budget"),
            },
            "leased_occupancy": {
                "value": occ.get("leased_occupancy_current"),
                "prior_week": occ.get("leased_occupancy_prior_week"),
                "budget": None,
            },
            "percent_collected": {
                "value": coll.get("percent_collected_current"),
                "budget": bud.get("collected_budget"),
            },
            "rent_growth": {
                "value": rent.get("rent_growth_current"),
                "prior": rent.get("rent_growth_prior"),
                "budget": rent_budget,
            },
            "trend_occupancy_30d": {
                "value": occ.get("trend_occupancy_30d"),
                "budget": bud.get("occupancy_budget"),
            },
            "staffing_vacancies": {
                "value": staff.get("value"),
                "prior_month": staff.get("prior_month"),
            },
            "noi_variance": {
                "value": fin.get("noi_variance_to_budget_ytd"),
                "as_of": fin.get("month_label"),
            },
            "capex_vs_budget": {
                "value": fin.get("capex_vs_budget_ytd"),
                "actual": fin.get("capex_actual_ytd"),
                "budget": fin.get("capex_budget_ytd"),
                "as_of": fin.get("month_label"),
            },
        },
        "marketing": {
            "traffic": box.get("traffic"),
            "tours": box.get("tours"),
            "applications": box.get("applications"),
            "tours_over_traffic": box.get("tours_over_traffic"),
            "conversions": box.get("conversions_apps_over_tours"),
            "period": box.get("period"),
        },
        "bottom_performers": bottom,
        "focus_areas": focus,
        "noi_distribution": fin.get("noi_distribution"),
        "net_income_distribution": fin.get("net_income_distribution"),
        "takeaways": take.get("items", []),
        "errors": errors,
        "status": "complete" if not errors else "complete_with_warnings",
    }

    OUT_FILE.parent.mkdir(parents=True, exist_ok=True)
    OUT_FILE.write_text(json.dumps(out, indent=2, default=str), encoding="utf-8")
    print(f"Wrote {OUT_FILE}")
    if errors:
        print("WARNINGS:")
        for e in errors:
            print("  -", e)
    return out

if __name__ == "__main__":
    main()
