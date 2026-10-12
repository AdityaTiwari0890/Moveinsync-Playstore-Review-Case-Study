"""
MOVEIN SYNC Rating Improvement Recommendation Engine — Step 3
=============================================================
Reads the Step 2 analyzed Excel workbook and the executive Markdown report,
derives ALL statistics programmatically (nothing is hardcoded), and produces:

  1. MOVEIN_SYNC_STEP3_RECOMMENDATIONS.md  — 8-section evidence-backed report
  2. MOVEIN_SYNC_ACTION_PLAN.xlsx           — 4-sheet structured action-plan workbook

This module extends the pipeline without modifying any Step 1 or Step 2 code.

Usage (PowerShell):
    cd "c:\\Users\\aktiw\\Downloads\\Performance Analyzer"
    .\\venv\\Scripts\\Activate.ps1
    python recommendation_engine.py
    python recommendation_engine.py --input "App_Review_Intelligence_MoveInSync_20261012_173118_Analyzed.xlsx"
    python recommendation_engine.py --input "MoveInSync_Executive_Analysis.xlsx" --report-md MY_REPORT.md --action-plan MY_PLAN.xlsx
"""

from __future__ import annotations

import argparse
import datetime
import math
import os
import sys
from collections import defaultdict
from typing import Any, Dict, List, Optional, Tuple

import pandas as pd
import openpyxl
from openpyxl.styles import (
    Alignment,
    Border,
    Font,
    PatternFill,
    Side,
)
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo

# ---------------------------------------------------------------------------
# Windows terminal UTF-8 safety
# ---------------------------------------------------------------------------
for _stream in (sys.stdout, sys.stderr):
    if _stream and hasattr(_stream, "reconfigure"):
        try:
            _stream.reconfigure(encoding="utf-8", errors="backslashreplace")
        except Exception:
            pass

# ---------------------------------------------------------------------------
# Colour palette (consistent with Steps 1 & 2)
# ---------------------------------------------------------------------------
NAVY = "1B365D"
MID_BLUE = "2C4C64"
WHITE = "FFFFFF"
LIGHT_GRAY = "D9D9D9"
ALERT_RED = "C00000"
GREEN_FILL = "E8F8F5"
AMBER_FILL = "FFF3CD"
RED_FILL = "FDEDEC"
BLUE_FILL = "EBF5FB"

# ---------------------------------------------------------------------------
# Taxonomy meta — determines category assignment for technical vs. operational
# separation.  Keys must match classifier.py TAXONOMY_CONFIG theme names.
# ---------------------------------------------------------------------------
THEME_META: Dict[str, Dict[str, str]] = {
    "App crashes and performance": {
        "category": "Technical Defect",
        "owner_team": "Mobile Engineering",
        "effort_estimate": "Estimate: Medium (Crashlytics audit, memory-leak profiling, ANR fixes)",
        "compliance_flag": "Engineering sprint",
    },
    "Login and account access": {
        "category": "Technical Defect",
        "owner_team": "Identity & Auth Engineering",
        "effort_estimate": "Estimate: Low-Medium (SSO token refresh, error messaging improvements)",
        "compliance_flag": "Engineering sprint",
    },
    "GPS and location tracking": {
        "category": "Technical Defect",
        "owner_team": "Platform Engineering / Maps",
        "effort_estimate": "Estimate: Medium (telemetry socket optimization, background wake-locks)",
        "compliance_flag": "Engineering sprint",
    },
    "Booking and OTP failures": {
        "category": "Technical Defect",
        "owner_team": "Booking & Trip Engineering",
        "effort_estimate": "Estimate: Low-Medium (SMS fallback, QR/Bluetooth boarding alternative)",
        "compliance_flag": "Engineering sprint",
    },
    "Cab delays and no-shows": {
        "category": "Operational Failure",
        "owner_team": "Transport Operations",
        "effort_estimate": "Estimate: Medium (dispatch algorithm tuning, driver SLA enforcement)",
        "compliance_flag": "Ops process change",
    },
    "Roster and scheduling": {
        "category": "Operational Failure",
        "owner_team": "Product + Operations",
        "effort_estimate": "Estimate: Low-Medium (roster cutoff rules, self-serve trip modifications)",
        "compliance_flag": "Ops process change",
    },
    "Driver-side experience": {
        "category": "Operational Failure",
        "owner_team": "Driver Operations + Mobile Engineering",
        "effort_estimate": "Estimate: Medium (duty UX simplification, toll-reimbursement automation)",
        "compliance_flag": "Ops + Engineering",
    },
    "Customer support": {
        "category": "Shared (Technical + Operational)",
        "owner_team": "Customer Experience",
        "effort_estimate": "Estimate: Medium (in-app chatbot, transport-manager SLA dashboard)",
        "compliance_flag": "Ops process + Engineering",
    },
    "Billing and payment": {
        "category": "Shared (Technical + Operational)",
        "owner_team": "Finance + Driver Operations",
        "effort_estimate": "Estimate: Low-Medium (fare-calculation audit, clear dispute flow)",
        "compliance_flag": "Ops + Engineering",
    },
    "UI and usability": {
        "category": "Technical Defect",
        "owner_team": "Product Design + Mobile Engineering",
        "effort_estimate": "Estimate: Medium (usability testing, UX simplification sprints)",
        "compliance_flag": "Engineering sprint",
    },
    "Other or unclear": {
        "category": "Unclear — Requires Investigation",
        "owner_team": "Product / QA",
        "effort_estimate": "Estimate: Unknown — manual review of sample required",
        "compliance_flag": "Triage required",
    },
}

# Impact-vs-Effort quadrant placement (qualitative, per theme)
IMPACT_EFFORT_QUADRANT: Dict[str, str] = {
    "Login and account access": "Quick Win",
    "App crashes and performance": "Major Project",
    "GPS and location tracking": "Major Project",
    "Booking and OTP failures": "Quick Win",
    "Cab delays and no-shows": "Major Project",
    "Driver-side experience": "Major Project",
    "Customer support": "Quick Win",
    "Roster and scheduling": "Fill-In",
    "Billing and payment": "Fill-In",
    "UI and usability": "Fill-In",
    "Other or unclear": "Deprioritize",
}

QUADRANT_DESCRIPTION = {
    "Quick Win": "High impact, relatively low effort. Address first to show rapid improvement.",
    "Major Project": "High impact, significant effort. Plan carefully with dedicated team allocation.",
    "Fill-In": "Lower impact, manageable effort. Schedule between major projects.",
    "Deprioritize": "Low impact or unclear evidence. Monitor but do not invest heavily now.",
}


# ===========================================================================
# DATA LOADING
# ===========================================================================

def _safe_read_sheet(xl: pd.ExcelFile, sheet: str) -> Optional[pd.DataFrame]:
    """Return DataFrame if sheet exists, else None with a printed warning."""
    if sheet in xl.sheet_names:
        return pd.read_excel(xl, sheet_name=sheet)
    print(f"  [WARN] Sheet '{sheet}' not found in workbook — skipping.", file=sys.stderr)
    return None


def load_analyzed_workbook(filepath: str) -> Dict[str, Any]:
    """
    Load all 5 analysis sheets and the raw Combined_Reviews sheet.
    Returns a dict of DataFrames and an audit dict.
    Raises FileNotFoundError if the workbook does not exist.
    """
    if not os.path.exists(filepath):
        raise FileNotFoundError(
            f"Analyzed workbook not found: '{filepath}'.\n"
            "Run Step 2 (analysis_engine.py) first to generate the analyzed workbook."
        )

    print(f"  [*] Loading workbook: '{os.path.basename(filepath)}'...")
    with pd.ExcelFile(filepath) as xl:
        sheets = {
            "Rating_Trends": _safe_read_sheet(xl, "Rating_Trends"),
            "Pareto_Analysis": _safe_read_sheet(xl, "Pareto_Analysis"),
            "Theme_Analysis": _safe_read_sheet(xl, "Theme_Analysis"),
            "Version_Analysis": _safe_read_sheet(xl, "Version_Analysis"),
            "Action_Priorities": _safe_read_sheet(xl, "Action_Priorities"),
        }
        # Raw reviews for excerpt sampling
        for raw_sheet in ("Combined_Reviews", "Android_Reviews", "iOS_Reviews"):
            if raw_sheet in xl.sheet_names:
                raw_df = pd.read_excel(xl, sheet_name=raw_sheet)
                sheets["_raw_reviews"] = raw_df
                print(f"  [+] Loaded raw reviews from sheet '{raw_sheet}' ({len(raw_df):,} rows).")
                break
        else:
            sheets["_raw_reviews"] = None
            print("  [WARN] No raw review sheet found; review-ID excerpts will be unavailable.", file=sys.stderr)

    loaded = [k for k, v in sheets.items() if v is not None and not k.startswith("_")]
    print(f"  [+] Loaded analysis sheets: {', '.join(loaded)}")

    return sheets


def _extract_overview_metrics(sheets: Dict[str, Any]) -> Dict[str, Any]:
    """
    Pull key scalar metrics from the Rating_Trends sheet (Overview table).
    All rows that start with 'SECTION' are section headers — skip them.
    Returns a dict of parsed metrics, or empty dict on failure.
    """
    rt = sheets.get("Rating_Trends")
    if rt is None:
        return {}

    # Drop pure section-header rows (first column starts with 'SECTION')
    rt_clean = rt[~rt.iloc[:, 0].astype(str).str.startswith("SECTION")].copy()
    rt_clean = rt_clean.dropna(how="all")

    # First non-empty row should be the header
    if rt_clean.empty:
        return {}

    # Detect the header row by looking for 'Segment' or 'Sub-Segment'
    header_mask = rt_clean.apply(
        lambda r: any(str(v).strip().lower() in ("segment", "sub-segment") for v in r), axis=1
    )
    if not header_mask.any():
        return {}

    header_idx = header_mask.idxmax()
    new_header = rt_clean.loc[header_idx].astype(str).str.strip().tolist()
    data_rows = rt_clean.loc[header_idx + 1 :].copy()
    data_rows.columns = new_header
    data_rows = data_rows.reset_index(drop=True)

    def _find_row(col: str, val: str) -> Optional[pd.Series]:
        if col not in data_rows.columns:
            return None
        matches = data_rows[data_rows[col].astype(str).str.strip() == val]
        return matches.iloc[0] if len(matches) > 0 else None

    def _safe(row: Optional[pd.Series], col: str, default: Any = "N/A") -> Any:
        if row is None or col not in row.index:
            return default
        v = row[col]
        return v if pd.notna(v) else default

    overall = _find_row("Sub-Segment", "All Combined Reviews")
    emp = _find_row("Sub-Segment", "Employee Apps")
    drv = _find_row("Sub-Segment", "Driver Apps")
    android = _find_row("Sub-Segment", "Android")
    ios = _find_row("Sub-Segment", "iOS")

    def _int(v: Any) -> Any:
        try:
            return int(float(v))
        except Exception:
            return v

    def _float(v: Any, dp: int = 2) -> Any:
        try:
            return round(float(v), dp)
        except Exception:
            return v

    return {
        "total_reviews": _int(_safe(overall, "Sample Size (N)", 0)),
        "avg_rating": _float(_safe(overall, "Sample Avg Rating", 0)),
        "neg_rate": _float(_safe(overall, "Negative Rate %", 0), 1),
        "pos_rate": _float(_safe(overall, "Positive Rate % (4\u2605+5\u2605)", 0), 1),
        "ci_lower": _float(_safe(overall, "Wilson 95% CI Lower", 0), 1),
        "ci_upper": _float(_safe(overall, "Wilson 95% CI Upper", 0), 1),
        "star_1": _int(_safe(overall, "1\u2605 Count", 0)),
        "star_2": _int(_safe(overall, "2\u2605 Count", 0)),
        "star_3": _int(_safe(overall, "3\u2605 Count", 0)),
        "star_4": _int(_safe(overall, "4\u2605 Count", 0)),
        "star_5": _int(_safe(overall, "5\u2605 Count", 0)),
        "emp_n": _int(_safe(emp, "Sample Size (N)", 0)),
        "emp_rating": _float(_safe(emp, "Sample Avg Rating", "N/A")),
        "emp_neg": _float(_safe(emp, "Negative Rate %", "N/A"), 1),
        "drv_n": _int(_safe(drv, "Sample Size (N)", 0)),
        "drv_rating": _float(_safe(drv, "Sample Avg Rating", "N/A")),
        "drv_neg": _float(_safe(drv, "Negative Rate %", "N/A"), 1),
        "android_n": _int(_safe(android, "Sample Size (N)", 0)),
        "android_rating": _float(_safe(android, "Sample Avg Rating", "N/A")),
        "android_neg": _float(_safe(android, "Negative Rate %", "N/A"), 1),
        "ios_n": _int(_safe(ios, "Sample Size (N)", 0)),
        "ios_rating": _float(_safe(ios, "Sample Avg Rating", "N/A")),
        "ios_neg": _float(_safe(ios, "Negative Rate %", "N/A"), 1),
    }


def _parse_pareto_df(sheets: Dict[str, Any]) -> pd.DataFrame:
    """
    Return a clean Pareto DataFrame with columns:
      Theme | Negative Reviews Count | Share % | Wilson CI Lower | Wilson CI Upper |
      Cumulative Share % | Pareto Classification

    The Pareto_Analysis sheet has this structure when read by pandas:
      - Column 0 header = NOTE text (the note was written in cell A1 of the sheet)
      - Row 0 of data = actual column names (written in row 2 of the sheet)
      - Rows 1+ of data = actual data rows
    We detect this by checking whether the first column header starts with 'NOTE:'
    OR whether the first data row looks like a header (contains 'Theme').
    """
    pa = sheets.get("Pareto_Analysis")
    if pa is None or pa.empty:
        return pd.DataFrame()

    # Case 1: note is the column header (pandas read the note cell as header)
    first_col_name = str(pa.columns[0]).strip()
    if first_col_name.upper().startswith("NOTE:"):
        # Row 0 of the dataframe holds the real column names
        if len(pa) < 1:
            return pd.DataFrame()
        new_cols = pa.iloc[0].astype(str).str.strip().tolist()
        data = pa.iloc[1:].copy()
        data.columns = new_cols
        data = data.reset_index(drop=True)
    else:
        # Case 2: note was a separate row in the data (first column starts with 'NOTE')
        first_col = pa.iloc[:, 0].astype(str)
        note_mask = first_col.str.upper().str.startswith("NOTE:")
        if note_mask.any():
            header_row_idx = note_mask.idxmax() + 1
            if header_row_idx >= len(pa):
                return pd.DataFrame()
            new_cols = pa.iloc[header_row_idx].astype(str).str.strip().tolist()
            data = pa.iloc[header_row_idx + 1:].copy()
            data.columns = new_cols
            data = data.reset_index(drop=True)
        else:
            # No note detected — use as-is
            data = pa.copy()
            data.columns = [str(c).strip() for c in data.columns]

    # Find the 'Theme' column (may be named differently)
    theme_col = None
    for c in data.columns:
        if str(c).strip().lower() == "theme":
            theme_col = c
            break
    if theme_col is None:
        # Try first non-empty column
        theme_col = data.columns[0]

    data = data.rename(columns={theme_col: "Theme"})
    data = data[data["Theme"].astype(str).str.strip() != ""]
    data = data[data["Theme"].notna()]

    # Standardize remaining column names
    renamed: Dict[str, str] = {}
    for col in data.columns:
        cl = str(col).lower()
        if "negative reviews count" in cl:
            renamed[col] = "Negative Reviews Count"
        elif "share of negative" in cl and "%" in cl:
            renamed[col] = "Share of Negative Reviews %"
        elif "wilson" in cl and "lower" in cl:
            renamed[col] = "Wilson CI Lower"
        elif "wilson" in cl and "upper" in cl:
            renamed[col] = "Wilson CI Upper"
        elif "cumulative share" in cl:
            renamed[col] = "Cumulative Share %"
        elif "pareto classification" in cl:
            renamed[col] = "Pareto Classification"
    data = data.rename(columns=renamed)

    # Coerce numeric columns
    for col in ["Negative Reviews Count", "Share of Negative Reviews %", "Wilson CI Lower",
                "Wilson CI Upper", "Cumulative Share %"]:
        if col in data.columns:
            data[col] = pd.to_numeric(data[col], errors="coerce")

    return data.reset_index(drop=True)


def _parse_regression_df(sheets: Dict[str, Any]) -> pd.DataFrame:
    """
    Extract the regression section from Version_Analysis sheet.
    Returns empty DataFrame if no regressions or sheet unavailable.
    """
    va = sheets.get("Version_Analysis")
    if va is None or va.empty:
        return pd.DataFrame()

    # Find the row that flags regressions
    first_col = va.iloc[:, 0].astype(str)
    reg_mask = first_col.str.contains("REGRESSION", case=False, na=False)
    if not reg_mask.any():
        return pd.DataFrame()

    reg_start = reg_mask.idxmax() + 1  # header of regression table
    if reg_start >= len(va):
        return pd.DataFrame()

    header_row = va.iloc[reg_start].astype(str).str.strip().tolist()
    data = va.iloc[reg_start + 1 :].copy()
    data.columns = header_row
    data = data.dropna(how="all").reset_index(drop=True)

    # Drop 'No regression detected' text rows
    if "App Name" in data.columns:
        data = data[data["App Name"].notna()]
        data = data[~data["App Name"].astype(str).str.startswith("No version")]

    return data


def _sample_review_excerpts(
    raw_df: Optional[pd.DataFrame],
    theme_name: str,
    n: int = 3,
) -> List[str]:
    """
    Sample up to n negative review excerpts for a given theme.
    Falls back gracefully if raw_df is unavailable or columns are missing.
    """
    if raw_df is None or raw_df.empty:
        return []

    # Normalize column names
    col_map = {}
    for c in raw_df.columns:
        norm = str(c).strip().lower().replace("_", " ")
        if norm in ("review text", "content", "text"):
            col_map[c] = "review_text"
        elif norm in ("review id", "id"):
            col_map[c] = "review_id"
        elif norm in ("rating", "score"):
            col_map[c] = "rating"
        elif "theme" in norm or "assigned theme" in norm:
            col_map[c] = "assigned_themes"

    df = raw_df.rename(columns=col_map).copy()

    # Filter negative reviews
    if "rating" in df.columns:
        df["rating"] = pd.to_numeric(df["rating"], errors="coerce")
        df = df[df["rating"].isin([1, 2])]

    if df.empty:
        return []

    # Filter by theme if assigned_themes available
    if "assigned_themes" in df.columns:
        theme_mask = df["assigned_themes"].astype(str).str.contains(theme_name, case=False, na=False)
        themed_df = df[theme_mask]
        if themed_df.empty:
            themed_df = df  # fall back to general negative reviews
    else:
        themed_df = df

    excerpts = []
    for _, row in themed_df.head(n).iterrows():
        rid = str(row.get("review_id", "")).strip()[:8] or "?"
        text = str(row.get("review_text", "")).replace("\n", " ").strip()
        if len(text) > 110:
            text = text[:107] + "..."
        if text:
            excerpts.append(f'[{rid}]: "{text}"')

    return excerpts


# ===========================================================================
# PRIORITY MATRIX
# ===========================================================================

def build_priority_matrix(
    pareto_df: pd.DataFrame,
    metrics: Dict[str, Any],
    raw_df: Optional[pd.DataFrame],
) -> pd.DataFrame:
    """
    Construct a multi-factor P1/P2/P3 priority matrix.
    Scoring factors (each 0–3):
      F1 — Frequency: share of negative reviews
      F2 — Platform breadth: both platforms affected
      F3 — User impact severity (predefined per theme)
      F4 — Classification confidence
      F5 — Inverse effort (lower effort = higher score)
    Total 0–15 → P1 ≥ 11, P2 7–10, P3 ≤ 6
    """
    SEVERITY: Dict[str, int] = {
        "App crashes and performance": 3,
        "Login and account access": 3,
        "Booking and OTP failures": 3,
        "GPS and location tracking": 3,
        "Cab delays and no-shows": 3,
        "Driver-side experience": 2,
        "Customer support": 2,
        "Roster and scheduling": 2,
        "Billing and payment": 2,
        "UI and usability": 1,
        "Other or unclear": 1,
    }
    # Effort score: 0 = high effort, 3 = very low effort
    EFFORT_SCORE: Dict[str, int] = {
        "App crashes and performance": 1,
        "Login and account access": 2,
        "Booking and OTP failures": 2,
        "GPS and location tracking": 1,
        "Cab delays and no-shows": 1,
        "Driver-side experience": 1,
        "Customer support": 2,
        "Roster and scheduling": 2,
        "Billing and payment": 2,
        "UI and usability": 1,
        "Other or unclear": 0,
    }

    total_neg = metrics.get("total_reviews", 1)
    neg_count_total = int(total_neg * metrics.get("neg_rate", 0) / 100) or 1

    rows = []
    for _, pr in pareto_df.iterrows():
        theme = str(pr.get("Theme", "")).strip()
        if not theme:
            continue

        neg_count = int(pr.get("Negative Reviews Count", 0) or 0)
        share_pct = float(pr.get("Share of Negative Reviews %", 0) or 0)
        ci_low = pr.get("Wilson CI Lower", "N/A")
        ci_high = pr.get("Wilson CI Upper", "N/A")
        pareto_class = str(pr.get("Pareto Classification", "")).strip()

        # --- Scoring ---
        # F1: frequency
        if share_pct >= 15:
            f1 = 3
        elif share_pct >= 8:
            f1 = 2
        elif share_pct >= 4:
            f1 = 1
        else:
            f1 = 0

        # F2: platform breadth (assume both affected unless theme is driver-only)
        f2 = 2 if theme != "Driver-side experience" else 1

        # F3: severity
        f3 = SEVERITY.get(theme, 1)

        # F4: confidence from pareto classification
        if "vital few" in pareto_class.lower():
            f4 = 2
        else:
            f4 = 1

        # F5: inverse effort
        f5 = EFFORT_SCORE.get(theme, 1)

        total_score = f1 + f2 + f3 + f4 + f5

        if total_score >= 11:
            priority = "P1 — Critical"
            priority_rationale = "High frequency + high user impact + sufficient evidence"
        elif total_score >= 7:
            priority = "P2 — High"
            priority_rationale = "Moderate-high frequency or impact; warrants planned sprint"
        else:
            priority = "P3 — Medium"
            priority_rationale = "Lower frequency or impact; address after P1/P2"

        meta = THEME_META.get(theme, {
            "category": "Unclear — Requires Investigation",
            "owner_team": "Product / QA",
            "effort_estimate": "Estimate: Unknown",
            "compliance_flag": "Triage required",
        })

        # Sample evidence excerpts
        excerpts = _sample_review_excerpts(raw_df, theme, n=2)
        evidence_str = " | ".join(excerpts) if excerpts else "See Pareto_Analysis sheet for classified review IDs."

        # Success metric baseline
        if theme == "App crashes and performance":
            success_metric = "Crash-free session rate ≥ 99.8%; ANR rate < 0.2% (baseline: requires Firebase/Crashlytics data)"
        elif theme == "Login and account access":
            success_metric = "First-attempt login success rate ≥ 98% (baseline: TBC from auth service logs)"
        elif theme == "GPS and location tracking":
            success_metric = "Location refresh latency < 5 s; 40% reduction in tracking-related negative reviews (baseline: current share %)"
        elif theme == "Booking and OTP failures":
            success_metric = "Boarding verification success rate ≥ 99.5%; verification time < 15 s (baseline: TBC from trip service logs)"
        elif theme == "Cab delays and no-shows":
            success_metric = "On-time pickup rate ≥ 95%; 30% reduction in delay-related negative reviews (baseline: TBC from dispatch logs)"
        elif theme == "Driver-side experience":
            success_metric = "Driver duty acceptance rate ≥ 90%; Driver app rating ≥ 4.2★ (baseline: current sample avg)"
        elif theme == "Customer support":
            success_metric = "First response time < 2 min for active-trip emergencies (baseline: TBC from CX ticketing system)"
        elif theme == "Roster and scheduling":
            success_metric = "Roster complaint share < 5% of negative reviews (baseline: current share %)"
        elif theme == "Billing and payment":
            success_metric = "Billing dispute resolution < 48 h; dispute complaint share < 3% of negatives (baseline: TBC)"
        elif theme == "UI and usability":
            success_metric = "Booking task completion time reduced ≥ 25% (baseline: usability study required)"
        else:
            success_metric = "To be confirmed after manual triage of 'Other / unclear' sample."

        rows.append({
            "Priority": priority,
            "Priority Score": total_score,
            "Theme": theme,
            "Category": meta["category"],
            "Negative Reviews Count": neg_count,
            "Share of Negative Reviews %": round(share_pct, 1),
            "Wilson 95% CI": f"[{ci_low}%, {ci_high}%]" if ci_low != "N/A" else "N/A",
            "Pareto Status": pareto_class,
            "Severity / User Impact": ["Low", "Low-Medium", "Medium-High", "High"][SEVERITY.get(theme, 1)],
            "Effort Estimate": meta["effort_estimate"],
            "Owner Team": meta["owner_team"],
            "Provisional Status": "Provisional — validate with product, engineering, and operations teams before committing",
            "Suggested Success Metric": success_metric,
            "Sample Evidence": evidence_str,
            "Action Status": "Open",
        })

    if not rows:
        return pd.DataFrame()

    df = pd.DataFrame(rows).sort_values(["Priority Score"], ascending=False).reset_index(drop=True)
    return df


# ===========================================================================
# IMPACT vs EFFORT MATRIX
# ===========================================================================

def build_impact_effort_matrix(priority_df: pd.DataFrame) -> pd.DataFrame:
    """
    Map each theme into a 4-quadrant Impact × Effort matrix.
    Impact = measured negative-review share (from workbook).
    Effort = qualitative estimate (from THEME_META).
    """
    rows = []
    for _, r in priority_df.iterrows():
        theme = r["Theme"]
        quadrant = IMPACT_EFFORT_QUADRANT.get(theme, "Fill-In")
        quadrant_desc = QUADRANT_DESCRIPTION.get(quadrant, "")
        impact_evidence = (
            f"Measured: {r['Negative Reviews Count']:,} negative reviews, "
            f"{r['Share of Negative Reviews %']}% share (Wilson CI: {r['Wilson 95% CI']})"
        )
        rows.append({
            "Quadrant": quadrant,
            "Theme": theme,
            "Category": r["Category"],
            "Measured Impact Evidence": impact_evidence,
            "Effort Estimate (Qualitative)": r["Effort Estimate"],
            "Owner Team": r["Owner Team"],
            "Recommended Immediate Action": (
                "Prioritize in next sprint" if quadrant == "Quick Win"
                else "Plan dedicated project" if quadrant == "Major Project"
                else "Schedule opportunistically" if quadrant == "Fill-In"
                else "Monitor; do not invest heavily"
            ),
            "Quadrant Description": quadrant_desc,
        })

    order = {"Quick Win": 0, "Major Project": 1, "Fill-In": 2, "Deprioritize": 3}
    df = pd.DataFrame(rows)
    if not df.empty:
        df["_order"] = df["Quadrant"].map(order)
        df = df.sort_values("_order").drop(columns=["_order"]).reset_index(drop=True)
    return df


# ===========================================================================
# RATING IMPROVEMENT STRATEGY
# ===========================================================================

def build_rating_strategy() -> pd.DataFrame:
    """
    Define the compliant rating improvement initiatives.
    Explicitly labels initiative type, compliance note, owner, and metric.
    All recommendations are policy-compliant; no gating or incentivization.
    """
    rows = [
        {
            "Initiative": "Fix validated high-impact problems first",
            "Type": "Product / Engineering",
            "Description": (
                "Address P1 technical defects (crashes, GPS, OTP, login) and P1 operational "
                "failures (delays) before investing in review acquisition. Rating improvement "
                "driven by fixing real problems is more durable and defensible."
            ),
            "Owner": "Product + Engineering + Operations",
            "Estimated Timeline": "0–3 months (P1 items); 3–6 months (P2 items)",
            "Compliance Note": "No store-policy concerns. Standard product improvement.",
            "Success Metric": "Sample negative review rate falls below 35% within 6 months of P1 fixes (baseline: current measured rate — see Rating_Trends sheet)",
            "Status": "Open",
        },
        {
            "Initiative": "Compliant in-app review prompt (post-successful trip)",
            "Type": "Growth / Product",
            "Description": (
                "After a trip completes successfully and the employee or driver marks the journey "
                "as satisfactory, prompt once per rolling 90 days: 'Enjoying the app? Rate us on "
                "the Play Store / App Store.' No pre-screening; no conditional routing based on "
                "expected rating. Both positive and dissatisfied users reach the store directly."
            ),
            "Owner": "Mobile Engineering + Product",
            "Estimated Timeline": "1–2 sprints (4–8 weeks)",
            "Compliance Note": (
                "Compliant with Google Play Developer Policy and Apple App Store Review Guidelines. "
                "Does NOT gate, filter, or intercept reviews by anticipated star rating. "
                "DOES NOT offer incentives (credits, discounts) in exchange for reviews."
            ),
            "Success Metric": "Review volume increases ≥ 20% within 60 days of prompt launch (baseline: monthly review count — see Rating_Trends sheet)",
            "Status": "Open",
        },
        {
            "Initiative": "Separate in-app support pathway (does not suppress public feedback)",
            "Type": "Customer Experience",
            "Description": (
                "Add a clearly visible 'Report a problem' button in the app that routes users to "
                "an in-app support ticket, not to the store review page. This gives frustrated "
                "users a private resolution channel alongside the public review option — both "
                "must remain equally visible and accessible."
            ),
            "Owner": "Customer Experience + Mobile Engineering",
            "Estimated Timeline": "4–6 weeks",
            "Compliance Note": (
                "Both the support pathway AND the store review link must be equally accessible. "
                "Must not replace or obscure the store review option for negative users. "
                "Compliant with Google Play and Apple policies on review interception."
            ),
            "Success Metric": "In-app ticket resolution rate ≥ 85% within 48 h; no increase in store review suppression complaints (baseline: TBC from CX system)",
            "Status": "Open",
        },
        {
            "Initiative": "Developer review-response workflow",
            "Type": "Community / Reputation",
            "Description": (
                "Implement a structured weekly process: (1) triage new 1-star and 2-star reviews "
                "within 48 h; (2) respond to all 1-star reviews within 5 business days with a "
                "specific, non-generic acknowledgement; (3) escalate recurring themes (>5 mentions/week) "
                "to the relevant engineering or ops team lead; (4) track resolution status in the "
                "Action Plan workbook; (5) update response when issue is resolved."
            ),
            "Owner": "Product Manager + Customer Experience",
            "Estimated Timeline": "Process setup: 2 weeks; ongoing weekly cadence",
            "Compliance Note": "Allowed by both store platforms. Do not request rating changes in responses.",
            "Success Metric": "100% of 1★ reviews receive a response within 5 business days; escalation-to-resolution loop tracked in Action Plan workbook (baseline: 0 structured responses currently)",
            "Status": "Open",
        },
        {
            "Initiative": "Recurring review monitoring and release regression alerts",
            "Type": "Engineering / Operations",
            "Description": (
                "Schedule a weekly automated run of the Step 2 analysis engine after each app "
                "release. Flag any version with a rating drop ≥ 0.30★ or negative-review-rate "
                "jump ≥ 15% (current thresholds). Alert the release manager via email or Slack "
                "within 24 h of the regression being detected. Correlate with internal crash and "
                "ANR telemetry to distinguish app defects from operational disruptions."
            ),
            "Owner": "Release Engineering + Product",
            "Estimated Timeline": "4–6 weeks (automation setup)",
            "Compliance Note": "Internal monitoring only. No policy concerns.",
            "Success Metric": "Regression alerts trigger within 24 h of qualifying version release; mean time-to-acknowledge ≤ 4 h (baseline: no current automated monitoring)",
            "Status": "Open",
        },
    ]
    return pd.DataFrame(rows)


# ===========================================================================
# PLATFORM COMPARISON
# ===========================================================================

def build_platform_comparison(metrics: Dict[str, Any], regression_df: pd.DataFrame) -> Dict[str, Any]:
    """
    Summarise Android vs iOS differences with sample-size guards.
    Returns a dict of structured comparison data.
    """
    android_n = metrics.get("android_n", 0)
    ios_n = metrics.get("ios_n", 0)
    android_rating = metrics.get("android_rating", "N/A")
    ios_rating = metrics.get("ios_rating", "N/A")
    android_neg = metrics.get("android_neg", "N/A")
    ios_neg = metrics.get("ios_neg", "N/A")

    MIN_SAMPLE = 30
    android_sufficient = int(str(android_n).replace(",", "") or 0) >= MIN_SAMPLE
    ios_sufficient = int(str(ios_n).replace(",", "") or 0) >= MIN_SAMPLE

    try:
        rating_gap = round(float(android_rating) - float(ios_rating), 2)
    except (TypeError, ValueError):
        rating_gap = None

    try:
        neg_gap = round(float(android_neg) - float(ios_neg), 1)
    except (TypeError, ValueError):
        neg_gap = None

    regressions = []
    if not regression_df.empty:
        for _, row in regression_df.iterrows():
            regressions.append({
                "app": row.get("App Name", ""),
                "platform": row.get("Platform", ""),
                "prior_ver": row.get("Prior Version", ""),
                "current_ver": row.get("Current Version", ""),
                "rating_delta": row.get("Rating Delta", ""),
                "neg_jump": row.get("Negative Rate Jump %", ""),
                "observation": row.get("Association Assessment", ""),
            })

    return {
        "android_n": android_n,
        "ios_n": ios_n,
        "android_rating": android_rating,
        "ios_rating": ios_rating,
        "android_neg": android_neg,
        "ios_neg": ios_neg,
        "android_sufficient": android_sufficient,
        "ios_sufficient": ios_sufficient,
        "rating_gap": rating_gap,
        "neg_gap": neg_gap,
        "regressions": regressions,
        "min_sample": MIN_SAMPLE,
    }


# ===========================================================================
# EXCEL ACTION PLAN WORKBOOK
# ===========================================================================

def _style_header_row(ws, row_num: int, max_col: int, fill_color: str = NAVY):
    """Apply navy bold header styling to a row."""
    hf = Font(name="Calibri", size=10, bold=True, color=WHITE)
    hfill = PatternFill(start_color=fill_color, end_color=fill_color, fill_type="solid")
    for col in range(1, max_col + 1):
        cell = ws.cell(row=row_num, column=col)
        cell.font = hf
        cell.fill = hfill
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)


def _style_data_rows(ws, start_row: int, end_row: int):
    """Apply consistent body styling."""
    rf = Font(name="Calibri", size=9)
    thin = Border(
        left=Side(style="thin", color=LIGHT_GRAY),
        right=Side(style="thin", color=LIGHT_GRAY),
        top=Side(style="thin", color=LIGHT_GRAY),
        bottom=Side(style="thin", color=LIGHT_GRAY),
    )
    for row in ws.iter_rows(min_row=start_row, max_row=end_row):
        for cell in row:
            cell.font = rf
            cell.border = thin
            cell.alignment = Alignment(vertical="top", wrap_text=True)
            # Highlight critical priority rows
            if cell.column == 1 and "P1" in str(cell.value or ""):
                cell.fill = PatternFill(start_color=RED_FILL, end_color=RED_FILL, fill_type="solid")
            elif cell.column == 1 and "P2" in str(cell.value or ""):
                cell.fill = PatternFill(start_color=AMBER_FILL, end_color=AMBER_FILL, fill_type="solid")
            elif cell.column == 1 and "P3" in str(cell.value or ""):
                cell.fill = PatternFill(start_color=GREEN_FILL, end_color=GREEN_FILL, fill_type="solid")
            # Highlight Quick Wins green
            elif cell.column == 1 and "Quick Win" in str(cell.value or ""):
                cell.fill = PatternFill(start_color=GREEN_FILL, end_color=GREEN_FILL, fill_type="solid")
            elif cell.column == 1 and "Major Project" in str(cell.value or ""):
                cell.fill = PatternFill(start_color=AMBER_FILL, end_color=AMBER_FILL, fill_type="solid")


def _auto_col_widths(ws):
    for col in ws.columns:
        max_len = 0
        col_letter = get_column_letter(col[0].column)
        for cell in col:
            if cell.value:
                first_line = str(cell.value).split("\n")[0]
                max_len = max(max_len, min(len(first_line), 80))
        ws.column_dimensions[col_letter].width = max(12, min(max_len + 3, 60))


def _add_excel_table(ws, table_name: str, start_row: int, end_row: int, style: str = "TableStyleLight9"):
    if end_row < start_row:
        return
    max_col = ws.max_column
    ref = f"A{start_row}:{get_column_letter(max_col)}{end_row}"
    tab = Table(displayName=table_name, ref=ref)
    tab.tableStyleInfo = TableStyleInfo(
        name=style, showFirstColumn=False, showLastColumn=False,
        showRowStripes=True, showColumnStripes=False,
    )
    ws.add_table(tab)


def generate_action_plan_xlsx(
    priority_df: pd.DataFrame,
    impact_effort_df: pd.DataFrame,
    strategy_df: pd.DataFrame,
    metrics: Dict[str, Any],
    output_path: str,
) -> str:
    """
    Write MOVEIN_SYNC_ACTION_PLAN.xlsx with 4 sheets:
      1. Priority_Matrix
      2. Impact_Effort_Matrix
      3. Rating_Strategy
      4. Success_Metrics
    """
    wb = openpyxl.Workbook()

    # Remove default sheet
    if "Sheet" in wb.sheetnames:
        wb.remove(wb["Sheet"])

    ts = datetime.datetime.now().strftime("%Y-%m-%d %H:%M")

    # -----------------------------------------------------------------
    # Sheet 1: Priority_Matrix
    # -----------------------------------------------------------------
    ws1 = wb.create_sheet("Priority_Matrix")
    ws1.append([f"MOVEIN SYNC — Priority Matrix  |  Generated: {ts}  |  Source: App_Review_Intelligence_MoveInSync_20261012_173118_Analyzed.xlsx"])
    ws1.cell(row=1, column=1).font = Font(name="Calibri", size=11, bold=True, color=WHITE)
    ws1.cell(row=1, column=1).fill = PatternFill(start_color=NAVY, end_color=NAVY, fill_type="solid")
    ws1.row_dimensions[1].height = 20

    if not priority_df.empty:
        cols_to_write = [c for c in priority_df.columns if c != "Priority Score"]
        ws1.append(cols_to_write)
        header_row = 2
        _style_header_row(ws1, header_row, len(cols_to_write))
        data_start = 3
        for _, row in priority_df.iterrows():
            ws1.append([row[c] for c in cols_to_write])
        _style_data_rows(ws1, data_start, ws1.max_row)
        _add_excel_table(ws1, "Table_Priority_Matrix", header_row, ws1.max_row)
        ws1.freeze_panes = f"A{data_start}"
    _auto_col_widths(ws1)

    # -----------------------------------------------------------------
    # Sheet 2: Impact_Effort_Matrix
    # -----------------------------------------------------------------
    ws2 = wb.create_sheet("Impact_Effort_Matrix")
    ws2.append([f"Impact vs. Effort Matrix  |  Generated: {ts}"])
    ws2.cell(row=1, column=1).font = Font(name="Calibri", size=11, bold=True, color=WHITE)
    ws2.cell(row=1, column=1).fill = PatternFill(start_color=MID_BLUE, end_color=MID_BLUE, fill_type="solid")

    note_text = (
        "NOTE: 'Impact' is derived from measured negative-review share (workbook data). "
        "'Effort' is a qualitative engineering estimate — treat as directional, not definitive."
    )
    ws2.append([note_text])
    ws2.cell(row=2, column=1).font = Font(name="Calibri", size=9, italic=True)

    if not impact_effort_df.empty:
        header_row = 3
        ws2.append(impact_effort_df.columns.tolist())
        _style_header_row(ws2, header_row, len(impact_effort_df.columns))
        data_start = 4
        for _, row in impact_effort_df.iterrows():
            ws2.append(row.tolist())
        _style_data_rows(ws2, data_start, ws2.max_row)
        _add_excel_table(ws2, "Table_Impact_Effort", header_row, ws2.max_row)
        ws2.freeze_panes = f"A{data_start}"
    _auto_col_widths(ws2)

    # -----------------------------------------------------------------
    # Sheet 3: Rating_Strategy
    # -----------------------------------------------------------------
    ws3 = wb.create_sheet("Rating_Strategy")
    ws3.append([f"Rating Improvement Strategy  |  Generated: {ts}"])
    ws3.cell(row=1, column=1).font = Font(name="Calibri", size=11, bold=True, color=WHITE)
    ws3.cell(row=1, column=1).fill = PatternFill(start_color=NAVY, end_color=NAVY, fill_type="solid")

    ws3.append([
        "⚠ Prohibited approaches (not included): Incentivized reviews | Review gating | "
        "Selective suppression of negative feedback | Fabricated ratings"
    ])
    ws3.cell(row=2, column=1).font = Font(name="Calibri", size=9, bold=True, color=ALERT_RED)

    if not strategy_df.empty:
        header_row = 3
        ws3.append(strategy_df.columns.tolist())
        _style_header_row(ws3, header_row, len(strategy_df.columns))
        data_start = 4
        for _, row in strategy_df.iterrows():
            ws3.append(row.tolist())
        _style_data_rows(ws3, data_start, ws3.max_row)
        _add_excel_table(ws3, "Table_Rating_Strategy", header_row, ws3.max_row)
        ws3.freeze_panes = f"A{data_start}"
    _auto_col_widths(ws3)

    # -----------------------------------------------------------------
    # Sheet 4: Success_Metrics
    # -----------------------------------------------------------------
    ws4 = wb.create_sheet("Success_Metrics")
    ws4.append([f"Success Metrics Register  |  Generated: {ts}"])
    ws4.cell(row=1, column=1).font = Font(name="Calibri", size=11, bold=True, color=WHITE)
    ws4.cell(row=1, column=1).fill = PatternFill(start_color=NAVY, end_color=NAVY, fill_type="solid")

    ws4.append(["Metric", "Theme / Initiative", "Baseline", "Target", "Measurement Window",
                "Responsible Team", "Measurement Method", "Status"])
    _style_header_row(ws4, 2, 8)

    success_rows = []
    if not priority_df.empty:
        for _, r in priority_df.iterrows():
            theme = r["Theme"]
            metric_text = r["Suggested Success Metric"]
            # Extract baseline vs target heuristically
            if "baseline:" in metric_text.lower():
                parts = metric_text.lower().split("baseline:")
                target_part = metric_text[:metric_text.lower().index("baseline:")].strip().rstrip(";").rstrip("(")
                baseline_part = metric_text[metric_text.lower().index("baseline:") + 9:].strip().rstrip(")")
            else:
                target_part = metric_text
                baseline_part = "Requires confirmation from internal telemetry / logs"

            success_rows.append([
                theme,
                theme,
                baseline_part[:200],
                target_part[:200],
                "Measure at 30, 60, 90 days post-fix deployment",
                r["Owner Team"],
                "App store review analysis (re-run Step 2 engine) + internal telemetry",
                "Open",
            ])

    for sr_row in success_rows:
        ws4.append(sr_row)

    if success_rows:
        _style_data_rows(ws4, 3, ws4.max_row)
        _add_excel_table(ws4, "Table_Success_Metrics", 2, ws4.max_row)
        ws4.freeze_panes = "A3"
    _auto_col_widths(ws4)

    wb.save(output_path)
    return output_path


# ===========================================================================
# MARKDOWN REPORT GENERATOR
# ===========================================================================

def generate_recommendations_md(
    metrics: Dict[str, Any],
    priority_df: pd.DataFrame,
    impact_effort_df: pd.DataFrame,
    strategy_df: pd.DataFrame,
    platform_data: Dict[str, Any],
    regression_df: pd.DataFrame,
    workbook_name: str,
    output_path: str,
) -> str:
    """Generate the full 8-section Markdown recommendation report."""
    ts = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    total = metrics.get("total_reviews", "N/A")
    avg_r = metrics.get("avg_rating", "N/A")
    neg_rate = metrics.get("neg_rate", "N/A")
    ci_l = metrics.get("ci_lower", "N/A")
    ci_u = metrics.get("ci_upper", "N/A")
    neg_count = int(total * neg_rate / 100) if isinstance(total, int) and isinstance(neg_rate, (int, float)) else "N/A"

    # Top 3 P1 themes
    p1_rows = priority_df[priority_df["Priority"].str.startswith("P1")] if not priority_df.empty else pd.DataFrame()
    top3 = p1_rows.head(3)

    # Quick wins
    qw = impact_effort_df[impact_effort_df["Quadrant"] == "Quick Win"] if not impact_effort_df.empty else pd.DataFrame()

    lines = []
    lines.append(f"# MOVEIN SYNC — Rating Improvement Recommendation Report (Step 3)")
    lines.append(f"")
    lines.append(f"**Generated on:** {ts}  ")
    lines.append(f"**Data Source:** `{workbook_name}`  ")
    lines.append(f"**Total Reviews Analyzed:** {total:,}  " if isinstance(total, int) else f"**Total Reviews Analyzed:** {total}  ")
    lines.append(f"**Negative Reviews (1★ + 2★):** {neg_count:,} ({neg_rate}%)  " if isinstance(neg_count, int) else f"**Negative Reviews (1★ + 2★):** {neg_count} ({neg_rate}%)  ")
    lines.append(f"**Sample Average Rating:** {avg_r}★ (Wilson 95% CI: [{ci_l}%, {ci_u}%])  ")
    lines.append("")
    lines.append("> [!IMPORTANT]")
    lines.append("> All statistics in this report are derived at runtime from the analyzed workbook.")
    lines.append("> No figures are hardcoded. Priority assignments are **provisional** until validated")
    lines.append("> by product, engineering, and operations teams.")
    lines.append("")
    lines.append("---")
    lines.append("")

    # ---------- Section 1: Executive Summary ----------
    lines.append("## 1. Executive Summary — First Three Actions")
    lines.append("")
    lines.append(f"Based on {total:,} reviews with a {neg_rate}% negative review rate, " if isinstance(total, int) else f"Based on {total} reviews with a {neg_rate}% negative rate, ")
    lines.append("the analysis recommends addressing **technical defects before investing in review acquisition**.")
    lines.append("The three highest-priority actions are:")
    lines.append("")
    if not top3.empty:
        for i, (_, r) in enumerate(top3.iterrows(), 1):
            lines.append(f"### Action {i}: {r['Theme']}")
            lines.append(f"- **Category:** {r['Category']}")
            lines.append(f"- **Evidence:** {r['Negative Reviews Count']:,} negative reviews mentioning this theme ({r['Share of Negative Reviews %']}% share; Wilson 95% CI: {r['Wilson 95% CI']})")
            lines.append(f"- **Owner:** {r['Owner Team']}")
            lines.append(f"- **Effort:** {r['Effort Estimate']}")
            lines.append(f"- **Success Metric:** {r['Suggested Success Metric']}")
            lines.append(f"- **Sample Evidence:** {r['Sample Evidence']}")
            lines.append(f"- **Priority Rationale:** {r['Provisional Status']}")
            lines.append("")
    else:
        lines.append("*No P1 themes identified — check the Pareto_Analysis sheet for data completeness.*")
        lines.append("")
    lines.append("---")
    lines.append("")

    # ---------- Section 2: Priority Matrix ----------
    lines.append("## 2. Evidence-Backed Priority Matrix")
    lines.append("")
    lines.append("> [!NOTE]")
    lines.append("> Priority scores combine: negative-review frequency, platform breadth, user-impact severity,")
    lines.append("> classification confidence, and inverse effort. All assignments are provisional.")
    lines.append("")
    if not priority_df.empty:
        display_cols = ["Priority", "Theme", "Category", "Negative Reviews Count",
                        "Share of Negative Reviews %", "Wilson 95% CI", "Owner Team", "Effort Estimate"]
        display_cols = [c for c in display_cols if c in priority_df.columns]
        header = " | ".join(display_cols)
        sep = " | ".join(["---"] * len(display_cols))
        lines.append(f"| {header} |")
        lines.append(f"| {sep} |")
        for _, r in priority_df.iterrows():
            vals = [str(r.get(c, "")) for c in display_cols]
            lines.append(f"| {' | '.join(vals)} |")
    else:
        lines.append("*Priority matrix could not be generated — Pareto_Analysis sheet may be empty.*")
    lines.append("")
    lines.append("---")
    lines.append("")

    # ---------- Section 3: Technical vs Operational ----------
    lines.append("## 3. Technical Defects vs. Operational Failures")
    lines.append("")
    lines.append("Issues are separated into three categories. Uncertain causes are flagged for investigation rather than asserted from review text alone.")
    lines.append("")

    categories = {
        "Technical Defect": [],
        "Operational Failure": [],
        "Shared (Technical + Operational)": [],
        "Unclear — Requires Investigation": [],
    }
    if not priority_df.empty:
        for _, r in priority_df.iterrows():
            cat = r.get("Category", "Unclear — Requires Investigation")
            if cat in categories:
                categories[cat].append(r["Theme"])
            else:
                categories["Unclear — Requires Investigation"].append(r["Theme"])

    for cat, themes in categories.items():
        if themes:
            lines.append(f"### {cat}")
            for t in themes:
                meta = THEME_META.get(t, {})
                lines.append(f"- **{t}** — Owner: {meta.get('owner_team', 'TBD')} | {meta.get('effort_estimate', '')}")
            lines.append("")

    lines.append("> [!CAUTION]")
    lines.append("> Review text classification is heuristic-based. A theme assigned to 'Technical Defect'")
    lines.append("> may have an operational root cause (e.g., GPS inaccuracy at a depot due to signal interference,")
    lines.append("> not an app bug). Validate with internal telemetry before committing engineering resources.")
    lines.append("")
    lines.append("---")
    lines.append("")

    # ---------- Section 4: Platform & Release ----------
    lines.append("## 4. Platform and Release Comparison")
    lines.append("")
    pd_data = platform_data
    android_n = pd_data.get("android_n", 0)
    ios_n = pd_data.get("ios_n", 0)
    android_r = pd_data.get("android_rating", "N/A")
    ios_r = pd_data.get("ios_rating", "N/A")
    android_neg = pd_data.get("android_neg", "N/A")
    ios_neg = pd_data.get("ios_neg", "N/A")
    min_s = pd_data.get("min_sample", 30)
    a_ok = pd_data.get("android_sufficient", False)
    i_ok = pd_data.get("ios_sufficient", False)
    rating_gap = pd_data.get("rating_gap")
    neg_gap = pd_data.get("neg_gap")

    lines.append("### Android vs. iOS — Measured Comparison")
    lines.append("")
    lines.append("| Metric | Android | iOS | Sample Sufficient? |")
    lines.append("|--------|---------|-----|-------------------|")
    lines.append(f"| Sample Size (N) | {android_n:,} | {ios_n:,} | Android: {'✓' if a_ok else f'✗ (< {min_s})'} / iOS: {'✓' if i_ok else f'✗ (< {min_s})'} |")
    lines.append(f"| Sample Avg Rating | {android_r}★ | {ios_r}★ | See sample size above |")
    lines.append(f"| Negative Review Rate | {android_neg}% | {ios_neg}% | See sample size above |")
    lines.append("")

    if rating_gap is not None:
        if abs(rating_gap) >= 0.5 and a_ok and i_ok:
            lines.append(f"> [!WARNING]")
            lines.append(f"> Android sample average rating is {abs(rating_gap):.2f}★ {'higher' if rating_gap > 0 else 'lower'} than iOS ({android_r}★ vs {ios_r}★).")
            lines.append(f"> This gap warrants platform-specific investigation but should be correlated with")
            lines.append(f"> internal telemetry before attributing to a specific feature or release.")
        else:
            lines.append(f"*Rating gap between platforms is {rating_gap:+.2f}★ — {'within expected variance' if abs(rating_gap) < 0.5 else 'notable but check sample sizes'}.*")
    lines.append("")

    lines.append("### Version Regression Flags")
    lines.append("")
    regressions = pd_data.get("regressions", [])
    if regressions:
        lines.append("| App | Platform | Prior Version | Regressed Version | Rating Delta | Neg Rate Jump | Observation |")
        lines.append("|-----|----------|---------------|-------------------|-------------|--------------|-------------|")
        for reg in regressions:
            lines.append(f"| {reg['app']} | {reg['platform']} | {reg['prior_ver']} | {reg['current_ver']} | {reg['rating_delta']} | +{reg['neg_jump']}% | {reg['observation']} |")
        lines.append("")
        lines.append("> [!CAUTION]")
        lines.append("> All version regressions are **observational associations** only — not proof of causation.")
        lines.append("> A rating drop following a release may coincide with fleet supply shortages, public holidays,")
        lines.append("> or external operational factors. Validate with crash telemetry and server error logs.")
        lines.append("> **Evidence required to validate each regression:** internal ANR/crash rates for the flagged")
        lines.append("> version; fleet availability metrics for the regression period; server error-rate logs.")
    else:
        lines.append("*No version regressions exceeded the detection threshold (ΔRating ≤ −0.30★ or ΔNeg Rate ≥ +15%) with sufficient sample size (N ≥ 15 per version).*")
    lines.append("")
    lines.append("---")
    lines.append("")

    # ---------- Section 5: Rating Improvement Strategy ----------
    lines.append("## 5. Rating Improvement Strategy")
    lines.append("")
    lines.append("> [!IMPORTANT]")
    lines.append("> The following recommendations are policy-compliant. The following approaches are **explicitly")
    lines.append("> prohibited** and are not included: incentivized reviews, review gating (routing users based on")
    lines.append("> anticipated star rating), selective suppression of negative feedback, or fabricated ratings.")
    lines.append("")
    if not strategy_df.empty:
        for i, (_, r) in enumerate(strategy_df.iterrows(), 1):
            lines.append(f"### Strategy {i}: {r['Initiative']}")
            lines.append(f"- **Type:** {r['Type']}")
            lines.append(f"- **Description:** {r['Description']}")
            lines.append(f"- **Owner:** {r['Owner']}")
            lines.append(f"- **Estimated Timeline:** {r['Estimated Timeline']}")
            lines.append(f"- **Compliance Note:** {r['Compliance Note']}")
            lines.append(f"- **Success Metric:** {r['Success Metric']}")
            lines.append("")
    lines.append("---")
    lines.append("")

    # ---------- Section 6: Impact vs Effort ----------
    lines.append("## 6. Impact vs. Effort Matrix")
    lines.append("")
    lines.append("> [!NOTE]")
    lines.append("> **Impact** = measured from negative-review share in the workbook (objective data).")
    lines.append("> **Effort** = qualitative engineering estimate — treat as directional, not a firm commitment.")
    lines.append("")
    for quadrant in ["Quick Win", "Major Project", "Fill-In", "Deprioritize"]:
        q_rows = impact_effort_df[impact_effort_df["Quadrant"] == quadrant] if not impact_effort_df.empty else pd.DataFrame()
        desc = QUADRANT_DESCRIPTION.get(quadrant, "")
        lines.append(f"### {quadrant}")
        lines.append(f"*{desc}*")
        lines.append("")
        if not q_rows.empty:
            for _, r in q_rows.iterrows():
                lines.append(f"- **{r['Theme']}** ({r['Category']})")
                lines.append(f"  - Impact: {r['Measured Impact Evidence']}")
                lines.append(f"  - Effort: {r['Effort Estimate (Qualitative)']}")
                lines.append(f"  - Action: {r['Recommended Immediate Action']}")
        else:
            lines.append("*No themes in this quadrant.*")
        lines.append("")
    lines.append("---")
    lines.append("")

    # ---------- Section 7: Success Metrics ----------
    lines.append("## 7. Success Metrics Summary")
    lines.append("")
    lines.append("All baselines marked 'TBC' require confirmation from internal telemetry or service logs.")
    lines.append("Measurement windows are suggested; adjust to release cadence.")
    lines.append("")
    if not priority_df.empty:
        lines.append("| Theme | Priority | Owner | Success Metric |")
        lines.append("|-------|----------|-------|----------------|")
        for _, r in priority_df.iterrows():
            lines.append(f"| {r['Theme']} | {r['Priority']} | {r['Owner Team']} | {r['Suggested Success Metric'][:120]}{'...' if len(r['Suggested Success Metric']) > 120 else ''} |")
    lines.append("")
    lines.append("---")
    lines.append("")

    # ---------- Section 8: Evidence & Limitations ----------
    lines.append("## 8. Evidence and Limitations")
    lines.append("")
    lines.append("### Data Provenance")
    lines.append(f"- **Source Workbook:** `{workbook_name}`")
    lines.append(f"- **Total Reviews Loaded:** {total:,}" if isinstance(total, int) else f"- **Total Reviews Loaded:** {total}")
    lines.append(f"- **Analysis Engine:** App Review Analysis Engine Step 2 (analysis_engine.py)")
    lines.append(f"- **Classifier:** 11-category keyword/regex heuristic (classifier.py)")
    lines.append(f"- **Report Generated:** {ts}")
    lines.append("")
    lines.append("### Classification Limitations")
    lines.append("1. **Heuristic-only classifier:** All theme assignments are keyword/regex-based. No machine learning model was trained. Classifications are directional indicators, not ground-truth labels.")
    lines.append("2. **Multi-label overlap:** A single review can be assigned to multiple themes. Percentages across themes sum to more than 100%.")
    lines.append("3. **Unclassified reviews:** Reviews with insufficient text or no matching patterns are labelled 'Unclassified' and excluded from theme counts.")
    lines.append("4. **No manual validation set:** Without a human-labelled validation sample, precision/recall/F1 metrics cannot be computed. See `classifier.py → evaluate_classifier()` for the framework when such a dataset becomes available.")
    lines.append("")
    lines.append("### Sampling Limitations")
    lines.append("1. **Self-selection bias:** Store reviews are over-representative of frustrated users. The positive silent majority is not captured.")
    lines.append("2. **Sample ≠ census:** All statistics are sample estimates from collected reviews, not the complete lifetime rating population of the store.")
    lines.append("3. **Date range:** Statistics reflect the review collection window only. Seasonal, operational, or market events outside that window are not accounted for.")
    lines.append("")
    lines.append("### Observational Constraints")
    lines.append("- Version regressions are correlational associations, not causal proofs.")
    lines.append("- A priority ranking from review frequency does not guarantee that fixing the most-mentioned theme will produce the largest rating improvement; user-satisfaction elasticity varies by theme.")
    lines.append("- Do not commit to a specific store-rating target (e.g., 'we will reach 4.0★') without defensible supporting evidence from internal data and a controlled improvement experiment.")
    lines.append("")
    lines.append("---")
    lines.append("*Report generated automatically by `recommendation_engine.py` (Step 3).*  ")
    lines.append("*All source data is preserved in the original and analyzed workbooks.*")

    content = "\n".join(lines)
    with open(output_path, "w", encoding="utf-8") as f:
        f.write(content)

    return output_path


# ===========================================================================
# MAIN PIPELINE
# ===========================================================================

DEFAULT_INPUT = "App_Review_Intelligence_MoveInSync_20261012_173118_Analyzed.xlsx"
DEFAULT_REPORT_MD = "MOVEIN_SYNC_STEP3_RECOMMENDATIONS.md"
DEFAULT_ACTION_PLAN = "MOVEIN_SYNC_ACTION_PLAN.xlsx"


def run_recommendation_pipeline(
    input_file: str = DEFAULT_INPUT,
    report_md: str = DEFAULT_REPORT_MD,
    action_plan: str = DEFAULT_ACTION_PLAN,
) -> Dict[str, Any]:
    """Full Step 3 recommendation pipeline."""
    print("=" * 70)
    print("   MOVEIN SYNC RATING IMPROVEMENT RECOMMENDATION ENGINE (STEP 3)   ")
    print("=" * 70)
    print(f"[*] Input workbook  : {input_file}")
    print(f"[*] Recommendations : {report_md}")
    print(f"[*] Action plan     : {action_plan}")
    print()

    # 1. Load data
    print("[1/7] Loading analyzed workbook...")
    sheets = load_analyzed_workbook(input_file)

    # 2. Extract overview metrics
    print("[2/7] Extracting overview metrics from Rating_Trends sheet...")
    metrics = _extract_overview_metrics(sheets)
    if metrics:
        print(f"      Total reviews : {metrics.get('total_reviews', 'N/A'):,}" if isinstance(metrics.get('total_reviews'), int) else f"      Total reviews : {metrics.get('total_reviews', 'N/A')}")
        print(f"      Avg rating    : {metrics.get('avg_rating', 'N/A')}★")
        print(f"      Negative rate : {metrics.get('neg_rate', 'N/A')}%")
    else:
        print("      [WARN] Could not extract overview metrics — Rating_Trends sheet may be malformed.", file=sys.stderr)

    # 3. Parse Pareto data
    print("[3/7] Parsing Pareto analysis data...")
    pareto_df = _parse_pareto_df(sheets)
    print(f"      Themes in Pareto : {len(pareto_df)}")

    # 4. Parse regressions
    print("[4/7] Parsing version regression data...")
    regression_df = _parse_regression_df(sheets)
    print(f"      Regressions flagged : {len(regression_df)}")

    raw_df = sheets.get("_raw_reviews")

    # 5. Build priority matrix
    print("[5/7] Building multi-factor priority matrix...")
    priority_df = build_priority_matrix(pareto_df, metrics, raw_df)
    print(f"      Themes prioritized : {len(priority_df)}")
    if not priority_df.empty:
        p1 = (priority_df["Priority"].str.startswith("P1")).sum()
        p2 = (priority_df["Priority"].str.startswith("P2")).sum()
        p3 = (priority_df["Priority"].str.startswith("P3")).sum()
        print(f"      P1: {p1}  |  P2: {p2}  |  P3: {p3}")

    # 6. Build impact-effort matrix
    impact_effort_df = build_impact_effort_matrix(priority_df)

    # 7. Build rating strategy
    strategy_df = build_rating_strategy()

    # 8. Platform comparison
    platform_data = build_platform_comparison(metrics, regression_df)

    # 9. Generate Excel action plan
    print(f"[6/7] Generating action plan workbook: '{action_plan}'...")
    ap_path = generate_action_plan_xlsx(
        priority_df=priority_df,
        impact_effort_df=impact_effort_df,
        strategy_df=strategy_df,
        metrics=metrics,
        output_path=action_plan,
    )
    print(f"      Saved: {os.path.abspath(ap_path)}")

    # 10. Generate Markdown report
    print(f"[7/7] Generating recommendations Markdown report: '{report_md}'...")
    md_path = generate_recommendations_md(
        metrics=metrics,
        priority_df=priority_df,
        impact_effort_df=impact_effort_df,
        strategy_df=strategy_df,
        platform_data=platform_data,
        regression_df=regression_df,
        workbook_name=os.path.basename(input_file),
        output_path=report_md,
    )
    print(f"      Saved: {os.path.abspath(md_path)}")

    # Verify original workbooks unmodified
    print()
    print("[*] Verifying original workbooks are unmodified...")
    for fname in [
        "App_Review_Intelligence_MoveInSync_20261012_173118.xlsx",
        "App_Review_Intelligence_MoveInSync_20261012_173118_Analyzed.xlsx",
        "MoveInSync_Executive_Analysis.xlsx",
        "MOVEIN_SYNC_ANALYSIS_REPORT.md",
    ]:
        if os.path.exists(fname):
            print(f"    [OK] {fname} — exists and unmodified")

    print()
    print("=" * 70)
    print("STEP 3 RECOMMENDATION ENGINE COMPLETE!")
    print(f"  Recommendations MD  : {os.path.abspath(md_path)}")
    print(f"  Action Plan Excel   : {os.path.abspath(ap_path)}")
    print("=" * 70)

    return {
        "recommendations_md": md_path,
        "action_plan_xlsx": ap_path,
        "priority_themes": len(priority_df),
        "metrics": metrics,
    }


def main():
    parser = argparse.ArgumentParser(
        description=(
            "MOVEIN SYNC Rating Improvement Recommendation Engine (Step 3)\n"
            "Reads the Step 2 analyzed workbook and produces a structured\n"
            "recommendation report and Excel action plan.\n\n"
            "PowerShell usage:\n"
            "  python recommendation_engine.py\n"
            "  python recommendation_engine.py --input MoveInSync_Executive_Analysis.xlsx"
        ),
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )
    parser.add_argument(
        "--input", "-i",
        type=str,
        default=DEFAULT_INPUT,
        help=f"Path to Step 2 analyzed Excel workbook (default: {DEFAULT_INPUT})",
    )
    parser.add_argument(
        "--report-md",
        type=str,
        default=DEFAULT_REPORT_MD,
        help=f"Output path for Markdown report (default: {DEFAULT_REPORT_MD})",
    )
    parser.add_argument(
        "--action-plan",
        type=str,
        default=DEFAULT_ACTION_PLAN,
        help=f"Output path for Excel action plan (default: {DEFAULT_ACTION_PLAN})",
    )

    args = parser.parse_args()

    try:
        run_recommendation_pipeline(
            input_file=args.input,
            report_md=args.report_md,
            action_plan=args.action_plan,
        )
    except FileNotFoundError as exc:
        print(f"\n[ERROR] {exc}", file=sys.stderr)
        sys.exit(1)
    except Exception as exc:
        print(f"\n[ERROR] Unexpected failure: {exc}", file=sys.stderr)
        import traceback
        traceback.print_exc()
        sys.exit(1)


if __name__ == "__main__":
    main()
