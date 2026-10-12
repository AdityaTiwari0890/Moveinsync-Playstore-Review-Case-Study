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
