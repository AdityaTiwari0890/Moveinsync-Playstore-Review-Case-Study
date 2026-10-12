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
