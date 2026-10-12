"""
App Review Analysis Engine - Step 2
Enterprise analytical pipeline for MOVEIN SYNC mobile application reviews.
Performs data normalization, statistical rating calculations with Wilson confidence intervals,
11-category multi-label complaint classification, Pareto 80/20 analysis, version regression tracking,
Action-Priority matrix generation, Excel enhancements with embedded openpyxl charts, and executive reporting.
"""

from typing import List, Dict, Any, Tuple, Optional
import os
import sys
import math
import datetime
import argparse
from collections import defaultdict
import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.chart import BarChart, LineChart, Reference
from openpyxl.worksheet.table import Table, TableStyleInfo

# Configure UTF-8 encoding for Windows terminal safety
if sys.stdout and hasattr(sys.stdout, "reconfigure"):
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="backslashreplace")
    except Exception:
        pass
if sys.stderr and hasattr(sys.stderr, "reconfigure"):
    try:
        sys.stderr.reconfigure(encoding="utf-8", errors="backslashreplace")
    except Exception:
        pass

from classifier import ReviewClassifier, TAXONOMY_CONFIG


# Color Scheme
NAVY_HEADER = "1B365D"
SECTION_HEADER = "2C4C64"
WHITE_TEXT = "FFFFFF"
BORDER_GRAY = "D9D9D9"
ALERT_RED = "C00000"
FLAG_FILL = "FDEDEC"
PARETO_VITAL_FILL = "E8F8F5"


def wilson_confidence_interval(count: int, n: int, confidence: float = 0.95) -> Tuple[float, float]:
    """
    Calculate the Wilson score confidence interval for a proportion.
    Formula:
      (p_hat + z^2/(2n) +/- z * sqrt((p_hat*(1-p_hat)/n) + z^2/(4n^2))) / (1 + z^2/n)
    Returns:
      (ci_lower_pct, ci_upper_pct) as percentages rounded to 1 decimal place.
    """
    if n <= 0:
        return (0.0, 0.0)

    p_hat = count / n
    z = 1.95996  # 95% confidence level

    denominator = 1 + (z ** 2) / n
    center = p_hat + (z ** 2) / (2 * n)
    margin = z * math.sqrt((p_hat * (1 - p_hat) / n) + (z ** 2) / (4 * (n ** 2)))

    lower = max(0.0, (center - margin) / denominator)
    upper = min(1.0, (center + margin) / denominator)

    return (round(lower * 100, 1), round(upper * 100, 1))


def load_and_normalize_workbook(filepath: str) -> Tuple[pd.DataFrame, Dict[str, Any]]:
    """
    Load reviews from an Excel workbook, normalize fields, deduplicate reviews,
    and safely handle missing data.
    """
    if not os.path.exists(filepath):
        raise FileNotFoundError(f"Input workbook not found at: '{filepath}'. Please check the path and try again.")

    with pd.ExcelFile(filepath) as excel_file:
        sheet_names = excel_file.sheet_names

        # Determine review source sheet
        if "Combined_Reviews" in sheet_names:
            df_raw = pd.read_excel(excel_file, sheet_name="Combined_Reviews")
        elif "Android_Reviews" in sheet_names or "iOS_Reviews" in sheet_names:
            dfs = []
            if "Android_Reviews" in sheet_names:
                dfs.append(pd.read_excel(excel_file, sheet_name="Android_Reviews"))
            if "iOS_Reviews" in sheet_names:
                dfs.append(pd.read_excel(excel_file, sheet_name="iOS_Reviews"))
            df_raw = pd.concat(dfs, ignore_index=True)
        else:
            raise ValueError(f"Workbook '{filepath}' does not contain 'Combined_Reviews', 'Android_Reviews', or 'iOS_Reviews'.")

    # Column Normalization
    col_mapping = {}
    for c in df_raw.columns:
        norm = str(c).strip().lower().replace("_", " ")
        if norm in ("app name", "app"):
            col_mapping[c] = "app_name"
        elif norm in ("store id", "package id", "app id"):
            col_mapping[c] = "store_id"
        elif norm in ("platform", "os"):
            col_mapping[c] = "platform"
        elif norm in ("app role", "role"):
            col_mapping[c] = "app_role"
        elif norm in ("review id", "id"):
            col_mapping[c] = "review_id"
        elif norm in ("reviewer name", "user", "user name"):
            col_mapping[c] = "reviewer_name"
        elif norm in ("rating", "score"):
            col_mapping[c] = "rating"
