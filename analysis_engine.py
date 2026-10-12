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
        elif norm in ("review title", "title"):
            col_mapping[c] = "review_title"
        elif norm in ("review text", "content", "text"):
            col_mapping[c] = "review_text"
        elif norm in ("review date", "date", "at"):
            col_mapping[c] = "review_date"
        elif norm in ("app version", "version"):
            col_mapping[c] = "app_version"
        elif norm in ("developer reply", "developer response"):
            col_mapping[c] = "developer_reply"
        elif norm in ("helpful votes", "thumbs up count", "votes"):
            col_mapping[c] = "helpful_votes"
        elif norm in ("source url", "url"):
            col_mapping[c] = "source_url"

    df = df_raw.rename(columns=col_mapping).copy()

    # Ensure critical columns exist with defaults
    if "app_name" not in df.columns:
        df["app_name"] = "MoveInSync"
    if "platform" not in df.columns:
        df["platform"] = "Unknown"
    if "app_role" not in df.columns:
        df["app_role"] = "employee"
    if "review_text" not in df.columns:
        df["review_text"] = ""
    if "review_title" not in df.columns:
        df["review_title"] = ""
    if "app_version" not in df.columns:
        df["app_version"] = "Unknown"

    initial_count = len(df)

    # Coerce ratings safely
    df["rating"] = pd.to_numeric(df["rating"], errors="coerce")
    invalid_ratings = df["rating"].isnull().sum()
    # Filter valid ratings 1..5
    df = df[df["rating"].between(1, 5)].copy()
    df["rating"] = df["rating"].astype(int)

    # Deduplicate reviews
    if "review_id" in df.columns:
        df["review_id"] = df["review_id"].astype(str)
        duplicates_count = df.duplicated(subset=["review_id"]).sum()
        df = df.drop_duplicates(subset=["review_id"]).copy()
    else:
        duplicates_count = df.duplicated(subset=["reviewer_name", "review_date", "review_text"]).sum()
        df = df.drop_duplicates(subset=["reviewer_name", "review_date", "review_text"]).copy()

    # Normalize Dates
    df["review_date_parsed"] = pd.to_datetime(df["review_date"], errors="coerce")
    df["year_month"] = df["review_date_parsed"].dt.strftime("%Y-%m")

    # Normalize Versions
    df["app_version_clean"] = df["app_version"].fillna("Unknown / Not Specified").astype(str).str.strip()
    df.loc[df["app_version_clean"].isin(["", "nan", "None", "null"]), "app_version_clean"] = "Unknown / Not Specified"

    audit_meta = {
        "input_filepath": filepath,
        "initial_rows": initial_count,
        "valid_rows": len(df),
        "invalid_ratings_dropped": int(invalid_ratings),
        "duplicates_dropped": int(duplicates_count),
        "apps_found": df["app_name"].value_counts().to_dict(),
        "platforms_found": df["platform"].value_counts().to_dict(),
        "roles_found": df["app_role"].value_counts().to_dict(),
    }

    return df, audit_meta


def compute_rating_analysis(df: pd.DataFrame, min_sample: int = 15) -> Dict[str, Any]:
    """
    Calculate rating distributions, negative rates with Wilson confidence intervals,
    and monthly trends with sample size safeguards.
    """
    def _calc_stats(sub_df: pd.DataFrame, segment_name: str, sub_segment: str) -> Dict[str, Any]:
        n = len(sub_df)
        if n == 0:
            return {
                "Segment": segment_name, "Sub-Segment": sub_segment, "Sample Size (N)": 0,
                "Sample Size Flag": "No Data", "Sample Avg Rating": None,
                "1★ Count": 0, "1★ %": 0.0, "2★ Count": 0, "2★ %": 0.0,
                "3★ Count": 0, "3★ %": 0.0, "4★ Count": 0, "4★ %": 0.0,
                "5★ Count": 0, "5★ %": 0.0,
                "Negative Count (1★+2★)": 0, "Negative Rate %": 0.0,
                "Wilson 95% CI Lower": 0.0, "Wilson 95% CI Upper": 0.0,
                "Positive Rate % (4★+5★)": 0.0,
            }

        ratings = sub_df["rating"]
        avg_r = round(ratings.mean(), 2)
        star_counts = ratings.value_counts().to_dict()

        c1 = star_counts.get(1, 0)
        c2 = star_counts.get(2, 0)
        c3 = star_counts.get(3, 0)
        c4 = star_counts.get(4, 0)
        c5 = star_counts.get(5, 0)

        neg_count = c1 + c2
        neg_rate = round(neg_count / n * 100, 1)
        ci_low, ci_high = wilson_confidence_interval(neg_count, n)
        pos_rate = round((c4 + c5) / n * 100, 1)

        flag = "Sufficient Sample" if n >= min_sample else f"Low Sample (N < {min_sample})"

        return {
            "Segment": segment_name,
            "Sub-Segment": sub_segment,
            "Sample Size (N)": n,
            "Sample Size Flag": flag,
            "Sample Avg Rating": avg_r,
            "1★ Count": c1, "1★ %": round(c1 / n * 100, 1),
