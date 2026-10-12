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
            "2★ Count": c2, "2★ %": round(c2 / n * 100, 1),
            "3★ Count": c3, "3★ %": round(c3 / n * 100, 1),
            "4★ Count": c4, "4★ %": round(c4 / n * 100, 1),
            "5★ Count": c5, "5★ %": round(c5 / n * 100, 1),
            "Negative Count (1★+2★)": neg_count,
            "Negative Rate %": neg_rate,
            "Wilson 95% CI Lower": ci_low,
            "Wilson 95% CI Upper": ci_high,
            "Positive Rate % (4★+5★)": pos_rate,
        }

    # 1. Segment Tables
    overview_rows = [
        _calc_stats(df, "Overall", "All Combined Reviews"),
        _calc_stats(df[df["app_role"] == "employee"], "App Role", "Employee Apps"),
        _calc_stats(df[df["app_role"] == "driver"], "App Role", "Driver Apps"),
        _calc_stats(df[df["app_role"] == "competitor"], "App Role", "Competitor Apps"),
        _calc_stats(df[df["platform"].str.lower() == "android"], "Platform", "Android"),
        _calc_stats(df[df["platform"].str.lower() == "ios"], "Platform", "iOS"),
    ]

    # App specific rows
    for app in df["app_name"].unique():
        overview_rows.append(_calc_stats(df[df["app_name"] == app], "Application", str(app)))

    # 2. Monthly Trajectory
    monthly_rows = []
    df_with_dates = df.dropna(subset=["year_month"]).sort_values("year_month")
    for ym, group in df_with_dates.groupby("year_month"):
        n_m = len(group)
        neg_m = (group["rating"].isin([1, 2])).sum()
        ci_l, ci_u = wilson_confidence_interval(neg_m, n_m)
        monthly_rows.append({
            "Year-Month": ym,
            "Sample Size (N)": n_m,
            "Sample Size Flag": "Sufficient" if n_m >= min_sample else f"Low (N < {min_sample})",
            "Sample Avg Rating": round(group["rating"].mean(), 2),
            "1★ Count": (group["rating"] == 1).sum(),
            "2★ Count": (group["rating"] == 2).sum(),
            "Negative Reviews Count": neg_m,
            "Negative Rate %": round(neg_m / n_m * 100, 1),
            "Wilson 95% CI Lower": ci_l,
            "Wilson 95% CI Upper": ci_u,
        })

    return {
        "overview_table": pd.DataFrame(overview_rows),
        "monthly_table": pd.DataFrame(monthly_rows),
    }


def compute_theme_and_pareto_analysis(
    df: pd.DataFrame, classifier: ReviewClassifier
) -> Tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    """
    Apply multi-label classification to all reviews, compute theme statistics across all data,
    and generate the Pareto 80/20 analysis on 1-star and 2-star negative reviews.
    """
    classified_records = []
    for _, row in df.iterrows():
        text = str(row.get("review_text") or "")
        title = str(row.get("review_title") or "")
        full_text = f"{title} {text}".strip()
        rating = int(row.get("rating", 0))

        c_res = classifier.classify_text(full_text, rating=rating)
        classified_records.append({
            "assigned_themes": c_res["themes"],
            "classification_method": c_res["method"],
            "evidence": str(c_res["evidence"]),
            "heuristic_confidence": c_res["heuristic_confidence"],
            "is_unclassified": c_res["is_unclassified"],
            "audit_note": c_res["audit_note"],
        })

    df_classified = df.copy()
    for k in ["assigned_themes", "classification_method", "evidence", "heuristic_confidence", "is_unclassified", "audit_note"]:
        df_classified[k] = [r[k] for r in classified_records]

    # --- Table 1: Full Theme Breakdown Across All Reviews ---
    total_reviews = len(df_classified)
    all_theme_rows = []

    for item in classifier.taxonomy:
        t_name = item["theme"]
        t_desc = item["description"]

        # Filter reviews where t_name is in assigned_themes
        matching_mask = df_classified["assigned_themes"].apply(lambda themes: t_name in themes)
        t_df = df_classified[matching_mask]
        t_count = len(t_df)

        emp_count = len(t_df[t_df["app_role"] == "employee"])
        drv_count = len(t_df[t_df["app_role"] == "driver"])
        and_count = len(t_df[t_df["platform"].str.lower() == "android"])
        ios_count = len(t_df[t_df["platform"].str.lower() == "ios"])

        avg_r = round(t_df["rating"].mean(), 2) if t_count > 0 else None
        neg_count = len(t_df[t_df["rating"].isin([1, 2])])
        neg_rate = round(neg_count / t_count * 100, 1) if t_count > 0 else 0.0

        # Sample excerpts
        excerpts = []
        for _, ex_row in t_df[t_df["rating"].isin([1, 2])].head(3).iterrows():
            txt = str(ex_row.get("review_text") or "").replace("\n", " ").strip()
            if len(txt) > 100:
                txt = txt[:97] + "..."
            if txt:
                excerpts.append(f'[{ex_row.get("review_id", "")[:8]}]: "{txt}"')
        sample_quotes = " | ".join(excerpts) if excerpts else "N/A"

        all_theme_rows.append({
            "Theme": t_name,
            "Description": t_desc,
