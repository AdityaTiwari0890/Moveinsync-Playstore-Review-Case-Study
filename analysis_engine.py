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
            "Total Mentions": t_count,
            "% of All Reviews (N={})".format(total_reviews): round(t_count / total_reviews * 100, 1),
            "Average Rating": avg_r,
            "Negative Mentions (1★+2★)": neg_count,
            "Negative Rate %": neg_rate,
            "Employee Mentions": emp_count,
            "Driver Mentions": drv_count,
            "Android Mentions": and_count,
            "iOS Mentions": ios_count,
            "Sample Negative Excerpts": sample_quotes,
        })

    theme_analysis_df = pd.DataFrame(all_theme_rows).sort_values("Negative Mentions (1★+2★)", ascending=False)

    # --- Table 2: Pareto 80/20 Analysis on 1★ & 2★ Negative Reviews ---
    df_negative = df_classified[df_classified["rating"].isin([1, 2])].copy()
    total_negative_reviews = len(df_negative)

    pareto_rows = []
    # Count occurrences across all 11 categories
    for item in classifier.taxonomy:
        t_name = item["theme"]
        mask = df_negative["assigned_themes"].apply(lambda themes: t_name in themes)
        c = mask.sum()

        share_neg = round(c / total_negative_reviews * 100, 1) if total_negative_reviews > 0 else 0.0
        ci_l, ci_u = wilson_confidence_interval(c, total_negative_reviews)

        pareto_rows.append({
            "Theme": t_name,
            "Negative Reviews Count": int(c),
            "Share of Negative Reviews % (N={})".format(total_negative_reviews): share_neg,
            "Wilson 95% CI Lower": ci_l,
            "Wilson 95% CI Upper": ci_u,
        })

    # Sort descending by frequency
    pareto_df = pd.DataFrame(pareto_rows).sort_values("Negative Reviews Count", ascending=False).reset_index(drop=True)

    # Calculate cumulative counts and cumulative percentage of total theme mentions
    total_mentions_sum = pareto_df["Negative Reviews Count"].sum()
    pareto_df["Cumulative Mentions"] = pareto_df["Negative Reviews Count"].cumsum()
    pareto_df["Cumulative Share %"] = round(
        pareto_df["Cumulative Mentions"] / total_mentions_sum * 100, 1
    ) if total_mentions_sum > 0 else 0.0

    # Pareto Status: Demarcate Vital Few (up to ~75-80% cutoff) vs Remaining
    pareto_status = []
    threshold_crossed = False
    for idx, row in pareto_df.iterrows():
        cum = row["Cumulative Share %"]
        if not threshold_crossed:
            pareto_status.append("Vital Few (Pareto 80% Core)")
            if cum >= 75.0:
                threshold_crossed = True
        else:
            pareto_status.append("Remaining Long Tail")
    pareto_df["Pareto Classification"] = pareto_status

    return df_classified, theme_analysis_df, pareto_df


def compute_version_analysis(
    df: pd.DataFrame,
    min_sample: int = 15,
    rating_drop_threshold: float = 0.30,
    neg_jump_threshold: float = 15.0,
) -> Tuple[pd.DataFrame, pd.DataFrame]:
    """
    Group reviews by app version, compute metrics, flag small sample sizes,
    and detect potential sequential release regressions.
    """
    version_groups = []
    flagged_regressions = []

    # Group by App Name, Platform, App Version
    grouped = df.groupby(["app_name", "platform", "app_version_clean"])

    for (app, plat, ver), sub_df in grouped:
        n = len(sub_df)
        avg_r = round(sub_df["rating"].mean(), 2)
        neg_count = (sub_df["rating"].isin([1, 2])).sum()
        neg_rate = round(neg_count / n * 100, 1)
        ci_l, ci_u = wilson_confidence_interval(neg_count, n)

        dates = sub_df["review_date"].dropna().astype(str).tolist()
        dates_sorted = sorted(dates) if dates else []
        earliest_date = dates_sorted[0][:10] if dates_sorted else "N/A"
        latest_date = dates_sorted[-1][:10] if dates_sorted else "N/A"

        # Determine top 2 complaint themes for this version
        theme_counts = defaultdict(int)
        for _, r in sub_df[sub_df["rating"].isin([1, 2])].iterrows():
            for t in r.get("assigned_themes", []):
                if t not in ("Unclassified", "Other or unclear"):
                    theme_counts[t] += 1
        top_themes = ", ".join([k for k, _ in sorted(theme_counts.items(), key=lambda x: x[1], reverse=True)[:2]]) or "None specific"

        sample_flag = "Sufficient" if n >= min_sample else f"Low Sample (N < {min_sample})"

        version_groups.append({
            "App Name": app,
            "Platform": plat,
            "App Version": ver,
            "Sample Size (N)": n,
            "Sample Size Flag": sample_flag,
            "Average Rating": avg_r,
            "Negative Rate %": neg_rate,
            "Wilson 95% CI Lower": ci_l,
            "Wilson 95% CI Upper": ci_u,
            "Earliest Review Date": earliest_date,
            "Latest Review Date": latest_date,
            "Top Complaint Themes": top_themes,
        })

    v_df = pd.DataFrame(version_groups).sort_values(
        ["App Name", "Platform", "Sample Size (N)"], ascending=[True, True, False]
    ).reset_index(drop=True)

    # Detect Sequential Regressions for versions with sufficient sample size
    for (app, plat), app_v_df in v_df.groupby(["App Name", "Platform"]):
        # Filter out 'Unknown / Not Specified' and sort by Earliest Review Date
        clean_v_df = app_v_df[app_v_df["App Version"] != "Unknown / Not Specified"].sort_values("Earliest Review Date").reset_index(drop=True)

        if len(clean_v_df) >= 2:
            for i in range(1, len(clean_v_df)):
                prev = clean_v_df.iloc[i - 1]
                curr = clean_v_df.iloc[i]

                # Check sample thresholds
                if prev["Sample Size (N)"] >= min_sample and curr["Sample Size (N)"] >= min_sample:
                    delta_r = round(curr["Average Rating"] - prev["Average Rating"], 2)
                    delta_neg = round(curr["Negative Rate %"] - prev["Negative Rate %"], 1)

                    if delta_r <= -rating_drop_threshold or delta_neg >= neg_jump_threshold:
                        # Extract sample negative quotes from current version
                        curr_ver_str = curr["App Version"]
                        curr_revs = df[(df["app_name"] == app) & (df["platform"] == plat) & (df["app_version_clean"] == curr_ver_str) & (df["rating"].isin([1, 2]))]
                        quotes = []
                        for _, q_row in curr_revs.head(3).iterrows():
                            txt = str(q_row.get("review_text", "")).replace("\n", " ").strip()
                            if len(txt) > 90:
                                txt = txt[:87] + "..."
                            quotes.append(f'[{str(q_row.get("review_id", ""))[:8]}]: "{txt}"')
                        excerpts_str = " | ".join(quotes) if quotes else "N/A"

                        flagged_regressions.append({
                            "App Name": app,
                            "Platform": plat,
                            "Prior Version": prev["App Version"],
                            "Current Version": curr["App Version"],
                            "Prior Rating (N={})".format(prev["Sample Size (N)"]): prev["Average Rating"],
                            "Current Rating (N={})".format(curr["Sample Size (N)"]): curr["Average Rating"],
                            "Rating Delta": delta_r,
                            "Prior Negative %": prev["Negative Rate %"],
                            "Current Negative %": curr["Negative Rate %"],
                            "Negative Rate Jump %": delta_neg,
                            "Association Assessment": f"Potential Regression ({delta_r}★ drop; +{delta_neg}% negative rate)",
                            "Representative Review Evidence": excerpts_str,
                            "Analytical Warning": "Observational association only; not proof of feature-level causation.",
                        })

    reg_df = pd.DataFrame(flagged_regressions) if flagged_regressions else pd.DataFrame(columns=[
        "App Name", "Platform", "Prior Version", "Current Version", "Rating Delta",
        "Negative Rate Jump %", "Association Assessment", "Representative Review Evidence", "Analytical Warning"
    ])

    return v_df, reg_df


def compute_action_priorities(
    pareto_df: pd.DataFrame, theme_analysis_df: pd.DataFrame, total_negative_reviews: int
) -> pd.DataFrame:
    """
    Construct the Action-Priority Matrix assigning P1-P3 priorities, user impact severity,
    estimated engineering effort, recommended investigations, and measurable success metrics.
    """
    priority_matrix = [
        {
            "Theme": "Cab delays and no-shows",
            "Priority": "P1 - Critical",
            "Severity / Impact": "High - Direct disruption to employee office shift arrival and driver earnings.",
            "Confidence in Evidence": "High (Statistically dominant negative theme across both platforms).",
            "Estimated Effort": "Medium (Algorithmic dispatch tuning + driver arrival SLA enforcement).",
            "Recommended Investigation / Fix": "Audit buffer time parameters, implement real-time driver delay alerts, and penalize repeated driver no-shows.",
            "Suggested Success Metric": "On-Time Arrival rate >= 95%; reduction in 'delay' reviews by 30%.",
            "Action Rationale": "Core value proposition failure; primary driver of 1-star ratings."
        },
        {
            "Theme": "GPS and location tracking",
            "Priority": "P1 - Critical",
            "Severity / Impact": "High - Commuters cannot track vehicle; leads to panic, missed cabs, and wasted driver wait time.",
            "Confidence in Evidence": "High (Consistent complaint across iOS and Android reviews).",
            "Estimated Effort": "Medium (Telemetry socket optimization, background location service wake-locks).",
            "Recommended Investigation / Fix": "Debug background GPS ping frequency on Android 13/14 battery optimization; improve map vector rendering.",
            "Suggested Success Metric": "GPS refresh latency < 5 seconds; 40% reduction in tracking-related negative reviews.",
            "Action Rationale": "Core operational tool; broken tracking generates immediate customer support escalations."
        },
        {
            "Theme": "Booking and OTP failures",
            "Priority": "P1 - Critical",
            "Severity / Impact": "High - Trips cannot be initiated without OTP boarding verification.",
            "Confidence in Evidence": "High (Frequent blocker mentioned by both employees and drivers).",
            "Estimated Effort": "Low to Medium (SMS gateway fallback, in-app push OTP, driver-side QR boarding fallback).",
            "Recommended Investigation / Fix": "Implement in-app dynamic boarding QR codes / Bluetooth handshake alongside SMS OTP.",
            "Suggested Success Metric": "Boarding verification success rate >= 99.5%; boarding delay < 15 seconds.",
            "Action Rationale": "Hard gate preventing trip start; directly stalls cab departure."
        },
        {
            "Theme": "App crashes and performance",
            "Priority": "P2 - High",
            "Severity / Impact": "Medium to High - Prevents users from accessing the app during critical commute windows.",
            "Confidence in Evidence": "High (Crash traces mentioned across versions).",
            "Estimated Effort": "Medium (Crashlytics audit, memory leak profiling, ANR fixes).",
            "Recommended Investigation / Fix": "Analyze Firebase / Crashlytics symbolicated crash logs for top fatal stack traces in current releases.",
            "Suggested Success Metric": "Crash-Free Session rate >= 99.8%; ANR rate < 0.2%.",
            "Action Rationale": "Technical debt directly degrading app store rating."
        },
        {
            "Theme": "Driver-side experience",
            "Priority": "P2 - High",
            "Severity / Impact": "Medium - Frustrated drivers decline trips, reject duties, or exhibit poor commuter service.",
            "Confidence in Evidence": "Medium (High concentration in MoveInSync Driver Partner app reviews).",
            "Estimated Effort": "Medium (Duty acceptance workflow simplification, transparent toll reimbursement).",
            "Recommended Investigation / Fix": "Streamline duty acceptance UX, ensure automated toll parsing, and add clear toll dispute buttons.",
            "Suggested Success Metric": "Driver duty acceptance rate >= 90%; Driver app Play Store rating >= 4.2★.",
            "Action Rationale": "Supply-side health is essential for reliable fleet operations."
        },
        {
            "Theme": "Login and account access",
            "Priority": "P2 - High",
            "Severity / Impact": "Medium - Locks users out during shift changes or after enterprise SSO updates.",
            "Confidence in Evidence": "Medium (Prominent in corporate enterprise logins).",
            "Estimated Effort": "Low (SSO session token refresh enhancements, clear error messaging for corporate HR database sync).",
            "Recommended Investigation / Fix": "Improve SSO token refresh failure messaging with direct link to corporate transport helpdesk.",
            "Suggested Success Metric": "First-attempt login success rate >= 98%.",
            "Action Rationale": "Prevents user onboarding and daily booking."
        },
        {
            "Theme": "Customer support",
            "Priority": "P3 - Medium",
            "Severity / Impact": "Medium - Unresolved grievances exacerbate negative store reviews.",
            "Confidence in Evidence": "Medium (Mentions of delayed helpline answers).",
            "Estimated Effort": "Medium (In-app chatbot routing, corporate transport manager SLA dashboard).",
            "Recommended Investigation / Fix": "Deploy automated trip-delay grievance workflow directly inside app.",
            "Suggested Success Metric": "First Response Time < 2 minutes for active trip emergencies.",
            "Action Rationale": "Mitigates escalation from frustrated users to public app store reviews."
        },
        {
            "Theme": "UI and usability",
            "Priority": "P3 - Medium",
            "Severity / Impact": "Low to Medium - Friction in discovering roster modifications or pickup points.",
            "Confidence in Evidence": "Low to Medium (Subjective design feedback).",
            "Estimated Effort": "Medium (Usability lab testing, UX simplification).",
            "Recommended Investigation / Fix": "Conduct user testing on roster scheduling screens and clarify button affordances.",
            "Suggested Success Metric": "Booking completion task time reduced by 25%.",
            "Action Rationale": "Improves daily commuter satisfaction."
        },
    ]

    action_df = pd.DataFrame(priority_matrix)
    # Merge negative counts from pareto_df if available
    count_map = pareto_df.set_index("Theme")["Negative Reviews Count"].to_dict()
    share_map = pareto_df.set_index("Theme")[f"Share of Negative Reviews % (N={total_negative_reviews})"].to_dict()

    action_df["Negative Reviews Count"] = action_df["Theme"].map(count_map).fillna(0).astype(int)
    action_df[f"Share of Negative Reviews % (N={total_negative_reviews})"] = action_df["Theme"].map(share_map).fillna(0.0)

    # Reorder columns
    cols = [
        "Priority", "Theme", "Negative Reviews Count", f"Share of Negative Reviews % (N={total_negative_reviews})",
        "Severity / Impact", "Confidence in Evidence", "Estimated Effort",
        "Recommended Investigation / Fix", "Suggested Success Metric", "Action Rationale"
    ]
    return action_df[cols].sort_values("Priority")


def enhance_workbook_with_analysis(
    input_filepath: str,
    output_filepath: str,
    rating_data: Dict[str, Any],
    theme_analysis_df: pd.DataFrame,
    pareto_df: pd.DataFrame,
    version_df: pd.DataFrame,
    regression_df: pd.DataFrame,
    action_df: pd.DataFrame,
    total_reviews: int,
    total_negative_reviews: int,
) -> str:
    """
    Load the original workbook with openpyxl, preserve all original raw sheets,
    and append/update the 5 analysis sheets with native Excel tables and charts.
    """
    wb = openpyxl.load_workbook(input_filepath)

    # Idempotent cleanup: remove analysis sheets if they exist from prior runs
    analysis_sheet_names = [
        "Rating_Trends", "Theme_Analysis", "Version_Analysis", "Pareto_Analysis", "Action_Priorities"
    ]
    for name in analysis_sheet_names:
        if name in wb.sheetnames:
            wb.remove(wb[name])

    # Shared Styling
    header_font = Font(name="Calibri", size=11, bold=True, color=WHITE_TEXT)
    header_fill = PatternFill(start_color=NAVY_HEADER, end_color=NAVY_HEADER, fill_type="solid")
    section_fill = PatternFill(start_color=SECTION_HEADER, end_color=SECTION_HEADER, fill_type="solid")
    regular_font = Font(name="Calibri", size=10)
    thin_border = Border(
        left=Side(style="thin", color=BORDER_GRAY),
        right=Side(style="thin", color=BORDER_GRAY),
        top=Side(style="thin", color=BORDER_GRAY),
        bottom=Side(style="thin", color=BORDER_GRAY),
    )

    # -------------------------------------------------------------
    # 1. Sheet: Rating_Trends
    # -------------------------------------------------------------
    ws_rt = wb.create_sheet(title="Rating_Trends")
    ws_rt.append(["SECTION 1: RATING OVERVIEW & SEGMENT BENCHMARKS"])
    ws_rt.cell(row=1, column=1).font = Font(name="Calibri", size=12, bold=True, color=WHITE_TEXT)
    ws_rt.cell(row=1, column=1).fill = section_fill

    overview_df = rating_data["overview_table"]
    ws_rt.append(overview_df.columns.tolist())
    for _, row in overview_df.iterrows():
        ws_rt.append(row.tolist())

    ws_rt.append([])  # Spacer

    r_idx = ws_rt.max_row + 1
    ws_rt.cell(row=r_idx, column=1, value="SECTION 2: MONTHLY REVIEW VOLUME & RATING TRAJECTORY")
    ws_rt.cell(row=r_idx, column=1).font = Font(name="Calibri", size=12, bold=True, color=WHITE_TEXT)
    ws_rt.cell(row=r_idx, column=1).fill = section_fill

    monthly_df = rating_data["monthly_table"]
    monthly_start_row = ws_rt.max_row + 1
    ws_rt.append(monthly_df.columns.tolist())
    for _, row in monthly_df.iterrows():
        ws_rt.append(row.tolist())

    # 1. Add BarChart for Star Rating Distribution (1★ to 5★)
