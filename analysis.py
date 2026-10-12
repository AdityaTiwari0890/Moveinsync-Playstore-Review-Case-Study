"""
Business Intelligence & Review Analysis Module
Calculates descriptive ratings statistics, monthly and version trajectories,
segment comparisons (Employee vs. Driver, Android vs. iOS, MoveInSync vs. Competitors),
and extracts audit-verifiable issue/positive themes for executive reporting.
"""

from typing import List, Dict, Any, Tuple
import datetime
import re
from collections import defaultdict


# Business issue and positive theme taxonomy
# Classification is transparently rule-based to allow manual verification
THEME_DEFINITIONS = [
    {
        "name": "App Crashes & Instability",
        "type": "Complaint",
        "keywords": [r"\bcrash\b", r"\bcrashes\b", r"\bcrashing\b", r"\bfreeze\b", r"\bfrozen\b",
                     r"\bhang\b", r"\bforce close\b", r"\bblack screen\b", r"\bstopped working\b"],
        "max_rating": 3,
    },
    {
        "name": "Login & OTP Authentication",
        "type": "Complaint",
        "keywords": [r"\botp\b", r"\blogin\b", r"\blog in\b", r"\bsign in\b", r"\bpassword\b",
                     r"\bverification code\b", r"\bcannot log\b", r"\bregister\b"],
        "max_rating": 3,
    },
    {
        "name": "GPS & Live Tracking Inaccuracy",
        "type": "Complaint",
        "keywords": [r"\bgps\b", r"\blocation\b", r"\blive tracking\b", r"\bmap\b", r"\bwrong location\b",
                     r"\btracking not working\b", r"\bnavigation\b"],
        "max_rating": 3,
    },
    {
        "name": "Cab Delays & Shift Timing",
        "type": "Complaint",
        "keywords": [r"\blate\b", r"\bdelay\b", r"\bdelayed\b", r"\bwaiting\b", r"\bpickup time\b",
                     r"\bdrop time\b", r"\bcancelled\b", r"\bcanceled\b", r"\bno show\b", r"\bmissed cab\b"],
        "max_rating": 3,
    },
    {
        "name": "UI Navigation & Slow Response",
        "type": "Complaint",
        "keywords": [r"\bui\b", r"\bux\b", r"\bslow\b", r"\bglitch\b", r"\blag\b", r"\bconfusing\b",
                     r"\bbad update\b", r"\bworse update\b", r"\bunresponsive\b"],
        "max_rating": 3,
    },
    {
        "name": "Battery Drain & Device Heating",
        "type": "Complaint",
        "keywords": [r"\bbattery\b", r"\bdrain\b", r"\bdraining\b", r"\bheating\b", r"\boverheating\b",
                     r"\bbattery consumption\b"],
        "max_rating": 3,
    },
    {
        "name": "Driver Conduct & Safety Concerns",
        "type": "Complaint",
        "keywords": [r"\brude\b", r"\brash driving\b", r"\bunsafe\b", r"\bdriver behavior\b",
                     r"\breckless\b", r"\bunprofessional\b", r"\barrogant\b", r"\bsafety\b"],
        "max_rating": 3,
    },
    {
        "name": "Customer Support & Grievances",
        "type": "Complaint",
        "keywords": [r"\bsupport\b", r"\bcustomer care\b", r"\bhelpline\b", r"\bno response\b",
                     r"\bunresolved\b", r"\bhelpdesk\b", r"\bcomplaint\b"],
        "max_rating": 3,
    },
    # Positive Themes
    {
        "name": "Punctual & Reliable Commute",
        "type": "Positive",
        "keywords": [r"\bon time\b", r"\bpunctual\b", r"\breliable\b", r"\bsmooth commute\b",
                     r"\btimely\b", r"\bconvenient\b", r"\bgood service\b", r"\bgreat service\b"],
        "min_rating": 4,
    },
    {
        "name": "Intuitive & User-Friendly UI",
        "type": "Positive",
        "keywords": [r"\beasy to use\b", r"\buser friendly\b", r"\bsimple\b", r"\bclean ui\b",
                     r"\bgreat app\b", r"\bgood app\b", r"\bseamless\b", r"\blove this app\b"],
        "min_rating": 4,
    },
    {
        "name": "Helpful & Courteous Drivers",
        "type": "Positive",
        "keywords": [r"\bpolite driver\b", r"\bgood driver\b", r"\bhelpful driver\b",
                     r"\bcourteous\b", r"\bwell behaved\b", r"\bprofessional driver\b"],
        "min_rating": 4,
    },
    {
        "name": "Safety & Peace of Mind",
        "type": "Positive",
        "keywords": [r"\bsafe\b", r"\bsecurity\b", r"\bsos\b", r"\bwomen safety\b",
                     r"\bsecure commute\b", r"\bpeace of mind\b"],
        "min_rating": 4,
    },
]


def calculate_sample_summary(reviews: List[Dict[str, Any]], requested_count: int = 500) -> Dict[str, Any]:
    """
    Calculate rating distributions and descriptive statistics for a collection of reviews.
    """
    total = len(reviews)
    if total == 0:
        return {
            "total_reviews": 0,
            "requested_count": requested_count,
            "target_reached": "No",
            "avg_rating": None,
            "star_1_count": 0, "star_1_pct": 0.0,
            "star_2_count": 0, "star_2_pct": 0.0,
            "star_3_count": 0, "star_3_pct": 0.0,
            "star_4_count": 0, "star_4_pct": 0.0,
            "star_5_count": 0, "star_5_pct": 0.0,
            "negative_count": 0, "negative_pct": 0.0,
            "positive_count": 0, "positive_pct": 0.0,
            "earliest_date": None,
            "most_recent_date": None,
            "collection_date": datetime.datetime.now().strftime("%Y-%m-%d"),
        }

    ratings = [int(r.get("rating") or 0) for r in reviews if r.get("rating")]
    avg_rating = round(sum(ratings) / len(ratings), 2) if ratings else 0.0

    counts = {s: ratings.count(s) for s in range(1, 6)}
    negative_count = counts[1] + counts[2]
    positive_count = counts[4] + counts[5]

    dates = [r.get("review_date") for r in reviews if r.get("review_date")]
    dates_sorted = sorted(dates) if dates else []
    earliest_date = dates_sorted[0][:10] if dates_sorted else None
    most_recent_date = dates_sorted[-1][:10] if dates_sorted else None

    return {
        "total_reviews": total,
        "requested_count": requested_count,
        "target_reached": "Yes" if total >= requested_count else "No",
        "avg_rating": avg_rating,
        "star_1_count": counts[1],
        "star_1_pct": round(counts[1] / total * 100, 1),
        "star_2_count": counts[2],
        "star_2_pct": round(counts[2] / total * 100, 1),
        "star_3_count": counts[3],
        "star_3_pct": round(counts[3] / total * 100, 1),
        "star_4_count": counts[4],
        "star_4_pct": round(counts[4] / total * 100, 1),
        "star_5_count": counts[5],
        "star_5_pct": round(counts[5] / total * 100, 1),
        "negative_count": negative_count,
        "negative_pct": round(negative_count / total * 100, 1),
        "positive_count": positive_count,
        "positive_pct": round(positive_count / total * 100, 1),
        "earliest_date": earliest_date,
        "most_recent_date": most_recent_date,
        "collection_date": datetime.datetime.now().strftime("%Y-%m-%d"),
    }


def calculate_monthly_trends(reviews: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
    """
    Calculate monthly review volume, average rating, and negative review percentage.
    """
    groups: Dict[Tuple[str, str, str], List[Dict[str, Any]]] = defaultdict(list)

    for r in reviews:
        r_date = r.get("review_date")
        if not r_date or len(r_date) < 7:
            continue
        ym = r_date[:7]  # 'YYYY-MM'
        app_name = r.get("app_name", "Unknown")
        platform = r.get("platform", "Unknown")
        groups[(ym, app_name, platform)].append(r)

    results = []
    for (ym, app_name, platform), revs in sorted(groups.items(), key=lambda x: x[0][0], reverse=True):
        ratings = [r.get("rating", 0) for r in revs if r.get("rating")]
        avg_r = round(sum(ratings) / len(ratings), 2) if ratings else 0.0
        neg = sum(1 for r in ratings if r in (1, 2))
        results.append({
            "year_month": ym,
            "app_name": app_name,
            "platform": platform,
            "review_count": len(revs),
            "avg_rating": avg_r,
            "negative_reviews_count": neg,
            "negative_percentage": round(neg / len(revs) * 100, 1) if revs else 0.0,
        })
    return results


def calculate_version_trends(reviews: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
    """
    Calculate rating trends grouped by app version where version data is available.
    """
    groups: Dict[Tuple[str, str, str], List[Dict[str, Any]]] = defaultdict(list)

    for r in reviews:
        ver = r.get("app_version")
        if not ver or ver.lower() in ("none", "null", "unknown", "varies with device"):
            continue
        app_name = r.get("app_name", "Unknown")
        platform = r.get("platform", "Unknown")
        groups[(app_name, platform, ver)].append(r)

    results = []
    for (app_name, platform, ver), revs in sorted(groups.items(), key=lambda x: len(x[1]), reverse=True):
        ratings = [r.get("rating", 0) for r in revs if r.get("rating")]
        avg_r = round(sum(ratings) / len(ratings), 2) if ratings else 0.0
        neg = sum(1 for r in ratings if r in (1, 2))
        results.append({
            "app_name": app_name,
            "platform": platform,
            "app_version": ver,
            "review_count": len(revs),
            "avg_rating": avg_r,
            "negative_percentage": round(neg / len(revs) * 100, 1) if revs else 0.0,
        })
    return results


def calculate_segment_comparisons(reviews: List[Dict[str, Any]]) -> Dict[str, Any]:
    """
    Calculate comparative metrics across segments:
    - Employee vs Driver apps
    - Android vs iOS
    - MoveInSync vs Competitors
    """
    # 1. Employee vs Driver
    employee_revs = [r for r in reviews if r.get("app_role") == "employee"]
    driver_revs = [r for r in reviews if r.get("app_role") == "driver"]
    competitor_revs = [r for r in reviews if r.get("app_role") == "competitor"]

    # 2. Android vs iOS
    android_revs = [r for r in reviews if r.get("platform") == "Android"]
    ios_revs = [r for r in reviews if r.get("platform") == "iOS"]

    def _quick_stats(rev_sublist: List[Dict[str, Any]]) -> Dict[str, Any]:
        cnt = len(rev_sublist)
        if cnt == 0:
            return {"count": 0, "avg_rating": None, "negative_pct": 0.0, "positive_pct": 0.0}
        ratings = [r.get("rating", 0) for r in rev_sublist if r.get("rating")]
        avg = round(sum(ratings) / len(ratings), 2) if ratings else 0.0
        neg = sum(1 for r in ratings if r in (1, 2))
        pos = sum(1 for r in ratings if r in (4, 5))
        return {
            "count": cnt,
            "avg_rating": avg,
            "negative_pct": round(neg / cnt * 100, 1),
            "positive_pct": round(pos / cnt * 100, 1),
        }

    return {
        "employee_vs_driver": {
            "employee": _quick_stats(employee_revs),
            "driver": _quick_stats(driver_revs),
        },
        "platform_comparison": {
            "android": _quick_stats(android_revs),
            "ios": _quick_stats(ios_revs),
        },
        "moveinsync_vs_competitors": {
            "moveinsync": _quick_stats([r for r in reviews if r.get("app_role") in ("employee", "driver")]),
            "competitors": _quick_stats(competitor_revs),
        },
    }


def extract_issue_themes(reviews: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
    """
    Extract frequency of recurring complaint themes and positive themes from reviews.
    Uses rule-based regex keyword matching for transparent, auditable classification.
    """
    total_reviews = len(reviews)
    if total_reviews == 0:
        return []

    theme_matches: Dict[str, List[Dict[str, Any]]] = defaultdict(list)

    for r in reviews:
        text = f"{r.get('review_title') or ''} {r.get('review_text') or ''}".lower()
        rating = int(r.get("rating") or 0)

        for theme in THEME_DEFINITIONS:
            # Check rating constraints
            if "max_rating" in theme and rating > theme["max_rating"]:
                continue
            if "min_rating" in theme and rating < theme["min_rating"]:
                continue

            matched = False
            for pat in theme["keywords"]:
                if re.search(pat, text, re.IGNORECASE):
                    matched = True
                    break

            if matched:
                theme_matches[theme["name"]].append(r)

    results = []
    for theme in THEME_DEFINITIONS:
        name = theme["name"]
        matched_revs = theme_matches.get(name, [])
        count = len(matched_revs)
        pct = round(count / total_reviews * 100, 1)

        ratings = [r.get("rating", 0) for r in matched_revs if r.get("rating")]
        avg_rating = round(sum(ratings) / len(ratings), 2) if ratings else None

        # Sample up to 5 representative review IDs
        rep_ids = ", ".join([str(r.get("review_id", "")) for r in matched_revs[:5]])

        # Sample up to 3 excerpts (cleaned and truncated to 120 chars)
        excerpts = []
        for r in matched_revs[:3]:
            txt = (r.get("review_text") or "").replace("\n", " ").strip()
            if len(txt) > 120:
                txt = txt[:117] + "..."
            if txt:
                excerpts.append(f'"{txt}"')
        excerpt_str = " | ".join(excerpts) if excerpts else "None"

        results.append({
            "theme": name,
            "type": theme["type"],
            "review_count": count,
            "percentage_of_total": pct,
            "average_rating": avg_rating,
            "representative_review_ids": rep_ids,
            "example_excerpts": excerpt_str,
            "classification_method": "Automated regex rule-based match (manually auditable)",
        })

    # Sort themes by review count descending
    results.sort(key=lambda x: x["review_count"], reverse=True)
    return results
