"""
Complaint & Issue Classifier Module
Implements a 11-category taxonomy with multi-label support, evidence tracing,
heuristic confidence estimation, Unclassified handling, and F1 evaluation.
"""

from typing import List, Dict, Any, Tuple, Optional, Set
import re
from collections import defaultdict


# Configurable 11-category taxonomy with domain-specific keyword/regex patterns
TAXONOMY_CONFIG = [
    {
        "theme": "Cab delays and no-shows",
        "description": "Late cab arrivals, delayed pickups, missed shifts, driver cancellations, or no-shows",
        "patterns": [
            r"\bdelay\b", r"\bdelayed\b", r"\blate\b", r"\bwaiting\b", r"\bwait\b",
            r"\bno[- ]show\b", r"\bcancelled\b", r"\bcanceled\b", r"\bcancellation\b",
            r"\bpickup time\b", r"\bdrop time\b", r"\bmissed cab\b", r"\bcab not arrive\b",
            r"\bwaiting for cab\b", r"\blate night\b", r"\breach late\b", r"\bshift timing\b",
        ],
    },
    {
        "theme": "Booking and OTP failures",
        "description": "Issues generating, receiving, or verifying trip OTPs, booking errors, and confirmation glitches",
        "patterns": [
            r"\botp\b", r"\bbooking\b", r"\bbook cab\b", r"\bcan'?t book\b", r"\bunable to book\b",
            r"\bconfirmation\b", r"\botp not received\b", r"\bwrong otp\b", r"\bverification code\b",
            r"\btrip otp\b", r"\bboard otp\b", r"\bstart trip otp\b", r"\bgenerate otp\b",
        ],
    },
    {
        "theme": "GPS and location tracking",
        "description": "Inaccurate live tracking, incorrect map location, GPS route errors, or lost driver navigation",
        "patterns": [
            r"\bgps\b", r"\blocation\b", r"\blive tracking\b", r"\bmap\b", r"\bmaps\b",
            r"\bwrong location\b", r"\bnavigation\b", r"\btracking\b", r"\bwrong route\b",
            r"\broute\b", r"\bnot moving\b", r"\blocation wrong\b", r"\bdriver location\b",
        ],
    },
    {
        "theme": "App crashes and performance",
        "description": "Application force closures, app hanging, freezing, black screens, or slow responsiveness",
        "patterns": [
            r"\bcrash\b", r"\bcrashes\b", r"\bcrashing\b", r"\bfreeze\b", r"\bfrozen\b",
            r"\bfreezing\b", r"\bhang\b", r"\bhanging\b", r"\bforce close\b", r"\bblack screen\b",
            r"\bwhite screen\b", r"\bstopped working\b", r"\bnot opening\b", r"\bunresponsive\b",
            r"\bbug\b", r"\bbugs\b", r"\bglitch\b", r"\bglitches\b", r"\blag\b", r"\blagging\b",
        ],
    },
    {
        "theme": "Login and account access",
        "description": "Inability to sign in, account lockouts, password resets, session expiration, or registration issues",
        "patterns": [
            r"\blogin\b", r"\blog in\b", r"\bsign in\b", r"\bsignin\b", r"\bpassword\b",
            r"\baccount\b", r"\bregister\b", r"\bregistration\b", r"\bcannot log\b",
            r"\bunable to log\b", r"\blogout\b", r"\blog out\b", r"\bsession expire\b",
            r"\blogin issue\b", r"\bauth\b", r"\bauthentication\b",
        ],
    },
    {
        "theme": "UI and usability",
        "description": "Confusing interface, unintuitive buttons, difficult navigation, or cluttered screen layout",
        "patterns": [
            r"\bui\b", r"\bux\b", r"\binterface\b", r"\bconfusing\b", r"\bcomplicated\b",
            r"\buser[- ]friendly\b", r"\bnavigation\b", r"\bbutton\b", r"\bfont\b",
            r"\bdesign\b", r"\bbad update\b", r"\bworse update\b", r"\bhard to use\b",
            r"\bdifficult to use\b", r"\bnot user friendly\b",
        ],
    },
    {
        "theme": "Roster and scheduling",
        "description": "Shift rostering, ad-hoc scheduling, roster cutoff, trip cancellation, or schedule modification",
        "patterns": [
            r"\broster\b", r"\brostering\b", r"\bschedule\b", r"\bscheduling\b",
            r"\bshift\b", r"\bshifts\b", r"\bad[- ]hoc\b", r"\bweekly roster\b",
            r"\bmonthly roster\b", r"\bcutoff\b", r"\bcut off\b", r"\bchange shift\b",
            r"\bweek off\b", r"\bcalendar\b", r"\bmodify trip\b",
        ],
    },
    {
        "theme": "Driver-side experience",
        "description": "Driver partner complaints: duty allocation, passenger no-show, toll/parking, driver navigation, or app lock",
        "patterns": [
            r"\bdriver app\b", r"\bduty\b", r"\bduties\b", r"\btrip sheet\b",
            r"\bpassenger\b", r"\bpassengers\b", r"\bpartner app\b", r"\bdriver partner\b",
            r"\bdriver behavior\b", r"\brude driver\b", r"\brash driving\b", r"\bpartner\b",
            r"\bdriver duty\b", r"\ballocation\b", r"\baccept trip\b", r"\bstart duty\b",
        ],
    },
    {
        "theme": "Billing and payment",
        "description": "Fare calculations, deductions, trip payment, claims, toll reimbursement, or payout discrepancies",
        "patterns": [
            r"\bbilling\b", r"\bbill\b", r"\bpayment\b", r"\bpayments\b", r"\bmoney\b",
            r"\bdeduction\b", r"\bdeducted\b", r"\bfare\b", r"\bpayout\b", r"\bpayouts\b",
            r"\bcharge\b", r"\bcharged\b", r"\bcharges\b", r"\breimbursement\b", r"\btoll\b",
        ],
    },
    {
        "theme": "Customer support",
        "description": "Helpdesk responsiveness, unresolved tickets, helpline numbers, or escalation grievances",
        "patterns": [
            r"\bsupport\b", r"\bcustomer care\b", r"\bhelpdesk\b", r"\bhelpline\b",
            r"\bticket\b", r"\bno response\b", r"\bunresolved\b", r"\bescalation\b",
            r"\bexecutive\b", r"\bcall center\b", r"\bcomplaint\b", r"\bcare team\b",
        ],
    },
    {
        "theme": "Other or unclear",
        "description": "Vague commentary, non-specific low ratings, or feedback not mapping to specific functional areas",
        "patterns": [
            r"\bbad\b", r"\bworst\b", r"\bterrible\b", r"\bhorrible\b", r"\buseless\b",
            r"\bwaste\b", r"\bpathetic\b", r"\bbakwas\b", r"\bghatiya\b", r"\bpoor\b",
        ],
    },
]


class ReviewClassifier:
    """
    Multi-label review classifier supporting keyword/regex patterns,
    evidence tracing, heuristic confidence, and an explicit Unclassified state.
    """

    def __init__(self, taxonomy: Optional[List[Dict[str, Any]]] = None):
        self.taxonomy = taxonomy or TAXONOMY_CONFIG

    def classify_text(self, text: Optional[str], rating: Optional[int] = None) -> Dict[str, Any]:
        """
        Classify a single review text string into 0 or more themes.

        Returns:
            Dict containing:
                - 'themes': List[str] assigned theme names
                - 'method': str ('Keyword-Rule-Based')
                - 'evidence': Dict[str, List[str]] matched tokens per theme
                - 'heuristic_confidence': str ('High', 'Medium', 'Low', 'None')
                - 'is_unclassified': bool
                - 'audit_note': str explanation
        """
        cleaned_text = str(text or "").strip()

        # Handle empty or insufficient text
        if not cleaned_text or len(cleaned_text) < 4:
            return {
                "themes": ["Unclassified"],
                "method": "Keyword-Rule-Based",
                "evidence": {},
                "heuristic_confidence": "None",
                "is_unclassified": True,
                "audit_note": "Insufficient or missing review text; no defect invented.",
            }

        text_lower = cleaned_text.lower()
        matched_themes: List[str] = []
        evidence_dict: Dict[str, List[str]] = {}
        total_tokens = 0

        # Evaluate the 10 specific functional themes first
        for item in self.taxonomy:
            theme_name = item["theme"]
            if theme_name == "Other or unclear":
                continue  # Handled as fallback

            matched_tokens: List[str] = []
            for pattern in item["patterns"]:
                matches = re.findall(pattern, text_lower, re.IGNORECASE)
                if matches:
                    matched_tokens.extend([m.lower() for m in matches])

            if matched_tokens:
                # Remove duplicate tokens while preserving order
                unique_tokens = list(dict.fromkeys(matched_tokens))
                matched_themes.append(theme_name)
                evidence_dict[theme_name] = unique_tokens
                total_tokens += len(unique_tokens)

        # If no specific functional theme matched
        if not matched_themes:
            # Check for generic complaint keywords
            other_item = next((item for item in self.taxonomy if item["theme"] == "Other or unclear"), None)
            matched_other_tokens = []
            if other_item:
                for pattern in other_item["patterns"]:
                    matches = re.findall(pattern, text_lower, re.IGNORECASE)
                    if matches:
                        matched_other_tokens.extend([m.lower() for m in matches])

            if matched_other_tokens:
                unique_other = list(dict.fromkeys(matched_other_tokens))
                return {
                    "themes": ["Other or unclear"],
                    "method": "Keyword-Rule-Based",
                    "evidence": {"Other or unclear": unique_other},
                    "heuristic_confidence": "Medium",
                    "is_unclassified": False,
                    "audit_note": "Generic sentiment detected without specific defect details.",
                }
            else:
                return {
                    "themes": ["Unclassified"],
                    "method": "Keyword-Rule-Based",
                    "evidence": {},
                    "heuristic_confidence": "None",
                    "is_unclassified": True,
                    "audit_note": "No matching taxonomy patterns found in text.",
                }

        # Calculate heuristic confidence (Note: heuristic score, NOT statistical probability)
        # High: 3+ matched tokens or 2+ distinct specific categories
        # Medium: 2 matched tokens
        # Low: 1 matched token
        if total_tokens >= 3 or len(matched_themes) >= 2:
            confidence = "High"
        elif total_tokens == 2:
            confidence = "Medium"
        else:
            confidence = "Low"

        return {
            "themes": matched_themes,
            "method": "Keyword-Rule-Based",
            "evidence": evidence_dict,
            "heuristic_confidence": confidence,
            "is_unclassified": False,
            "audit_note": f"Matched {len(matched_themes)} themes via {total_tokens} tokens.",
        }


def evaluate_classifier(
    ground_truth_labels: List[List[str]],
    predicted_labels: List[List[str]],
    taxonomy_names: Optional[List[str]] = None,
) -> Dict[str, Any]:
    """
    Calculate per-theme and macro-averaged Precision, Recall, and F1 Score
    given a list of true labels and predicted labels for a validation sample.
    """
    if len(ground_truth_labels) != len(predicted_labels):
        raise ValueError("Length of ground truth and predicted labels must match.")

    if not taxonomy_names:
        taxonomy_names = [t["theme"] for t in TAXONOMY_CONFIG] + ["Unclassified"]

    stats: Dict[str, Dict[str, int]] = {
        name: {"tp": 0, "fp": 0, "fn": 0, "tn": 0} for name in taxonomy_names
    }

    n_samples = len(ground_truth_labels)

    for true_set, pred_set in zip(ground_truth_labels, predicted_labels):
        true_themes = set(true_set)
        pred_themes = set(pred_set)

        for name in taxonomy_names:
            in_true = name in true_themes
            in_pred = name in pred_themes

            if in_true and in_pred:
                stats[name]["tp"] += 1
            elif not in_true and in_pred:
                stats[name]["fp"] += 1
            elif in_true and not in_pred:
                stats[name]["fn"] += 1
            else:
                stats[name]["tn"] += 1

    per_theme_results: Dict[str, Dict[str, Any]] = {}
    precisions = []
