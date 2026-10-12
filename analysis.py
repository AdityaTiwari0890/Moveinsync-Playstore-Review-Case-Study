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
