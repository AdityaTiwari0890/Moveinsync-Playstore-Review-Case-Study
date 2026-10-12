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
