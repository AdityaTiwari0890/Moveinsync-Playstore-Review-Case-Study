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
