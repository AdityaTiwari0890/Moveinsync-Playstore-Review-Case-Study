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
