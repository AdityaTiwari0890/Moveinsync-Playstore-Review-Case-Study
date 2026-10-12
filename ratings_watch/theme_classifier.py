from __future__ import annotations

from typing import Any, Dict, List, Optional

from classifier import ReviewClassifier


class RatingsWatchClassifier:
    """Thin wrapper over the existing ReviewClassifier for Step 4 monitoring."""

    def __init__(self):
