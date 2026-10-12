"""
ratings_watch/normalization.py
Map incoming review dicts (from scrapers or Excel import) to DB schema.
"""

from __future__ import annotations
import json
import re
from datetime import datetime, timezone
from typing import Any, Dict, List, Optional


def _clean_str(v: Any) -> Optional[str]:
    """Return stripped string or None."""
