from __future__ import annotations

from collections import Counter
from datetime import datetime
from typing import Any, Dict, List, Optional

import pandas as pd
import sqlite3

from ratings_watch.database import get_connection


def _safe_float(value: Any, default: float = 0.0) -> float:
    try:
        return float(value)
    except (TypeError, ValueError):
