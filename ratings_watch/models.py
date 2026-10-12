from __future__ import annotations

from dataclasses import dataclass, field
from datetime import datetime
from typing import Any, Dict, List, Optional


@dataclass
class ReviewRecord:
    review_id: str
    platform: str
    app_name: str
    app_role: str = "employee"
    rating: int = 0
    review_text: str = ""
    review_title: str = ""
    review_date: Optional[str] = None
    app_version: str = "Unknown"
