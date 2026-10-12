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
    review_source: str = "import"
    collection_timestamp: Optional[str] = None
    store_id: Optional[str] = None
    theme: Optional[str] = None
    theme_confidence: str = "unknown"
    extra: Dict[str, Any] = field(default_factory=dict)


@dataclass
class RatingSnapshot:
    snapshot_id: Optional[int] = None
    platform: str = "unknown"
    app_identifier: str = ""
    official_rating: Optional[float] = None
    source: str = "manual"
    observed_at: Optional[str] = None
    metadata: Dict[str, Any] = field(default_factory=dict)

