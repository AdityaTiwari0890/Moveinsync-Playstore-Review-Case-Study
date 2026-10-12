from __future__ import annotations

from datetime import datetime, timezone
from typing import Any, Dict, List, Optional

from ratings_watch.database import bulk_upsert_reviews, save_rating_snapshot, start_collection_run, finish_collection_run
from ratings_watch.normalization import normalize_reviews
from ratings_watch.review_sources import AppleAppSource, GooglePlaySource
from ratings_watch.theme_classifier import RatingsWatchClassifier


class CollectionService:
    def __init__(self, db_path: str):
        self.db_path = db_path
        self.classifier = RatingsWatchClassifier()

    def collect_platform(self, platform: str, app_id: str, app_name: str = "MoveInSync", count: int = 200, **kwargs) -> Dict[str, Any]:
        run_id = start_collection_run(platform, db_path=self.db_path)
        summary = {"fetched": 0, "inserted": 0, "updated": 0, "skipped": 0, "failed": 0, "status": "running"}
        try:
            platform_key = platform.lower()
            if platform_key == "android":
                source = GooglePlaySource()
            elif platform_key == "ios":
                source = AppleAppSource()
            else:
