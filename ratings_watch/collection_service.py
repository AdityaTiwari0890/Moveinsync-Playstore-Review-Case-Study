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
                raise ValueError(f"Platform '{platform}' is not supported; choose Android or iOS.")
            source_reviews = source.fetch_reviews(
                app_id,
                app_name=app_name,
                app_role=kwargs.get("app_role", "employee"),
                count=count,
                country=kwargs.get("country", "in"),
                lang=kwargs.get("lang", "en"),
            )
            raw_reviews = source_reviews
            summary["fetched"] = len(raw_reviews)
            summary["source_errors_observed"] = bool(getattr(source, "had_errors", False))
            normalized = normalize_reviews(raw_reviews)
            summary["rejected"] = len(raw_reviews) - len(normalized)
            classified = []
            for review in normalized:
                review.setdefault("app_name", app_name)
                review.setdefault("app_role", "employee")
                review.setdefault("platform", platform)
                review["collection_timestamp"] = datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")
                if review.get("review_text"):
                    classification = self.classifier.classify(review["review_text"], rating=review.get("rating"))
                    review["assigned_themes"] = classification["themes"]
                    review["classification_method"] = classification["classification_method"]
                    review["heuristic_confidence"] = classification["heuristic_confidence"]
                    review["is_unclassified"] = classification["is_unclassified"]
                classified.append(review)
