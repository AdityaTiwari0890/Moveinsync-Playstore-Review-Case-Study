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
            result = bulk_upsert_reviews(classified, run_id=run_id, is_imported=False, db_path=self.db_path)
            summary.update(result)
            summary["failed"] += summary["rejected"]
            summary["status"] = (
                "partial"
                if summary["rejected"] or result.get("failed") or summary["source_errors_observed"] or not summary["fetched"]
                else "success"
            )
            if not summary["fetched"]:
                summary["source_message"] = "The source adapter returned no reviews; verify the app ID, availability and collection logs."

            try:
                from metadata import get_app_store_metadata, get_play_store_metadata

                metadata = (
                    get_play_store_metadata(app_id, country=kwargs.get("country", "in"), lang=kwargs.get("lang", "en"))
                    if platform_key == "android"
                    else get_app_store_metadata(app_id, country=kwargs.get("country", "in"))
                )
                if metadata.get("status") == "SUCCESS" and metadata.get("store_rating") is not None:
                    snapshot_id = save_rating_snapshot(
                        platform="Android" if platform_key == "android" else "iOS",
                        store_id=app_id,
                        app_name=app_name,
                        official_rating=metadata["store_rating"],
                        rating_source=f"{platform_key}_public_store_metadata",
                        total_ratings_count=metadata.get("store_ratings_count"),
                        run_id=run_id,
                        metadata={"store_url": metadata.get("store_url"), "source_status": metadata.get("status")},
                        db_path=self.db_path,
                    )
                    summary["rating_snapshot_id"] = snapshot_id
                    summary["rating_snapshot_status"] = "recorded_from_public_store_metadata"
                else:
                    summary["rating_snapshot_status"] = "unavailable"
                    summary["rating_snapshot_message"] = metadata.get("limitation_note") or "Source did not return a verified published rating."
            except Exception as metadata_error:
                summary["rating_snapshot_status"] = "unavailable"
                summary["rating_snapshot_message"] = f"Store metadata lookup failed: {metadata_error}"

            finish_collection_run(
                run_id,
                summary["status"],
                summary,
                notes=f"Collected {summary['fetched']} source reviews; store rating snapshot: {summary.get('rating_snapshot_status', 'unavailable')}",
                db_path=self.db_path,
            )
            return summary
        except Exception as exc:
            summary["failed"] = max(summary["failed"], 1)
            summary["status"] = "failed"
            summary["error"] = str(exc)
            finish_collection_run(run_id, "failed", summary, error=str(exc), db_path=self.db_path)
            return summary
