from __future__ import annotations

from abc import ABC, abstractmethod
from typing import Any, Dict, List


class BaseReviewSource(ABC):
    @abstractmethod
    def fetch_reviews(self, app_id: str, app_name: str, app_role: str, count: int, **kwargs) -> List[Dict[str, Any]]:
        raise NotImplementedError


class GooglePlaySource(BaseReviewSource):
    """Adapter over the project's existing paginated Google Play collector."""

    def __init__(self) -> None:
        self.had_errors = False

    def fetch_reviews(self, app_id: str, app_name: str, app_role: str, count: int, **kwargs) -> List[Dict[str, Any]]:
        from play_store_scraper import scrape_play_store_reviews

        def capture(level: str, _message: str) -> None:
            if level.upper() == "ERROR":
                self.had_errors = True

        results = scrape_play_store_reviews(
            package_id=app_id,
            app_name=app_name,
            app_role=app_role,
            count=count,
            country=kwargs.get("country", "in"),
            lang=kwargs.get("lang", "en"),
            log_callback=capture,
        )
        return results


class AppleAppSource(BaseReviewSource):
    """Adapter over Apple's public customer-review RSS collection already used by Step 1."""

    def __init__(self) -> None:
        self.had_errors = False

    def fetch_reviews(self, app_id: str, app_name: str, app_role: str, count: int, **kwargs) -> List[Dict[str, Any]]:
        from app_store_scraper import scrape_app_store_reviews

        def capture(level: str, _message: str) -> None:
            if level.upper() in {"ERROR", "WARNING"}:
                self.had_errors = True

        results = scrape_app_store_reviews(
            app_id=app_id,
            app_name=app_name,
            app_role=app_role,
            count=count,
            country=kwargs.get("country", "in"),
            log_callback=capture,
        )
        return results
