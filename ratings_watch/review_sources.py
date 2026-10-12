from __future__ import annotations

from abc import ABC, abstractmethod
from typing import Any, Dict, List


class BaseReviewSource(ABC):
    @abstractmethod
    def fetch_reviews(self, app_id: str, app_name: str, app_role: str, count: int, **kwargs) -> List[Dict[str, Any]]:
        raise NotImplementedError


class GooglePlaySource(BaseReviewSource):
    """Adapter over the project's existing paginated Google Play collector."""
