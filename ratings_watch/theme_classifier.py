from __future__ import annotations

from typing import Any, Dict, List, Optional

from classifier import ReviewClassifier


class RatingsWatchClassifier:
    """Thin wrapper over the existing ReviewClassifier for Step 4 monitoring."""

    def __init__(self):
        self._classifier = ReviewClassifier()

    def classify(self, text: Optional[str], rating: Optional[int] = None) -> Dict[str, Any]:
        result = self._classifier.classify_text(text, rating=rating)
        themes = result.get("themes") or []
        if not isinstance(themes, list):
            themes = [themes]
        is_unclassified = bool(result.get("is_unclassified")) or "Unclassified" in themes
        if is_unclassified:
            themes = []
        return {
            "themes": themes,
            "classification_method": result.get("method", "Keyword-Rule-Based"),
            "heuristic_confidence": result.get("heuristic_confidence", "None"),
            "is_unclassified": is_unclassified,
            "evidence": result.get("evidence", {}),
        }

    def classify_reviews(self, reviews: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
        out = []
        for review in reviews:
            text = review.get("review_text") or review.get("Review Text") or ""
            rating = review.get("rating")
            if rating is None:
                rating = review.get("Rating")
            item = self.classify(text, rating)
            merged = dict(review)
            merged["assigned_themes"] = item["themes"]
            merged["classification_method"] = item["classification_method"]
            merged["heuristic_confidence"] = item["heuristic_confidence"]
            merged["is_unclassified"] = item["is_unclassified"]
            out.append(merged)
        return out
