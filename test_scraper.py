"""
Acceptance & Integration Test Suite for App Review Intelligence Scraper
Tests Play Store scraper, App Store scraper, Metadata collector,
Business Analysis calculations, and Excel 6-sheet export integrity.
"""

import unittest
import os
import datetime
import openpyxl

from play_store_scraper import search_play_store, scrape_play_store_reviews
from app_store_scraper import search_app_store, scrape_app_store_reviews
from metadata import get_play_store_metadata, get_app_store_metadata
from analysis import (
    calculate_sample_summary,
    calculate_monthly_trends,
    calculate_segment_comparisons,
    extract_issue_themes,
)
from excel_exporter import export_intelligence_workbook


class TestAppReviewIntelligenceScraper(unittest.TestCase):

    def test_01_play_store_search(self):
        """Verify Play Store candidate search finds MoveInSync."""
        candidates = search_play_store("MoveInSync", limit=5)
        self.assertGreater(len(candidates), 0, "Play Store search should return candidates")
        app_ids = [c["app_id"] for c in candidates]
        self.assertTrue(
            any("moveinsync" in str(aid).lower() for aid in app_ids),
            f"Expected MoveInSync in search candidates, got: {app_ids}"
        )

    def test_02_play_store_reviews_collection(self):
        """Verify Play Store review scraping retrieves >= 100 reviews, sorted newest first, deduplicated."""
        reviews = scrape_play_store_reviews(
            package_id="com.moveinsync.ets",
            app_name="MoveInSync",
            app_role="employee",
            count=100,
        )
        self.assertGreaterEqual(len(reviews), 100, f"Expected at least 100 reviews, got {len(reviews)}")

        # Check sorting: newest to oldest
        dates = [r["review_date"] for r in reviews if r.get("review_date")]
        self.assertEqual(dates, sorted(dates, reverse=True), "Reviews should be sorted newest first")

        # Check deduplication
        review_ids = [r["review_id"] for r in reviews]
        self.assertEqual(len(review_ids), len(set(review_ids)), "All review IDs must be unique (no duplicates)")

        # Verify schema fields
        first = reviews[0]
        self.assertIn("app_name", first)
        self.assertIn("store_id", first)
        self.assertEqual(first["platform"], "Android")
        self.assertEqual(first["app_role"], "employee")
        self.assertIn("rating", first)
        self.assertIn(first["rating"], [1, 2, 3, 4, 5])
        self.assertIn("review_text", first)
        self.assertIn("collection_timestamp", first)

    def test_03_app_store_search_and_reviews(self):
        """Verify Apple App Store search and reviews collection."""
        candidates = search_app_store("MoveInSync", limit=5)
        self.assertGreater(len(candidates), 0, "App Store search should return candidates")

        # Test scraping with a widely reviewed app (Uber - 368677368) to verify pagination >= 100 reviews
        reviews = scrape_app_store_reviews(
