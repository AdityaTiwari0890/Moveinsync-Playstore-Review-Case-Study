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
            app_id="368677368",
            app_name="Uber",
            app_role="competitor",
            count=100,
            country="us",
        )
        self.assertGreaterEqual(len(reviews), 100, f"Expected at least 100 reviews for Uber, got {len(reviews)}")

        # Check sorting: newest to oldest
        dates = [r["review_date"] for r in reviews if r.get("review_date")]
        self.assertEqual(dates, sorted(dates, reverse=True), "iOS Reviews should be sorted newest first")

        # Check deduplication
        review_ids = [r["review_id"] for r in reviews]
        self.assertEqual(len(review_ids), len(set(review_ids)), "All iOS review IDs must be unique")

    def test_04_metadata_retrieval(self):
        """Verify store metadata extraction and limitation reporting."""
        meta_android = get_play_store_metadata("com.moveinsync.ets")
        self.assertEqual(meta_android["status"], "SUCCESS")
        self.assertEqual(meta_android["title"], "MoveInSync")
        self.assertIsNotNone(meta_android["store_rating"])
        self.assertIn("limitation_note", meta_android)

        meta_ios = get_app_store_metadata("962949521")
        self.assertEqual(meta_ios["status"], "SUCCESS")
        self.assertIn("MoveInSync", meta_ios["title"])
        self.assertIsNotNone(meta_ios["store_rating"])

    def test_05_business_analysis(self):
        """Verify business intelligence calculations and theme extraction."""
        sample_reviews = [
            {"rating": 1, "review_title": "Bad app", "review_text": "Frequent crash and login otp error", "review_date": "2026-10-10 10:00:00", "app_role": "employee", "platform": "Android", "review_id": "r1"},
            {"rating": 1, "review_title": "Worst", "review_text": "Late cab delay and wrong gps location", "review_date": "2026-10-09 10:00:00", "app_role": "employee", "platform": "Android", "review_id": "r2"},
            {"rating": 2, "review_title": "Slow", "review_text": "Battery drain and slow ui glitch", "review_date": "2026-09-15 10:00:00", "app_role": "driver", "platform": "Android", "review_id": "r3"},
            {"rating": 5, "review_title": "Great", "review_text": "Very smooth commute and punctual on time", "review_date": "2026-09-10 10:00:00", "app_role": "employee", "platform": "iOS", "review_id": "r4"},
            {"rating": 5, "review_title": "Love it", "review_text": "Easy to use clean ui and safe", "review_date": "2026-08-01 10:00:00", "app_role": "competitor", "platform": "iOS", "review_id": "r5"},
        ]
        summary = calculate_sample_summary(sample_reviews, requested_count=5)
        self.assertEqual(summary["total_reviews"], 5)
        self.assertEqual(summary["target_reached"], "Yes")
        self.assertEqual(summary["avg_rating"], 2.8)
        self.assertEqual(summary["star_1_count"], 2)
        self.assertEqual(summary["negative_pct"], 60.0)
        self.assertEqual(summary["positive_pct"], 40.0)

        # Monthly trends
        monthly = calculate_monthly_trends(sample_reviews)
        self.assertGreater(len(monthly), 0)

        # Segment comparisons
        comps = calculate_segment_comparisons(sample_reviews)
        self.assertIn("employee_vs_driver", comps)
        self.assertIn("platform_comparison", comps)
        self.assertIn("moveinsync_vs_competitors", comps)

        # Theme extraction
        themes = extract_issue_themes(sample_reviews)
        theme_names = [t["theme"] for t in themes if t["review_count"] > 0]
        self.assertTrue(any("Crash" in t for t in theme_names), "Should extract crash theme")
        self.assertTrue(any("Punctual" in t or "Smooth" in t for t in theme_names), "Should extract punctual theme")

    def test_06_excel_export_and_workbook_integrity(self):
        """Verify that export_intelligence_workbook creates a valid 6-sheet workbook with tables."""
        test_file = "test_intelligence_export.xlsx"
        if os.path.exists(test_file):
            os.remove(test_file)

        mock_android = [{
            "app_name": "MoveInSync", "store_id": "com.moveinsync.ets", "platform": "Android",
            "app_role": "employee", "review_id": "rev_and_1", "reviewer_name": "John Doe",
            "rating": 5, "review_text": "Excellent commute service", "review_date": "2026-10-12 12:00:00",
            "app_version": "12.0.0", "developer_reply": None, "developer_reply_date": None,
            "helpful_votes": 3, "collection_timestamp": "2026-10-12 12:05:00",
            "source_url": "https://play.google.com/store/apps/details?id=com.moveinsync.ets"
        }]
        mock_ios = [{
            "app_name": "MoveInSync", "store_id": "962949521", "platform": "iOS",
            "app_role": "employee", "review_id": "rev_ios_1", "reviewer_name": "Jane Doe",
            "rating": 4, "review_title": "Good app", "review_text": "Reliable routing",
            "review_date": "2026-10-11 10:00:00", "app_version": "5.4.1",
            "developer_reply": None, "developer_reply_date": None, "helpful_votes": 1,
            "collection_timestamp": "2026-10-12 12:05:00", "source_url": "https://apps.apple.com/app/id962949521"
        }]
        mock_combined = mock_android + mock_ios
        mock_meta = [get_play_store_metadata("com.moveinsync.ets")]
        mock_sm = [calculate_sample_summary(mock_android, 100)]
        mock_sm[0].update({"app_name": "MoveInSync", "platform": "Android", "app_role": "employee"})
        mock_mt = calculate_monthly_trends(mock_combined)
        mock_sc = calculate_segment_comparisons(mock_combined)
        mock_th = extract_issue_themes(mock_combined)
        mock_logs = [{
            "timestamp": "2026-10-12 12:00:00", "app_name": "MoveInSync", "platform": "Android",
            "store_id": "com.moveinsync.ets", "status": "SUCCESS", "requested_count": 100,
            "retrieved_count": 1, "data_source": "google-play-scraper",
            "message": "Sample test log", "limitation_note": ""
        }]

        export_intelligence_workbook(
            android_reviews=mock_android,
            ios_reviews=mock_ios,
            combined_reviews=mock_combined,
            metadata_list=mock_meta,
            summary_metrics=mock_sm,
            monthly_trends=mock_mt,
            segment_comparisons=mock_sc,
            issue_themes=mock_th,
            collection_logs=mock_logs,
            output_filepath=test_file,
        )

        self.assertTrue(os.path.exists(test_file), "Excel file should exist")

        # Load and verify sheets using openpyxl
        wb = openpyxl.load_workbook(test_file)
        expected_sheets = [
            "Android_Reviews",
            "iOS_Reviews",
            "Combined_Reviews",
            "App_Summary",
            "Issue_Themes",
            "Collection_Log",
        ]
        self.assertEqual(wb.sheetnames, expected_sheets, f"Sheets must match {expected_sheets}")

        # Check table creation on Android_Reviews
        ws_and = wb["Android_Reviews"]
        self.assertGreater(len(ws_and.tables), 0, "Android_Reviews sheet should have an Excel Table")

        # Check table creation on Issue_Themes
        ws_th = wb["Issue_Themes"]
        self.assertGreater(len(ws_th.tables), 0, "Issue_Themes sheet should have an Excel Table")

        wb.close()
        # Clean up test export
        try:
            os.remove(test_file)
        except Exception:
            pass


if __name__ == "__main__":
    unittest.main()
