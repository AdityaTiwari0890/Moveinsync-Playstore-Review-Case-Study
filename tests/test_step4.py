import os
import sqlite3
import tempfile
import unittest
from unittest.mock import ANY, patch

import pandas as pd

from ratings_watch.database import close_all_connections, init_db, save_rating_snapshot, upsert_alert
from ratings_watch.metrics import get_rating_summary, get_negative_review_share
from ratings_watch.monitor import import_existing_workbook
from ratings_watch.alert_engine import collect_alerts
from ratings_watch.llm_assistant import LLMAssistant
from ratings_watch.normalization import normalize_review
from ratings_watch.collection_service import CollectionService
from ratings_watch.review_sources import AppleAppSource
from ratings_watch.theme_classifier import RatingsWatchClassifier


class Step4Tests(unittest.TestCase):
    def setUp(self):
        self.tempdir = tempfile.TemporaryDirectory()
        self.db_path = os.path.join(self.tempdir.name, "ratings_watch.db")
        init_db(self.db_path)

    def tearDown(self):
        close_all_connections()
        self.tempdir.cleanup()

    def test_database_init_and_import(self):
        workbook = os.path.join(self.tempdir.name, "sample.xlsx")
        df = pd.DataFrame([
            {"App Name": "MoveInSync", "Platform": "Android", "Review ID": "r1", "Rating": 5, "Review Text": "good app", "Review Date": "2026-01-05", "App Version": "1.0"},
            {"App Name": "MoveInSync", "Platform": "Android", "Review ID": "r1", "Rating": 5, "Review Text": "good app", "Review Date": "2026-01-05", "App Version": "1.0"},
            {"App Name": "MoveInSync", "Platform": "Android", "Review ID": "r2", "Rating": 1, "Review Text": "bad app", "Review Date": "2026-01-06", "App Version": "1.1"},
        ])
        with pd.ExcelWriter(workbook) as writer:
            df.to_excel(writer, sheet_name="Combined_Reviews", index=False)

        result = import_existing_workbook(workbook, self.db_path)
        self.assertEqual(result["inserted"], 2)
        conn = sqlite3.connect(self.db_path)
        total = conn.execute("SELECT COUNT(*) FROM reviews").fetchone()[0]
        self.assertEqual(total, 2)
        conn.close()

    def test_rating_metrics_and_negative_share(self):
        conn = sqlite3.connect(self.db_path)
        conn.execute("INSERT INTO reviews (review_id, platform, app_name, rating, review_text, review_date, first_collected_at, last_updated_at) VALUES (?, ?, ?, ?, ?, ?, ?, ?)", ("a1", "Android", "MoveInSync", 5, "good", "2026-01-05", "2026-01-01T00:00:00Z", "2026-01-01T00:00:00Z"))
