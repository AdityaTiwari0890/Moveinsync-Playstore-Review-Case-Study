"""
Automated Test Suite for Step 2 App Review Analysis Engine
Validates:
1. Percentage calculations and rounding precision
2. Wilson score confidence interval boundaries
3. Multiple-theme counting and overlapping share mathematics
4. Handling missing fields (ratings, dates, versions, text)
5. Duplicate review detection and deduplication
6. Date grouping and monthly trends
7. Pareto cumulative shares and vital few cutoff logic
8. Classification validation metrics (Precision, Recall, F1 score)
9. Workbook idempotence, sheet preservation, and chart attachment
"""

import unittest
import os
import openpyxl
import pandas as pd

from classifier import ReviewClassifier, evaluate_classifier
from analysis_engine import (
    wilson_confidence_interval,
    load_and_normalize_workbook,
    compute_rating_analysis,
    compute_theme_and_pareto_analysis,
    compute_version_analysis,
    compute_action_priorities,
    enhance_workbook_with_analysis,
)


class TestAnalysisEngine(unittest.TestCase):

    def test_01_wilson_confidence_interval(self):
        """Verify Wilson score confidence interval calculation against hand-calculated bounds."""
        # Test case: 50 successes out of 100 trials (50%)
        # Expected 95% Wilson CI is approx [40.4%, 59.6%]
        ci_low, ci_high = wilson_confidence_interval(50, 100)
        self.assertAlmostEqual(ci_low, 40.4, delta=0.5)
        self.assertAlmostEqual(ci_high, 59.6, delta=0.5)

        # Test case: 0 out of 50 trials (0%)
        ci_low_0, ci_high_0 = wilson_confidence_interval(0, 50)
        self.assertEqual(ci_low_0, 0.0)
        self.assertLess(ci_high_0, 8.0)

        # Edge case: N <= 0
        self.assertEqual(wilson_confidence_interval(0, 0), (0.0, 0.0))

    def test_02_multiple_theme_classification_and_evidence(self):
        """Verify multi-label classification, evidence tracing, and heuristic confidence."""
        clf = ReviewClassifier()

        # Review with both GPS tracking issues and cab delay
        sample_text = "Driver was 30 minutes late and live tracking was totally frozen on the map"
        res = clf.classify_text(sample_text, rating=1)

        self.assertIn("Cab delays and no-shows", res["themes"])
        self.assertIn("GPS and location tracking", res["themes"])
        self.assertFalse(res["is_unclassified"])
        self.assertEqual(res["heuristic_confidence"], "High")
        self.assertIn("Cab delays and no-shows", res["evidence"])
        self.assertIn("GPS and location tracking", res["evidence"])

        # Unclassified test case (insufficient text)
        empty_res = clf.classify_text("", rating=1)
        self.assertIn("Unclassified", empty_res["themes"])
        self.assertTrue(empty_res["is_unclassified"])
        self.assertEqual(empty_res["heuristic_confidence"], "None")

    def test_03_classifier_evaluation_metrics(self):
        """Verify precision, recall, and F1 score against a hand-calculated validation set."""
        # 3 samples:
        # Sample 1: True: ['Cab delays and no-shows'], Pred: ['Cab delays and no-shows'] -> TP
        # Sample 2: True: ['GPS and location tracking'], Pred: ['Cab delays and no-shows'] -> Cab FP, GPS FN
        # Sample 3: True: ['GPS and location tracking'], Pred: ['GPS and location tracking'] -> GPS TP
        truth = [
            ["Cab delays and no-shows"],
            ["GPS and location tracking"],
            ["GPS and location tracking"],
        ]
        pred = [
            ["Cab delays and no-shows"],
            ["Cab delays and no-shows"],
            ["GPS and location tracking"],
        ]

        metrics = evaluate_classifier(truth, pred)
        cab_metrics = metrics["per_theme"]["Cab delays and no-shows"]
        gps_metrics = metrics["per_theme"]["GPS and location tracking"]

        # Cab: TP=1, FP=1, FN=0 -> Precision = 1/(1+1) = 0.5, Recall = 1/(1+0) = 1.0, F1 = 2*0.5*1/1.5 = 0.6667
        self.assertEqual(cab_metrics["true_positives"], 1)
        self.assertEqual(cab_metrics["false_positives"], 1)
        self.assertEqual(cab_metrics["precision"], 0.5)
        self.assertEqual(cab_metrics["recall"], 1.0)
        self.assertAlmostEqual(cab_metrics["f1_score"], 0.6667, delta=0.001)

        # GPS: TP=1, FP=0, FN=1 -> Precision = 1/(1+0) = 1.0, Recall = 1/(1+1) = 0.5, F1 = 0.6667
        self.assertEqual(gps_metrics["true_positives"], 1)
        self.assertEqual(gps_metrics["false_negatives"], 1)
        self.assertEqual(gps_metrics["precision"], 1.0)
        self.assertEqual(gps_metrics["recall"], 0.5)
        self.assertAlmostEqual(gps_metrics["f1_score"], 0.6667, delta=0.001)

        # Macro F1 should be ~0.6667
        self.assertAlmostEqual(metrics["macro_f1_score"], 0.6667, delta=0.001)

    def test_04_pareto_overlapping_shares_and_cutoff(self):
        """Verify that Pareto analysis calculates cumulative shares and identifies vital few."""
        test_df = pd.DataFrame([
            {"review_text": "Late cab delay and wrong gps route", "rating": 1, "app_role": "employee", "platform": "Android"},
            {"review_text": "Delay in pickup and driver rude", "rating": 1, "app_role": "employee", "platform": "Android"},
            {"review_text": "Otp not received and late cab", "rating": 2, "app_role": "employee", "platform": "iOS"},
            {"review_text": "App crash continuously", "rating": 1, "app_role": "driver", "platform": "Android"},
