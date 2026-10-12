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

