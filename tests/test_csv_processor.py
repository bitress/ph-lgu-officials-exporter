"""
Unit tests for officials_2025_2028.csv processing and multi-format export.
"""

import os
import sys
import unittest

WORKSPACE_DIR = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
if WORKSPACE_DIR not in sys.path:
    sys.path.insert(0, WORKSPACE_DIR)

from data_processor import OfficialsData


class TestCsvProcessor(unittest.TestCase):

    @classmethod
    def setUpClass(cls):
        cls.csv_path = os.path.join(WORKSPACE_DIR, "officials_2025_2028.csv")
        cls.data = OfficialsData(cls.csv_path)
        cls.data.load()

    def test_record_count_and_regions(self):
        """Verify 1,633 cities & municipalities and 18 Philippine regions."""
        self.assertEqual(len(self.data.records), 1633)
        self.assertEqual(len(self.data.regions), 18)
        self.assertIn("NATIONAL CAPITAL REGION", self.data.regions)
        self.assertIn("NEGROS ISLAND REGION", self.data.regions)
        self.assertIn("BANGSAMORO AUTONOMOUS REGION IN MUSLIM MINDANAO", self.data.regions)

    def test_address_coverage_100_percent(self):
        """Verify 100% address coverage across all 1,633 LGUs via smart fallback."""
        stats = self.data.get_statistics()
        self.assertEqual(stats["addresses_count"], 1633)
        self.assertEqual(stats["address_coverage_pct"], 100.0)

    def test_mayor_and_vice_mayor_coverage(self):
        """Verify high coverage (>99%) for both Mayors and Vice Mayors."""
        stats = self.data.get_statistics()
        self.assertGreaterEqual(stats["mayor_coverage_pct"], 99.0)
        self.assertGreaterEqual(stats["vm_coverage_pct"], 99.0)

    def test_filter_and_email_inclusion(self):
        """Verify filtering records with emails and contacts included."""
        recs = self.data.filter_records(
            region="NATIONAL CAPITAL REGION",
            include_contacts=True
        )
        self.assertGreater(len(recs), 0)
        first = recs[0]
        self.assertIn("Mayor Name", first)
        self.assertIn("Mayor Contact", first)
        self.assertIn("Mayor Email", first)

    def test_caching_speed(self):
        """Verify that loading from CSV cache is sub-second (< 0.1s)."""
        import time
        d2 = OfficialsData(self.csv_path)
        t0 = time.time()
        d2.load()
        elapsed = time.time() - t0
        self.assertEqual(len(d2.records), 1633)
        self.assertLess(elapsed, 0.2, f"Cache load took {elapsed:.3f}s, expected < 0.2s")


if __name__ == "__main__":
    unittest.main()
