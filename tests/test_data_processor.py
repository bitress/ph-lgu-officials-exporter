"""
Unit tests for data_processor.py
Tests parsing of Officials_List.xlsx, caching, data cleaning, address fallback,
multi-format exports (CSV, XLSX, JSON), and rolling backup creation.
"""

import os
import sys
import time
import json
import unittest

# Ensure workspace root is in sys.path
WORKSPACE_DIR = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
if WORKSPACE_DIR not in sys.path:
    sys.path.insert(0, WORKSPACE_DIR)

from data_processor import (
    OfficialsData,
    clean_token,
    clean_suffix,
    clean_address,
    clean_text_encoding,
    format_official_name,
    format_address_title_case,
    title_case_text
)


class TestDataProcessor(unittest.TestCase):

    @classmethod
    def setUpClass(cls):
        cls.excel_path = os.path.join(WORKSPACE_DIR, "officials_2025_2028.xlsx")
        cls.data = OfficialsData(cls.excel_path)
        cls.data.load()

    def test_record_count_and_regions(self):
        """Verify 1,633 LGUs (1,484 municipalities and 149 cities) across 18 Philippine regions."""
        self.assertEqual(len(self.data.records), 1633)
        self.assertEqual(len(self.data.regions), 18)
        self.assertIn("NATIONAL CAPITAL REGION", self.data.regions)
        self.assertIn("NEGROS ISLAND REGION", self.data.regions)

    def test_lgu_type_breakdown_and_filtering(self):
        """Verify correct classification and filtering of Municipalities (1,484) vs Cities (149)."""
        stats = self.data.get_statistics()
        self.assertEqual(stats["total_lgus"], 1633)
        self.assertEqual(stats["total_municipalities"], 1484)
        self.assertEqual(stats["total_cities"], 149)

        cities = self.data.filter_records(lgu_type="city")
        self.assertEqual(len(cities), 149)

        munis = self.data.filter_records(lgu_type="municipality")
        self.assertEqual(len(munis), 1484)

        all_recs = self.data.filter_records(lgu_type="all")
        self.assertEqual(len(all_recs), 1633)

        # NCR has 16 Cities and 1 Municipality (Pateros)
        ncr_cities = self.data.filter_records(region="NCR", lgu_type="city")
        self.assertEqual(len(ncr_cities), 16)
        ncr_munis = self.data.filter_records(region="NCR", lgu_type="municipality")
        self.assertEqual(len(ncr_munis), 1)
        self.assertEqual(ncr_munis[0]["Town"], "Pateros")

    def test_address_coverage_100_percent(self):
        """Verify 100% address coverage across all 1,633 LGUs via fallback."""
        stats = self.data.get_statistics()
        self.assertEqual(stats["addresses_count"], 1633)
        self.assertEqual(stats["address_coverage_pct"], 100.0)

    def test_caching_speed(self):
        """Verify that loading from cache is sub-second (typically < 0.1s)."""
        d2 = OfficialsData(self.excel_path)
        t0 = time.time()
        d2.load()
        elapsed = time.time() - t0
        self.assertEqual(len(d2.records), 1633)
        self.assertLess(elapsed, 0.5, f"Cache load took {elapsed:.3f}s, expected < 0.5s")

    def test_clean_suffix(self):
        """Test suffix normalization and removal of junk tokens."""
        # Clean standard suffixes
        self.assertEqual(clean_suffix("JR"), "Jr.")
        self.assertEqual(clean_suffix("jr."), "Jr.")
        self.assertEqual(clean_suffix("JR,"), "Jr.")
        self.assertEqual(clean_suffix("JR.I"), "Jr.")
        self.assertEqual(clean_suffix("SR"), "Sr.")
        self.assertEqual(clean_suffix("III"), "III")
        self.assertEqual(clean_suffix("iv"), "IV")
        self.assertEqual(clean_suffix("JR IV"), "Jr. IV")
        self.assertEqual(clean_suffix("jr-ii"), "Jr. II")

        # Invalid junk suffixes should become empty
        self.assertEqual(clean_suffix("N/A"), "")
        self.assertEqual(clean_suffix("N\\A"), "")
        self.assertEqual(clean_suffix("N/A/"), "")
        self.assertEqual(clean_suffix("''N/A"), "")
        self.assertEqual(clean_suffix("-"), "")
        self.assertEqual(clean_suffix("`"), "")
        self.assertEqual(clean_suffix("MIR"), "")
        self.assertEqual(clean_suffix("MR"), "")
        self.assertEqual(clean_suffix("DR."), "")
        self.assertEqual(clean_suffix("TAN"), "")
        self.assertEqual(clean_suffix(None), "")

    def test_clean_text_encoding_mojibake(self):
        """Test encoding and mojibake repairs."""
        self.assertEqual(clean_text_encoding("Dasmari\ufffdas"), "Dasmariñas")
        self.assertEqual(clean_text_encoding("ParaÃ±aque"), "Parañaque")
        self.assertEqual(clean_text_encoding("PE\ufffdARRUBIA"), "PEÑARRUBIA")

    def test_format_address_title_case(self):
        """Test address title casing and preservation of LGU terms."""
        raw = "LGU-MALUNGON, POBLACION, MALUNGON, SARANGANI"
        formatted = format_address_title_case(raw)
        self.assertIn("LGU-Malungon", formatted)
        self.assertIn("Poblacion", formatted)

        raw2 = "BARANGAY HALL,, POBLACION NATL HIGHWAY"
        formatted2 = format_address_title_case(raw2)
        self.assertNotIn(",,", formatted2)
        self.assertIn("Brgy.", formatted2)
        self.assertIn("Nat'l", formatted2)

    def test_format_official_name(self):
        """Test name formatting with titles, initials, and casing styles."""
        # Title case
        name = format_official_name("JUAN", "DELA CRUZ", "SANTOS", "JR.")
        self.assertEqual(name, "Juan D. Santos Jr.")

        # Full middle name
        name_full = format_official_name("JUAN", "DELA CRUZ", "SANTOS", "JR.", middle_as_initial=False)
        self.assertEqual(name_full, "Juan Dela Cruz Santos Jr.")

        # Uppercase
        name_upper = format_official_name("JUAN", "DELA CRUZ", "SANTOS", "JR.", name_case="upper")
        self.assertEqual(name_upper, "JUAN D. SANTOS JR.")

        # Dirty suffix ignored
        name_clean_suf = format_official_name("JOHNNY", "ATIENZA", "SAMBO", "N\\A")
        self.assertEqual(name_clean_suf, "Johnny A. Sambo")

    def test_export_csv_utf8_sig(self):
        """Verify CSV export includes UTF-8 BOM and correct headers."""
        out_csv = os.path.join(WORKSPACE_DIR, "tests", "test_output.csv")
        recs = self.data.filter_records(region="REGION 3", province="PAMPANGA")
        self.assertGreater(len(recs), 0)

        count = self.data.export_csv(out_csv, recs, include_region=True, include_contacts=True)
        self.assertEqual(count, len(recs))

        # Check BOM
        with open(out_csv, "rb") as f:
            header_bytes = f.read(3)
            self.assertEqual(header_bytes, b'\xef\xbb\xbf')

        # Clean up
        if os.path.exists(out_csv):
            os.remove(out_csv)

    def test_export_xlsx(self):
        """Verify Excel export creates valid styled workbook."""
        out_xlsx = os.path.join(WORKSPACE_DIR, "tests", "test_output.xlsx")
        recs = self.data.filter_records(region="NCR")
        count = self.data.export_xlsx(out_xlsx, recs, include_region=True)
        self.assertEqual(count, len(recs))

        import openpyxl
        wb = openpyxl.load_workbook(out_xlsx, read_only=True)
        ws = wb.active
        self.assertEqual(ws.title, "Municipal Officials")
        rows = list(ws.iter_rows(values_only=True))
        self.assertEqual(len(rows), len(recs) + 1)
        self.assertIn("Mayor Name", rows[0])
        self.assertIn("Region", rows[0])
        wb.close()

        # Clean up
        if os.path.exists(out_xlsx):
            os.remove(out_xlsx)

    def test_export_json(self):
        """Verify JSON export creates parseable array with requested fields."""
        out_json = os.path.join(WORKSPACE_DIR, "tests", "test_output.json")
        recs = self.data.filter_records(region="CAR")
        count = self.data.export_json(out_json, recs, include_region=True, include_contacts=True)
        self.assertEqual(count, len(recs))

        with open(out_json, "r", encoding="utf-8") as f:
            data = json.load(f)
        self.assertEqual(len(data), len(recs))
        self.assertIn("Mayor Name", data[0])
        self.assertIn("Mayor Contact", data[0])

        # Clean up
        if os.path.exists(out_json):
            os.remove(out_json)

    def test_export_xlsx_with_verification(self):
        """Verify Excel export includes Contact Verified column and green row highlights for verified LGUs."""
        out_xlsx = os.path.join(WORKSPACE_DIR, "tests", "test_verified_output.xlsx")
        recs = self.data.filter_records(region="NCR")
        # In NCR, Province is recorded as 'Pateros' for Pateros
        vmap = {"Pateros|Pateros": "2026-10-09T10:00:00"}
        count = self.data.export_xlsx(
            out_xlsx, recs,
            include_region=True,
            include_verified_col=True,
            verification_map=vmap
        )
        self.assertEqual(count, len(recs))

        import openpyxl
        wb = openpyxl.load_workbook(out_xlsx)
        ws = wb.active
        headers = [c.value for c in ws[1]]
        self.assertIn("Contact Verified", headers)
        v_idx = headers.index("Contact Verified")

        # Find Pateros row
        pateros_found = False
        for row in ws.iter_rows(min_row=2):
            vals = [c.value for c in row]
            if "Pateros" in vals:
                pateros_found = True
                self.assertEqual(vals[v_idx], "Yes")
                # Check green fill
                self.assertIsNotNone(row[0].fill)
                self.assertIsNotNone(row[0].fill.start_color)
                self.assertIn("D8F4D8", str(row[0].fill.start_color.rgb))
            else:
                self.assertEqual(vals[v_idx], "No")

        self.assertTrue(pateros_found)
        wb.close()

        if os.path.exists(out_xlsx):
            os.remove(out_xlsx)

    def test_export_csv_with_verification(self):
        """Verify CSV export includes Contact Verified column."""
        out_csv = os.path.join(WORKSPACE_DIR, "tests", "test_verified_output.csv")
        recs = self.data.filter_records(region="NCR")
        vmap = {"Pateros|Pateros": "2026-10-09T10:00:00"}
        count = self.data.export_csv(
            out_csv, recs,
            include_verified_col=True,
            verification_map=vmap
        )
        self.assertEqual(count, len(recs))

        with open(out_csv, "r", encoding="utf-8-sig") as f:
            lines = f.readlines()
        self.assertIn("Contact Verified", lines[0])

        if os.path.exists(out_csv):
            os.remove(out_csv)


if __name__ == "__main__":
    unittest.main()

