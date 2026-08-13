"""
Data Loader Extraction Unit Test Suite
--------------------------------------
Description: Dispatches temporary mock data files across various formats (JSON, XLSX)
             to confirm uniform data parsing and column-stripping normalization.

Author: Joseph Adogeri
Version: 1.0.0
Since: 2026-08-13
File: tests/test_data_loader.py
License: MIT
"""

import os
import sys
import json
import unittest
import pandas as pd

# Direct absolute path tracking injection to safely import from src/
sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), "..", "src")))
from src.utils.data_loader import load_raw_records


class TestDataLoader(unittest.TestCase):

    def setUp(self):
        """Builds an isolated execution environment with sample data files."""
        self.test_dir = os.path.abspath(os.path.dirname(__file__))
        self.json_path = os.path.join(self.test_dir, "mock_data.json")
        self.xlsx_path = os.path.join(self.test_dir, "mock_data.xlsx")

        self.sample_data = [
            {"PARCELID": "10020034", "TAX_YEAR": 2026, " REASON ": "Value Correction"},
            {"PARCELID": "50060078", "TAX_YEAR": 2026, " REASON ": "Clerical Error"}
        ]

        # Write out temporary JSON target file
        with open(self.json_path, "w", encoding="utf-8") as f:
            json.dump(self.sample_data, f)

        # Write out temporary XLSX target spreadsheet file using pandas
        df = pd.DataFrame(self.sample_data)
        df.to_excel(self.xlsx_path, index=False, engine="openpyxl")

    def tearDown(self):
        """Wipes mock assets off the local directory to keep the workspace clean."""
        for path in [self.json_path, self.xlsx_path]:
            if os.path.exists(path):
                os.remove(path)

    def test_load_records_from_json(self):
        """Confirms JSON file streams parse down into expected lists."""
        records = load_raw_records(self.json_path)
        self.assertEqual(len(records), 2)
        self.assertEqual(records[0]["PARCELID"], "10020034")

    def test_load_records_from_xlsx_with_whitespace_stripping(self):
        """Verifies Excel column headers are cleanly stripped of exterior spacing errors."""
        records = load_raw_records(self.xlsx_path)
        self.assertEqual(len(records), 2)

        # Verify that column headers were stripped of leading/trailing spaces
        first_row = records[0]
        self.assertIn("REASON", first_row)
        self.assertNotIn(" REASON ", first_row)
        self.assertEqual(first_row["REASON"], "Value Correction")


if __name__ == "__main__":
    unittest.main()
