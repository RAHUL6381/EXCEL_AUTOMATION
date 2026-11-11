"""
test_automation.py
Automated tests for Excel Financial Summary Automation Tool.
"""

import os
import shutil
import unittest
import openpyxl
from excel_processor import (
    process_excel_summary,
    extract_financial_data,
    validate_file_paths,
    ExcelAutomationError
)

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
SAMPLE_SOURCE = os.path.join(BASE_DIR, "sample_source_23-10-2026.xlsx")
TEMPLATE_FILE = os.path.join(BASE_DIR, "template.xlsx")
OUTPUT_FILE = os.path.join(BASE_DIR, "test_output.xlsx")


class TestExcelAutomation(unittest.TestCase):

    def setUp(self):
        if os.path.exists(OUTPUT_FILE):
            os.remove(OUTPUT_FILE)

    def tearDown(self):
        if os.path.exists(OUTPUT_FILE):
            try:
                os.remove(OUTPUT_FILE)
            except Exception:
                pass

    def test_file_validation_missing_source(self):
        with self.assertRaises(ExcelAutomationError) as ctx:
            validate_file_paths("non_existent_source.xlsx", TEMPLATE_FILE, OUTPUT_FILE)
        self.assertIn("Source file not found", str(ctx.exception))

    def test_file_validation_overwrite_source_prevention(self):
        with self.assertRaises(ExcelAutomationError) as ctx:
            validate_file_paths(SAMPLE_SOURCE, TEMPLATE_FILE, SAMPLE_SOURCE)
        self.assertIn("Accidental overwrite prevented", str(ctx.exception))

    def test_file_validation_overwrite_template_prevention(self):
        with self.assertRaises(ExcelAutomationError) as ctx:
            validate_file_paths(SAMPLE_SOURCE, TEMPLATE_FILE, TEMPLATE_FILE)
        self.assertIn("Accidental overwrite prevented", str(ctx.exception))

    def test_data_extraction(self):
        extracted = extract_financial_data(SAMPLE_SOURCE)
        self.assertEqual(extracted["source_sheet"], "KSM & SDM")
        self.assertIsNotNone(extracted["date_value"])
        cells = extracted["cells"]

        # KSM
        self.assertEqual(cells["B3"]["value"], "21.07Crore")
        self.assertEqual(cells["B5"]["value"], "19.30Crore ")
        self.assertEqual(cells["B9"]["value"], "1.49 crore  negative")
        self.assertEqual(cells["B11"]["value"], "14.87 lakhs negative")

        # SDM
        self.assertEqual(cells["E3"]["value"], "5.88Crore")
        self.assertEqual(cells["E5"]["value"], "5.45 Crore ")
        self.assertEqual(cells["E9"]["value"], "51.80 lakhs negative")
        self.assertEqual(cells["E11"]["value"], "18.30 lakhs   positive ")

        # KSM HUF
        self.assertEqual(cells["H3"]["value"], "45.54 lakhs ")
        self.assertEqual(cells["H5"]["value"], "Nil ")
        self.assertEqual(cells["H9"]["value"], "3.30 lakhs negative")
        self.assertEqual(cells["H11"]["value"], "0.38 lakhs negative")

    def test_end_to_end_processing(self):
        res = process_excel_summary(SAMPLE_SOURCE, TEMPLATE_FILE, OUTPUT_FILE)
        self.assertTrue(res["success"])
        self.assertTrue(os.path.exists(OUTPUT_FILE))
        self.assertEqual(res["replacement_count"], 15)

        # Inspect generated workbook
        wb = openpyxl.load_workbook(OUTPUT_FILE)
        ws = wb.active

        # Check section titles and headers
        self.assertEqual(ws["A1"].value, "KSM")
        self.assertEqual(ws["A2"].value, "TODAY DATE")
        self.assertEqual(ws["B2"].value, "STOCK VALUE")
        self.assertEqual(ws["C2"].value, "MTF DEBIT")
        self.assertEqual(ws["D2"].value, "BOOKED PNL")
        self.assertEqual(ws["E2"].value, "MTM")

        # Row 3 (KSM values)
        # Check date in A3
        self.assertIn(str(ws["A3"].value), ["2026-10-23", "2026-10-23 00:00:00", "23-10-2026"])
        self.assertEqual(ws["B3"].value, "21.07Crore")
        self.assertEqual(ws["C3"].value, "19.30Crore ")
        self.assertEqual(ws["D3"].value, "14.87 lakhs negative")
        self.assertEqual(ws["E3"].value, "1.49 crore  negative")

        # Row 6 & 7 (SDM header)
        self.assertEqual(ws["A6"].value, "SDM")
        self.assertEqual(ws["A7"].value, "TODAY DATE")
        self.assertEqual(ws["B7"].value, "STOCK VALUE")
        self.assertEqual(ws["C7"].value, "MTF DEBIT")
        self.assertEqual(ws["D7"].value, "BOOKED PNL")
        self.assertEqual(ws["E7"].value, "MTM")

        # Row 8 (SDM values)
        self.assertEqual(ws["B8"].value, "5.88Crore")
        self.assertEqual(ws["C8"].value, "5.45 Crore ")
        self.assertEqual(ws["D8"].value, "18.30 lakhs   positive ")
        self.assertEqual(ws["E8"].value, "51.80 lakhs negative")

        # Row 11 & 12 (KSM HUF header)
        self.assertEqual(ws["A11"].value, "KSM HUF")
        self.assertEqual(ws["A12"].value, "TODAY DATE")
        self.assertEqual(ws["B12"].value, "STOCK VALUE")
        self.assertEqual(ws["C12"].value, "MTF DEBIT")
        self.assertEqual(ws["D12"].value, "BOOKED PNL")
        self.assertEqual(ws["E12"].value, "MTM")

        # Row 13 (KSM HUF values)
        self.assertEqual(ws["B13"].value, "45.54 lakhs ")
        self.assertEqual(ws["C13"].value, "Nil ")
        self.assertEqual(ws["D13"].value, "0.38 lakhs negative")
        self.assertEqual(ws["E13"].value, "3.30 lakhs negative")

        wb.close()

    def test_ntf_debit_typo_correction(self):
        # Create a temporary template with 'NTF DEBIT' typo
        temp_tpl = os.path.join(BASE_DIR, "temp_ntf_tpl.xlsx")
        shutil.copyfile(TEMPLATE_FILE, temp_tpl)
        wb = openpyxl.load_workbook(temp_tpl)
        ws = wb.active
        ws["C2"].value = "NTF DEBIT"
        ws["C7"].value = "NTF DEBIT"
        ws["C12"].value = "NTF DEBIT"
        wb.save(temp_tpl)
        wb.close()

        try:
            res = process_excel_summary(SAMPLE_SOURCE, temp_tpl, OUTPUT_FILE)
            self.assertTrue(res["success"])
            wb_out = openpyxl.load_workbook(OUTPUT_FILE)
            ws_out = wb_out.active
            self.assertEqual(ws_out["C2"].value, "MTF DEBIT")
            self.assertEqual(ws_out["C7"].value, "MTF DEBIT")
            self.assertEqual(ws_out["C12"].value, "MTF DEBIT")
            wb_out.close()
        finally:
            if os.path.exists(temp_tpl):
                os.remove(temp_tpl)


if __name__ == "__main__":
    unittest.main()
