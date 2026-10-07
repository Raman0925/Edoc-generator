"""Opt-in real Office smoke test; normal CI has no Microsoft Excel installed.

Windows PowerShell: $env:EDOC_RUN_EXCEL_TESTS='1'; python -m unittest discover -s tests -v
"""

import os
import sys
import tempfile
import unittest
from pathlib import Path

from pypdf import PdfReader

from edoc.config import Mapping, Profile
from edoc.engine import run_batch
from edoc.excel import ExcelSession


@unittest.skipUnless(sys.platform == "win32" and os.getenv("EDOC_RUN_EXCEL_TESTS") == "1",
                     "Requires Windows, installed Excel and EDOC_RUN_EXCEL_TESTS=1")
class RealExcelSmokeTest(unittest.TestCase):
    def test_real_excel_mapping_save_and_pdf_export(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            inputs = root / "input"
            inputs.mkdir()
            template = root / "template.xlsx"
            report = inputs / "ACC-01.xlsx"
            with ExcelSession() as session:
                book = session.app.Workbooks.Add()
                try:
                    sheet = book.Worksheets.Item(1)
                    sheet.Name = "Approval"
                    sheet.Range("A1").Value2 = "Company approval"
                    sheet.Range("A2").Value2 = "Serial number"
                    sheet.Range("B2").Value2 = "PLACEHOLDER"
                    sheet.Range("A3").Value2 = "Sensitivity"
                    sheet.Range("B3").NumberFormat = "0.000000000"
                    sheet.PageSetup.PrintArea = "$A$1:$B$3"
                    sheet.Columns("A:B").ColumnWidth = 24
                    book.SaveAs(Filename=str(template), FileFormat=51)
                finally:
                    book.Close(SaveChanges=False)
                    book = None
                book = session.app.Workbooks.Add()
                try:
                    sheet = book.Worksheets.Item(1)
                    sheet.Name = "Report"
                    sheet.Range("A1").Value2 = "Calibration certificate"
                    sheet.Range("B4").NumberFormat = "00000"
                    sheet.Range("B4").Value2 = 12
                    sheet.Range("C12").Value2 = 1.123456789
                    sheet.Range("C12").NumberFormat = "0.000000000"
                    sheet.PageSetup.PrintArea = "$A$1:$C$12"
                    sheet.Columns("A:C").ColumnWidth = 24
                    book.SaveAs(Filename=str(report), FileFormat=51)
                finally:
                    book.Close(SaveChanges=False)
                    book = None
            profile = Profile(template=str(template), input_dir=str(inputs), output_dir=str(root / "output"),
                              approval_sheet="Approval", report_sheet="Report", mappings=[
                                  Mapping("Serial", "Report", "B4", "Approval", "B2", "text"),
                                  Mapping("Sensitivity", "Report", "C12", "Approval", "B3")])
            result = run_batch(profile)
            self.assertEqual(result.succeeded, 1, str(result.results))
            output = Path(result.results[0].output)
            self.assertEqual(len(PdfReader(str(output)).pages), 2)
            with ExcelSession() as session:
                with session.open(output.parent / "approval.xlsx") as book:
                    sheet = book.Worksheets.Item("Approval")
                    self.assertEqual(sheet.Range("B2").Value2, "00012")
                    self.assertAlmostEqual(sheet.Range("B3").Value2, 1.123456789, places=9)
                with session.open(report) as book:
                    self.assertEqual(book.Worksheets.Item("Report").Range("B4").Value2, 12)
