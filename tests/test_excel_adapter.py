import datetime as dt
import unittest
from types import SimpleNamespace

from edoc.config import Mapping, ValidationError
from edoc.excel import ExcelSession


class Cell:
    def __init__(self, raw=None, typed=None, text="", formula=False):
        self.Value2 = raw
        self.Value = typed if typed is not None else raw
        self.Text = text
        self.HasFormula = formula
        self.NumberFormat = "0.000000"
        self.Address = "$F$6"
        self.MergeCells = False

    def ClearContents(self):
        self.Value2 = None


class ExcelAdapterTests(unittest.TestCase):
    def setUp(self):
        self.session = ExcelSession()
        self.session.app = SimpleNamespace(WorksheetFunction=SimpleNamespace(IsError=lambda _: False))
        self.mapping = Mapping("Sensitivity", "Report", "B4", "Approval", "F6")
        self.cell = Cell()
        self.sheet = SimpleNamespace(Name="Approval", ProtectContents=False, Range=lambda _: self.cell)
        self.book = SimpleNamespace(Worksheets=SimpleNamespace(Item=lambda _: self.sheet))

    def test_numeric_value_uses_unrounded_value2(self):
        self.cell = Cell(raw=1.123456789, typed=1.1235)
        self.assertEqual(self.session.read(self.book, self.mapping), 1.123456789)

    def test_date_is_copied_as_date(self):
        date = dt.datetime(2026, 10, 7, tzinfo=dt.timezone.utc)
        self.cell = Cell(raw=46302, typed=date)
        self.assertEqual(self.session.read(self.book, self.mapping), date.replace(tzinfo=None))

    def test_displayed_text_keeps_leading_zeros(self):
        self.mapping.mode = "text"
        self.cell = Cell(raw=12, text="00012")
        self.assertEqual(self.session.read(self.book, self.mapping), "00012")

    def test_hash_display_is_rejected(self):
        self.mapping.mode = "text"
        self.cell = Cell(raw=12, text="#####")
        with self.assertRaisesRegex(ValidationError, "Widen"):
            self.session.read(self.book, self.mapping)

    def test_excel_error_rejected(self):
        self.session.app.WorksheetFunction.IsError = lambda _: True
        with self.assertRaisesRegex(ValidationError, "Excel error"):
            self.session.read(self.book, self.mapping)

    def test_formula_target_requires_explicit_permission(self):
        self.cell.HasFormula = True
        with self.assertRaisesRegex(ValidationError, "target has a formula"):
            self.session.check_target(self.book, self.mapping)
        self.mapping.allow_formula_overwrite = True
        self.assertEqual(self.session.check_target(self.book, self.mapping), "approval!$F$6")

    def test_protected_sheet_rejected(self):
        self.sheet.ProtectContents = True
        with self.assertRaisesRegex(ValidationError, "protected"):
            self.session.check_target(self.book, self.mapping)

    def test_literal_string_assignment_preserves_format(self):
        self.session.write(self.book, self.mapping, "=1+1")
        self.assertEqual(self.cell.Value2, "=1+1")
        self.assertEqual(self.cell.NumberFormat, "0.000000")

    def test_optional_blank_clears_placeholder(self):
        self.cell.Value2 = "placeholder"
        self.session.write(self.book, self.mapping, None)
        self.assertIsNone(self.cell.Value2)

    def test_merged_cell_uses_top_left(self):
        top_left = Cell(raw="ACC-01")
        self.cell.MergeCells = True
        self.cell.MergeArea = SimpleNamespace(Cells=lambda row, col: top_left)
        self.assertEqual(self.session.read(self.book, self.mapping), "ACC-01")
