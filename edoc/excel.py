"""Microsoft Excel COM adapter. Use on a dedicated worker thread only."""

from __future__ import annotations

import datetime as dt
import gc
import sys
from contextlib import contextmanager
from pathlib import Path
from typing import Iterator

from .config import Mapping, ValidationError


class ExcelSession:
    """Own a separate Excel instance; never attach to or kill the user's Excel."""

    def __enter__(self) -> ExcelSession:
        if sys.platform != "win32":
            raise ValidationError("Excel processing requires Windows and installed desktop Microsoft Excel.")
        try:
            import pythoncom
            import win32com.client
        except ImportError as exc:
            raise ValidationError("Excel support is missing. Run setup_windows.bat or use the Windows executable.") from exc
        self.pythoncom = pythoncom
        self.pythoncom.CoInitialize()
        self.app = None
        try:
            self.app = win32com.client.DispatchEx("Excel.Application")
            self.app.Visible = False
            self.app.DisplayAlerts = False
            self.app.EnableEvents = False
            self.app.AskToUpdateLinks = False
            self.app.AutomationSecurity = 3  # msoAutomationSecurityForceDisable
            # Establish manual mode before opening a source report. Excel needs
            # an open workbook to accept this setting; the bootstrap is private.
            bootstrap = self.app.Workbooks.Add()
            try:
                self.app.Calculation = -4135  # xlCalculationManual
            finally:
                bootstrap.Close(SaveChanges=False)
                bootstrap = None
        except Exception as exc:
            self.__exit__(None, None, None)
            raise ValidationError("Could not start desktop Microsoft Excel. Open Excel once and finish its setup.") from exc
        return self

    def __exit__(self, *_args) -> None:
        try:
            if self.app is not None:
                try:
                    self.app.Quit()
                finally:
                    self.app = None
                    gc.collect()
        finally:
            self.pythoncom.CoUninitialize()

    @contextmanager
    def open(self, path: Path, writable: bool = False) -> Iterator:
        book = None
        try:
            book = self.app.Workbooks.Open(
                Filename=str(path.resolve()), UpdateLinks=0, ReadOnly=not writable,
                Password="", WriteResPassword="", IgnoreReadOnlyRecommended=True,
                Notify=False, AddToMru=False, CorruptLoad=0,
            )
            # Excel cannot set calculation mode before its first workbook is open.
            self.app.Calculation = -4135  # xlCalculationManual
            if writable and book.ReadOnly:
                raise ValidationError("The copied approval template opened read-only.")
            if book.LinkSources(1) or book.LinkSources(2):
                raise ValidationError("Workbook contains external links. Use a self-contained workbook.")
            if book.Connections.Count:
                raise ValidationError("Workbook contains data connections. Use a self-contained workbook.")
            if any(int(sheet.Type) in (3, 4) for sheet in book.Sheets):
                raise ValidationError("Legacy Excel macro sheets are not supported.")
            yield book
        finally:
            if book is not None:
                book.Close(SaveChanges=False)
                book = None

    @staticmethod
    def sheets(book) -> list[str]:
        return [sheet.Name for sheet in book.Worksheets]

    @staticmethod
    def _sheet(book, name: str):
        try:
            return book.Worksheets.Item(name)
        except Exception as exc:
            raise ValidationError(f"Worksheet '{name}' was not found. Available: {', '.join(ExcelSession.sheets(book))}") from exc

    def _cell(self, book, sheet: str, address: str):
        cell = self._sheet(book, sheet).Range(address)
        if cell.MergeCells:
            cell = cell.MergeArea.Cells(1, 1)
        return cell

    def read(self, book, mapping: Mapping):
        cell = self._cell(book, mapping.source_sheet, mapping.source_cell)
        if self.app.WorksheetFunction.IsError(cell):
            raise ValidationError(f"{mapping.label}: source cell contains an Excel error.")
        raw = cell.Value2
        if raw is None or (isinstance(raw, str) and not raw.strip()):
            if mapping.required:
                raise ValidationError(f"{mapping.label}: {mapping.source_sheet}!{mapping.source_cell} is empty.")
            return None
        if mapping.mode == "text":
            value = str(cell.Text)
            if value and set(value) == {"#"}:
                raise ValidationError(f"{mapping.label}: Excel displays ####. Widen that source column or use Value mode.")
            return value
        # Value2 avoids Currency's four-decimal rounding; Value recognizes dates.
        typed = cell.Value
        if isinstance(typed, dt.datetime):
            return typed.replace(tzinfo=None)
        return raw

    def check_target(self, book, mapping: Mapping) -> str:
        sheet = self._sheet(book, mapping.target_sheet)
        if sheet.ProtectContents:
            raise ValidationError(f"Worksheet '{mapping.target_sheet}' is protected. Use an unlocked template copy.")
        cell = self._cell(book, mapping.target_sheet, mapping.target_cell)
        if cell.HasFormula and not mapping.allow_formula_overwrite:
            raise ValidationError(f"{mapping.label}: target has a formula. Enable formula replacement for this mapping only if intended.")
        return f"{sheet.Name.casefold()}!{cell.Address}"

    def write(self, book, mapping: Mapping, value) -> None:
        cell = self._cell(book, mapping.target_sheet, mapping.target_cell)
        if value is None:
            cell.ClearContents()
        elif isinstance(value, str):
            # Source strings are always data, including strings beginning '='.
            # Retain the template's number format after assigning a literal.
            number_format = cell.NumberFormat
            cell.NumberFormat = "@"
            cell.Value2 = value
            cell.NumberFormat = number_format
        elif isinstance(value, dt.datetime):
            cell.Value = value
        else:
            cell.Value2 = value

    def calculate(self, book) -> None:
        # Only the approval workbook is open here. Excel's dependency scheduler
        # handles cross-sheet formulas better than calculating sheets in order.
        self.app.CalculateFullRebuild()

    def save(self, book) -> None:
        # Save the deliverable in normal automatic calculation mode, then return
        # this private Excel instance to manual before the next report is opened.
        self.app.Calculation = -4105  # xlCalculationAutomatic
        try:
            book.Save()
        finally:
            self.app.Calculation = -4135

    def export_pdf(self, book, destination: Path, sheet: str = "") -> None:
        target = self._sheet(book, sheet) if sheet.strip() else book
        target.ExportAsFixedFormat(
            Type=0, Filename=str(destination.resolve()), Quality=0,
            IncludeDocProperties=False, IgnorePrintAreas=False, OpenAfterPublish=False,
        )
        if not destination.is_file() or not destination.stat().st_size:
            raise ValidationError("Excel did not create a PDF. Check the workbook's print area and page setup.")
