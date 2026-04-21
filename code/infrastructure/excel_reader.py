"""Excel reading operations using xlrd library."""

from __future__ import annotations

from pathlib import Path
from typing import Any, List, Tuple


class ExcelReader:
    """
    Handles reading Excel files using xlrd library.

    This class isolates all Excel reading operations, making the code
    testable without real Excel files (via mocking) and centralizing
    xlrd-specific logic.
    """

    def __init__(self, xlrd_module):
        """
        Initialize ExcelReader with xlrd module.

        Args:
            xlrd_module: The xlrd module (injected for testability)
        """
        self.xlrd = xlrd_module

    def read_source_rows(
        self,
        file_path: Path,
        sheet_name: str
    ) -> Tuple[List[List[Any]], int]:
        """
        Read all rows from a source Excel file.

        Args:
            file_path: Path to the .xls file
            sheet_name: Name of the sheet to read

        Returns:
            Tuple of (rows, datemode):
            - rows: List of lists containing cell values
            - datemode: Excel datemode for date conversion (0 or 1)

        Raises:
            ValueError: If sheet is not found
            IOError: If file cannot be opened
        """
        try:
            wb = self.xlrd.open_workbook(str(file_path), formatting_info=False)
        except Exception as e:
            raise IOError(f"Cannot open workbook {file_path}: {e}") from e

        if sheet_name not in wb.sheet_names():
            raise ValueError(f"Sheet '{sheet_name}' not found in {file_path}")

        ws = wb.sheet_by_name(sheet_name)
        rows: List[List[Any]] = []

        for r in range(ws.nrows):
            row = [
                ws.cell_value(r, c) if c < ws.ncols else None
                for c in range(ws.ncols)
            ]
            rows.append(row)

        datemode = int(getattr(wb, "datemode", 0))
        return rows, datemode

    def get_sheet_names(self, file_path: Path) -> List[str]:
        """
        Get list of all sheet names in an Excel file.

        Args:
            file_path: Path to the .xls file

        Returns:
            List of sheet names

        Raises:
            IOError: If file cannot be opened
        """
        try:
            wb = self.xlrd.open_workbook(str(file_path), formatting_info=False)
            return wb.sheet_names()
        except Exception as e:
            raise IOError(f"Cannot open workbook {file_path}: {e}") from e

    def get_cell_value(
        self,
        file_path: Path,
        sheet_name: str,
        row: int,
        col: int
    ) -> Any:
        """
        Get value of a specific cell.

        Args:
            file_path: Path to the .xls file
            sheet_name: Name of the sheet
            row: Row index (0-based)
            col: Column index (0-based)

        Returns:
            Cell value

        Raises:
            ValueError: If sheet not found or indices out of bounds
            IOError: If file cannot be opened
        """
        try:
            wb = self.xlrd.open_workbook(str(file_path), formatting_info=False)
        except Exception as e:
            raise IOError(f"Cannot open workbook {file_path}: {e}") from e

        if sheet_name not in wb.sheet_names():
            raise ValueError(f"Sheet '{sheet_name}' not found")

        ws = wb.sheet_by_name(sheet_name)

        if row >= ws.nrows or col >= ws.ncols:
            raise ValueError(f"Cell ({row}, {col}) out of bounds")

        return ws.cell_value(row, col)
