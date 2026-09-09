"""
High-level programmatic API for xlsx-viewer spreadsheet and formula engine.
"""

from __future__ import annotations

from pathlib import Path
from typing import Any, Callable

from .formulas import FormulaEngine
from .models import CellPosition, WorkbookData


def evaluate_formula(
    formula: str, context: dict[str, Any] | Callable[[int, int, str | None], Any] | None = None
) -> Any:
    """
    Evaluate an Excel formula string and return the computed value.

    Args:
        formula: Excel formula string starting with or without '=', e.g. '=SUM(10, 20, 30)' or '=VLOOKUP("A", A1:B10, 2, FALSE)'
        context: Optional cell value lookup dictionary (e.g. {'A1': 100, 'B1': 200}) or callable (row, col, sheet) -> value.

    Returns:
        The computed result (int, float, str, bool, or Excel error string like '#DIV/0!').

    Examples:
        >>> from xlsx_viewer import evaluate_formula
        >>> evaluate_formula("=SUM(1, 2, 3)")
        6
        >>> evaluate_formula("=IF(A1 > 50, 'PASS', 'FAIL')", context={'A1': 75})
        'PASS'
    """
    if not formula.startswith("="):
        formula = "=" + formula

    if callable(context):
        resolver = context
    elif isinstance(context, dict):
        # Build coordinate resolver
        cell_map = {}
        for k, v in context.items():
            pos = CellPosition.from_excel(k)
            cell_map[(pos.row, pos.col)] = v

        def resolver(r, c, s=None):
            return cell_map.get((r, c), 0)
    else:

        def resolver(r, c, s=None):
            return 0

    engine = FormulaEngine(resolver)
    return engine.evaluate(formula)


def load_workbook(path: str | Path, create_if_missing: bool = True) -> WorkbookData:
    """
    Load an Excel (.xlsx) file into memory data models.
    If create_if_missing is True and the file does not exist, an empty workbook
    with a default 'Sheet1' will be created automatically.

    Args:
        path: Path to the .xlsx file.
        create_if_missing: If True, initialize a new empty workbook when path is not found.

    Returns:
        WorkbookData instance containing sheets, cells, formulas, and formatting.
    """
    from openpyxl import load_workbook as _load_wb

    path = Path(path)
    if not path.exists():
        if create_if_missing:
            wb_data = WorkbookData(file_path=str(path))
            wb_data.add_sheet("Sheet1")
            return wb_data
        raise FileNotFoundError(f"Workbook not found: {path}")

    wb_data = WorkbookData(file_path=str(path))
    pyxl_wb = _load_wb(path, data_only=False)

    for sheet_name in pyxl_wb.sheetnames:
        pyxl_sheet = pyxl_wb[sheet_name]
        sheet_data = wb_data.add_sheet(sheet_name)

        for row_idx, row in enumerate(pyxl_sheet.iter_rows(values_only=False)):
            for col_idx, cell in enumerate(row):
                if cell.value is not None:
                    val = cell.value
                    formula = None
                    if isinstance(val, str) and val.startswith("="):
                        formula = val
                    sheet_data.set_cell(row_idx, col_idx, val, formula=formula)

    return wb_data


def save_workbook(wb: WorkbookData, path: str | Path | None = None) -> None:
    """
    Save a WorkbookData instance to disk as an .xlsx file.

    Args:
        wb: WorkbookData instance to save.
        path: Optional destination file path. If omitted, wb.file_path is used.
    """
    wb.save(path)


def launch_viewer(path: str | Path | None = None) -> None:
    """
    Launch the Excel Viewer Pro modern desktop GUI application.

    Args:
        path: Optional path to an .xlsx file to open immediately upon launch.
    """
    from .gui import main as gui_main

    gui_main(str(path) if path else None)
