"""
xlsx-viewer: Modern Excel File Viewer, Data Models, and 129-Function Formula Engine with SIMD Acceleration.
"""

from __future__ import annotations

from .api import evaluate_formula, launch_viewer, load_workbook
from .asm import (
    asm_avg,
    asm_fv,
    asm_max,
    asm_min,
    asm_pmt,
    asm_pv,
    asm_str_hash,
    asm_sum,
    asm_sumproduct,
    is_asm_available,
)
from .formatting import ConditionalFormattingEngine, NumberFormatter
from .formulas import (
    FUNCTION_METADATA,
    CellRef,
    FormulaEngine,
    parse_range,
    shift_formula_references,
)
from .models import (
    CellComment,
    CellData,
    CellPosition,
    CellRange,
    CellStyle,
    SheetData,
    UndoAction,
    WorkbookData,
)

__version__ = "1.0.1"
__author__ = "eminsk"
__license__ = "MIT"

__all__ = [
    # Top-Level Helpers
    "evaluate_formula",
    "load_workbook",
    "launch_viewer",
    "__version__",
    # Formula Engine
    "FormulaEngine",
    "FUNCTION_METADATA",
    "CellRef",
    "parse_range",
    "shift_formula_references",
    # Data Models
    "WorkbookData",
    "SheetData",
    "CellData",
    "CellPosition",
    "CellRange",
    "CellStyle",
    "CellComment",
    "UndoAction",
    # Formatting
    "ConditionalFormattingEngine",
    "NumberFormatter",
    # Hardware SIMD Acceleration
    "is_asm_available",
    "asm_sum",
    "asm_avg",
    "asm_min",
    "asm_max",
    "asm_sumproduct",
    "asm_pmt",
    "asm_pv",
    "asm_fv",
    "asm_str_hash",
]
