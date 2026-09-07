"""
Comprehensive Unit Tests for xlsx_viewer Library API.
Tests headless formula evaluation, SIMD assembly acceleration, data models, and CLI.
"""

import sys
import unittest
from pathlib import Path

# Ensure package is importable
root_dir = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(root_dir))

import xlsx_viewer as xv
from xlsx_viewer.models import CellPosition, CellRange, WorkbookData


class TestXlsxViewerLibrary(unittest.TestCase):
    def test_version_and_metadata(self):
        self.assertEqual(xv.__version__, "1.0.0")
        self.assertGreaterEqual(len(xv.FUNCTION_METADATA), 120)

    def test_evaluate_arithmetic(self):
        self.assertEqual(xv.evaluate_formula("=2+3*4"), 14)
        self.assertEqual(xv.evaluate_formula("=(2+3)*4"), 20)
        self.assertEqual(xv.evaluate_formula("=2^8"), 256)
        self.assertEqual(xv.evaluate_formula("=100/4"), 25)
        self.assertEqual(xv.evaluate_formula("=10/0"), "#DIV/0!")

    def test_evaluate_math_and_logic(self):
        self.assertEqual(xv.evaluate_formula("=SQRT(81)"), 9.0)
        self.assertEqual(xv.evaluate_formula("=ABS(-42)"), 42)
        self.assertEqual(xv.evaluate_formula('=IF(10 > 5, "YES", "NO")'), "YES")
        self.assertEqual(xv.evaluate_formula('=IF(10 < 5, "YES", "NO")'), "NO")

    def test_evaluate_string_functions(self):
        self.assertEqual(xv.evaluate_formula('=UPPER("hello")'), "HELLO")
        self.assertEqual(xv.evaluate_formula('=LOWER("WORLD")'), "world")
        self.assertEqual(xv.evaluate_formula('=CONCATENATE("Excel", " ", "Pro")'), "Excel Pro")

    def test_evaluate_with_context(self):
        context = {
            "A1": 100,
            "A2": 250,
            "B1": 50,
        }
        self.assertEqual(xv.evaluate_formula("=A1 + A2 + B1", context=context), 400)
        self.assertEqual(
            xv.evaluate_formula('=IF(A2 > 200, "HIGH", "LOW")', context=context), "HIGH"
        )

    def test_simd_math_engine(self):
        # Even if native DLL is unavailable on non-Windows/CI, fallback must produce identical results
        data = [1.0, 2.0, 3.0, 4.0, 5.0]
        self.assertEqual(xv.asm_sum(data), 15.0)
        self.assertEqual(xv.asm_avg(data), 3.0)
        self.assertEqual(xv.asm_min(data), 1.0)
        self.assertEqual(xv.asm_max(data), 5.0)
        self.assertEqual(xv.asm_sumproduct([1.0, 2.0], [10.0, 20.0]), 50.0)

        # Financial PMT calculation
        pmt = xv.asm_pmt(0.05 / 12, 360, -250000)
        self.assertAlmostEqual(pmt, 1342.05, places=1)

    def test_data_models(self):
        pos = CellPosition.from_excel("D15")
        self.assertEqual((pos.row, pos.col), (14, 3))
        self.assertEqual(pos.to_excel(), "D15")

        rng = CellRange.from_excel("B2:E6")
        self.assertEqual(rng.row_count, 5)
        self.assertEqual(rng.col_count, 4)
        self.assertTrue(rng.contains(3, 2))
        self.assertFalse(rng.contains(0, 0))

        wb = WorkbookData()
        sheet = wb.add_sheet("TestSheet")
        sheet.set_cell(0, 0, "Hello")
        self.assertEqual(sheet.get_cell_value(0, 0), "Hello")
        self.assertIn("TestSheet", wb.sheet_names)

    def test_cli_headless_execution(self):
        from xlsx_viewer.cli import main as cli_main

        # Test --info
        ret_info = cli_main(["--info"])
        self.assertEqual(ret_info, 0)

        # Test --list-formulas
        ret_list = cli_main(["--list-formulas"])
        self.assertEqual(ret_list, 0)

        # Test --calc
        ret_calc = cli_main(["--calc", "=SUM(10, 20, 30)"])
        self.assertEqual(ret_calc, 0)


if __name__ == "__main__":
    unittest.main()
