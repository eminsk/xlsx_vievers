"""
Unit and integration test suite for Excel Viewer Pro.
"""

import unittest
from pathlib import Path
from models import CellPosition, CellRange, CellStyle, SheetData
from formulas import FormulaEngine, FormulaStore, shift_formula_references, CellRef, parse_range
from formatting import NumberFormatter, ConditionalFormattingEngine
from openpyxl import Workbook


class TestExcelViewerPro(unittest.TestCase):

    def test_cell_and_range_models(self):
        p1 = CellPosition(0, 0)
        self.assertEqual(p1.to_excel(), "A1")

        p2 = CellPosition.from_excel("C10")
        self.assertEqual((p2.row, p2.col), (9, 2))

        rng = CellRange(p1, p2)
        self.assertEqual(rng.to_excel(), "A1:C10")
        self.assertTrue(rng.contains(5, 1))
        self.assertFalse(rng.contains(10, 3))
        self.assertEqual(rng.row_count, 10)
        self.assertEqual(rng.col_count, 3)

    def test_formula_arithmetic_and_precedence(self):
        engine = FormulaEngine(lambda r, c, s=None: 0)
        self.assertEqual(engine.evaluate("=2+3*4"), 14)
        self.assertEqual(engine.evaluate("=(2+3)*4"), 20)
        self.assertEqual(engine.evaluate("=2^3"), 8)
        self.assertEqual(engine.evaluate("=10/2"), 5)
        self.assertEqual(engine.evaluate("=10/0"), "#DIV/0!")

    def test_formula_lookup_and_math_functions(self):
        data = {
            (0, 0): "Name", (0, 1): "Salary", (0, 2): "Dept",
            (1, 0): "Alice", (1, 1): 50000, (1, 2): "IT",
            (2, 0): "Bob", (2, 1): 60000, (2, 2): "HR",
            (3, 0): "Charlie", (3, 1): 70000, (3, 2): "IT",
        }
        engine = FormulaEngine(lambda r, c, s=None: data.get((r, c), 0))

        # SUM, AVERAGE, COUNT
        self.assertEqual(engine.evaluate("=SUM(B2:B4)"), 180000)
        self.assertEqual(engine.evaluate("=AVERAGE(B2:B4)"), 60000)
        self.assertEqual(engine.evaluate("=COUNT(B2:B4)"), 3)

        # VLOOKUP
        self.assertEqual(engine.evaluate('=VLOOKUP("Bob", A2:C4, 2, FALSE)'), 60000)
        self.assertEqual(engine.evaluate('=VLOOKUP("Charlie", A2:C4, 3, FALSE)'), "IT")

        # INDEX and MATCH
        self.assertEqual(engine.evaluate('=INDEX(B2:B4, 2)'), 60000)
        self.assertEqual(engine.evaluate('=MATCH("Charlie", A2:A4, 0)'), 3)

        # SUMIFS & COUNTIFS
        self.assertEqual(engine.evaluate('=COUNTIF(C2:C4, "IT")'), 2)
        self.assertEqual(engine.evaluate('=SUMIF(C2:C4, "IT", B2:B4)'), 120000)

        # Logical IF & IFERROR
        self.assertEqual(engine.evaluate('=IF(B2>55000, "High", "Low")'), "Low")
        self.assertEqual(engine.evaluate('=IF(B3>55000, "High", "Low")'), "High")
        self.assertEqual(engine.evaluate('=IFERROR(10/0, "Error Occurred")'), "Error Occurred")

        # Text functions
        self.assertEqual(engine.evaluate('=CONCATENATE("Hello", " ", "World")'), "Hello World")
        self.assertEqual(engine.evaluate('=UPPER("excel pro")'), "EXCEL PRO")
        self.assertEqual(engine.evaluate('=LEFT("Antigravity", 4)'), "Anti")
        self.assertEqual(engine.evaluate('=RIGHT("Antigravity", 7)'), "gravity")
        self.assertEqual(engine.evaluate('=LEN("Google")'), 6)

    def test_reference_shifting(self):
        f1 = "=A1+$B$1+C$2+$D3"
        shifted = shift_formula_references(f1, row_delta=2, col_delta=1)
        self.assertEqual(shifted, "=B3+$B$1+D$2+$D5")

    def test_number_formatting(self):
        self.assertEqual(NumberFormatter.format_value(1234.56, "$#,##0.00"), "$1,234.56")
        self.assertEqual(NumberFormatter.format_value(1234.5, "#,##0.00 ₽"), "1,234.50 ₽")
        self.assertEqual(NumberFormatter.format_value(0.155, "0.00%"), "15.50%")
        self.assertEqual(NumberFormatter.format_value(0.15, "0%"), "15%")
        self.assertEqual(NumberFormatter.format_value(5000, "#,##0"), "5,000")

    def test_conditional_formatting_engine(self):
        rule_gt = {"type": "greater_than", "value": 50, "bg_color": "#00FF00", "fg_color": "#000000"}
        bg, fg = ConditionalFormattingEngine.evaluate_rule(75, rule_gt)
        self.assertEqual(bg, "#00FF00")

        bg, fg = ConditionalFormattingEngine.evaluate_rule(25, rule_gt)
        self.assertIsNone(bg)

        # 3-color scale
        rule_scale = {"type": "color_scale", "scale": "green_yellow_red"}
        bg, fg = ConditionalFormattingEngine.evaluate_rule(50, rule_scale, [0, 50, 100])
        self.assertIsNotNone(bg)

    def test_openpyxl_roundtrip(self):
        wb = Workbook()
        ws = wb.active
        ws.title = "Sales"
        ws["A1"] = "Item"
        ws["B1"] = "Price"
        ws["A2"] = "Widget"
        ws["B2"] = 25.50
        ws["A3"] = "Total"
        ws["B3"] = "=SUM(B2:B2)"

        test_path = Path("test_output.xlsx")
        try:
            wb.save(test_path)
            self.assertTrue(test_path.exists())
        finally:
            if test_path.exists():
                test_path.unlink()


if __name__ == "__main__":
    unittest.main()
