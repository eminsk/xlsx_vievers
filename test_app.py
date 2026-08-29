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
        self.assertEqual(NumberFormatter.format_value(1234.5, "#,##0.00 ₽"), "1 234.50 ₽")
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

    def test_deep_chain_memoization(self):
        # 200 sequential dependent cells: A1 = 10, A2 = A1 + 5, A3 = A2 + 5...
        formulas = {}
        values = {(0, 0): 10}
        for r in range(1, 200):
            formulas[(r, 0)] = f"=A{r}+5"

        calc_cache = {}
        evaluating = set()

        def get_val(r, c, sheet=None):
            key = (sheet or "Sheet1", r, c)
            if key in calc_cache:
                return calc_cache[key]
            if key in evaluating:
                return 0
            if (r, c) in values:
                return values[(r, c)]
            f = formulas.get((r, c))
            if f:
                evaluating.add(key)
                try:
                    res = engine.evaluate(f, sheet or "Sheet1")
                finally:
                    evaluating.remove(key)
                calc_cache[key] = res
                return res
            return 0

        engine = FormulaEngine(get_val)
        import time
        t0 = time.time()
        res = get_val(199, 0) # A200
        t1 = time.time()
        # 10 + 199 * 5 = 1005
        self.assertEqual(res, 1005)
        self.assertLess(t1 - t0, 0.2) # Must compute in < 200ms

    def test_cross_sheet_unicode_references(self):
        sheets = {
            "Дашборд и Сводка": {(4, 2): 100}, # C5
            "Форекс (5 дней)": {(0, 0): "=MIN('Дашборд и Сводка'!$C$5, 50)"}
        }
        def get_val(r, c, sheet=None):
            target = sheet or "Дашборд и Сводка"
            return sheets.get(target, {}).get((r, c), 0)

        engine = FormulaEngine(get_val)
        res = engine.evaluate("='Дашборд и Сводка'!$C$5", "Форекс (5 дней)")
        self.assertEqual(res, 100)
        res_min = engine.evaluate("=MIN('Дашборд и Сводка'!$C$5, 50)", "Форекс (5 дней)")
        self.assertEqual(res_min, 50)

    def test_circular_reference_protection(self):
        # A1 = =B1, B1 = =A1
        formulas = {(0, 0): "=B1", (0, 1): "=A1"}
        evaluating = set()
        def get_val(r, c, sheet=None):
            key = (sheet, r, c)
            if key in evaluating:
                return 0
            evaluating.add(key)
            try:
                f = formulas.get((r, c))
                return engine.evaluate(f, sheet) if f else 0
            finally:
                evaluating.remove(key)

        engine = FormulaEngine(get_val)
        res = engine.evaluate("=A1", "Sheet1")
        self.assertEqual(res, 0)

    def test_extended_number_formatting(self):
        self.assertEqual(NumberFormatter.format_value(0.1234, "#,##0.0%"), "12.3%")
        self.assertEqual(NumberFormatter.format_value(0.095, "#,##0.0%"), "9.5%")
        self.assertEqual(NumberFormatter.format_value(-50.5, "$#,##0.00"), "-$50.50")
        self.assertEqual(NumberFormatter.format_value(15264010, "$#,##0.00"), "$15,264,010.00")
        self.assertEqual(NumberFormatter.format_value(100, "0.00"), "100.00")

    def test_new_financial_functions(self):
        engine = FormulaEngine(lambda r, c, s=None: 0)
        # PMT: rate=0.05/12, nper=360, pv=200000 -> approx -1073.64
        pmt = engine.evaluate("=PMT(0.05/12, 360, 200000)")
        self.assertAlmostEqual(pmt, -1073.6432, places=2)

        # FV: rate=0.05/12, nper=120, pmt=-100, pv=-1000 -> approx 17175.24
        fv = engine.evaluate("=FV(0.05/12, 120, -100, -1000)")
        self.assertAlmostEqual(fv, 17175.24, places=1)

        # PV: rate=0.08/12, nper=240, pmt=-500, fv=0 -> approx 59777.15
        pv = engine.evaluate("=PV(0.08/12, 240, -500, 0)")
        self.assertAlmostEqual(pv, 59777.15, places=1)

        # NPER: rate=0.06/12, pmt=-250, pv=10000 -> approx 44.74
        nper = engine.evaluate("=NPER(0.06/12, -250, 10000)")
        self.assertAlmostEqual(nper, 44.74, places=1)

    def test_new_date_and_time_functions(self):
        engine = FormulaEngine(lambda r, c, s=None: 0)
        self.assertEqual(engine.evaluate('=TIME(14, 30, 15)'), "14:30:15")
        self.assertEqual(engine.evaluate('=DATEDIF("2020-01-01", "2023-01-01", "Y")'), 3)
        self.assertEqual(engine.evaluate('=DATEDIF("2020-01-01", "2020-05-01", "M")'), 4)
        self.assertEqual(engine.evaluate('=DATEDIF("2020-01-01", "2020-01-15", "D")'), 14)
        self.assertEqual(engine.evaluate('=EDATE("2023-01-15", 2)'), "2023-03-15")
        self.assertEqual(engine.evaluate('=EOMONTH("2023-01-15", 1)'), "2023-02-28")
        self.assertEqual(engine.evaluate('=HOUR("14:30:15")'), 14)
        self.assertEqual(engine.evaluate('=MINUTE("14:30:15")'), 30)
        self.assertEqual(engine.evaluate('=SECOND("14:30:15")'), 15)

    def test_sumproduct_lookup_rank_row_col(self):
        data = {
            (0, 0): 10, (0, 1): 2,
            (1, 0): 20, (1, 1): 3,
            (2, 0): 30, (2, 1): 4,
        }
        engine = FormulaEngine(lambda r, c, s=None: data.get((r, c), 0))
        # SUMPRODUCT: 10*2 + 20*3 + 30*4 = 20 + 60 + 120 = 200
        self.assertEqual(engine.evaluate("=SUMPRODUCT(A1:A3, B1:B3)"), 200)

        # RANK
        self.assertEqual(engine.evaluate("=RANK(20, A1:A3)"), 2)
        self.assertEqual(engine.evaluate("=RANK(30, A1:A3)"), 1)

        # ROW & COLUMN
        self.assertEqual(engine.evaluate("=ROW(C5)"), 5)
        self.assertEqual(engine.evaluate("=COLUMN(C5)"), 3)

        # LOOKUP
        self.assertEqual(engine.evaluate("=LOOKUP(25, A1:A3, B1:B3)"), 3)

    def test_sheet_data_matrix_invariants(self):
        s = SheetData(name="Test", col_count=3, row_count=3)
        s.set_cell_value(5, 5, "Expanded")
        self.assertEqual(s.row_count, 6)
        self.assertEqual(s.col_count, 6)
        self.assertEqual(s.get_cell_value(5, 5), "Expanded")
        self.assertIsNone(s.get_cell_value(10, 10))

    def test_additional_formula_functions(self):
        data = {
            (0, 0): "Alpha", (0, 1): "Beta",
            (1, 0): "", (1, 1): 42,
            (2, 0): None, (2, 1): 99,
        }
        engine = FormulaEngine(lambda r, c, s=None: data.get((r, c)))
        self.assertEqual(engine.evaluate("=COUNTBLANK(A1:B3)"), 2)
        self.assertEqual(engine.evaluate("=ROWS(A1:B3)"), 3)
        self.assertEqual(engine.evaluate("=COLUMNS(A1:B3)"), 2)
        self.assertEqual(engine.evaluate('=WEEKDAY("2023-01-01")'), 1) # Sunday = 1

    def test_accounting_and_serial_date_formatting(self):
        self.assertEqual(NumberFormatter.format_value(0, "_($* #,##0.00_);_($* (#,##0.00);_($* \"-\"??_);_(@_)"), "$ -")
        self.assertEqual(NumberFormatter.format_value(-1234.56, "_($* #,##0.00_);_($* (#,##0.00);_($* \"-\"??_);_(@_)"), "($1,234.56)")
        self.assertEqual(NumberFormatter.format_value(44927, "yyyy-mm-dd"), "2023-01-01")


if __name__ == "__main__":
    unittest.main()
