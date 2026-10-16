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
        self.assertTrue(xv.__version__.startswith("1."))
        self.assertGreaterEqual(len(xv.FUNCTION_METADATA), 120)

    def test_evaluate_arithmetic(self):
        self.assertEqual(xv.evaluate_formula("=2+3*4"), 14)
        self.assertEqual(xv.evaluate_formula("=(2+3)*4"), 20)
        self.assertEqual(xv.evaluate_formula("=SUM(10, 20) + MAX(5, 15)"), 45)
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

    def test_launch_viewer_and_cli_default(self):
        try:
            import xlsx_viewer.gui  # noqa: F401
        except Exception:
            self.skipTest(
                "GUI dependencies (ttkbootstrap/tkinter/Pillow) not available in this runtime"
            )

        from unittest.mock import MagicMock, patch

        from xlsx_viewer.cli import main as cli_main

        with patch("xlsx_viewer.gui.ExcelViewerPro") as mock_cls:
            mock_app = MagicMock()
            mock_cls.return_value = mock_app

            # 1. Test launch_viewer() with no args
            xv.launch_viewer()
            mock_app.run.assert_called()

            # 2. Test launch_viewer("test.xlsx") with arg
            mock_app.reset_mock()
            xv.launch_viewer("test.xlsx")
            mock_app.run.assert_called()

            # 3. Test cli_main([]) default GUI launch
            mock_app.reset_mock()
            ret = cli_main([])
            self.assertEqual(ret, 0)
            mock_app.run.assert_called()

    def test_multithreading_and_no_gil(self):
        import concurrent.futures

        def worker(i):
            return xv.evaluate_formula(f"=SUM({i}, 10, 20) * 2")

        with concurrent.futures.ThreadPoolExecutor(max_workers=8) as ex:
            results = list(ex.map(worker, range(50)))

        self.assertEqual(len(results), 50)
        self.assertEqual(results[0], 60.0)
        self.assertEqual(results[10], 80.0)

    def test_load_and_save_workbook(self):
        import tempfile

        with tempfile.TemporaryDirectory() as tmpdir:
            test_path = Path(tmpdir) / "nonexistent.xlsx"

            # 1. create_if_missing=True (default) should succeed and create Sheet1
            wb = xv.load_workbook(test_path)
            self.assertIn("Sheet1", wb.sheet_names)
            sheet = wb.get_sheet("Sheet1")
            sheet.set_cell(0, 0, "TestValue")

            # 2. Save workbook to disk
            wb.save()
            self.assertTrue(test_path.exists())

            # 3. Reload existing workbook
            wb_reloaded = xv.load_workbook(test_path)
            self.assertEqual(wb_reloaded.get_sheet("Sheet1").get_cell(0, 0).value, "TestValue")

            # 4. create_if_missing=False on missing file should raise FileNotFoundError
            missing = Path(tmpdir) / "missing.xlsx"
            with self.assertRaises(FileNotFoundError):
                xv.load_workbook(missing, create_if_missing=False)

    def test_load_legacy_xls(self):
        sample_xls = Path(__file__).resolve().parent.parent / "examples" / "sample_legacy.xls"
        if not sample_xls.exists():
            try:
                import tempfile

                import xlwt

                with tempfile.TemporaryDirectory() as tmpdir:
                    sample_xls = Path(tmpdir) / "test.xls"
                    wb = xlwt.Workbook()
                    ws = wb.add_sheet("TestSheet")
                    ws.write(0, 0, "Hello")
                    ws.write(0, 1, 123.45)
                    wb.save(str(sample_xls))
            except ImportError:
                self.skipTest("Neither sample_legacy.xls exists nor xlwt is installed")

        try:
            import xlrd  # noqa: F401
        except ImportError:
            self.skipTest("xlrd not installed")

        loaded = xv.load_workbook(sample_xls)
        self.assertGreater(len(loaded.sheet_names), 0)
        sh = loaded.get_sheet()
        self.assertIsNotNone(sh)
        self.assertGreater(sh.row_count, 0)

    def test_cli_to_csv(self):
        import tempfile

        from xlsx_viewer.cli import main

        with tempfile.TemporaryDirectory() as tmpdir:
            xlsx_path = Path(tmpdir) / "test.xlsx"
            csv_path = Path(tmpdir) / "test.csv"
            wb = xv.load_workbook(xlsx_path)
            sh = wb.get_sheet("Sheet1")
            sh.set_cell(0, 0, "Name")
            sh.set_cell(0, 1, "Score")
            sh.set_cell(1, 0, "Alice")
            sh.set_cell(1, 1, 95)
            wb.save(xlsx_path)

            ret = main(["--to-csv", str(xlsx_path), str(csv_path)])
            self.assertEqual(ret, 0)
            self.assertTrue(csv_path.exists())
            content = csv_path.read_text(encoding="utf-8")
            self.assertIn("Name,Score", content)
            self.assertIn("Alice,95", content)

    def test_load_html_table_as_xls(self):
        import tempfile

        with tempfile.TemporaryDirectory() as tmpdir:
            html_xls = Path(tmpdir) / "export.xls"
            html_xls.write_text(
                "<table><tr><th>Col1</th><th>Col2</th></tr><tr><td>Alpha</td><td>100</td></tr></table>",
                encoding="utf-8",
            )
            wb = xv.load_workbook(html_xls)
            self.assertIn("Sheet1", wb.sheet_names)
            sh = wb.get_sheet("Sheet1")
            self.assertEqual(sh.get_cell_value(0, 0), "Col1")
            self.assertEqual(sh.get_cell_value(1, 0), "Alpha")

    def test_workbook_data_custom_sheet_name(self):
        from xlsx_viewer.models import WorkbookData

        wb = WorkbookData()
        wb.add_sheet("FinanceReport")
        self.assertEqual(wb.active_sheet_name, "FinanceReport")
        sh = wb.get_sheet()
        self.assertIsNotNone(sh)
        self.assertEqual(sh.name, "FinanceReport")

    def test_downloads_prais_xlsx_if_present(self):
        prais_path = Path(r"D:\Downloads\prais.xlsx")
        if prais_path.exists():
            wb = xv.load_workbook(prais_path)
            self.assertGreater(len(wb.sheet_names), 0)
            sh = wb.get_sheet()
            self.assertIsNotNone(sh)
            self.assertGreater(sh.row_count, 10)

    def test_downloads_financial_plan_xlsx_if_present(self):
        plan_path = Path(r"D:\Downloads\ФИНАНСОВЫЙ ПЛАН И КАЛЬКУЛЯТОР ДОХОДНОСТИ НА 1 ГОД.xlsx")
        if plan_path.exists():
            wb = xv.load_workbook(plan_path)
            self.assertGreater(len(wb.sheet_names), 0)
            sh = wb.get_sheet()
            self.assertIsNotNone(sh)
            self.assertGreater(sh.row_count, 5)

    def test_downloads_tablica_backup_if_present(self):
        bkp_path = Path(r"D:\Downloads\Таблица_backup.xlsx")
        if bkp_path.exists():
            wb = xv.load_workbook(bkp_path)
            self.assertGreater(len(wb.sheet_names), 0)
            sh = wb.get_sheet()
            self.assertIsNotNone(sh)


if __name__ == "__main__":
    unittest.main()
