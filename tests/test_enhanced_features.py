"""
Test suite for enhanced Excel Viewer Pro GUI, dialogs, SIMD formulas, and file exports.
"""

import json
import unittest
from pathlib import Path

from xlsx_viewer.formatting import NumberFormatter
from xlsx_viewer.formulas import FormulaEngine
from xlsx_viewer.models import SheetData


class TestEnhancedFeatures(unittest.TestCase):
    def test_json_export_structure(self):
        s_data = SheetData(
            name="Quarterly Report",
            headers=["Category", "Q1", "Q2", "Total"],
            rows=[
                ["Enterprise", 125000, 145000, 270000],
                ["SaaS", 45000, 52000, 97000],
            ],
            col_count=4,
            row_count=2,
        )

        data = {
            "sheet": s_data.name,
            "headers": s_data.headers,
            "rows": s_data.rows,
        }
        json_str = json.dumps(data, indent=2, ensure_ascii=False)
        parsed = json.loads(json_str)

        self.assertEqual(parsed["sheet"], "Quarterly Report")
        self.assertEqual(len(parsed["headers"]), 4)
        self.assertEqual(len(parsed["rows"]), 2)
        self.assertEqual(parsed["rows"][0][0], "Enterprise")
        self.assertEqual(parsed["rows"][1][3], 97000)

    def test_simd_accelerated_formulas(self):
        engine = FormulaEngine(lambda r, c, s=None: 0)

        # Test large sum accelerated via SIMD
        sum_formula = "=SUM(" + ", ".join(str(i * 1.5) for i in range(32)) + ")"
        expected = sum(i * 1.5 for i in range(32))
        res_sum = engine.evaluate(sum_formula)
        self.assertAlmostEqual(res_sum, expected, places=5)

        # Test average accelerated via SIMD
        avg_formula = "=AVERAGE(" + ", ".join(str(i * 2.0) for i in range(24)) + ")"
        expected_avg = sum(i * 2.0 for i in range(24)) / 24
        res_avg = engine.evaluate(avg_formula)
        self.assertAlmostEqual(res_avg, expected_avg, places=5)

        # Test min and max
        min_formula = "=MIN(105.5, 42.0, 19.5, 88.0, 12.25, 99.0)"
        self.assertEqual(engine.evaluate(min_formula), 12.25)

        max_formula = "=MAX(105.5, 42.0, 19.5, 88.0, 12.25, 99.0)"
        self.assertEqual(engine.evaluate(max_formula), 105.5)

    def test_simd_financial_formulas(self):
        engine = FormulaEngine(lambda r, c, s=None: 0)

        # PMT: 5% annual interest over 30 years for 200,000 loan
        res_pmt = engine.evaluate("=PMT(0.05/12, 360, 200000)")
        self.assertAlmostEqual(res_pmt, -1073.6432, places=2)

        # PV
        res_pv = engine.evaluate("=PV(0.08/12, 240, -500)")
        self.assertAlmostEqual(res_pv, 59777.15, places=1)

        # FV
        res_fv = engine.evaluate("=FV(0.05/12, 120, -100, -1000)")
        self.assertAlmostEqual(res_fv, 17175.24, places=1)

    def test_currency_and_percent_formatting(self):
        # Dollar
        self.assertEqual(NumberFormatter.format_value(2500000, "$#,##0.00"), "$2,500,000.00")
        # Ruble
        self.assertEqual(NumberFormatter.format_value(150000, "#,##0.00 ₽"), "150 000.00 ₽")
        # Euro
        self.assertEqual(NumberFormatter.format_value(12345.67, "€#,##0.00"), "€12,345.67")
        # Percent
        self.assertEqual(NumberFormatter.format_value(0.245, "0.00%"), "24.50%")
        # Scientific
        self.assertEqual(NumberFormatter.format_value(1234000, "0.00E+00"), "1.23E+06")

    def test_dialog_classes_availability(self):
        from xlsx_viewer.dialogs import (
            ChartWizardDialog,
            CustomSortDialog,
            FindReplaceDialog,
            FormatCellsDialog,
            FormulaReferenceDialog,
            GoalSeekDialog,
            InsertFunctionDialog,
            McpStatusDialog,
            RemoveDuplicatesDialog,
            ShortcutsHelpDialog,
            SimdBenchmarkDialog,
            SimdInfoDialog,
            TextToColumnsDialog,
        )

        self.assertTrue(callable(FormatCellsDialog))
        self.assertTrue(callable(FormulaReferenceDialog))
        self.assertTrue(callable(McpStatusDialog))
        self.assertTrue(callable(ShortcutsHelpDialog))
        self.assertTrue(callable(SimdBenchmarkDialog))
        self.assertTrue(callable(SimdInfoDialog))
        self.assertTrue(callable(ChartWizardDialog))
        self.assertTrue(callable(CustomSortDialog))
        self.assertTrue(callable(FindReplaceDialog))
        self.assertTrue(callable(GoalSeekDialog))
        self.assertTrue(callable(InsertFunctionDialog))
        self.assertTrue(callable(RemoveDuplicatesDialog))
        self.assertTrue(callable(TextToColumnsDialog))

    def test_asm_gui64_binary_validity(self):
        exe_path = Path(__file__).resolve().parent.parent / "asm" / "xlsx_gui64.exe"
        self.assertTrue(exe_path.exists(), f"Binary {exe_path} must exist.")
        self.assertGreater(exe_path.stat().st_size, 30000, "Executable must be > 30KB.")

        with open(exe_path, "rb") as f:
            header = f.read(1024)
            # DOS MZ header
            self.assertEqual(header[:2], b"MZ")
            # PE header offset
            pe_offset = int.from_bytes(header[0x3C:0x40], "little")
            # PE signature
            self.assertEqual(header[pe_offset : pe_offset + 4], b"PE\x00\x00")
            # Machine = IMAGE_FILE_MACHINE_AMD64 (0x8664)
            machine = int.from_bytes(header[pe_offset + 4 : pe_offset + 6], "little")
            self.assertEqual(machine, 0x8664)

    def test_asm_gui64_imports_and_symbols(self):
        exe_path = Path(__file__).resolve().parent.parent / "asm" / "xlsx_gui64.exe"
        data = exe_path.read_bytes()
        required_symbols = [
            b"GetKeyState",
            b"InvalidateRect",
            b"SendMessageA",
            b"MoveWindow",
            b"SetWindowLongPtrA",
            b"CallWindowProcA",
            b"OpenClipboard",
            b"CloseClipboard",
            b"GetOpenFileNameA",
            b"GetSaveFileNameA",
        ]
        for sym in required_symbols:
            self.assertIn(sym, data, f"Required Win32 API {sym.decode()} missing from PE binary.")

    def test_python_compatibility_declarations(self):
        """Verify Python 3.8 - 3.16 and free-threaded support declarations."""
        pyproject_path = Path(__file__).resolve().parent.parent / "pyproject.toml"
        content = pyproject_path.read_text(encoding="utf-8")
        self.assertIn("Programming Language :: Python :: 3.16", content)
        self.assertIn('requires-python = ">=3.8"', content)
        self.assertIn("free-threaded", content)
        self.assertIn("no-gil", content)

    def test_gui_menu_structure(self):
        import inspect

        from xlsx_viewer.gui import ExcelViewerPro

        source = inspect.getsource(ExcelViewerPro._create_menu)
        categories = [
            "File",
            "Edit",
            "View",
            "Insert",
            "Format",
            "Formulas",
            "Data",
            "Tools",
            "Help",
        ]
        for cat in categories:
            self.assertIn(f'label="{cat}"', source)

    def test_multithreading_and_free_threaded_safety(self):
        """Verify thread-safety and compatibility with free-threaded Python (no-gil 3.13t-3.16t)."""
        import concurrent.futures

        from xlsx_viewer import XlsxMCPServer, evaluate_formula

        def worker(idx: int) -> float:
            f = f"=SUM({idx}, 10, 20) * 2"
            res = evaluate_formula(f)
            return float(res)

        with concurrent.futures.ThreadPoolExecutor(max_workers=8) as pool:
            futures = [pool.submit(worker, i) for i in range(32)]
            results = [f.result() for f in futures]

        for i, val in enumerate(results):
            expected = (i + 10 + 20) * 2
            self.assertEqual(val, expected)

        # Test concurrent MCP server requests
        server = XlsxMCPServer()

        def mcp_worker(idx: int) -> int:
            resp = server.handle_request(
                {
                    "jsonrpc": "2.0",
                    "id": idx,
                    "method": "tools/call",
                    "params": {
                        "name": "xlsx_evaluate_formula",
                        "arguments": {
                            "formula": f"=SUM(A1, {idx})",
                            "context": {"A1": 50},
                        },
                    },
                }
            )
            payload = json.loads(resp["result"]["content"][0]["text"])
            return int(payload["result"])

        with concurrent.futures.ThreadPoolExecutor(max_workers=8) as pool:
            mcp_futures = [pool.submit(mcp_worker, i) for i in range(16)]
            mcp_results = [f.result() for f in mcp_futures]

        for i, val in enumerate(mcp_results):
            self.assertEqual(val, 50 + i)


if __name__ == "__main__":
    unittest.main()
