"""
Unit and benchmark test suite for Excel Viewer Pro ASM Acceleration Engine.
"""

from __future__ import annotations

import sys
import time
import unittest
from pathlib import Path

# Add current directory to path
sys.path.insert(0, str(Path(__file__).parent.resolve()))
import asm_bridge


class TestAsmEngine(unittest.TestCase):
    def test_asm_availability(self):
        self.assertTrue(asm_bridge.asm_is_available(), "Native 64-bit Assembly DLL must be loaded")

    def test_vector_sum_avg_min_max(self):
        nums = [10.5, 20.25, 30.75, 40.0, 50.5, 60.0, 70.25, 80.0]
        # Sum: 362.25
        self.assertAlmostEqual(asm_bridge.asm_sum(nums), sum(nums), places=5)
        # Avg: 362.25 / 8 = 45.28125
        self.assertAlmostEqual(asm_bridge.asm_avg(nums), sum(nums) / len(nums), places=5)
        # Min: 10.5
        self.assertAlmostEqual(asm_bridge.asm_min(nums), 10.5, places=5)
        # Max: 80.0
        self.assertAlmostEqual(asm_bridge.asm_max(nums), 80.0, places=5)

    def test_vector_sumproduct(self):
        a = [1.0, 2.0, 3.0, 4.0, 5.0]
        b = [10.0, 20.0, 30.0, 40.0, 50.0]
        # 1*10 + 2*20 + 3*30 + 4*40 + 5*50 = 10 + 40 + 90 + 160 + 250 = 550
        res = asm_bridge.asm_sumproduct(a, b)
        self.assertAlmostEqual(res, 550.0, places=5)

    def test_financial_pmt_pv_fv(self):
        # PMT: Rate=0.05/12, Nper=360, PV=200000 -> approx -1073.64
        pmt = asm_bridge.asm_pmt(0.05 / 12, 360, 200000.0)
        self.assertAlmostEqual(pmt, -1073.6432, places=2)

        # PV: Rate=0.08/12, Nper=240, PMT=-500 -> approx 59777.15
        pv = asm_bridge.asm_pv(0.08 / 12, 240, -500.0)
        self.assertAlmostEqual(pv, 59777.15, places=1)

        # FV: Rate=0.05/12, Nper=120, PMT=-100, PV=-1000 -> approx 17175.24
        fv = asm_bridge.asm_fv(0.05 / 12, 120, -100.0, -1000.0)
        self.assertAlmostEqual(fv, 17175.24, places=1)

    def test_string_hash(self):
        h1 = asm_bridge.asm_str_hash("A1:Z100")
        h2 = asm_bridge.asm_str_hash("A1:Z100")
        h3 = asm_bridge.asm_str_hash("A1:Z101")
        self.assertEqual(h1, h2)
        self.assertNotEqual(h1, h3)
        self.assertIsInstance(h1, int)

    def test_gui_executable_exists_and_valid(self):
        exe_path = Path(__file__).parent / "xlsx_gui64.exe"
        self.assertTrue(exe_path.exists(), "xlsx_gui64.exe must exist")
        self.assertGreater(exe_path.stat().st_size, 30000, "Executable must be > 30KB")
        with open(exe_path, "rb") as f:
            header = f.read(2)
            self.assertEqual(header, b"MZ", "Must be valid Windows PE executable")

    def test_gui_mouse_and_keyboard_selection(self):
        import ctypes
        import subprocess

        user32 = ctypes.windll.user32
        exe_path = Path(__file__).parent / "xlsx_gui64.exe"
        proc = subprocess.Popen([str(exe_path)])
        time.sleep(1.0)
        hwnd_main = 0
        try:
            def cb_m(h, _):
                nonlocal hwnd_main
                cls = ctypes.create_unicode_buffer(256)
                user32.GetClassNameW(h, cls, 256)
                if "ExcelViewerPro_ASM64" in cls.value:
                    hwnd_main = h
                    return False
                return True

            user32.EnumWindows(ctypes.WINFUNCTYPE(ctypes.c_bool, ctypes.c_void_p, ctypes.c_void_p)(cb_m), 0)
            self.assertTrue(hwnd_main, "Main window must be found")

            h_grid, h_name, h_header = 0, 0, 0
            def cb_c(h, _):
                nonlocal h_grid, h_name, h_header
                cid = user32.GetWindowLongW(h, -12)
                cls = ctypes.create_unicode_buffer(256)
                user32.GetClassNameW(h, cls, 256)
                if cid == 1301: h_grid = h
                elif cid == 1201: h_name = h
                elif "header" in cls.value.lower(): h_header = h
                return True

            user32.EnumChildWindows(hwnd_main, ctypes.WINFUNCTYPE(ctypes.c_bool, ctypes.c_void_p, ctypes.c_void_p)(cb_c), 0)
            self.assertTrue(h_grid and h_name and h_header, "Child controls must be found")

            def gt(hw):
                l = user32.SendMessageW(hw, 0xE, 0, 0)
                b = ctypes.create_unicode_buffer(l + 2)
                user32.SendMessageW(hw, 0xD, l + 1, ctypes.byref(b))
                return b.value

            user32.SetForegroundWindow(hwnd_main)
            user32.SetFocus(h_grid)
            time.sleep(0.1)

            # 1. Navigation down and right
            user32.SendMessageW(h_grid, 0x100, 0x28, 0) # VK_DOWN
            user32.SendMessageW(h_grid, 0x101, 0x28, 0)
            user32.SendMessageW(h_grid, 0x100, 0x27, 0) # VK_RIGHT
            user32.SendMessageW(h_grid, 0x101, 0x27, 0)
            time.sleep(0.05)
            self.assertEqual(gt(h_name), "B2")

            # 2. Shift + Down and Shift + Right (2D range)
            user32.SendMessageW(h_grid, 0x100, 0x10, 0) # VK_SHIFT
            user32.SendMessageW(h_grid, 0x100, 0x28, 0) # VK_DOWN
            user32.SendMessageW(h_grid, 0x101, 0x28, 0)
            user32.SendMessageW(h_grid, 0x100, 0x27, 0) # VK_RIGHT
            user32.SendMessageW(h_grid, 0x101, 0x27, 0)
            user32.SendMessageW(h_grid, 0x101, 0x10, 0)
            time.sleep(0.05)
            self.assertEqual(gt(h_name), "B2:C3")

            # 3. Ctrl + A (Select All)
            user32.SendMessageW(h_grid, 0x100, 0x11, 0) # VK_CONTROL
            user32.SendMessageW(h_grid, 0x100, ord("A"), 0)
            user32.SendMessageW(h_grid, 0x101, ord("A"), 0)
            user32.SendMessageW(h_grid, 0x101, 0x11, 0)
            time.sleep(0.05)
            self.assertEqual(gt(h_name), "A1:Z30")

            # 4. Mouse click and drag
            user32.SendMessageW(h_grid, 0x201, 1, (50 << 16) | 120)
            user32.SendMessageW(h_grid, 0x200, 1, (110 << 16) | 250)
            user32.SendMessageW(h_grid, 0x202, 0, (110 << 16) | 250)
            time.sleep(0.05)
            self.assertTrue(":" in gt(h_name))

            # 5. Column Header Click
            user32.SendMessageW(h_header, 0x201, 1, (10 << 16) | 320)
            user32.SendMessageW(h_header, 0x202, 0, (10 << 16) | 320)
            time.sleep(0.05)
            self.assertEqual(gt(h_name), "C1:C30")

            # 6. Row Header Click (Col 0 in Grid at row 5)
            user32.SendMessageW(h_grid, 0x201, 1, (110 << 16) | 25)
            user32.SendMessageW(h_grid, 0x202, 0, (110 << 16) | 25)
            time.sleep(0.05)
            self.assertTrue(gt(h_name).startswith("A") and ":Z" in gt(h_name))
        finally:
            if hwnd_main:
                user32.PostMessageW(hwnd_main, 0x0010, 0, 0) # WM_CLOSE
            time.sleep(0.2)
            if proc.poll() is None:
                proc.kill()

    def test_gui_loaded_file_selection_and_tabs(self):
        import ctypes
        import subprocess

        user32 = ctypes.windll.user32
        exe_path = Path(__file__).parent / "xlsx_gui64.exe"
        sample_path = Path(__file__).parent.parent / "examples" / "sample_financial_plan.xlsx"
        if not sample_path.exists():
            return

        proc = subprocess.Popen([str(exe_path), str(sample_path)])
        time.sleep(1.2)
        hwnd_main = 0
        try:
            def cb_m(h, _):
                nonlocal hwnd_main
                cls = ctypes.create_unicode_buffer(256)
                user32.GetClassNameW(h, cls, 256)
                if "ExcelViewerPro_ASM64" in cls.value:
                    hwnd_main = h
                    return False
                return True

            user32.EnumWindows(ctypes.WINFUNCTYPE(ctypes.c_bool, ctypes.c_void_p, ctypes.c_void_p)(cb_m), 0)
            self.assertTrue(hwnd_main, "Main window must be found on loaded file")

            h_grid = user32.GetDlgItem(hwnd_main, 1301)
            h_name = user32.GetDlgItem(hwnd_main, 1201)
            h_ribbon = user32.GetDlgItem(hwnd_main, 1100)

            def gt(hw):
                l = user32.SendMessageW(hw, 0xE, 0, 0)
                b = ctypes.create_unicode_buffer(l + 2)
                user32.SendMessageW(hw, 0xD, l + 1, ctypes.byref(b))
                return b.value

            # 1. Selection must initialize to A1 on file open
            self.assertEqual(gt(h_name), "A1")

            # 2. Arrow keys immediately move selection
            user32.SendMessageW(h_grid, 0x100, 0x28, 0) # VK_DOWN
            user32.SendMessageW(h_grid, 0x101, 0x28, 0)
            time.sleep(0.05)
            self.assertEqual(gt(h_name), "A2")

            user32.SendMessageW(h_grid, 0x100, 0x27, 0) # VK_RIGHT
            user32.SendMessageW(h_grid, 0x101, 0x27, 0)
            time.sleep(0.05)
            self.assertEqual(gt(h_name), "B2")

            # 3. Shift + Down selects 2D range
            user32.SendMessageW(h_grid, 0x100, 0x10, 0) # VK_SHIFT
            user32.SendMessageW(h_grid, 0x100, 0x28, 0) # VK_DOWN
            user32.SendMessageW(h_grid, 0x101, 0x28, 0)
            user32.SendMessageW(h_grid, 0x101, 0x10, 0)
            time.sleep(0.05)
            self.assertEqual(gt(h_name), "B2:B3")

            # 4. Ribbon tab switching (Click Insert at 80, 10)
            lparam_ins = (10 << 16) | 80
            user32.SendMessageW(h_ribbon, 0x0201, 1, lparam_ins)
            user32.SendMessageW(h_ribbon, 0x0202, 0, lparam_ins)
            time.sleep(0.15)
            cur_tab = user32.SendMessageW(h_ribbon, 0x1300 + 11, 0, 0)
            self.assertEqual(cur_tab, 1)

            h_ins_btn = user32.GetDlgItem(hwnd_main, 1400)
            h_bold_btn = user32.GetDlgItem(hwnd_main, 1125)
            self.assertEqual(user32.IsWindowVisible(h_ins_btn), 1)
            self.assertEqual(user32.IsWindowVisible(h_bold_btn), 0)

            # Switch back to Home (Click at 25, 10)
            lparam_home = (10 << 16) | 25
            user32.SendMessageW(h_ribbon, 0x0201, 1, lparam_home)
            user32.SendMessageW(h_ribbon, 0x0202, 0, lparam_home)
            time.sleep(0.15)
            cur_tab0 = user32.SendMessageW(h_ribbon, 0x1300 + 11, 0, 0)
            self.assertEqual(cur_tab0, 0)
            self.assertEqual(user32.IsWindowVisible(h_bold_btn), 1)

            # 5. Bold command executes cleanly
            user32.SendMessageW(hwnd_main, 0x0111, 1125, 0) # ID_BTN_BOLD
            time.sleep(0.05)
        finally:
            if hwnd_main:
                user32.PostMessageW(hwnd_main, 0x0010, 0, 0)
            time.sleep(0.2)
            if proc.poll() is None:
                proc.kill()


def run_benchmark():
    print("=" * 70)
    print("  EXCEL VIEWER PRO — SIMD SSE2 x64 ASSEMBLY PERFORMANCE BENCHMARK")
    print("=" * 70)

    n = 1_000_000
    print(f"Generating array of {n:,} floating-point numbers...")
    data = [(i % 100) * 0.25 for i in range(n)]

    # 1. Pure Python sum()
    t0 = time.perf_counter()
    py_sum = sum(data)
    t_py = (time.perf_counter() - t0) * 1000.0

    # 2. Native x64 Assembly SIMD SSE2
    t0 = time.perf_counter()
    asm_sum_val = asm_bridge.asm_sum(data)
    t_asm = (time.perf_counter() - t0) * 1000.0

    print(f"\n[1] Pure Python sum():        {t_py:8.2f} ms | Result: {py_sum:.2f}")
    print(f"[2] Native x64 SIMD SSE2 Sum: {t_asm:8.2f} ms | Result: {asm_sum_val:.2f}")

    diff = abs(py_sum - asm_sum_val)
    print(f"\nPrecision Delta: {diff:g} (Exact match: {diff < 1e-5})")
    print(f"Throughput:      {(n / (t_asm / 1000.0)) / 1_000_000:.2f} Million Doubles / sec")
    print("=" * 70)


if __name__ == "__main__":
    # Run Unit Tests
    suite = unittest.TestLoader().loadTestsFromTestCase(TestAsmEngine)
    runner = unittest.TextTestRunner(verbosity=2)
    result = runner.run(suite)

    if result.wasSuccessful():
        print()
        run_benchmark()
