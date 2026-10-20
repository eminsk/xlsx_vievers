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

            # 6. Mouse click on Row 23 (A23) must select ONLY A23, NOT from A1
            user32.SendMessageW(h_grid, 0x0201, 1, (340 << 16) | 100) # LBUTTONDOWN at Row 23
            user32.SendMessageW(h_grid, 0x0202, 0, (340 << 16) | 100) # LBUTTONUP
            time.sleep(0.05)
            self.assertEqual(gt(h_name), "A23")

            # 7. Mouse drag from Row 11 to Row 16 must select A11:A16, NOT starting from A1
            user32.SendMessageW(h_grid, 0x0201, 1, (160 << 16) | 100) # LBUTTONDOWN at Row 11
            time.sleep(0.02)
            self.assertEqual(gt(h_name), "A11")
            user32.SendMessageW(h_grid, 0x0200, 1, (240 << 16) | 100) # MOUSEMOVE to Row 16
            time.sleep(0.02)
            self.assertEqual(gt(h_name), "A11:A16")
            user32.SendMessageW(h_grid, 0x0202, 0, (240 << 16) | 100) # LBUTTONUP
            time.sleep(0.05)
            self.assertEqual(gt(h_name), "A11:A16")
        finally:
            if hwnd_main:
                user32.PostMessageW(hwnd_main, 0x0010, 0, 0)
            time.sleep(0.2)
            if proc.poll() is None:
                proc.kill()

    def test_gui_icon_embedded_in_pe_and_taskbar(self):
        import ctypes
        from ctypes import wintypes
        import subprocess

        shell32 = ctypes.windll.shell32
        user32 = ctypes.windll.user32
        exe_path = Path(__file__).parent / "xlsx_gui64.exe"

        # 1. Verify PE Embedded Resource Icon
        h_icon_large = wintypes.HICON()
        h_icon_small = wintypes.HICON()
        num = shell32.ExtractIconExW(str(exe_path), 0, ctypes.byref(h_icon_large), ctypes.byref(h_icon_small), 1)
        self.assertGreater(num, 0, "xlsx_gui64.exe must have embedded PE icon")
        self.assertNotEqual(h_icon_large.value, 0, "Large icon handle must not be null")
        self.assertNotEqual(h_icon_small.value, 0, "Small icon handle must not be null")

        # 2. Verify Window Icon at Runtime
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

            # Check WM_GETICON or class icons
            WM_GETICON = 0x007F
            ICON_BIG = 1
            ICON_SMALL = 0
            GCLP_HICON = -14
            GCLP_HICONSM = -34

            h_big = user32.SendMessageW(hwnd_main, WM_GETICON, ICON_BIG, 0)
            if not h_big:
                GetClassLongPtr = getattr(user32, "GetClassLongPtrW", user32.GetClassLongW)
                h_big = GetClassLongPtr(hwnd_main, GCLP_HICON)

            self.assertNotEqual(h_big, 0, "Window must have big icon for taskbar / Alt+Tab")
        finally:
            if hwnd_main:
                user32.PostMessageW(hwnd_main, 0x0010, 0, 0)
            time.sleep(0.2)
            if proc.poll() is None:
                proc.kill()

    def test_ribbon_tab_styling_and_buttons(self):
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

            ctrls = {}
            ctrls_h = {}
            def cb_c(h, _):
                cid = user32.GetWindowLongW(h, -12)
                buf = ctypes.create_unicode_buffer(256)
                user32.GetWindowTextW(h, buf, 256)
                ctrls[cid] = buf.value
                ctrls_h[cid] = h
                return True

            user32.EnumChildWindows(hwnd_main, ctypes.WINFUNCTYPE(ctypes.c_bool, ctypes.c_void_p, ctypes.c_void_p)(cb_c), 0)

            # Check Insert Tab buttons exist and have polished labels
            self.assertIn(1401, ctrls) # ID_BTN_INS_TABLE
            self.assertEqual(ctrls[1401], "Table")
            self.assertIn(1402, ctrls) # ID_BTN_INS_PIVOT
            self.assertEqual(ctrls[1402], "PivotTable")

            # Check Data Tab buttons
            self.assertIn(1421, ctrls) # ID_BTN_DATA_CSV
            self.assertIn(ctrls[1421], ("From File", "From CSV"))
            self.assertIn(1424, ctrls) # ID_BTN_DATA_SORT_ASC
            self.assertEqual(ctrls[1424], "Sort A-Z")

            # Check Formulas Tab buttons
            self.assertIn(1441, ctrls) # ID_BTN_FORM_FX
            self.assertEqual(ctrls[1441], "Insert fx")
            self.assertIn(1442, ctrls) # ID_BTN_FORM_AUTOSUM
            self.assertEqual(ctrls[1442], "AutoSum")

            # Check View Tab buttons
            self.assertIn(1461, ctrls) # ID_BTN_VIEW_NORMAL
            self.assertEqual(ctrls[1461], "Normal")
            self.assertIn(1464, ctrls) # ID_BTN_VIEW_GRIDLINES
            self.assertEqual(ctrls[1464], "Gridlines")

            # Check BS_OWNERDRAW (0x000B) on modern ribbon buttons
            BS_OWNERDRAW = 0x000B
            for btn_id in [1401, 1421, 1441, 1461, 1111]:
                h_btn = ctrls_h.get(btn_id)
                self.assertIsNotNone(h_btn, f"Button {btn_id} must exist")
                style = user32.GetWindowLongW(h_btn, -16) # GWL_STYLE
                self.assertEqual(style & 0xF, BS_OWNERDRAW, f"Button {btn_id} must have BS_OWNERDRAW")

            # Check that tab switching does NOT corrupt active cell coordinates (e.g. 4294967286)
            h_tabs = ctrls_h.get(1100) # ID_RIBBON_TABS
            h_name = ctrls_h.get(1201) # ID_EDIT_NAME_BOX
            if h_tabs and h_name:
                user32.SendMessageW(h_tabs, 0x130C, 1, 0) # TCM_SETCURSEL
                buf_name = ctypes.create_unicode_buffer(64)
                user32.GetWindowTextW(h_name, buf_name, 64)
                self.assertNotIn("42949672", buf_name.value)
        finally:
            if hwnd_main:
                user32.PostMessageW(hwnd_main, 0x0010, 0, 0)
            time.sleep(0.2)
            if proc.poll() is None:
                proc.kill()

    def test_gui_formula_display_and_editing(self):
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
            self.assertTrue(hwnd_main, "Main window must be found")

            h_grid = user32.GetDlgItem(hwnd_main, 1301)
            h_name = user32.GetDlgItem(hwnd_main, 1201)
            h_form = user32.GetDlgItem(hwnd_main, 1203)
            self.assertTrue(h_grid and h_name and h_form, "Controls must exist")

            def gt_a(hw):
                l = user32.SendMessageA(hw, 0xE, 0, 0)
                b = ctypes.create_string_buffer(l + 2)
                user32.SendMessageA(hw, 0xD, l + 1, ctypes.byref(b))
                return b.value

            def gt_w(hw):
                l = user32.SendMessageW(hw, 0xE, 0, 0)
                b = ctypes.create_unicode_buffer(l + 2)
                user32.SendMessageW(hw, 0xD, l + 1, ctypes.byref(b))
                return b.value

            def nav_to(target_row, target_col):
                user32.SendMessageW(h_grid, 0x100, 0x24, 0) # Home -> Col A
                user32.SendMessageW(h_grid, 0x101, 0x24, 0)
                for _ in range(5):
                    user32.SendMessageW(h_grid, 0x100, 0x21, 0) # PageUp -> Row 0
                    user32.SendMessageW(h_grid, 0x101, 0x21, 0)
                for _ in range(target_row):
                    user32.SendMessageW(h_grid, 0x100, 0x28, 0) # Down
                    user32.SendMessageW(h_grid, 0x101, 0x28, 0)
                for _ in range(target_col - 1):
                    user32.SendMessageW(h_grid, 0x100, 0x27, 0) # Right
                    user32.SendMessageW(h_grid, 0x101, 0x27, 0)
                time.sleep(0.04)

            # 1. A1 initial display has cell text and no formula '='
            self.assertEqual(gt_w(h_name), "A1")
            raw_a1 = gt_a(h_form)
            self.assertFalse(raw_a1.startswith(b"="), "A1 should be plain text, not a formula")

            # 2. C22 has underlying formula ='Форекс (5 дней)'!D184
            nav_to(21, 3)
            self.assertEqual(gt_w(h_name), "C22")
            form_c22 = gt_a(h_form)
            self.assertTrue(form_c22.startswith(b"="), "C22 must have formula starting with '='")
            self.assertIn(b"!D184", form_c22, "C22 formula must reference cell D184")
            self.assertEqual(form_c22.decode("cp1251", errors="replace"), "='Форекс (5 дней)'!D184")

            # 3. B27 has formula =SUM(B15:B26)
            nav_to(26, 2)
            self.assertEqual(gt_w(h_name), "B27")
            form_b27 = gt_a(h_form)
            self.assertEqual(form_b27.decode("cp1251", errors="replace"), "=SUM(B15:B26)")

            # 4. D27 has formula =MAX(D15:D26)
            nav_to(26, 4)
            self.assertEqual(gt_w(h_name), "D27")
            form_d27 = gt_a(h_form)
            self.assertEqual(form_d27.decode("cp1251", errors="replace"), "=MAX(D15:D26)")

            # 5. Enter formula =B21+B22 in cell B26
            nav_to(25, 2)
            self.assertEqual(gt_w(h_name), "B26")
            user32.SendMessageA(h_form, 0xC, 0, b"=B21+B22")
            user32.SendMessageW(h_form, 0x100, 0x0D, 0) # Enter
            user32.SendMessageW(h_form, 0x101, 0x0D, 0)
            time.sleep(0.05)
            self.assertEqual(gt_a(h_form), b"=B21+B22")

            # Move away to B27, then move back up to B26 to verify formula persistence
            user32.SendMessageW(h_grid, 0x100, 0x28, 0) # Down to B27
            user32.SendMessageW(h_grid, 0x101, 0x28, 0)
            time.sleep(0.04)
            self.assertEqual(gt_w(h_name), "B27")
            self.assertEqual(gt_a(h_form), b"=SUM(B15:B26)")

            user32.SendMessageW(h_grid, 0x100, 0x26, 0) # Up to B26
            user32.SendMessageW(h_grid, 0x101, 0x26, 0)
            time.sleep(0.04)
            self.assertEqual(gt_w(h_name), "B26")
            self.assertEqual(gt_a(h_form), b"=B21+B22")

            # 6. Overwrite with plain number 999
            user32.SendMessageA(h_form, 0xC, 0, b"999")
            user32.SendMessageW(h_form, 0x100, 0x0D, 0) # Enter
            user32.SendMessageW(h_form, 0x101, 0x0D, 0)
            time.sleep(0.05)
            self.assertEqual(gt_a(h_form), b"999")

            # Move away and back to verify plain text persistence and formula cleared
            user32.SendMessageW(h_grid, 0x100, 0x28, 0) # Down to B27
            user32.SendMessageW(h_grid, 0x101, 0x28, 0)
            time.sleep(0.04)
            user32.SendMessageW(h_grid, 0x100, 0x26, 0) # Up to B26
            user32.SendMessageW(h_grid, 0x101, 0x26, 0)
            time.sleep(0.04)
            self.assertEqual(gt_a(h_form), b"999")
        finally:
            if hwnd_main:
                user32.PostMessageW(hwnd_main, 0x0010, 0, 0)
            time.sleep(0.2)
            if proc.poll() is None:
                proc.kill()

    def test_gui_multi_sheet_tabs_and_switching(self):
        import ctypes
        import subprocess

        user32 = ctypes.windll.user32
        exe_path = Path(__file__).parent / "xlsx_gui64.exe"
        sample_path = Path(__file__).parent.parent / "examples" / "sample_financial_plan.xlsx"
        if not sample_path.exists():
            return

        proc = subprocess.Popen([str(exe_path), str(sample_path)])
        time.sleep(1.5)
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

            h_grid = user32.GetDlgItem(hwnd_main, 1301)
            h_tabs = user32.GetDlgItem(hwnd_main, 1302)
            h_name = user32.GetDlgItem(hwnd_main, 1201)
            h_form = user32.GetDlgItem(hwnd_main, 1203)
            self.assertTrue(h_grid and h_tabs, "Grid and Sheet tabs must exist")

            def click_tab(x):
                lparam = (12 << 16) | x
                user32.SendMessageW(h_tabs, 0x0201, 1, lparam)
                user32.SendMessageW(h_tabs, 0x0202, 0, lparam)
                time.sleep(0.25)

            def gt_w(hw):
                l = user32.SendMessageW(hw, 0xE, 0, 0)
                b = ctypes.create_unicode_buffer(l + 2)
                user32.SendMessageW(hw, 0xD, l + 1, ctypes.byref(b))
                return b.value

            def gt_a(hw):
                l = user32.SendMessageA(hw, 0xE, 0, 0)
                b = ctypes.create_string_buffer(l + 2)
                user32.SendMessageA(hw, 0xD, l + 1, ctypes.byref(b))
                return b.value

            def nav_to(target_row, target_col):
                user32.SendMessageW(h_grid, 0x100, 0x24, 0) # Home -> Col A
                user32.SendMessageW(h_grid, 0x101, 0x24, 0)
                for _ in range(10):
                    user32.SendMessageW(h_grid, 0x100, 0x21, 0) # PageUp -> Row 0
                    user32.SendMessageW(h_grid, 0x101, 0x21, 0)
                for _ in range(target_row):
                    user32.SendMessageW(h_grid, 0x100, 0x28, 0) # Down
                    user32.SendMessageW(h_grid, 0x101, 0x28, 0)
                for _ in range(target_col - 1):
                    user32.SendMessageW(h_grid, 0x100, 0x27, 0) # Right
                    user32.SendMessageW(h_grid, 0x101, 0x27, 0)
                time.sleep(0.05)

            # 1. Sheet 0 is initially loaded (27 rows)
            count0 = user32.SendMessageW(h_grid, 0x1004, 0, 0)
            self.assertEqual(count0, 27, "Sheet 0 must have 27 rows")
            self.assertEqual(gt_w(h_name), "A1")

            # 2. Click Tab 1 (Forex) -> 308 rows
            click_tab(150)
            count1 = user32.SendMessageW(h_grid, 0x1004, 0, 0)
            cur_tab1 = user32.SendMessageW(h_tabs, 0x1300 + 11, 0, 0)
            self.assertEqual(cur_tab1, 1, "Tab 1 must be active")
            self.assertEqual(count1, 308, "Sheet 1 must have 308 rows")
            self.assertEqual(gt_w(h_name), "A1")
            txt1 = gt_a(h_form).decode("cp1251", errors="replace")
            self.assertIn("ФОРЕКС", txt1.upper())

            # Navigate to D6 on Sheet 1 (formula: ='Дашборд и Сводка'!$C$5)
            nav_to(5, 4)
            self.assertEqual(gt_w(h_name), "D6")
            form_d6 = gt_a(h_form).decode("cp1251", errors="replace")
            self.assertEqual(form_d6, "='Дашборд и Сводка'!$C$5")

            # Navigate to E6 on Sheet 1 (formula: =MIN('Дашборд и Сводка'!$C$9, ...))
            nav_to(5, 5)
            self.assertEqual(gt_w(h_name), "E6")
            form_e6 = gt_a(h_form).decode("cp1251", errors="replace")
            self.assertTrue(form_e6.startswith("=MIN("))

            # 3. Click Tab 2 (Crypto) -> 415 rows
            click_tab(250)
            count2 = user32.SendMessageW(h_grid, 0x1004, 0, 0)
            cur_tab2 = user32.SendMessageW(h_tabs, 0x1300 + 11, 0, 0)
            self.assertEqual(cur_tab2, 2, "Tab 2 must be active")
            self.assertEqual(count2, 415, "Sheet 2 must have 415 rows")
            self.assertEqual(gt_w(h_name), "A1")
            txt2 = gt_a(h_form).decode("cp1251", errors="replace")
            self.assertIn("КРИПТОВАЛЮТА", txt2.upper())

            # Navigate to D37 on Sheet 2 (formula: =D6)
            nav_to(36, 4)
            self.assertEqual(gt_w(h_name), "D37")
            form_d37 = gt_a(h_form).decode("cp1251", errors="replace")
            self.assertEqual(form_d37, "=D6")

            # 4. Click back to Tab 0 (Dashboard) -> 27 rows
            click_tab(50)
            count0_back = user32.SendMessageW(h_grid, 0x1004, 0, 0)
            cur_tab0 = user32.SendMessageW(h_tabs, 0x1300 + 11, 0, 0)
            self.assertEqual(cur_tab0, 0, "Tab 0 must be active again")
            self.assertEqual(count0_back, 27, "Sheet 0 must return to having 27 rows")
            self.assertEqual(gt_w(h_name), "A1")
            txt0 = gt_a(h_form).decode("cp1251", errors="replace")
            self.assertIn("ПЛАН", txt0.upper())
        finally:
            if hwnd_main:
                user32.PostMessageW(hwnd_main, 0x0010, 0, 0)
            time.sleep(0.2)
            if proc.poll() is None:
                proc.kill()

    def test_gui_fx_wizard_129_functions_and_evaluation(self):
        import ctypes
        from ctypes import wintypes
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

            # 1. Open FX wizard via ID_BTN_FORM_FX (1441)
            user32.SendMessageW(hwnd_main, 0x0111, 1441, 0)
            time.sleep(0.4)

            hwnd_wiz = 0
            def cb_w(h, _):
                nonlocal hwnd_wiz
                cls = ctypes.create_unicode_buffer(256)
                user32.GetClassNameW(h, cls, 256)
                if "FxWizardClass" in cls.value:
                    hwnd_wiz = h
                    return False
                return True

            user32.EnumWindows(ctypes.WINFUNCTYPE(ctypes.c_bool, ctypes.c_void_p, ctypes.c_void_p)(cb_w), 0)
            self.assertTrue(hwnd_wiz, "FxWizard dialog window must be opened")

            h_list = user32.GetDlgItem(hwnd_wiz, 2001)
            h_ok = user32.GetDlgItem(hwnd_wiz, 2002)
            h_cancel = user32.GetDlgItem(hwnd_wiz, 2003)
            h_search = user32.GetDlgItem(hwnd_wiz, 2004)
            h_cat = user32.GetDlgItem(hwnd_wiz, 2006)

            self.assertTrue(h_list and h_ok and h_cancel and h_search and h_cat, "All wizard controls must exist")

            # 2. Check 129 functions in "All" category
            count_all = user32.SendMessageW(h_list, 0x018B, 0, 0) # LB_GETCOUNT
            self.assertEqual(count_all, 129, "Wizard must have all 129 Excel functions loaded")

            # 3. Filter by Financial category (cat index 7)
            user32.SendMessageW(h_cat, 0x014E, 7, 0) # CB_SETCURSEL = 7
            user32.SendMessageW(hwnd_wiz, 0x0111, (1 << 16) | 2006, h_cat) # CBN_SELCHANGE
            time.sleep(0.1)
            count_fin = user32.SendMessageW(h_list, 0x018B, 0, 0)
            self.assertEqual(count_fin, 5, "Financial category must contain 5 financial functions (FV, NPER, PMT, PV, RATE)")

            # 4. Search for "COUNT"
            user32.SendMessageW(h_cat, 0x014E, 0, 0) # Back to All
            user32.SendMessageW(hwnd_wiz, 0x0111, (1 << 16) | 2006, h_cat)
            user32.SendMessageA(h_search, 0x000C, 0, b"COUNT")
            user32.SendMessageW(hwnd_wiz, 0x0111, (0x0300 << 16) | 2004, h_search) # EN_CHANGE
            time.sleep(0.1)
            count_cnt = user32.SendMessageW(h_list, 0x018B, 0, 0)
            self.assertGreaterEqual(count_cnt, 5, "Search for COUNT should find COUNT, COUNTA, COUNTIF, etc.")

            # 5. Insert selected function
            user32.SendMessageW(hwnd_wiz, 0x0111, 2002, h_ok) # Click Insert
            time.sleep(0.2)

            # Verify formula inserted in formula bar
            h_form = user32.GetDlgItem(hwnd_main, 1203)
            def gt_a(hw):
                l = user32.SendMessageA(hw, 0xE, 0, 0)
                b = ctypes.create_string_buffer(l + 2)
                user32.SendMessageA(hw, 0xD, l + 1, ctypes.byref(b))
                return b.value

            form_val = gt_a(h_form)
            self.assertTrue(form_val.startswith(b"="), "Inserted formula must start with '='")
            self.assertIn(b"COUNT", form_val)

            # 6. Test formula entry and persistence: ABS, SQRT, PI
            user32.SendMessageA(h_form, 0x000C, 0, b"=ABS(-125.5)")
            user32.SendMessageW(h_form, 0x0100, 0x0D, 0) # Enter
            user32.SendMessageW(h_form, 0x0101, 0x0D, 0)
            time.sleep(0.1)
            self.assertEqual(gt_a(h_form), b"=ABS(-125.5)")

            # Test SQRT(144)
            user32.SendMessageA(h_form, 0x000C, 0, b"=SQRT(144)")
            user32.SendMessageW(h_form, 0x0100, 0x0D, 0)
            user32.SendMessageW(h_form, 0x0101, 0x0D, 0)
            time.sleep(0.1)
            self.assertEqual(gt_a(h_form), b"=SQRT(144)")

            # Test PI()
            user32.SendMessageA(h_form, 0x000C, 0, b"=PI()")
            user32.SendMessageW(h_form, 0x0100, 0x0D, 0)
            user32.SendMessageW(h_form, 0x0101, 0x0D, 0)
            time.sleep(0.1)
            self.assertEqual(gt_a(h_form), b"=PI()")
        finally:
            if hwnd_main:
                user32.PostMessageW(hwnd_main, 0x0010, 0, 0)
            time.sleep(0.2)
            if proc.poll() is None:
                proc.kill()

    def test_gui_font_selection_per_cell(self):
        import ctypes
        import subprocess

        user32 = ctypes.windll.user32
        user32.GetDlgItem.restype = ctypes.c_void_p
        user32.GetDlgItem.argtypes = [ctypes.c_void_p, ctypes.c_int]
        exe_path = Path(__file__).parent / "xlsx_gui64.exe"
        proc = subprocess.Popen([str(exe_path)])

        hwnd_main = 0
        def enum_cb(h, _):
            nonlocal hwnd_main
            cls = ctypes.create_unicode_buffer(256)
            user32.GetClassNameW(h, cls, 256)
            if "ExcelViewerPro_ASM64" in cls.value:
                hwnd_main = h
                return False
            return True

        for _ in range(30):
            user32.EnumWindows(ctypes.WINFUNCTYPE(ctypes.c_bool, ctypes.c_void_p, ctypes.c_void_p)(enum_cb), 0)
            if hwnd_main:
                break
            time.sleep(0.1)

        self.assertTrue(hwnd_main, "Main window must be found")
        try:
            h_grid = user32.GetDlgItem(hwnd_main, 1301)
            h_name = user32.GetDlgItem(hwnd_main, 1201)
            h_combo_size = user32.GetDlgItem(hwnd_main, 1122)
            h_combo_font = user32.GetDlgItem(hwnd_main, 1121)
            self.assertTrue(h_grid and h_name and h_combo_size and h_combo_font)

            def get_text(h):
                buf = ctypes.create_unicode_buffer(256)
                user32.SendMessageW(h, 0x000D, 256, ctypes.byref(buf))
                return buf.value

            def get_combo_sel(h):
                return user32.SendMessageW(h, 0x0147, 0, 0)

            self.assertEqual(get_text(h_name), "A1")
            self.assertEqual(get_combo_sel(h_combo_size), 2) # default 10pt

            # 1. Increase font size of cell A1 (A+)
            user32.SendMessageW(hwnd_main, 0x0111, 1123, 0) # ID_BTN_INC_FONT
            time.sleep(0.08)
            self.assertEqual(get_combo_sel(h_combo_size), 4) # 12pt

            # 2. Change font family of cell A1 to Calibri (index 1)
            user32.SendMessageW(h_combo_font, 0x014E, 1, 0)
            user32.SendMessageW(hwnd_main, 0x0111, (1 << 16) | 1121, h_combo_font)
            time.sleep(0.08)

            # 3. Move to cell A2: must have default 10pt and Segoe UI!
            user32.SendMessageW(h_grid, 0x0100, 0x28, 0) # Down
            user32.SendMessageW(h_grid, 0x0101, 0x28, 0)
            time.sleep(0.08)
            self.assertEqual(get_text(h_name), "A2")
            self.assertEqual(get_combo_sel(h_combo_size), 2) # default 10pt
            self.assertEqual(get_combo_sel(h_combo_font), 0) # default Segoe UI

            # 4. Move to cell B2: must also have default 10pt and Segoe UI!
            user32.SendMessageW(h_grid, 0x0100, 0x27, 0) # Right
            user32.SendMessageW(h_grid, 0x0101, 0x27, 0)
            time.sleep(0.08)
            self.assertEqual(get_text(h_name), "B2")
            self.assertEqual(get_combo_sel(h_combo_size), 2)
            self.assertEqual(get_combo_sel(h_combo_font), 0)

            # 5. Move back to cell A1: custom 12pt and Calibri restored!
            user32.SendMessageW(h_grid, 0x0100, 0x25, 0) # Left
            user32.SendMessageW(h_grid, 0x0101, 0x25, 0)
            time.sleep(0.04)
            user32.SendMessageW(h_grid, 0x0100, 0x26, 0) # Up
            user32.SendMessageW(h_grid, 0x0101, 0x26, 0)
            time.sleep(0.04)
            self.assertEqual(get_text(h_name), "A1")
            self.assertEqual(get_combo_sel(h_combo_size), 4) # 12pt
            self.assertEqual(get_combo_sel(h_combo_font), 1) # Calibri
        finally:
            if hwnd_main:
                user32.PostMessageW(hwnd_main, 0x0010, 0, 0)
            time.sleep(0.2)
            if proc.poll() is None:
                proc.kill()

    def test_gui_color_buttons_and_universal_file_menu(self):
        import ctypes
        import subprocess

        exe_path = Path(__file__).parent / "xlsx_gui64.exe"
        self.assertTrue(exe_path.exists())

        # 1. Verify PE Binary contains universal file menu strings and ChooseColorA import
        pe_bytes = exe_path.read_bytes()
        self.assertIn(b"&Open File (Excel, CSV)...", pe_bytes)
        self.assertIn(b"&Save File\tCtrl+S", pe_bytes)
        self.assertIn(b"Save &As...\tCtrl+Shift+S", pe_bytes)
        self.assertIn(b"ChooseColorA", pe_bytes)

        # 2. Verify live UI controls for Fill Color and Text Color
        user32 = ctypes.windll.user32
        user32.GetDlgItem.restype = ctypes.c_void_p
        user32.GetDlgItem.argtypes = [ctypes.c_void_p, ctypes.c_int]
        user32.GetWindowLongW.restype = ctypes.c_long
        user32.GetWindowLongW.argtypes = [ctypes.c_void_p, ctypes.c_int]

        proc = subprocess.Popen([str(exe_path)])
        hwnd_main = 0

        def enum_cb(h, _):
            nonlocal hwnd_main
            cls = ctypes.create_unicode_buffer(256)
            user32.GetClassNameW(h, cls, 256)
            if "ExcelViewerPro_ASM64" in cls.value:
                hwnd_main = h
                return False
            return True

        try:
            for _ in range(30):
                user32.EnumWindows(ctypes.WINFUNCTYPE(ctypes.c_bool, ctypes.c_void_p, ctypes.c_void_p)(enum_cb), 0)
                if hwnd_main:
                    break
                time.sleep(0.1)

            self.assertTrue(hwnd_main, "Main window must be found")

            # ID_BTN_FILL_COLOR = 1130, ID_BTN_TEXT_COLOR = 1131
            h_fill = user32.GetDlgItem(hwnd_main, 1130)
            h_text = user32.GetDlgItem(hwnd_main, 1131)
            self.assertTrue(h_fill, "Fill Color button must exist")
            self.assertTrue(h_text, "Text Color button must exist")

            # Verify buttons have BS_OWNERDRAW style (0x0B)
            GWL_STYLE = -16
            style_fill = user32.GetWindowLongW(h_fill, GWL_STYLE)
            style_text = user32.GetWindowLongW(h_text, GWL_STYLE)
            self.assertEqual(style_fill & 0x0F, 0x0B, "Fill Color button must be BS_OWNERDRAW")
            self.assertEqual(style_text & 0x0F, 0x0B, "Text Color button must be BS_OWNERDRAW")
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
