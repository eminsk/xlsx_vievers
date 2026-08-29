"""
Unit and benchmark test suite for Excel Viewer Pro ASM Acceleration Engine.
"""

from __future__ import annotations

import unittest
import time
import sys
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
