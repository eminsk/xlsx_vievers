"""
High-Performance Assembly Acceleration Bridge for Excel Viewer Pro.
Interfaces with 64-bit SIMD / FPU native assembly engine (xlsx_math64.dll)
and dynamic in-memory JIT machine code execution via Windows VirtualAlloc.
"""

from __future__ import annotations

import ctypes
import os
import sys
import math
from pathlib import Path
from typing import Sequence, Any, Iterable


# =============================================================================
# DLL / Native Library Loader
# =============================================================================

_DLL_PATH = Path(__file__).parent / "xlsx_math64.dll"
_dll: ctypes.CDLL | None = None
_is_available = False

if sys.maxsize > 2**32 and _DLL_PATH.exists():
    try:
        _dll = ctypes.CDLL(str(_DLL_PATH.resolve()))

        # Configure C ABI Signatures
        # 1. vec_sum_f64: (const double* arr, uint64_t count) -> double
        _dll.vec_sum_f64.restype = ctypes.c_double
        _dll.vec_sum_f64.argtypes = [ctypes.POINTER(ctypes.c_double), ctypes.c_uint64]

        # 2. vec_avg_f64: (const double* arr, uint64_t count) -> double
        _dll.vec_avg_f64.restype = ctypes.c_double
        _dll.vec_avg_f64.argtypes = [ctypes.POINTER(ctypes.c_double), ctypes.c_uint64]

        # 3. vec_min_f64: (const double* arr, uint64_t count) -> double
        _dll.vec_min_f64.restype = ctypes.c_double
        _dll.vec_min_f64.argtypes = [ctypes.POINTER(ctypes.c_double), ctypes.c_uint64]

        # 4. vec_max_f64: (const double* arr, uint64_t count) -> double
        _dll.vec_max_f64.restype = ctypes.c_double
        _dll.vec_max_f64.argtypes = [ctypes.POINTER(ctypes.c_double), ctypes.c_uint64]

        # 5. vec_sumproduct_f64: (const double* a, const double* b, uint64_t count) -> double
        _dll.vec_sumproduct_f64.restype = ctypes.c_double
        _dll.vec_sumproduct_f64.argtypes = [
            ctypes.POINTER(ctypes.c_double),
            ctypes.POINTER(ctypes.c_double),
            ctypes.c_uint64
        ]

        # 6. fast_pmt_f64: (double rate, double nper, double pv, double fv, int64_t type) -> double
        _dll.fast_pmt_f64.restype = ctypes.c_double
        _dll.fast_pmt_f64.argtypes = [
            ctypes.c_double, ctypes.c_double, ctypes.c_double, ctypes.c_double, ctypes.c_int64
        ]

        # 7. fast_pv_f64: (double rate, double nper, double pmt, double fv, int64_t type) -> double
        _dll.fast_pv_f64.restype = ctypes.c_double
        _dll.fast_pv_f64.argtypes = [
            ctypes.c_double, ctypes.c_double, ctypes.c_double, ctypes.c_double, ctypes.c_int64
        ]

        # 8. fast_fv_f64: (double rate, double nper, double pmt, double pv, int64_t type) -> double
        _dll.fast_fv_f64.restype = ctypes.c_double
        _dll.fast_fv_f64.argtypes = [
            ctypes.c_double, ctypes.c_double, ctypes.c_double, ctypes.c_double, ctypes.c_int64
        ]

        # 9. fast_str_hash: (const char* str, uint64_t len) -> uint64_t
        _dll.fast_str_hash.restype = ctypes.c_uint64
        _dll.fast_str_hash.argtypes = [ctypes.c_char_p, ctypes.c_uint64]

        # 10. fast_count_nonblank: (const uint64_t* ptr_arr, uint64_t count) -> uint64_t
        _dll.fast_count_nonblank.restype = ctypes.c_uint64
        _dll.fast_count_nonblank.argtypes = [ctypes.POINTER(ctypes.c_uint64), ctypes.c_uint64]

        _is_available = True
    except Exception as e:
        _is_available = False


def asm_is_available() -> bool:
    """Return True if native 64-bit Assembly engine is loaded and active."""
    return _is_available


import array

# =============================================================================
# Helper: Array conversions to Ctypes contiguous buffers
# =============================================================================

def _to_double_array(values: Sequence[Any]) -> tuple[Any, int]:
    """Convert Python numeric sequence to contiguous ctypes c_double array."""
    if isinstance(values, array.array) and values.typecode == 'd':
        addr, n = values.buffer_info()
        if n == 0:
            return (ctypes.c_double * 0)(), 0
        return ctypes.cast(addr, ctypes.POINTER(ctypes.c_double)), n

    clean_nums: list[float] = []
    for v in values:
        if v is None or v == "":
            continue
        try:
            clean_nums.append(float(v))
        except (ValueError, TypeError):
            pass

    n = len(clean_nums)
    if n == 0:
        return (ctypes.c_double * 0)(), 0

    arr = array.array('d', clean_nums)
    addr, _ = arr.buffer_info()
    # Keep reference to avoid GC during call
    ptr = ctypes.cast(addr, ctypes.POINTER(ctypes.c_double))
    ptr._keepalive = arr
    return ptr, n


# =============================================================================
# High-Level Vector Math Acceleration
# =============================================================================

def asm_sum(values: Sequence[Any]) -> float:
    """Ultra-fast SIMD SSE2 vector addition of numbers."""
    if _dll is not None:
        buf, n = _to_double_array(values)
        if n == 0:
            return 0.0
        return float(_dll.vec_sum_f64(buf, n))
    
    # Pure Python fallback
    nums = [float(v) for v in values if v is not None and v != ""]
    return sum(nums) if nums else 0.0


def asm_avg(values: Sequence[Any]) -> float:
    """Ultra-fast SIMD vector average calculation."""
    if _dll is not None:
        buf, n = _to_double_array(values)
        if n == 0:
            return 0.0
        return float(_dll.vec_avg_f64(buf, n))
    
    nums = [float(v) for v in values if v is not None and v != ""]
    return sum(nums) / len(nums) if nums else 0.0


def asm_min(values: Sequence[Any]) -> float:
    """Ultra-fast SIMD vector minimum finding."""
    if _dll is not None:
        buf, n = _to_double_array(values)
        if n == 0:
            return 0.0
        return float(_dll.vec_min_f64(buf, n))
    
    nums = [float(v) for v in values if v is not None and v != ""]
    return min(nums) if nums else 0.0


def asm_max(values: Sequence[Any]) -> float:
    """Ultra-fast SIMD vector maximum finding."""
    if _dll is not None:
        buf, n = _to_double_array(values)
        if n == 0:
            return 0.0
        return float(_dll.vec_max_f64(buf, n))
    
    nums = [float(v) for v in values if v is not None and v != ""]
    return max(nums) if nums else 0.0


def asm_sumproduct(values_a: Sequence[Any], values_b: Sequence[Any]) -> float:
    """Ultra-fast SIMD dot product (SUMPRODUCT)."""
    if _dll is not None:
        buf_a, n_a = _to_double_array(values_a)
        buf_b, n_b = _to_double_array(values_b)
        n = min(n_a, n_b)
        if n == 0:
            return 0.0
        return float(_dll.vec_sumproduct_f64(buf_a, buf_b, n))
    
    nums_a = [float(v) for v in values_a if v is not None and v != ""]
    nums_b = [float(v) for v in values_b if v is not None and v != ""]
    n = min(len(nums_a), len(nums_b))
    return sum(nums_a[i] * nums_b[i] for i in range(n))


# =============================================================================
# High-Level Financial Math Acceleration
# =============================================================================

def asm_pmt(rate: float, nper: float, pv: float, fv: float = 0.0, type_: int = 0) -> float:
    """Ultra-fast Excel PMT calculation via native FPU/SIMD."""
    if _dll is not None:
        return float(_dll.fast_pmt_f64(rate, nper, pv, fv, type_))
    
    if nper == 0:
        return 0.0
    if rate == 0:
        return -(pv + fv) / nper
    pvif = (1 + rate) ** nper
    pmt = (rate / (pvif - 1)) * -(pv * pvif + fv)
    if type_ == 1:
        pmt = pmt / (1 + rate)
    return pmt


def asm_pv(rate: float, nper: float, pmt: float, fv: float = 0.0, type_: int = 0) -> float:
    """Ultra-fast Excel PV calculation via native FPU/SIMD."""
    if _dll is not None:
        return float(_dll.fast_pv_f64(rate, nper, pmt, fv, type_))
    
    if rate == 0:
        return -(pmt * nper + fv)
    pvif = (1 + rate) ** nper
    fact = (1 + rate * type_) * (pvif - 1) / rate
    return -(fv + pmt * fact) / pvif


def asm_fv(rate: float, nper: float, pmt: float, pv: float = 0.0, type_: int = 0) -> float:
    """Ultra-fast Excel FV calculation via native FPU/SIMD."""
    if _dll is not None:
        return float(_dll.fast_fv_f64(rate, nper, pmt, pv, type_))
    
    if rate == 0:
        return -(pv + pmt * nper)
    pvif = (1 + rate) ** nper
    fact = (1 + rate * type_) * (pvif - 1) / rate
    return -(pv * pvif + pmt * fact)


def asm_str_hash(text: str) -> int:
    """64-bit FNV-1a String Hashing for fast cache indexing."""
    raw = text.encode("utf-8")
    if _dll is not None:
        return int(_dll.fast_str_hash(raw, len(raw)))
    
    # Pure Python FNV-1a 64-bit
    h = 14695981039346656037
    for byte in raw:
        h ^= byte
        h = (h * 1099511628211) & 0xFFFFFFFFFFFFFFFF
    return h
