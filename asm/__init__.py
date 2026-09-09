"""
Assembly SIMD SSE2 Acceleration for Excel Viewer.
"""

from .asm_bridge import (
    asm_avg,
    asm_fv,
    asm_is_available,
    asm_max,
    asm_min,
    asm_pmt,
    asm_pv,
    asm_str_hash,
    asm_sum,
    asm_sumproduct,
)

# Aliases for convenience
is_asm_available = asm_is_available
vec_sum_f64 = asm_sum
vec_avg_f64 = asm_avg
vec_min_f64 = asm_min
vec_max_f64 = asm_max
vec_sumproduct_f64 = asm_sumproduct
fast_pmt_f64 = asm_pmt

# Additional aliases matching README documentation
simd_sum = asm_sum
simd_avg = asm_avg
simd_min = asm_min
simd_max = asm_max
simd_sumproduct = asm_sumproduct

__all__ = [
    "asm_is_available",
    "is_asm_available",
    "asm_sum",
    "vec_sum_f64",
    "simd_sum",
    "asm_avg",
    "vec_avg_f64",
    "simd_avg",
    "asm_min",
    "vec_min_f64",
    "simd_min",
    "asm_max",
    "vec_max_f64",
    "simd_max",
    "asm_sumproduct",
    "vec_sumproduct_f64",
    "simd_sumproduct",
    "asm_pmt",
    "fast_pmt_f64",
    "asm_pv",
    "asm_fv",
    "asm_str_hash",
]
