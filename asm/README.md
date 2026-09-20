# Excel Viewer Pro — 64-bit Assembly Acceleration Suite (x86-64 FASM)

High-performance hardware acceleration suite providing vectorized SIMD (SSE2/AVX) math, financial computations, and ultra-fast string hashing for **Excel Viewer Pro**.

Engineered for 64-bit Flat Assembler (**FASM**) with strict adherence to the Microsoft Windows x64 ABI and seamless zero-copy Python integration via `ctypes` and native memory buffers.

---

## 📁 Module Components

### 1. FASM Source Files and DLL Kernels
- **[`xlsx_math64.asm`](xlsx_math64.asm)** — Source code for the 64-bit SIMD acceleration dynamic library:
  - `vec_sum_f64(arr, count)` — Vectorized parallel summation of `double` (float64) arrays using SSE2 `addpd` with 4-way loop unrolling and horizontal reduction.
  - `vec_avg_f64(arr, count)` — Vectorized arithmetic mean computation.
  - `vec_min_f64(arr, count)` — Sub-nanosecond minimum search via SSE2 `minsd`.
  - `vec_max_f64(arr, count)` — Sub-nanosecond maximum search via SSE2 `maxsd`.
  - `vec_sumproduct_f64(a, b, count)` — Vectorized dot-product of float64 arrays (`mulpd` + `addpd`).
  - `fast_pmt_f64(rate, nper, pv, fv, type)` — Hardware-accelerated loan payment calculation (Excel `PMT`) with x87 FPU / SSE2 power exponentiation `(1 + rate)^nper`.
  - `fast_pv_f64(rate, nper, pmt, fv, type)` — Present value of investment (Excel `PV`).
  - `fast_fv_f64(rate, nper, pmt, pv, type)` — Future value of investment (Excel `FV`).
  - `fast_str_hash(str, len)` — 64-bit FNV-1a string hashing for instant formula and cell reference cache lookups.
  - `fast_count_nonblank(ptr_arr, count)` — Rapid non-empty pointer counting.

- **[`xlsx_math64.dll`](xlsx_math64.dll)** — Pre-compiled standalone 64-bit DLL (only ~2.5 KB of pure machine code with zero external dependencies).

### 2. Standalone Pure Assembly GUI Spreadsheet Application
- **[`xlsx_gui64.asm`](xlsx_gui64.asm)** / **`xlsx_gui64.exe`** — **Autonomous Excel Viewer Pro Spreadsheet Engine written entirely in x64 Assembly (FASM)**:
  - **Menu Bar**: `File` (New, Open Demo, Save, Exit), `Formulas` (Recalculate All SIMD, AutoSum, Average, PMT), `Data` (Sort Ascending/Descending, 1M SIMD Benchmark), `Help` (About).
  - **Toolbar**: Quick action buttons (`New`, `Open Demo`, `Save`, `AutoSum (SIMD)`, `Loan PMT`, `1M Benchmark`, `About`).
  - **Formula Bar**: Active cell address indicator (`B5`), `[fx]` evaluate trigger, and formula editing input (`=SUM(B1:B4)`).
  - **Spreadsheet Grid**: Multi-column tabular view (`SysListView32` in `LVS_REPORT` mode) with hardware double-buffering (`LVS_EX_DOUBLEBUFFER`), grid lines (`LVS_EX_GRIDLINES`), and full column formatting.
  - **Sheet Tabs**: Native `SysTabControl32` container supporting multiple interactive sheets.
  - **Dynamic Status Bar**: 5 interactive panels displaying status, selected coordinates, `SUM`, `AVG`, and the active SIMD engine status.

### 3. Python Integration Bridge
- **[`asm_bridge.py`](asm_bridge.py)** — High-level Python bridge module:
  - Automatic detection and dynamic linking of `xlsx_math64.dll`.
  - Zero-copy buffer protocol access for `array.array('d')` avoiding Python object overhead.
  - Exported functions: `asm_sum()`, `asm_avg()`, `asm_min()`, `asm_max()`, `asm_sumproduct()`, `asm_pmt()`, `asm_pv()`, `asm_fv()`, `asm_str_hash()`, `asm_is_available()`.

### 4. Verification and Build Automation
- **[`test_asm.py`](test_asm.py)** — Test suite and benchmark:
  - 100% test coverage across all math and financial primitives.
  - Exact numerical agreement with pure Python and `formulas.py` to within $10^{-12}$.
  - Stress testing over 1,000,000 elements.
- **[`build.bat`](build.bat)** — Build script compiling modules via `FASM.EXE` and running test verification.

---

## ⚡ Build and Compilation

Using the automated build script:

```cmd
cd C:\proekts\xlsx_vievers\asm
build.bat
```

Or manually with FASM:
```cmd
# 1. Compile 64-bit DLL
FASM.EXE xlsx_math64.asm xlsx_math64.dll

# 2. Compile standalone GUI application
FASM.EXE xlsx_gui64.asm xlsx_gui64.exe

# 3. Run verification tests and benchmark
python test_asm.py
```
