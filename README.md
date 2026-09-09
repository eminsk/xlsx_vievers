# xlsx-viewer-pro (Excel Viewer Pro)

[![PyPI version](https://img.shields.io/pypi/v/xlsx-viewer-pro.svg)](https://pypi.org/project/xlsx-viewer-pro/)
[![Python Versions](https://img.shields.io/badge/python-3.10%20%7C%203.11%20%7C%203.12%20%7C%203.13%20%7C%203.14%20%7C%203.15-blue.svg)](https://pypi.org/project/xlsx-viewer-pro/)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
[![CI](https://github.com/eminsk/xlsx_vievers/actions/workflows/ci.yml/badge.svg)](https://github.com/eminsk/xlsx_vievers/actions)
[![Downloads](https://static.pepy.tech/badge/xlsx-viewer-pro)](https://pepy.tech/project/xlsx-viewer-pro)

A high-performance Python spreadsheet library, headless formula engine (129+ functions), native x64 SIMD SSE2 math engine, and modern Office Ribbon desktop application for `.xlsx`, `.xlsm`, `.csv`, and `.tsv` files. Compatible with **Python 3.10, 3.11, 3.12, 3.13, 3.14, and Python 3.15 (RC & Preview)**.

---

## ⚡ Key Highlights

- 🧮 **Headless Excel Formula Engine**: Evaluate Excel formulas in pure Python without launching any UI. Supports **129 functions** across Math, Trigonometry, Statistics, Finance (`PMT`, `PV`, `FV`, `NPV`, `IRR`), Lookups (`VLOOKUP`, `XLOOKUP`, `INDEX`, `MATCH`), Text, Logic, and Date/Time.
- 🚀 **Hardware SIMD SSE2 Acceleration**: Built-in 64-bit native assembly engine (`xlsx_math64.dll`) delivering up to **7.4+ million double-precision operations per second** with automatic pure-Python fallback on non-Windows/ARM platforms.
- 🛠️ **Dual-Use Architecture**: Use it as a lightweight Python library, a command-line tool (`xlsx-viewer-pro` / `xlsx-viewer` / `xv`), or a full-featured desktop spreadsheet application.
- 🎨 **Modern Desktop GUI**: Office Ribbon UI with dynamic real-time status bar statistics (`SUM`, `AVERAGE`, `COUNT`, `MIN`, `MAX`), 10+ chart types, conditional formatting, auto-filter, and goal seek.
- 🧩 **Zero Heavy Spreadsheet Engine Dependencies**: No LibreOffice, COM, or Excel installation required.

---

## 📦 Installation

Add to your project with `uv` (recommended):

```bash
uv add xlsx-viewer-pro
```

Or install via standard `pip`:

```bash
pip install xlsx-viewer-pro
```

You can also run the CLI instantly without installing using `uvx`:

```bash
uvx xlsx-viewer-pro --calc "=SUM(10, 20, 30) * 2"
```

#### ⚡ Zero C-Build Requirements & Python 3.15 Ready
`xlsx-viewer-pro` installs in milliseconds with **zero C build requirements** by default (`openpyxl` + `ttkbootstrap`), and bundles pre-compiled standalone Flat Assembler (FASM) 64-bit DLLs for SIMD SSE2 math.

* **Standard install (CLI, desktop viewer, 129 formula functions, SIMD SSE2):**
  ```bash
  uv add xlsx-viewer-pro
  # or
  pip install xlsx-viewer-pro
  ```
* **Optional charting support (Matplotlib Chart Wizard):**
  ```bash
  uv add "xlsx-viewer-pro[charts]"
  # or
  pip install "xlsx-viewer-pro[charts]"
  ```

#### 🧵 Free-Threaded (No-GIL / PEP 703) Verified
`xlsx-viewer-pro` is **100% verified on Python 3.13t, 3.14t, and 3.15t without GIL** (`-X gil=0`). All 129 formula functions and the native x64 SIMD SSE2 FASM engine are fully thread-safe, enabling parallel formula evaluation across all CPU cores without GIL bottlenecks via standard `ThreadPoolExecutor`.

---

## 🐍 Python Library Quickstart

### 1. Headless Formula Evaluation

Evaluate any Excel formula directly from Python:

```python
from xlsx_viewer import evaluate_formula

# Basic arithmetic & math
result = evaluate_formula("=SUM(10, 20, 30) * 2")
print(result)  # 120.0

# Financial formulas
monthly_payment = evaluate_formula("=PMT(0.05 / 12, 360, -250000)")
print(f"Monthly Mortgage: ${monthly_payment:.2f}")  # $1342.05

# Statistical & Logic
val = evaluate_formula("=IF(AVERAGE(85, 90, 92) >= 90, 'Honors', 'Standard')")
print(val)  # 'Honors'
```

### 2. Evaluating Formulas with Cell Context

Pass cell coordinate mappings directly using standard Excel notation:

```python
from xlsx_viewer import evaluate_formula

context = {
    "A1": 15000,
    "A2": 3200,
    "B1": 0.15,
    "B2": "Tier-1",
}

# Reference cells directly
net_profit = evaluate_formula("=(A1 - A2) * (1 - B1)", context=context)
print(f"Net Profit: ${net_profit:.2f}")  # $10030.00

# Text and conditional logic with context
status = evaluate_formula(
    '=IF(A1 > 10000, CONCAT(B2, " - High Volume"), "Normal")',
    context=context,
)
print(status)  # 'Tier-1 - High Volume'
```

### 3. Using the `FormulaEngine` Class

For dynamic resolution or integrating with custom data sources:

```python
from xlsx_viewer import FormulaEngine

# Custom resolver: fn(row: int, col: int, sheet: str | None) -> value
matrix = [
    [10, 20, 30],
    [40, 50, 60],
]


def custom_resolver(row: int, col: int, sheet: str | None = None):
    try:
        return matrix[row][col]
    except IndexError:
        return 0


engine = FormulaEngine(resolver=custom_resolver)
result = engine.evaluate("=SUM(A1:C2) + MAX(A1:C2)")
print(result)  # 210 + 60 = 270.0
```

### 4. SIMD Hardware-Accelerated Vector Math

Leverage native 64-bit SSE2 assembly routines for high-throughput calculations:

```python
from xlsx_viewer.asm import (
    is_asm_available,
    simd_avg,
    simd_max,
    simd_min,
    simd_sum,
    simd_sumproduct,
)

print("SIMD SSE2 Active:", is_asm_available())

data_a = [10.5, 20.25, 30.75, 40.0, 50.5, 60.0]
data_b = [1.0, 2.0, 3.0, 4.0, 5.0, 6.0]

print("Sum:", simd_sum(data_a))
print("Average:", simd_avg(data_a))
print("Min / Max:", simd_min(data_a), simd_max(data_a))
print("Sumproduct:", simd_sumproduct(data_a, data_b))
```

### 5. In-Memory Workbook Data Models

Load, inspect, and manipulate Excel workbooks programmatically. If you don't have an existing workbook, you can create one with `openpyxl` or inspect any existing `.xlsx` file:

```python
import openpyxl
from xlsx_viewer import load_workbook

# 1. Create a sample workbook if one doesn't exist
sample_wb = openpyxl.Workbook()
sample_sheet = sample_wb.active
sample_sheet.title = "Sheet1"
sample_sheet["A1"] = 100
sample_sheet["B1"] = 200
sample_wb.save("report.xlsx")

# 2. Load the workbook into xlsx-viewer
wb = load_workbook("report.xlsx")
print(f"Sheets: {wb.sheet_names}")

sheet = wb.get_sheet("Sheet1")
print(f"Dimensions: {sheet.max_row} rows x {sheet.max_col} cols")

# Read cell values and formulas
cell = sheet.get_cell(row=0, col=0)  # A1
if cell:
    print(f"A1 Value: {cell.value}, Formula: {cell.formula}")

# Modify cells in-memory
sheet.set_cell(row=0, col=1, value=1250.50)
```

---

## 💻 Command Line Interface (CLI)

> **Note:** The commands below are **terminal commands** (PowerShell / CMD / Bash), not Python code. Run them directly in your shell or with `uv run`.

`xlsx-viewer` comes with a CLI tool accessible as both `xlsx-viewer` and `xv`:

```bash
# Evaluate an Excel formula from the terminal
xv --calc "=SUM(10, 20, 30) * 2"
# 120.0

# Calculate financial loan payment
xv -c "=PMT(0.065 / 12, 360, -350000)"
# 2212.24

# List all 129 supported Excel formula functions
xv --list-formulas

# Inspect environment and hardware SIMD acceleration status
xv --info

# Open a workbook in the desktop GUI
xv sample_report.xlsx
```

*(To invoke the CLI from inside a Python script, use `from xlsx_viewer.cli import main; main(["--calc", "=SUM(10, 20)"])`)*


---

## 🖥️ Desktop Application Features

Launch the desktop spreadsheet GUI with:

```bash
xlsx-viewer
# or
python -m xlsx_viewer
```

### 📊 Office Ribbon Navigation
- **Home Tab**: Undo/Redo, Clipboard (Cut, Copy, Paste, Paste Special), Font styling, Cell fill & text colors, Cell borders, Alignment & Wrap Text, Number formats (`$`, `€`, `₽`, `%`, decimals), Conditional Formatting.
- **Insert Tab**: Chart Wizard (10+ chart types), Shapes, and Comments.
- **Data Tab**: Multi-level Sorting, AutoFilter with column search checkboxes, Remove Duplicates, Text-to-Columns, and Goal Seek.
- **Formulas Tab**: Function Wizard (`fx`), Precedence auditing, and calculation options (`F9`).
- **View Tab**: Gridline toggling, Zoom levels (50% – 200%), and Theme switcher.

### 📈 Chart Wizard (10+ Types)
Interactive charts powered by Matplotlib:
- Column & Bar Charts
- Line & Area Charts
- Pie & Donut Charts
- XY Scatter Plots & Histograms
- Professional color palettes: *Excel Classic, Modern Teal, Vibrant, Pastel, Monochrome*.

### ⌨️ Keyboard Shortcuts

| Shortcut | Action |
|----------|--------|
| `Ctrl + N` | New Workbook |
| `Ctrl + O` | Open File |
| `Ctrl + S` | Save File |
| `Ctrl + Shift + S` | Save As... |
| `Ctrl + Z` / `Ctrl + Y` | Undo / Redo |
| `Ctrl + C` / `Ctrl + X` / `Ctrl + V` | Copy / Cut / Paste |
| `Ctrl + A` | Select Entire Sheet |
| `Ctrl + F` / `Ctrl + H` | Find / Replace |
| `Ctrl + G` | Go to Cell |
| `F2` / Double Click | In-place Cell Editor |
| `F9` | Recalculate All Formulas |
| `Shift + Arrow Keys` | Expand Cell Selection |
| `Tab` / `Enter` | Move Right / Down |

---

## 🧮 Supported Formula Functions (129 Total)

<details>
<summary><strong>Click to expand full function list by category</strong></summary>

### 📐 Math & Trigonometry (33)
`SUM`, `SUMIF`, `SUMIFS`, `PRODUCT`, `ROUND`, `ROUNDUP`, `ROUNDDOWN`, `INT`, `TRUNC`, `ABS`, `MOD`, `POWER`, `SQRT`, `PI`, `RAND`, `RANDBETWEEN`, `CEILING`, `FLOOR`, `SIGN`, `SIN`, `COS`, `TAN`, `ASIN`, `ACOS`, `ATAN`, `DEGREES`, `RADIANS`, `EXP`, `LN`, `LOG`, `LOG10`, `FACT`, `SUMPRODUCT`

### 📊 Statistical (20)
`AVERAGE`, `AVERAGEA`, `AVERAGEIF`, `AVERAGEIFS`, `COUNT`, `COUNTA`, `COUNTBLANK`, `COUNTIF`, `COUNTIFS`, `MIN`, `MINIFS`, `MAX`, `MAXIFS`, `MEDIAN`, `MODE`, `STDEV`, `STDEVP`, `VAR`, `VARP`, `LARGE`, `SMALL`, `RANK`

### 💰 Financial (9)
`PMT`, `PV`, `FV`, `RATE`, `NPER`, `NPV`, `IRR`, `SLN`, `SYD`

### 🔍 Lookup & Reference (11)
`VLOOKUP`, `HLOOKUP`, `XLOOKUP`, `INDEX`, `MATCH`, `LOOKUP`, `CHOOSE`, `ROW`, `COLUMN`, `ROWS`, `COLUMNS`

### 🔤 Text (23)
`CONCAT`, `CONCATENATE`, `TEXTJOIN`, `LEFT`, `RIGHT`, `MID`, `LEN`, `TRIM`, `UPPER`, `LOWER`, `PROPER`, `EXACT`, `FIND`, `SEARCH`, `REPLACE`, `SUBSTITUTE`, `REPT`, `TEXT`, `VALUE`, `CHAR`, `CODE`, `CLEAN`, `T`

### ⚖️ Logical (11)
`IF`, `IFS`, `SWITCH`, `AND`, `OR`, `NOT`, `XOR`, `IFERROR`, `IFNA`, `TRUE`, `FALSE`

### 📅 Date & Time (11)
`TODAY`, `NOW`, `DATE`, `TIME`, `YEAR`, `MONTH`, `DAY`, `DAYS`, `EDATE`, `EOMONTH`, `WEEKDAY`

### ℹ️ Information (10)
`ISBLANK`, `ISNUMBER`, `ISTEXT`, `ISNONTEXT`, `ISLOGICAL`, `ISERROR`, `ISERR`, `ISNA`, `TYPE`, `N`

</details>

---

## 🏛️ Architecture

```
xlsx-viewer/
├── xlsx_viewer/            # Core Library Package
│   ├── __init__.py         # Public API exports
│   ├── api.py              # evaluate_formula, load_workbook, launch_gui
│   ├── cli.py              # CLI entrypoint (xlsx-viewer / xv)
│   ├── formulas.py         # Headless formula evaluation engine (129 functions)
│   ├── models.py           # In-memory workbook, sheet, cell, and formatting models
│   ├── formatting.py       # Number, date, currency, and conditional formatting
│   ├── widgets.py          # Tkinter / ttkbootstrap UI widgets (Ribbon, Sheet, Grid)
│   ├── dialogs.py          # Wizards (Chart, Function, Filter, Goal Seek)
│   ├── config.py           # Styling, themes, and configuration
│   └── asm/                # Hardware acceleration
│       ├── xlsx_math64.dll # Native x64 SSE2 assembly DLL
│       └── asm_bridge.py   # ctypes bridge with fallback
├── tests/                  # Integration and Unit Test Suite (29 tests)
│   ├── test_app.py
│   └── test_library_api.py
├── pyproject.toml          # PEP 517 / 621 Build Specification
└── README.md
```

---

## 🧪 Running Tests

Run the test suite with pytest:

```bash
uv run pytest -v
```

All 29 tests pass with 100% success rate on Python 3.10 through 3.14.

---

## 📄 License

This project is licensed under the [MIT License](LICENSE).

---

## 👤 Author

Developed by **eminsk** ([GitHub](https://github.com/eminsk)) • [M_N_Nik@yahoo.com](mailto:M_N_Nik@yahoo.com)
