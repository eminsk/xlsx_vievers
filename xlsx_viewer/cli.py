"""
Unified CLI for xlsx-viewer (Headless Formula Evaluation + Desktop Viewer).
"""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

from . import __version__
from .api import evaluate_formula, launch_viewer
from .formulas import FUNCTION_METADATA


def export_to_csv_with_styles(in_file: str | Path, out_file: str | Path) -> int:
    """Export spreadsheet to CSV with evaluated formulas, formatted numbers and binary .style file."""
    import csv
    import locale
    import struct
    from typing import Any

    from .api import load_workbook
    from .formatting import NumberFormatter
    from .formulas import FormulaEngine, FormulaStore

    in_path = Path(in_file)
    out_path = Path(out_file)
    out_path.parent.mkdir(parents=True, exist_ok=True)

    try:
        enc = locale.getpreferredencoding(False) or "utf-8"
    except Exception:
        enc = "utf-8"

    # Try rich openpyxl extraction if .xlsx
    if in_path.suffix.lower() == ".xlsx" and in_path.exists():
        try:
            import openpyxl

            wb_form = openpyxl.load_workbook(in_path, data_only=False)
            wb_vals = openpyxl.load_workbook(in_path, data_only=True)

            store = FormulaStore()
            sheets: dict[str, list[list[Any]]] = {}
            for name in wb_form.sheetnames:
                ws = wb_form[name]
                grid = []
                for r in range(1, ws.max_row + 1):
                    row = []
                    for c in range(1, ws.max_column + 1):
                        val = ws.cell(r, c).value
                        if isinstance(val, str) and val.startswith("="):
                            store.set(name, r - 1, c - 1, val)
                        row.append(val)
                    grid.append(row)
                sheets[name] = grid

            calc_cache: dict[tuple[str, int, int], Any] = {}

            def resolver(r: int, c: int, sheet_name: str | None = None) -> Any:
                s_name = sheet_name or wb_form.sheetnames[0]
                key = (s_name, r, c)
                if key in calc_cache:
                    return calc_cache[key]
                val = (
                    sheets[s_name][r][c]
                    if r < len(sheets[s_name]) and c < len(sheets[s_name][r])
                    else 0
                )
                if s_name in wb_vals.sheetnames:
                    ws_v = wb_vals[s_name]
                    if r < ws_v.max_row and c < ws_v.max_column:
                        cached_val = ws_v.cell(r + 1, c + 1).value
                        if cached_val is not None:
                            calc_cache[key] = cached_val
                            return cached_val
                formula = store.get(s_name, r, c)
                if formula:
                    calc_cache[key] = 0
                    try:
                        res = engine.evaluate(formula, s_name)
                    except Exception:
                        res = 0
                    calc_cache[key] = res
                    return res
                calc_cache[key] = val
                return val

            engine = FormulaEngine(resolver)
            active_name = wb_form.active.title if wb_form.active else wb_form.sheetnames[0]
            ws0 = wb_form[active_name]

            merged_fills: dict[tuple[int, int], str] = {}
            for rng in ws0.merged_cells.ranges:
                top_cell = ws0.cell(rng.min_row, rng.min_col)
                if top_cell.fill and top_cell.fill.start_color and top_cell.fill.start_color.rgb:
                    rgb = str(top_cell.fill.start_color.rgb)
                    if len(rgb) == 8:
                        rgb = rgb[2:]
                    if len(rgb) == 6 and rgb.upper() != "000000":
                        for mr in range(rng.min_row, rng.max_row + 1):
                            for mc in range(rng.min_col, rng.max_col + 1):
                                merged_fills[(mr - 1, mc)] = rgb

            records = []
            rows_out = []
            max_c = min(26, ws0.max_column)
            for r in range(ws0.max_row):
                row_vals = []
                for c in range(1, max_c + 1):
                    val = resolver(r, c - 1, active_name)
                    fmt = ws0.cell(r + 1, c).number_format
                    fmt_val = NumberFormatter.format_value(val, fmt)
                    if isinstance(fmt_val, str):
                        fmt_val = fmt_val.replace("\u2192", "->").replace("\u2014", "-")
                    row_vals.append(fmt_val)

                    # Extract Style
                    cell = ws0.cell(r + 1, c)
                    fill_rgb = merged_fills.get((r, c))
                    if (
                        not fill_rgb
                        and cell.fill
                        and cell.fill.start_color
                        and cell.fill.start_color.rgb
                    ):
                        rgb = str(cell.fill.start_color.rgb)
                        if len(rgb) == 8:
                            rgb = rgb[2:]
                        if len(rgb) == 6 and rgb.upper() != "000000":
                            fill_rgb = rgb
                    font_rgb = None
                    if cell.font and cell.font.color and cell.font.color.rgb:
                        try:
                            rgb = str(cell.font.color.rgb)
                            if len(rgb) == 8:
                                rgb = rgb[2:]
                            if len(rgb) == 6:
                                font_rgb = rgb
                        except Exception:
                            pass
                    bold = bool(cell.font and cell.font.bold)
                    flags = 1 if bold else 0
                    clr_bk = 0
                    if fill_rgb:
                        flags |= 2
                        clr_bk = (
                            int(fill_rgb[0:2], 16)
                            | (int(fill_rgb[2:4], 16) << 8)
                            | (int(fill_rgb[4:6], 16) << 16)
                        )
                    clr_fg = 0
                    if font_rgb:
                        flags |= 4
                        clr_fg = (
                            int(font_rgb[0:2], 16)
                            | (int(font_rgb[2:4], 16) << 8)
                            | (int(font_rgb[4:6], 16) << 16)
                        )
                    if flags != 0:
                        records.append((r, c, flags, clr_bk, clr_fg))
                rows_out.append(row_vals)

            with open(out_path, "w", newline="", encoding=enc, errors="replace") as f:
                writer = csv.writer(f)
                for row in rows_out:
                    writer.writerow(row)

            style_path = Path(f"{out_path}.style")
            with open(style_path, "wb") as sf:
                sf.write(struct.pack("<4sI", b"XLSS", len(records)))
                for r, c, fl, bk, fg in records:
                    sf.write(struct.pack("<HBBII", r, c, fl, bk, fg))

            return 0
        except Exception:
            pass

    # Fallback / Generic .xls / .csv loader
    wb_data = load_workbook(in_path)
    s_data = wb_data.get_sheet()
    if not s_data and wb_data.sheets:
        s_data = next(iter(wb_data.sheets.values()))
    if not s_data:
        print(f"Error: No sheets found in {in_path}", file=sys.stderr)
        return 1

    records = []
    rows_out = []
    for r_idx, row in enumerate(s_data.rows):
        row_vals = []
        for c_idx, val in enumerate(row):
            style = s_data.cell_styles.get((r_idx, c_idx))
            fmt = style.number_format if style else None
            fmt_val = NumberFormatter.format_value(val, fmt)
            if isinstance(fmt_val, str):
                fmt_val = fmt_val.replace("\u2192", "->").replace("\u2014", "-")
            row_vals.append(fmt_val)

            if style and c_idx < 26:
                flags = 1 if style.bold else 0
                clr_bk = 0
                if style.bg_color and style.bg_color.startswith("#"):
                    hex_c = style.bg_color.lstrip("#")
                    if len(hex_c) == 6:
                        flags |= 2
                        clr_bk = (
                            int(hex_c[0:2], 16)
                            | (int(hex_c[2:4], 16) << 8)
                            | (int(hex_c[4:6], 16) << 16)
                        )
                clr_fg = 0
                if style.fg_color and style.fg_color.startswith("#"):
                    hex_c = style.fg_color.lstrip("#")
                    if len(hex_c) == 6:
                        flags |= 4
                        clr_fg = (
                            int(hex_c[0:2], 16)
                            | (int(hex_c[2:4], 16) << 8)
                            | (int(hex_c[4:6], 16) << 16)
                        )
                if flags != 0:
                    records.append((r_idx, c_idx + 1, flags, clr_bk, clr_fg))
        rows_out.append(row_vals)

    with open(out_path, "w", newline="", encoding=enc, errors="replace") as f:
        writer = csv.writer(f)
        for row in rows_out:
            writer.writerow(row)

    if records:
        style_path = Path(f"{out_path}.style")
        with open(style_path, "wb") as sf:
            sf.write(struct.pack("<4sI", b"XLSS", len(records)))
            for r, c, fl, bk, fg in records:
                sf.write(struct.pack("<HBBII", r, c, fl, bk, fg))

    return 0


def main(argv: list[str] | None = None) -> int:
    """Main CLI entrypoint for xlsx-viewer."""
    if argv is None:
        argv = sys.argv[1:]

    parser = argparse.ArgumentParser(
        prog="xlsx-viewer",
        description="xlsx-viewer: Modern Excel File Viewer & 129-Function Formula Engine with SIMD Acceleration.",
        formatter_class=argparse.RawTextHelpFormatter,
        epilog=(
            "Examples:\n"
            "  # Launch GUI Viewer with file\n"
            "  xlsx-viewer report.xlsx\n\n"
            "  # Headless CLI formula evaluation\n"
            '  xlsx-viewer --calc "=SUM(10, 20, 30)*2"\n'
            '  xlsx-viewer --calc "=PMT(0.05/12, 360, -250000)"\n\n'
            "  # List all 129 supported Excel formula functions\n"
            "  xlsx-viewer --list-formulas\n"
        ),
    )

    parser.add_argument(
        "-v",
        "--version",
        action="version",
        version=f"%(prog)s {__version__}",
    )

    parser.add_argument(
        "file",
        nargs="?",
        default=None,
        help="Path to Excel (.xlsx) file to open in modern desktop GUI viewer.",
    )

    parser.add_argument(
        "-c",
        "--calc",
        metavar="FORMULA",
        help="Evaluate an Excel formula headless in the terminal (e.g. --calc '=SUM(1,2,3)').",
    )

    parser.add_argument(
        "--list-formulas",
        "--list-funcs",
        dest="list_formulas",
        action="store_true",
        help="Print all 129 supported Excel formula functions grouped by category.",
    )

    parser.add_argument(
        "--info",
        action="store_true",
        help="Display library information, active hardware ISA acceleration, and version.",
    )

    parser.add_argument(
        "--mcp",
        action="store_true",
        help="Start Model Context Protocol (MCP) Server over stdio for Claude/Cursor/Windsurf.",
    )

    parser.add_argument(
        "--to-csv",
        nargs=2,
        metavar=("INPUT", "OUTPUT"),
        help="Export first sheet of an Excel (.xlsx / .xls) file to CSV format.",
    )

    args = parser.parse_args(argv)

    if args.to_csv:
        in_file, out_file = args.to_csv
        try:
            return export_to_csv_with_styles(in_file, out_file)
        except Exception as e:
            print(f"Error converting to CSV: {e}", file=sys.stderr)
            return 1

    if args.mcp:
        from .mcp_server import XlsxMCPServer

        XlsxMCPServer().run_stdio()
        return 0

    if args.list_formulas:
        categories: dict[str, list[str]] = {}
        for name, meta in FUNCTION_METADATA.items():
            cat = (
                meta.get("cat", "General")
                if isinstance(meta, dict)
                else getattr(meta, "category", "General")
            )
            categories.setdefault(cat, []).append(name)

        print(
            f"xlsx-viewer {__version__} - Supported Formula Functions ({len(FUNCTION_METADATA)} total):\n"
        )
        for cat in sorted(categories):
            funcs = sorted(categories[cat])
            print(f"[{cat.upper()}] ({len(funcs)} functions):")
            print("  " + ", ".join(funcs))
            print()
        return 0

    if args.info:
        from .asm import is_asm_available

        print(f"xlsx-viewer version: {__version__}")
        print(f"Supported functions: {len(FUNCTION_METADATA)}")
        print(f"Native x64 SIMD SSE2: {'ACTIVE' if is_asm_available() else 'Pure Python Fallback'}")
        return 0

    if args.calc:
        res = evaluate_formula(args.calc)
        print(f"Result: {res}")
        return 0

    # Default action: Launch GUI Viewer
    launch_viewer(args.file)
    return 0


if __name__ == "__main__":
    sys.exit(main())
