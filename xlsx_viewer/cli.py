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
        import csv

        from .api import load_workbook

        in_file, out_file = args.to_csv
        try:
            wb_data = load_workbook(in_file)
            s_data = wb_data.get_sheet()
            if not s_data and wb_data.sheets:
                s_data = next(iter(wb_data.sheets.values()))
            if not s_data:
                print(f"Error: No sheets found in {in_file}", file=sys.stderr)
                return 1
            import locale

            try:
                enc = locale.getpreferredencoding(False) or "utf-8"
            except Exception:
                enc = "utf-8"

            Path(out_file).parent.mkdir(parents=True, exist_ok=True)
            with open(out_file, "w", newline="", encoding=enc, errors="replace") as f:
                writer = csv.writer(f)
                for row in s_data.rows:
                    writer.writerow([str(v) if v is not None else "" for v in row])
            return 0
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
