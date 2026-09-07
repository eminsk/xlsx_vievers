"""
Unified CLI for xlsx-viewer (Headless Formula Evaluation + Desktop Viewer).
"""

from __future__ import annotations

import argparse
import sys

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
        action="store_true",
        help="Print all 129 supported Excel formula functions grouped by category.",
    )

    parser.add_argument(
        "--info",
        action="store_true",
        help="Display library information, active hardware ISA acceleration, and version.",
    )

    args = parser.parse_args(argv)

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
