"""
Native Model Context Protocol (MCP) JSON-RPC 2.0 Server for xlsx-viewer-pro.
Exposes deterministic 129-function Excel formula evaluation (with SSE2 SIMD acceleration)
and headless .xlsx reading/writing to Claude Desktop, Cursor, Windsurf, Antigravity,
and any MCP-compatible AI agent over stdio.
"""

from __future__ import annotations

import json
import sys
from pathlib import Path
from typing import Any

from . import __version__
from .api import evaluate_formula, load_workbook, save_workbook
from .asm import is_asm_available
from .formulas import FUNCTION_METADATA, FormulaEngine
from .models import CellPosition

MCP_PROTOCOL_VERSION = "2024-11-05"

MCP_TOOLS_SCHEMA: list[dict[str, Any]] = [
    {
        "name": "xlsx_evaluate_formula",
        "description": (
            "Deterministically evaluate any of the 129+ supported Excel formulas "
            "(e.g. '=PMT(0.08/12, 60, -50000)', '=XLOOKUP(...)', '=IRR(...)', '=SUMIFS(...)') "
            "using the xlsx-viewer-pro formula & SIMD engine. Optionally pass cell values "
            "via `context` (e.g. {'A1': 100, 'B1': 250}) or reference an existing `.xlsx` file."
        ),
        "inputSchema": {
            "type": "object",
            "properties": {
                "formula": {"type": "string", "description": "Excel formula string (e.g. '=SUM(A1:A3)*1.2')"},
                "context": {
                    "type": "object",
                    "description": "Optional dictionary mapping cell refs (e.g. 'A1', 'B2') to values",
                },
                "file_path": {
                    "type": "string",
                    "description": "Optional path to an .xlsx workbook to resolve cell references from",
                },
                "sheet": {
                    "type": "string",
                    "description": "Optional worksheet name when `file_path` is provided",
                },
            },
            "required": ["formula"],
        },
    },
    {
        "name": "xlsx_read_workbook",
        "description": (
            "Read worksheet names, grid dimensions, raw values, and recalculated formula "
            "results from an Excel (.xlsx) file headlessly without requiring Microsoft Excel."
        ),
        "inputSchema": {
            "type": "object",
            "properties": {
                "file_path": {"type": "string", "description": "Path to the .xlsx file"},
                "sheet": {"type": "string", "description": "Optional sheet name (defaults to active sheet)"},
                "max_rows": {"type": "integer", "default": 100, "description": "Maximum rows to return"},
            },
            "required": ["file_path"],
        },
    },
    {
        "name": "xlsx_write_cells",
        "description": (
            "Create or update an Excel (.xlsx) file with a dictionary of cell coordinates "
            "and values/formulas (e.g. {'A1': 100, 'A2': 200, 'A3': '=SUM(A1:A2)'}), save to disk, "
            "and return the recalculated values."
        ),
        "inputSchema": {
            "type": "object",
            "properties": {
                "file_path": {"type": "string", "description": "Path to the .xlsx file to create or update"},
                "cells": {
                    "type": "object",
                    "description": "Mapping of Excel cell refs ('A1', 'B2') to numbers, strings, or '=FORMULA'",
                },
                "sheet": {"type": "string", "default": "Sheet1", "description": "Worksheet name"},
            },
            "required": ["file_path", "cells"],
        },
    },
    {
        "name": "xlsx_list_formulas",
        "description": "List all 129 supported Excel formula functions and hardware SIMD SSE2 acceleration status.",
        "inputSchema": {
            "type": "object",
            "properties": {},
        },
    },
]


class XlsxMCPServer:
    """Zero-dependency Model Context Protocol (MCP) JSON-RPC 2.0 server for xlsx-viewer-pro."""

    def handle_request(self, request: dict[str, Any]) -> dict[str, Any] | None:
        """Process a single JSON-RPC 2.0 MCP message."""
        method = request.get("method", "")
        req_id = request.get("id")
        params = request.get("params") or {}

        if req_id is None and method.startswith("notifications/"):
            return None

        try:
            if method == "initialize":
                return {
                    "jsonrpc": "2.0",
                    "id": req_id,
                    "result": {
                        "protocolVersion": MCP_PROTOCOL_VERSION,
                        "capabilities": {"tools": {}},
                        "serverInfo": {
                            "name": "xlsx-viewer-pro-mcp",
                            "version": __version__,
                            "simdAvailable": bool(is_asm_available()),
                        },
                    },
                }

            if method == "ping":
                return {"jsonrpc": "2.0", "id": req_id, "result": {}}

            if method == "tools/list":
                return {"jsonrpc": "2.0", "id": req_id, "result": {"tools": MCP_TOOLS_SCHEMA}}

            if method == "tools/call":
                tool_name = params.get("name", "")
                args = params.get("arguments") or {}
                output = self._call_tool(tool_name, args)
                return {
                    "jsonrpc": "2.0",
                    "id": req_id,
                    "result": {
                        "content": [
                            {
                                "type": "text",
                                "text": json.dumps(output, ensure_ascii=False, indent=2, default=str),
                            }
                        ],
                        "isError": False,
                    },
                }

            return {
                "jsonrpc": "2.0",
                "id": req_id,
                "error": {"code": -32601, "message": f"Method not found: {method}"},
            }
        except Exception as exc:
            return {
                "jsonrpc": "2.0",
                "id": req_id,
                "result": {
                    "content": [{"type": "text", "text": f"Error: {exc}"}],
                    "isError": True,
                },
            }

    def _build_sheet_resolver(self, wb, default_sheet):
        def resolver(r: int, c: int, sheet_name: str | None = None) -> Any:
            target_sheet = wb.get_sheet(sheet_name) if sheet_name else default_sheet
            if target_sheet is None:
                return 0
            val = target_sheet.get_cell_value(r, c)
            if isinstance(val, str) and val.startswith("="):
                eng = FormulaEngine(lambda rr, cc, ss=None: resolver(rr, cc, ss or target_sheet.name))
                return eng.evaluate(val)
            return 0 if (val is None or val == "") else val

        return resolver

    def _call_tool(self, name: str, args: dict[str, Any]) -> Any:
        if name == "xlsx_evaluate_formula":
            formula = str(args["formula"])
            file_path = args.get("file_path")
            if file_path:
                wb = load_workbook(file_path, create_if_missing=False)
                sheet = wb.get_sheet(args.get("sheet")) or wb.get_sheet()
                resolver = self._build_sheet_resolver(wb, sheet)
                result = evaluate_formula(formula, context=resolver)
            else:
                ctx = args.get("context")
                result = evaluate_formula(formula, context=ctx)
            return {
                "formula": formula,
                "result": result,
                "simd_active": bool(is_asm_available()),
            }

        if name == "xlsx_read_workbook":
            file_path = str(args["file_path"])
            wb = load_workbook(file_path, create_if_missing=False)
            sheet = wb.get_sheet(args.get("sheet")) or wb.get_sheet()
            if sheet is None:
                raise ValueError(f"No worksheet found in {file_path}")

            max_rows = int(args.get("max_rows", 100))
            resolver = self._build_sheet_resolver(wb, sheet)
            engine = FormulaEngine(resolver)

            cells_out: dict[str, Any] = {}
            formulas_out: dict[str, str] = {}
            for r_idx in range(min(sheet.row_count, max_rows)):
                row_vals = sheet.rows[r_idx]
                for c_idx, val in enumerate(row_vals):
                    if val is None or val == "":
                        continue
                    ref = CellPosition(r_idx, c_idx).to_excel()
                    if isinstance(val, str) and val.startswith("="):
                        formulas_out[ref] = val
                        cells_out[ref] = engine.evaluate(val)
                    else:
                        cells_out[ref] = val

            return {
                "file_path": str(Path(file_path).resolve()),
                "sheets": wb.sheet_names,
                "active_sheet": sheet.name,
                "rows": sheet.row_count,
                "cols": sheet.col_count,
                "cells": cells_out,
                "formulas": formulas_out,
            }

        if name == "xlsx_write_cells":
            file_path = str(args["file_path"])
            sheet_name = str(args.get("sheet", "Sheet1"))
            cells_map: dict[str, Any] = dict(args.get("cells") or {})

            wb = load_workbook(file_path, create_if_missing=True)
            sheet = wb.get_sheet(sheet_name) or wb.add_sheet(sheet_name)

            for ref_str, val in cells_map.items():
                pos = CellPosition.from_excel(ref_str)
                formula = val if (isinstance(val, str) and val.startswith("=")) else None
                sheet.set_cell(pos.row, pos.col, val, formula=formula)

            save_workbook(wb, file_path)

            resolver = self._build_sheet_resolver(wb, sheet)
            engine = FormulaEngine(resolver)
            evaluated: dict[str, Any] = {}
            for ref_str in cells_map:
                pos = CellPosition.from_excel(ref_str)
                raw = sheet.get_cell_value(pos.row, pos.col)
                if isinstance(raw, str) and raw.startswith("="):
                    evaluated[ref_str] = engine.evaluate(raw)
                else:
                    evaluated[ref_str] = raw

            return {
                "status": "saved",
                "file_path": str(Path(file_path).resolve()),
                "sheet": sheet.name,
                "evaluated_cells": evaluated,
            }

        if name == "xlsx_list_formulas":
            return {
                "version": __version__,
                "simd_sse2_active": bool(is_asm_available()),
                "total_functions": len(FUNCTION_METADATA),
                "functions": sorted(FUNCTION_METADATA.keys()),
            }

        raise ValueError(f"Unknown MCP tool: {name}")

    def run_stdio(self) -> None:
        """Run the xlsx-viewer-pro MCP JSON-RPC 2.0 server over standard input/output."""
        for raw_line in sys.stdin:
            line = raw_line.strip()
            if not line:
                continue
            try:
                req = json.loads(line)
            except json.JSONDecodeError:
                continue
            resp = self.handle_request(req)
            if resp is not None:
                sys.stdout.write(json.dumps(resp, ensure_ascii=False, default=str) + "\n")
                sys.stdout.flush()


def main_mcp(argv: list[str] | None = None) -> int:
    """CLI entry point for xlsx-mcp / xlsx-viewer --mcp."""
    import argparse

    parser = argparse.ArgumentParser(
        prog="xlsx-mcp",
        description="Start xlsx-viewer-pro Model Context Protocol (MCP) Server over stdio",
    )
    parser.parse_args(argv)
    server = XlsxMCPServer()
    server.run_stdio()
    return 0


if __name__ == "__main__":
    sys.exit(main_mcp())
