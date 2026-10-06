"""
Tests for xlsx-viewer-pro Model Context Protocol (MCP) JSON-RPC 2.0 Server.
"""

from __future__ import annotations

import json

from xlsx_viewer import XlsxMCPServer


def test_xlsx_mcp_server_lifecycle(tmp_path):
    server = XlsxMCPServer()

    # 1. initialize
    init_resp = server.handle_request({"jsonrpc": "2.0", "id": 1, "method": "initialize"})
    assert init_resp is not None
    assert init_resp["result"]["serverInfo"]["name"] == "xlsx-viewer-pro-mcp"

    # 2. tools/list
    list_resp = server.handle_request({"jsonrpc": "2.0", "id": 2, "method": "tools/list"})
    assert list_resp is not None
    tools = [t["name"] for t in list_resp["result"]["tools"]]
    assert "xlsx_evaluate_formula" in tools
    assert "xlsx_read_workbook" in tools
    assert "xlsx_write_cells" in tools
    assert "xlsx_list_formulas" in tools

    # 3. evaluate formula with context
    eval_resp = server.handle_request(
        {
            "jsonrpc": "2.0",
            "id": 3,
            "method": "tools/call",
            "params": {
                "name": "xlsx_evaluate_formula",
                "arguments": {
                    "formula": "=SUM(A1, B1) * 2",
                    "context": {"A1": 15, "B1": 25},
                },
            },
        }
    )
    assert eval_resp is not None
    assert eval_resp["result"]["isError"] is False
    payload = json.loads(eval_resp["result"]["content"][0]["text"])
    assert payload["result"] == 80

    # 4. write cells and formulas to .xlsx and verify recalculation
    wb_path = tmp_path / "financial_model.xlsx"
    write_resp = server.handle_request(
        {
            "jsonrpc": "2.0",
            "id": 4,
            "method": "tools/call",
            "params": {
                "name": "xlsx_write_cells",
                "arguments": {
                    "file_path": str(wb_path),
                    "cells": {
                        "A1": 100,
                        "A2": 200,
                        "A3": 300,
                        "B1": "=SUM(A1:A3)",
                        "B2": "=AVERAGE(A1:A3)",
                    },
                },
            },
        }
    )
    assert write_resp is not None
    w_data = json.loads(write_resp["result"]["content"][0]["text"])
    assert w_data["evaluated_cells"]["B1"] == 600
    assert w_data["evaluated_cells"]["B2"] == 200

    # 5. read workbook back
    read_resp = server.handle_request(
        {
            "jsonrpc": "2.0",
            "id": 5,
            "method": "tools/call",
            "params": {
                "name": "xlsx_read_workbook",
                "arguments": {"file_path": str(wb_path)},
            },
        }
    )
    assert read_resp is not None
    r_data = json.loads(read_resp["result"]["content"][0]["text"])
    assert r_data["cells"]["B1"] == 600
    assert r_data["formulas"]["B1"] == "=SUM(A1:A3)"


import tempfile
import unittest
from pathlib import Path


class TestXlsxMCPServer(unittest.TestCase):
    def test_mcp_lifecycle(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            test_xlsx_mcp_server_lifecycle(Path(tmpdir))

