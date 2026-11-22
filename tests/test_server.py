import json
import os
import sys
from pathlib import Path

import pytest
from mcp import ClientSession, StdioServerParameters
from mcp.client.stdio import stdio_client
from mcp.shared.memory import create_connected_server_and_client_session

from sheetmcp import build_server
from sheetmcp_core import Config

EXPECTED_TOOLS = {
    "quick_look",
    "clean",
    "manage_objects",
    "backups",
    "batch",
    "build_report",
    "server_info",
    "list_files",
    "describe",
    "read_range",
    "write_range",
    "manage_sheets",
    "insert_delete",
    "query",
    "transform",
    "aggregate",
    "combine",
    "convert",
    "profile",
    "formulas",
    "search",
    "diff",
    "format_range",
    "add_chart",
}
WRITE_TOOLS = {
    "write_range",
    "manage_sheets",
    "insert_delete",
    "convert",
    "format_range",
    "add_chart",
    "manage_objects",
    "backups",
    "build_report",
}


def _payload(result):
    if result.structuredContent:
        sc = result.structuredContent
        return sc.get("result", sc)
    return json.loads(result.content[0].text)


@pytest.mark.anyio
async def test_in_memory_tools_resources_prompts(data_dir: Path):
    server = build_server(Config(roots=[str(data_dir)]))
    async with create_connected_server_and_client_session(server._mcp_server) as client:
        tools = await client.list_tools()
        names = {t.name for t in tools.tools}
        assert names == EXPECTED_TOOLS
        for t in tools.tools:
            assert t.description and len(t.description) > 30
        r = await client.call_tool("query", {"sql": "SELECT count(*) AS n FROM sales"})
        assert not r.isError and _payload(r)["rows"] == [[40]]
        err = await client.call_tool("describe", {"path": "nope.xlsx"})
        assert err.isError and "Hint" in err.content[0].text
        res = await client.list_resources()
        uris = {str(x.uri) for x in res.resources}
        assert "sheet://files" in uris and "sheet://describe/sales.xlsx" in uris
        content = await client.read_resource("sheet://describe/sales.xlsx")
        assert "Summary" in content.contents[0].text
        prompts = await client.list_prompts()
        assert {"analyze_spreadsheet", "clean_data", "audit_formulas", "compare_files", "build_report"} <= {
            p.name for p in prompts.prompts
        }
        p = await client.get_prompt("analyze_spreadsheet", {"path": "sales.xlsx"})
        assert "quick_look" in p.messages[0].content.text


@pytest.mark.anyio
async def test_read_only_hides_write_tools(data_dir: Path):
    server = build_server(Config(roots=[str(data_dir)], read_only=True))
    async with create_connected_server_and_client_session(server._mcp_server) as client:
        names = {t.name for t in (await client.list_tools()).tools}
        assert names == EXPECTED_TOOLS - WRITE_TOOLS


@pytest.mark.anyio
async def test_stdio_subprocess(data_dir: Path):
    params = StdioServerParameters(
        command=sys.executable,
        args=["-m", "sheetmcp", str(data_dir)],
        env={**os.environ, "PYTHONUNBUFFERED": "1"},
    )
    async with stdio_client(params) as (read, write), ClientSession(read, write) as session:
        await session.initialize()
        names = {t.name for t in (await session.list_tools()).tools}
        assert names == EXPECTED_TOOLS
        r = await session.call_tool("read_range", {"path": "sales.xlsx", "sheet": "Summary", "range": "B6"})
        assert not r.isError
        assert _payload(r)["formulas"]["B6"] == "=SUM(B2:B5)"


@pytest.fixture()
def anyio_backend():
    return "asyncio"


@pytest.mark.anyio
async def test_streamable_http(data_dir: Path):
    import asyncio
    import socket

    from mcp.client.streamable_http import streamable_http_client

    with socket.socket() as s:
        s.bind(("127.0.0.1", 0))
        port = s.getsockname()[1]
    proc = await asyncio.create_subprocess_exec(
        sys.executable,
        "-m",
        "sheetmcp",
        str(data_dir),
        "--transport",
        "http",
        "--port",
        str(port),
        stdout=asyncio.subprocess.DEVNULL,
        stderr=asyncio.subprocess.DEVNULL,
    )
    try:
        for _ in range(100):
            try:
                _r, w = await asyncio.open_connection("127.0.0.1", port)
                w.close()
                break
            except OSError:
                await asyncio.sleep(0.1)
        async with (
            streamable_http_client(f"http://127.0.0.1:{port}/mcp") as (read, write, _),
            ClientSession(read, write) as session,
        ):
            await session.initialize()
            r = await session.call_tool("list_files", {})
            assert not r.isError and _payload(r)["count"] >= 7
    finally:
        proc.terminate()
        await proc.wait()


@pytest.mark.anyio
async def test_new_tools_over_mcp(data_dir: Path):
    server = build_server(Config(roots=[str(data_dir)]))
    async with create_connected_server_and_client_session(server._mcp_server) as client:
        r = await client.call_tool("quick_look", {})
        assert not r.isError and _payload(r)["file_count"] >= 7
        r = await client.call_tool("query", {})
        assert not r.isError and any(t["table"] == "sales" for t in _payload(r)["tables"])
        r = await client.call_tool(
            "build_report",
            {
                "source": "sales.xlsx",
                "output": "rep.xlsx",
                "group_by": ["Region"],
                "metrics": ["Revenue:sum"],
            },
        )
        assert not r.isError and _payload(r)["summary_rows"] == 4
        r = await client.call_tool("clean", {"source": "customers.csv", "operations": [{"op": "dedupe"}]})
        assert not r.isError and _payload(r)["rows_after"] == 12
        r = await client.call_tool(
            "batch",
            {
                "operations": [
                    {"tool": "write_range", "args": {"path": "customers.csv", "cells": {"A2": "Q"}}}
                ]
            },
        )
        assert not r.isError and _payload(r)["succeeded"] == 1
        r = await client.call_tool("backups", {"path": "customers.csv", "action": "undo"})
        assert not r.isError
        r = await client.call_tool("manage_objects", {"path": "sales.xlsx"})
        assert not r.isError and _payload(r)["tables"][0]["name"] == "SalesTable"


@pytest.mark.anyio
async def test_read_only_batch_rejects_writes(data_dir: Path):
    server = build_server(Config(roots=[str(data_dir)], read_only=True))
    async with create_connected_server_and_client_session(server._mcp_server) as client:
        r = await client.call_tool(
            "batch", {"operations": [{"tool": "write_range", "args": {"path": "x.xlsx", "cells": {"A1": 1}}}]}
        )
        assert r.isError and "read-only" in r.content[0].text
        r = await client.call_tool("batch", {"operations": [{"tool": "quick_look", "args": {}}]})
        assert not r.isError
