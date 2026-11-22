"""Talk to the SheetMCP server over stdio like an AI app would.

uv run python examples/mcp_client_demo.py
"""

import asyncio
import json
import sys
from pathlib import Path

from mcp import ClientSession, StdioServerParameters
from mcp.client.stdio import stdio_client

DATA = Path(__file__).parent / "data"


async def main() -> None:
    params = StdioServerParameters(command=sys.executable, args=["-m", "sheetmcp", str(DATA)])
    async with stdio_client(params) as (read, write), ClientSession(read, write) as session:
        await session.initialize()
        tools = await session.list_tools()
        print(f"{len(tools.tools)} tools:", ", ".join(t.name for t in tools.tools))
        r = await session.call_tool(
            "query",
            {"sql": "SELECT Region, sum(Revenue) AS revenue FROM sales GROUP BY 1", "format": "markdown"},
        )
        print(json.loads(r.content[0].text)["markdown"])


if __name__ == "__main__":
    asyncio.run(main())
