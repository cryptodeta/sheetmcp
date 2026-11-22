"""SheetMCP: the Excel/CSV MCP server."""

from sheetmcp_core import __version__

from .server import build_server

__all__ = ["__version__", "build_server"]
