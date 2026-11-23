"""SheetMCP: the Excel/CSV MCP server, spreadsheet engine (``sheetmcp.core``) and ``sheetcli``."""

from __future__ import annotations

from typing import TYPE_CHECKING, Any

__version__ = "0.3.0"

if TYPE_CHECKING:
    from .server import build_server

__all__ = ["__version__", "build_server"]


def __getattr__(name: str) -> Any:
    # Lazy so `import sheetmcp.core` and `sheetcli` do not load the MCP SDK.
    if name == "build_server":
        from .server import build_server

        return build_server
    raise AttributeError(f"module {__name__!r} has no attribute {name!r}")
