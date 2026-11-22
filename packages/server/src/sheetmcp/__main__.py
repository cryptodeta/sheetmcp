"""Command line entry point: `sheetmcp [ROOT ...] [--read-only] [--transport stdio|http]`."""

from __future__ import annotations

import argparse
import logging
import os
import sys

from sheetmcp_core import Config, SheetError, __version__


def parse_args(argv: list[str] | None = None) -> argparse.Namespace:
    ap = argparse.ArgumentParser(
        prog="sheetmcp",
        description="SheetMCP: Excel/CSV/Parquet MCP server. Formulas, SQL across files, safe writes. No API keys.",
    )
    ap.add_argument(
        "roots", nargs="*", help="Folder(s) the server may access (default: $SHEETMCP_ROOTS or current dir)."
    )
    ap.add_argument("--root", "-r", action="append", default=[], help="Allowed folder (repeatable).")
    ap.add_argument("--read-only", action="store_true", default=None, help="Disable every write tool.")
    ap.add_argument("--no-backup", action="store_true", help="Do not back up files before overwriting.")
    ap.add_argument("--max-rows", type=int, default=None, help="Default max rows per response (default 100).")
    ap.add_argument(
        "--transport",
        "-t",
        choices=["stdio", "http", "streamable-http", "sse"],
        default=os.environ.get("SHEETMCP_TRANSPORT", "stdio"),
    )
    ap.add_argument("--host", default=os.environ.get("SHEETMCP_HOST", "127.0.0.1"))
    ap.add_argument("--port", type=int, default=int(os.environ.get("SHEETMCP_PORT", "8000")))
    ap.add_argument("--log-level", default=os.environ.get("SHEETMCP_LOG_LEVEL", "WARNING"))
    ap.add_argument("--version", action="version", version=f"sheetmcp {__version__}")
    return ap.parse_args(argv)


def main(argv: list[str] | None = None) -> None:
    args = parse_args(argv)
    logging.basicConfig(
        level=args.log_level.upper(), stream=sys.stderr, format="%(levelname)s sheetmcp: %(message)s"
    )
    roots = [*args.roots, *args.root]
    overrides: dict[str, object] = {"read_only": args.read_only, "max_rows": args.max_rows}
    if roots:
        overrides["roots"] = roots
    if args.no_backup:
        overrides["backup"] = False
    try:
        cfg = Config.from_env(**overrides)
        from .server import build_server

        mcp = build_server(cfg, host=args.host, port=args.port)
    except SheetError as e:
        print(f"sheetmcp: {e}", file=sys.stderr)
        sys.exit(2)
    transport = {"http": "streamable-http"}.get(args.transport, args.transport)
    if transport != "stdio":
        print(
            f"sheetmcp {__version__} serving {transport} on http://{args.host}:{args.port}"
            f"{'/mcp' if transport == 'streamable-http' else '/sse'}  roots={cfg.roots}",
            file=sys.stderr,
        )
    mcp.run(transport=transport)  # type: ignore[arg-type]


if __name__ == "__main__":
    main()
