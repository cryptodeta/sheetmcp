"""`sheetcli` - run SheetMCP engine operations from a terminal.

Examples:
  sheetcli ls data/
  sheetcli describe data/sales.xlsx
  sheetcli read data/sales.xlsx -s Sales -r A1:F20 --md
  sheetcli query "SELECT region, sum(amount) FROM sales GROUP BY 1" --md
  sheetcli convert data/sales.xlsx out/sales.parquet
  sheetcli profile data/customers.csv
  sheetcli formulas data/model.xlsx errors
  sheetcli diff v1.xlsx v2.xlsx
  sheetcli search "acme" data/
"""

from __future__ import annotations

import argparse
import json
import sys
from typing import Any

from sheetmcp_core import Config, SheetError, Sheets, __version__


def _out(res: Any, md: bool = False) -> None:
    if md and isinstance(res, dict) and "markdown" in res:
        print(res["markdown"])
        extra = {k: v for k, v in res.items() if k not in ("markdown", "rows", "values")}
        if extra.get("truncated"):
            print(f"\n… truncated, next offset {extra.get('next_offset')}", file=sys.stderr)
        return
    print(json.dumps(res, indent=2, ensure_ascii=False, default=str))


def _tables(items: list[str] | None) -> dict[str, str] | None:
    if not items:
        return None
    out = {}
    for it in items:
        if "=" not in it:
            raise SheetError(f"Bad --table '{it}'.", "Use alias=path, e.g. -t sales=data/sales.xlsx")
        k, v = it.split("=", 1)
        out[k] = v
    return out


def build_parser() -> argparse.ArgumentParser:
    ap = argparse.ArgumentParser(
        prog="sheetcli",
        description="SheetMCP engine on the command line.",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=__doc__,
    )
    ap.add_argument(
        "--root", action="append", default=None, help="Allowed folder(s); default: current directory."
    )
    ap.add_argument("--read-only", action="store_true")
    ap.add_argument("--no-backup", action="store_true")
    ap.add_argument("--version", action="version", version=f"sheetcli {__version__}")
    sub = ap.add_subparsers(dest="cmd", required=True)

    p = sub.add_parser("ls", help="list spreadsheet files")
    p.add_argument("path", nargs="?", default=".")
    p.add_argument("-p", "--pattern")

    p = sub.add_parser("describe", help="sheets, columns, types, preview")
    p.add_argument("path")
    p.add_argument("-s", "--sheet")

    p = sub.add_parser("read", help="read a range (values + formulas)")
    p.add_argument("path")
    p.add_argument("-s", "--sheet")
    p.add_argument("-r", "--range")
    p.add_argument("-n", "--limit", type=int)
    p.add_argument("--offset", type=int, default=0)
    p.add_argument("--md", action="store_true", help="markdown table output")

    p = sub.add_parser("write", help="write cells: sheetcli write f.xlsx B2=42 C2='=B2*2'")
    p.add_argument("path")
    p.add_argument("assignments", nargs="+")
    p.add_argument("-s", "--sheet")
    p.add_argument("--dry-run", action="store_true")

    p = sub.add_parser("query", help="SQL across files")
    p.add_argument("sql")
    p.add_argument("-t", "--table", action="append", help="alias=path (repeatable)")
    p.add_argument("-n", "--limit", type=int)
    p.add_argument("-o", "--output")
    p.add_argument("--md", action="store_true")

    p = sub.add_parser("convert", help="convert between xlsx/csv/parquet/json/md")
    p.add_argument("input")
    p.add_argument("output")
    p.add_argument("-s", "--sheet")
    p.add_argument("--all-sheets", action="store_true")

    p = sub.add_parser("profile", help="stats and data quality report")
    p.add_argument("source")
    p.add_argument("-s", "--sheet")

    p = sub.add_parser("formulas", help="list | errors | trace | evaluate | stats")
    p.add_argument("path")
    p.add_argument("action", nargs="?", default="list")
    p.add_argument("-s", "--sheet")
    p.add_argument("-c", "--cell")
    p.add_argument("-f", "--formula")

    p = sub.add_parser("diff", help="compare two files/sheets")
    p.add_argument("a")
    p.add_argument("b", nargs="?")
    p.add_argument("-k", "--key", action="append")
    p.add_argument("--sheet-a")
    p.add_argument("--sheet-b")

    p = sub.add_parser("search", help="find text in all spreadsheets")
    p.add_argument("text")
    p.add_argument("path", nargs="?", default=".")
    p.add_argument("--regex", action="store_true")
    return ap


def _parse_value(v: str) -> Any:
    if v.startswith("="):
        return v
    try:
        return json.loads(v)
    except ValueError:
        return v


def main(argv: list[str] | None = None) -> int:
    args = build_parser().parse_args(argv)
    try:
        cfg = Config.from_env(read_only=args.read_only or None, **({"roots": args.root} if args.root else {}))
        if args.no_backup:
            cfg.backup = False
        s = Sheets(cfg)
        c = args.cmd
        if c == "ls":
            _out(s.list_files(args.path, pattern=args.pattern))
        elif c == "describe":
            _out(s.describe(args.path, args.sheet))
        elif c == "read":
            _out(
                s.read_range(
                    args.path,
                    args.sheet,
                    args.range,
                    offset=args.offset,
                    limit=args.limit,
                    format="markdown" if args.md else "json",
                ),
                args.md,
            )
        elif c == "write":
            cells = {}
            for a in args.assignments:
                if "=" not in a:
                    raise SheetError(f"Bad assignment '{a}'.", "Use CELL=VALUE, e.g. B2=42 or C2='=B2*2'")
                k, v = a.split("=", 1)
                cells[k] = _parse_value(v)
            _out(s.write_range(args.path, args.sheet, cells=cells, dry_run=args.dry_run))
        elif c == "query":
            _out(
                s.query(
                    args.sql,
                    _tables(args.table),
                    limit=args.limit,
                    output=args.output,
                    format="markdown" if args.md else "json",
                ),
                args.md,
            )
        elif c == "convert":
            _out(s.convert(args.input, args.output, args.sheet, args.all_sheets))
        elif c == "profile":
            _out(s.profile(args.source, args.sheet))
        elif c == "formulas":
            _out(s.formulas(args.path, args.action, args.sheet, args.cell, args.formula))
        elif c == "diff":
            _out(s.diff(args.a, args.b, args.sheet_a, args.sheet_b, args.key))
        elif c == "search":
            _out(s.search(args.text, args.path, regex=args.regex))
        return 0
    except SheetError as e:
        print(f"error: {e}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    sys.exit(main())
