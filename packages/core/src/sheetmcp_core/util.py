"""Small helpers: JSON-safe values, A1 notation, markdown tables."""

from __future__ import annotations

import datetime as dt
import math
import re
from decimal import Decimal
from typing import Any

from openpyxl.utils.cell import (
    column_index_from_string,
    get_column_letter,
    range_boundaries,
)

from .errors import SheetError

ERROR_VALUES = {
    "#NULL!",
    "#DIV/0!",
    "#VALUE!",
    "#REF!",
    "#NAME?",
    "#NUM!",
    "#N/A",
    "#GETTING_DATA",
    "#SPILL!",
    "#CALC!",
}


def jsonable(v: Any, max_chars: int | None = None) -> Any:
    """Convert a cell / DB value into something JSON can carry."""
    if v is None or isinstance(v, (bool, int)):
        return v
    if isinstance(v, float):
        return None if math.isnan(v) or math.isinf(v) else v
    if isinstance(v, Decimal):
        return float(v)
    if isinstance(v, dt.datetime):
        if v.hour == v.minute == v.second == v.microsecond == 0 and v.tzinfo is None:
            return v.date().isoformat()
        return v.isoformat()
    if isinstance(v, (dt.date, dt.time)):
        return v.isoformat()
    if isinstance(v, dt.timedelta):
        return str(v)
    if isinstance(v, bytes):
        return f"<{len(v)} bytes>"
    if isinstance(v, (list, tuple)):
        return [jsonable(x, max_chars) for x in v]
    if isinstance(v, dict):
        return {str(k): jsonable(x, max_chars) for k, x in v.items()}
    s = str(v)
    if max_chars and len(s) > max_chars:
        s = s[:max_chars] + f"…(+{len(s) - max_chars} chars)"
    return s


def col_letter(idx: int) -> str:
    return get_column_letter(idx)


def col_index(letter: str) -> int:
    return column_index_from_string(letter.upper())


_COL_RE = re.compile(r"^[A-Za-z]{1,3}$")


def parse_range(
    ref: str, max_row: int | None = None, max_col: int | None = None
) -> tuple[int, int, int, int]:
    """Return (min_col, min_row, max_col, max_row) for 'A1', 'A1:C9', 'A:C', '2:5'."""
    ref = ref.strip().replace("$", "")
    if "!" in ref:
        ref = ref.split("!", 1)[1]
    try:
        min_col, min_row, mx_col, mx_row = range_boundaries(ref)
    except (ValueError, TypeError) as e:
        raise SheetError(
            f"'{ref}' is not a valid cell range.",
            "Use A1 notation like 'B2', 'A1:D20', 'A:C' (whole columns) or '3:10' (whole rows).",
        ) from e
    min_col = min_col or 1
    min_row = min_row or 1
    mx_col = mx_col or max_col or min_col
    mx_row = mx_row or max_row or min_row
    return min_col, min_row, mx_col, mx_row


def cell_ref(col: int, row: int) -> str:
    return f"{get_column_letter(col)}{row}"


def md_escape(v: Any) -> str:
    s = "" if v is None else str(v)
    return s.replace("|", "\\|").replace("\n", " ")


def to_markdown(columns: list[str], rows: list[list[Any]] | list[tuple[Any, ...]]) -> str:
    head = "| " + " | ".join(md_escape(c) for c in columns) + " |"
    sep = "| " + " | ".join("---" for _ in columns) + " |"
    body = ["| " + " | ".join(md_escape(jsonable(v)) for v in r) + " |" for r in rows]
    return "\n".join([head, sep, *body])


def safe_identifier(name: str, used: set[str] | None = None) -> str:
    """Make a SQL-friendly table alias out of a filename or sheet name."""
    s = re.sub(r"\W+", "_", name.strip()).strip("_").lower() or "t"
    if s[0].isdigit():
        s = "t_" + s
    if used is not None:
        base, i = s, 2
        while s in used:
            s = f"{base}_{i}"
            i += 1
        used.add(s)
    return s


def qident(name: str) -> str:
    """Quote a SQL identifier for DuckDB."""
    return '"' + str(name).replace('"', '""') + '"'


def qstr(s: str) -> str:
    return "'" + str(s).replace("'", "''") + "'"


def is_error_value(v: Any) -> bool:
    return isinstance(v, str) and v in ERROR_VALUES


def unique_headers(raw: list[Any]) -> list[str]:
    out: list[str] = []
    seen: dict[str, int] = {}
    for i, h in enumerate(raw, start=1):
        name = str(jsonable(h)).strip() if h is not None and str(h).strip() != "" else f"column_{i}"
        if name in seen:
            seen[name] += 1
            name = f"{name}_{seen[name]}"
        else:
            seen[name] = 1
        out.append(name)
    return out
