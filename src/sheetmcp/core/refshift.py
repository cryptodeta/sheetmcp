"""Rewrite formula references after inserting/deleting rows or columns (openpyxl doesn't)."""

from __future__ import annotations

import re

from openpyxl.formula.tokenizer import Token, Tokenizer
from openpyxl.utils.cell import column_index_from_string, get_column_letter

from .evaluator import split_ref

_PART = re.compile(r"^(\$?)([A-Za-z]{1,3})?(\$?)(\d+)?$")


def _shift_num(n: int, idx: int, count: int) -> int | None:
    """New position of row/col ``n``. None when it is deleted."""
    if count >= 0:
        return n + count if n >= idx else n
    end = idx - count - 1  # last deleted
    if idx <= n <= end:
        return None
    return n + count if n > end else n


def shift_address(addr: str, axis: str, idx: int, count: int) -> str | None:
    parts = addr.split(":")
    if len(parts) > 2:
        return addr
    parsed = []
    for p in parts:
        m = _PART.match(p)
        if not m:
            return addr
        parsed.append(list(m.groups()))
    gi = 3 if axis == "row" else 1

    def num(g: list[str | None]) -> int | None:
        v = g[gi]
        if v is None:
            return None
        return int(v) if axis == "row" else column_index_from_string(v.upper())

    def put(g: list[str | None], n: int) -> None:
        g[gi] = str(n) if axis == "row" else get_column_letter(n)

    if len(parsed) == 1:
        n = num(parsed[0])
        if n is None:
            return addr
        new = _shift_num(n, idx, count)
        if new is None:
            return None
        put(parsed[0], new)
        return "".join(x or "" for x in parsed[0])
    a, b = parsed
    na, nb = num(a), num(b)
    if na is None or nb is None:
        return addr  # whole-column ref when shifting rows (or vice versa)
    if count >= 0:
        put(a, _shift_num(na, idx, count))  # type: ignore[arg-type]
        put(b, _shift_num(nb, idx, count))  # type: ignore[arg-type]
    else:
        end = idx - count - 1
        sa = _shift_num(na, idx, count)
        sb = _shift_num(nb, idx, count)
        if sa is None:
            sa = idx if nb > end else None
        if sb is None:
            sb = idx - 1 if na < idx else None
        if sa is None or sb is None or sa > sb:
            return None
        put(a, sa)
        put(b, sb)
    return "".join(x or "" for x in a) + ":" + "".join(x or "" for x in b)


def shift_formula(
    formula: str, formula_sheet: str, target_sheet: str, axis: str, idx: int, count: int
) -> str:
    if not isinstance(formula, str) or not formula.startswith("="):
        return formula
    try:
        toks = Tokenizer(formula).items
    except Exception:
        return formula
    out: list[str] = []
    pos = 1
    for t in toks:
        if t.type == Token.WSPACE:
            continue
        i = formula.find(t.value, pos)
        if i < 0:
            return formula
        out.append(formula[pos:i])
        val = t.value
        if t.type == Token.OPERAND and t.subtype == Token.RANGE and not val.startswith("["):
            sh, _ = split_ref(val, formula_sheet)
            if sh.lower() == target_sheet.lower():
                prefix, addr = (
                    (val.rsplit("!", 1)[0] + "!", val.rsplit("!", 1)[1]) if "!" in val else ("", val)
                )
                new = shift_address(addr, axis, idx, count)
                val = "#REF!" if new is None else prefix + new
        out.append(val)
        pos = i + len(t.value)
    out.append(formula[pos:])
    return "=" + "".join(out)
