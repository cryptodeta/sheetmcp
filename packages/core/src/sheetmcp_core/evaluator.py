"""A small, dependency-free Excel formula evaluator.

It covers the formulas people actually write (arithmetic, comparisons, text, ~90 common
functions incl. SUMIFS / VLOOKUP / XLOOKUP / INDEX-MATCH, cross-sheet refs and named ranges).
Unsupported functions return ``#NAME?`` with a clear message instead of guessing.
"""

from __future__ import annotations

import datetime as dt
import fnmatch
import math
import re
import statistics
from collections.abc import Iterable
from typing import Any

from openpyxl.formula.tokenizer import Token, Tokenizer
from openpyxl.utils.cell import range_boundaries
from openpyxl.utils.datetime import from_excel, to_excel

from .util import ERROR_VALUES, cell_ref


class XLErr(Exception):
    def __init__(self, code: str, detail: str = "") -> None:
        self.code = code
        self.detail = detail
        super().__init__(f"{code} {detail}".strip())


class Range:
    __slots__ = ("grid", "origin", "sheet")

    def __init__(self, grid: list[list[Any]], sheet: str = "", origin: tuple[int, int] = (1, 1)) -> None:
        self.grid = grid
        self.sheet = sheet
        self.origin = origin

    @property
    def rows(self) -> int:
        return len(self.grid)

    @property
    def cols(self) -> int:
        return len(self.grid[0]) if self.grid else 0

    def values(self) -> Iterable[Any]:
        for r in self.grid:
            yield from r


EMPTY = object()
INFIX = {"=": 1, "<>": 1, "<": 1, ">": 1, "<=": 1, ">=": 1, "&": 2, "+": 3, "-": 3, "*": 4, "/": 4, "^": 5}


# ------------------------------------------------------------------ parser
class _Parser:
    def __init__(self, formula: str) -> None:
        try:
            items = Tokenizer(formula).items
        except Exception as e:
            raise XLErr("#VALUE!", f"cannot parse formula: {e}") from e
        self.toks = [t for t in items if t.type != Token.WSPACE]
        self.i = 0

    def peek(self) -> Token | None:
        return self.toks[self.i] if self.i < len(self.toks) else None

    def next(self) -> Token:
        if self.i >= len(self.toks):
            raise XLErr("#VALUE!", "unexpected end of formula")
        t = self.toks[self.i]
        self.i += 1
        return t

    def parse(self) -> Any:
        node = self.expr(0)
        if self.peek() is not None:
            raise XLErr("#VALUE!", f"unexpected token {self.peek().value!r}")
        return node

    def expr(self, min_prec: int) -> Any:
        left = self.unary()
        while True:
            t = self.peek()
            if t is None or t.type != Token.OP_IN or t.value not in INFIX:
                break
            prec = INFIX[t.value]
            if prec < min_prec:
                break
            self.next()
            right = self.expr(prec + 1)
            left = ("bin", t.value, left, right)
        return left

    def unary(self) -> Any:
        t = self.peek()
        if t is not None and t.type == Token.OP_PRE:
            self.next()
            operand = self.unary()
            return ("neg", operand) if t.value == "-" else operand
        node = self.primary()
        while (t := self.peek()) is not None and t.type == Token.OP_POST and t.value == "%":
            self.next()
            node = ("pct", node)
        return node

    def primary(self) -> Any:
        t = self.next()
        if t.type == Token.OPERAND:
            if t.subtype == Token.NUMBER:
                v = float(t.value)
                return (
                    "lit",
                    int(v) if v.is_integer() and "." not in t.value and "e" not in t.value.lower() else v,
                )
            if t.subtype == Token.TEXT:
                return ("lit", t.value[1:-1].replace('""', '"'))
            if t.subtype == Token.LOGICAL:
                return ("lit", t.value.upper() == "TRUE")
            if t.subtype == Token.ERROR:
                return ("err", t.value)
            return ("ref", t.value)
        if t.type == Token.FUNC and t.subtype == Token.OPEN:
            name = t.value[:-1].upper()
            for pre in ("_XLFN._XLWS.", "_XLFN.", "_XLWS."):
                if name.startswith(pre):
                    name = name[len(pre) :]
            args: list[Any] = []
            nt = self.peek()
            if nt is not None and nt.type == Token.FUNC and nt.subtype == Token.CLOSE:
                self.next()
                return ("fn", name, args)
            while True:
                nt = self.peek()
                if nt is not None and (
                    (nt.type == Token.SEP and nt.subtype == Token.ARG)
                    or (nt.type == Token.FUNC and nt.subtype == Token.CLOSE)
                ):
                    args.append(("empty",))
                else:
                    args.append(self.expr(0))
                nt = self.next()
                if nt.type == Token.FUNC and nt.subtype == Token.CLOSE:
                    break
                if not (nt.type == Token.SEP and nt.subtype == Token.ARG):
                    raise XLErr("#VALUE!", f"unexpected {nt.value!r} in {name}()")
            return ("fn", name, args)
        if t.type == Token.PAREN and t.subtype == Token.OPEN:
            node = self.expr(0)
            c = self.next()
            if not (c.type == Token.PAREN and c.subtype == Token.CLOSE):
                raise XLErr("#VALUE!", "missing )")
            return node
        if t.type == Token.ARRAY and t.subtype == Token.OPEN:
            rows: list[list[Any]] = [[]]
            while True:
                nt = self.peek()
                if nt is not None and nt.type == Token.ARRAY and nt.subtype == Token.CLOSE:
                    self.next()
                    break
                rows[-1].append(self.expr(0))
                nt = self.next()
                if nt.type == Token.ARRAY and nt.subtype == Token.CLOSE:
                    break
                if nt.type == Token.SEP and nt.subtype == Token.ROW:
                    rows.append([])
            return ("arr", rows)
        raise XLErr("#VALUE!", f"unexpected token {t.value!r}")


_PARSE_CACHE: dict[str, Any] = {}


def parse_formula(formula: str) -> Any:
    f = formula if formula.startswith("=") else "=" + formula
    node = _PARSE_CACHE.get(f)
    if node is None:
        node = _Parser(f).parse()
        if len(_PARSE_CACHE) > 5000:
            _PARSE_CACHE.clear()
        _PARSE_CACHE[f] = node
    return node


def formula_refs(formula: str) -> list[str]:
    """All range operands referenced by a formula (raw text, e.g. 'Sheet2!A1:B3')."""
    try:
        return [
            t.value
            for t in Tokenizer(formula if formula.startswith("=") else "=" + formula).items
            if t.type == Token.OPERAND and t.subtype == Token.RANGE
        ]
    except Exception:
        return []


def split_ref(ref: str, default_sheet: str) -> tuple[str, str]:
    if "!" in ref:
        sh, addr = ref.rsplit("!", 1)
        sh = sh.strip()
        if sh.startswith("'") and sh.endswith("'"):
            sh = sh[1:-1].replace("''", "'")
        return sh, addr.replace("$", "")
    return default_sheet, ref.replace("$", "")


# ------------------------------------------------------------------ value helpers
def to_num(v: Any) -> float | int:
    if isinstance(v, Range):
        v = scalar(v)
    if v is None or v is EMPTY:
        return 0
    if isinstance(v, bool):
        return int(v)
    if isinstance(v, (int, float)):
        return v
    if isinstance(v, dt.datetime):
        return to_excel(v)
    if isinstance(v, dt.date):
        return to_excel(dt.datetime(v.year, v.month, v.day))
    if isinstance(v, dt.time):
        return to_excel(v)
    if isinstance(v, str):
        if v in ERROR_VALUES:
            raise XLErr(v)
        s = v.strip().replace(",", "")
        try:
            if s.endswith("%"):
                return float(s[:-1]) / 100
            return float(s) if s else 0
        except ValueError:
            raise XLErr("#VALUE!", f"'{v}' is not a number") from None
    raise XLErr("#VALUE!")


def to_str(v: Any) -> str:
    if isinstance(v, Range):
        v = scalar(v)
    if v is None or v is EMPTY:
        return ""
    if isinstance(v, bool):
        return "TRUE" if v else "FALSE"
    if isinstance(v, float):
        return (
            str(int(v))
            if v.is_integer()
            else repr(round(v, 15)).rstrip("0").rstrip(".")
            if "e" not in repr(v)
            else repr(v)
        )
    if isinstance(v, str) and v in ERROR_VALUES:
        raise XLErr(v)
    return str(v)


def to_bool(v: Any) -> bool:
    if isinstance(v, Range):
        v = scalar(v)
    if v is None or v is EMPTY:
        return False
    if isinstance(v, bool):
        return v
    if isinstance(v, (int, float)):
        return v != 0
    if isinstance(v, str):
        if v in ERROR_VALUES:
            raise XLErr(v)
        if v.upper() in {"TRUE", "FALSE"}:
            return v.upper() == "TRUE"
    raise XLErr("#VALUE!", f"'{v}' is not TRUE/FALSE")


def scalar(v: Any) -> Any:
    if isinstance(v, Range):
        if v.rows == 1 and v.cols == 1:
            return v.grid[0][0]
        raise XLErr("#VALUE!", "a multi-cell range was used where a single value is expected")
    if isinstance(v, str) and v in ERROR_VALUES:
        raise XLErr(v)
    return v


def _rank(v: Any) -> int:
    if isinstance(v, bool):
        return 2
    if isinstance(v, str):
        return 1
    return 0


def compare(a: Any, b: Any, op: str) -> bool:
    a, b = scalar(a), scalar(b)
    if a is None:
        a = "" if isinstance(b, str) else (False if isinstance(b, bool) else 0)
    if b is None:
        b = "" if isinstance(a, str) else (False if isinstance(a, bool) else 0)
    if isinstance(a, (dt.date, dt.time)):
        a = to_num(a)
    if isinstance(b, (dt.date, dt.time)):
        b = to_num(b)
    ra, rb = _rank(a), _rank(b)
    if ra != rb:
        x, y = ra, rb
    elif isinstance(a, str):
        x, y = a.lower(), b.lower()
    else:
        x, y = a, b
    return {"=": x == y, "<>": x != y, "<": x < y, ">": x > y, "<=": x <= y, ">=": x >= y}[op]


def make_criteria(c: Any):
    c = scalar(c)
    if isinstance(c, (int, float)) and not isinstance(c, bool):
        return lambda v: (
            (not isinstance(v, (str, bool)) and v is not None and to_num(v) == c)
            or (isinstance(v, str) and _is_numlike(v) and float(v) == c)
        )
    if isinstance(c, bool):
        return lambda v: v is c
    s = "" if c is None else str(c)
    m = re.match(r"^(<=|>=|<>|=|<|>)(.*)$", s, re.S)
    op, rhs = (m.group(1), m.group(2)) if m else ("=", s)
    num = float(rhs) if _is_numlike(rhs) else None
    if num is not None:

        def f(v: Any) -> bool:
            if v is None or (isinstance(v, (str, bool)) and not (isinstance(v, str) and _is_numlike(v))):
                return op == "<>"
            return compare(to_num(v), num, op)

        return f
    if op in {"=", "<>"}:
        if rhs == "":
            return (lambda v: v is None or v == "") if op == "=" else (lambda v: not (v is None or v == ""))
        if any(ch in rhs for ch in "*?"):
            pat = rhs.lower().replace("~*", "\x00").replace("~?", "\x01")
            pat = fnmatch.translate(pat).replace("\x00", r"\*").replace("\x01", r"\?")
            rx = re.compile(pat, re.S)
            return (
                (lambda v: bool(rx.match(to_str(v).lower())))
                if op == "="
                else (lambda v: not rx.match(to_str(v).lower()))
            )
        return (
            (lambda v: to_str(v).lower() == rhs.lower())
            if op == "="
            else (lambda v: to_str(v).lower() != rhs.lower())
        )
    return lambda v: isinstance(v, str) and compare(v, rhs, op)


def _is_numlike(s: str) -> bool:
    try:
        float(s)
        return True
    except (TypeError, ValueError):
        return False


def _date(v: Any) -> dt.datetime:
    v = scalar(v)
    if isinstance(v, dt.datetime):
        return v
    if isinstance(v, dt.date):
        return dt.datetime(v.year, v.month, v.day)
    if isinstance(v, str):
        try:
            return dt.datetime.fromisoformat(v)
        except ValueError:
            pass
    return from_excel(to_num(v))


# ------------------------------------------------------------------ evaluator
class Evaluator:
    """Evaluate formulas of an openpyxl workbook loaded with data_only=False."""

    def __init__(self, wb: Any, max_range_cells: int = 500_000) -> None:
        self.wb = wb
        self.memo: dict[tuple[str, str], Any] = {}
        self.stack: set[tuple[str, str]] = set()
        self.max_range_cells = max_range_cells
        self.unsupported: set[str] = set()

    # public -----------------------------------------------------------
    def evaluate_cell(self, sheet: str, coord: str) -> Any:
        """Return the computed value; Excel errors are returned as strings like '#DIV/0!'."""
        try:
            v = self._cell(sheet, coord.replace("$", "").upper())
            return v
        except XLErr as e:
            return e.code

    def evaluate_formula(self, formula: str, sheet: str) -> Any:
        try:
            r = self.ev(parse_formula(formula), sheet)
            if isinstance(r, Range):
                return r.grid[0][0] if r.rows == 1 and r.cols == 1 else [list(x) for x in r.grid]
            return r
        except XLErr as e:
            return e.code
        except RecursionError:
            return "#NUM!"

    # internals ---------------------------------------------------------
    def _ws(self, sheet: str):
        if sheet in self.wb.sheetnames:
            return self.wb[sheet]
        for n in self.wb.sheetnames:
            if n.lower() == sheet.lower():
                return self.wb[n]
        raise XLErr("#REF!", f"sheet '{sheet}' not found")

    def _cell(self, sheet: str, coord: str) -> Any:
        ws = self._ws(sheet)
        key = (ws.title, coord)
        if key in self.memo:
            return self.memo[key]
        raw = ws[coord].value
        f = getattr(raw, "text", None)  # ArrayFormula
        if isinstance(raw, str) and raw.startswith("=") and len(raw) > 1:
            f = raw
        if f:
            if key in self.stack:
                raise XLErr("#CIRC!", f"circular reference at {ws.title}!{coord}")
            self.stack.add(key)
            try:
                try:
                    v = self.ev(parse_formula(f), ws.title)
                    if isinstance(v, Range):
                        v = v.grid[0][0] if v.grid and v.grid[0] else None
                    if v is EMPTY:
                        v = 0
                except XLErr as e:
                    if e.code == "#CIRC!":
                        raise
                    v = e.code
            finally:
                self.stack.discard(key)
            self.memo[key] = v
            return v
        if raw is not None and raw.__class__.__name__ == "DataTableFormula":
            raise XLErr("#NAME?", "data table formulas are not supported")
        self.memo[key] = raw
        return raw

    def _value(self, sheet: str, coord: str) -> Any:
        v = self._cell(sheet, coord)
        return v

    def ref(self, text: str, sheet: str) -> Range:
        if text.startswith("["):
            raise XLErr("#REF!", "external workbook references are not supported")
        sh, addr = split_ref(text, sheet)
        if "[" in addr:
            raise XLErr("#NAME?", "structured table references (Table[Col]) are not supported yet")
        if not re.match(r"^[A-Za-z]{0,3}\d*(:[A-Za-z]{0,3}\d*)?$", addr) or addr in {"", ":"}:
            return self._named(text, sheet)
        try:
            min_col, min_row, max_col, max_row = range_boundaries(addr.upper())
        except ValueError:
            return self._named(text, sheet)
        ws = self._ws(sh)
        min_col = min_col or 1
        min_row = min_row or 1
        max_col = max_col or max(ws.max_column, 1)
        max_row = max_row or max(ws.max_row, 1)
        max_row = min(max_row, max(ws.max_row, min_row))
        max_col = min(max_col, max(ws.max_column, min_col))
        if (max_row - min_row + 1) * (max_col - min_col + 1) > self.max_range_cells:
            raise XLErr("#NUM!", "range too large to evaluate")
        grid = [
            [self._value(ws.title, cell_ref(c, r)) for c in range(min_col, max_col + 1)]
            for r in range(min_row, max_row + 1)
        ]
        return Range(grid, ws.title, (min_row, min_col))

    def _named(self, name: str, sheet: str) -> Range:
        ws = self._ws(sheet)
        dn = None
        scoped = getattr(ws, "defined_names", None)
        if scoped and name in scoped:
            dn = scoped[name]
        if dn is None:
            dns = self.wb.defined_names
            dn = dns.get(name) if hasattr(dns, "get") else None
            if dn is None:
                for k in list(dns.keys()) if hasattr(dns, "keys") else []:
                    if k.lower() == name.lower():
                        dn = dns[k]
                        break
        if dn is None:
            raise XLErr("#NAME?", f"unknown name '{name}'")
        dest = list(dn.destinations)
        if not dest:
            val = dn.attr_text
            return Range([[self.evaluate_formula(val, sheet)]])
        sh, coord = dest[0]
        return self.ref(f"'{sh}'!{coord}", sheet)

    def ev(self, node: Any, sheet: str) -> Any:
        kind = node[0]
        if kind == "lit":
            return node[1]
        if kind == "err":
            raise XLErr(node[1])
        if kind == "ref":
            return self.ref(node[1], sheet)
        if kind == "empty":
            return EMPTY
        if kind == "neg":
            return -to_num(self.ev(node[1], sheet))
        if kind == "pct":
            return to_num(self.ev(node[1], sheet)) / 100
        if kind == "arr":
            return Range([[scalar(self.ev(x, sheet)) for x in row] for row in node[1]])
        if kind == "bin":
            op = node[1]
            a = self.ev(node[2], sheet)
            b = self.ev(node[3], sheet)
            if op == "&":
                return to_str(a) + to_str(b)
            if op in {"=", "<>", "<", ">", "<=", ">="}:
                return compare(a, b, op)
            x, y = to_num(a), to_num(b)
            if op == "+":
                return _clean(x + y)
            if op == "-":
                return _clean(x - y)
            if op == "*":
                return _clean(x * y)
            if op == "/":
                if y == 0:
                    raise XLErr("#DIV/0!")
                return _clean(x / y)
            if op == "^":
                try:
                    return _clean(x**y)
                except (OverflowError, ZeroDivisionError, ValueError):
                    raise XLErr("#NUM!") from None
        if kind == "fn":
            name, args = node[1], node[2]
            fn = FUNCTIONS.get(name)
            if fn is None:
                self.unsupported.add(name)
                raise XLErr("#NAME?", f"function {name} is not supported by the built-in evaluator")
            return fn(self, args, sheet)
        raise XLErr("#VALUE!")


def _clean(v: float | int) -> float | int:
    if isinstance(v, float):
        if math.isnan(v) or math.isinf(v):
            raise XLErr("#NUM!")
        r = round(v, 12)
        return int(r) if r.is_integer() and abs(r) < 1e15 else v
    return v


# ------------------------------------------------------------------ functions
FUNCTIONS: dict[str, Any] = {}


def fn(*names: str, lazy: bool = False):
    def deco(f):
        for n in names:
            if lazy:
                FUNCTIONS[n] = f
            else:
                FUNCTIONS[n] = lambda ev, args, sheet, _f=f: _f(*[ev.ev(a, sheet) for a in args])
        return f

    return deco


def _nums(args: Iterable[Any], strict_direct: bool = True) -> list[float]:
    """Numbers for SUM-like functions: from ranges only numeric cells, direct args are coerced."""
    out: list[float] = []
    for a in args:
        if a is EMPTY:
            continue
        if isinstance(a, Range):
            for v in a.values():
                if isinstance(v, str) and v in ERROR_VALUES:
                    raise XLErr(v)
                if isinstance(v, (int, float)) and not isinstance(v, bool):
                    out.append(v)
                elif isinstance(v, (dt.date, dt.time)):
                    out.append(to_num(v))
        else:
            out.append(to_num(a))
    return out


def _flat(args: Iterable[Any]) -> list[Any]:
    out: list[Any] = []
    for a in args:
        if isinstance(a, Range):
            out.extend(a.values())
        elif a is not EMPTY:
            out.append(a)
    return out


def _opt(args: tuple[Any, ...], i: int, default: Any) -> Any:
    return default if len(args) <= i or args[i] is EMPTY else scalar(args[i])


@fn("SUM")
def _sum(*a):
    return _clean(math.fsum(_nums(a)))


@fn("AVERAGE")
def _avg(*a):
    n = _nums(a)
    if not n:
        raise XLErr("#DIV/0!")
    return _clean(math.fsum(n) / len(n))


@fn("MIN")
def _min(*a):
    n = _nums(a)
    return min(n) if n else 0


@fn("MAX")
def _max(*a):
    n = _nums(a)
    return max(n) if n else 0


@fn("COUNT")
def _count(*a):
    c = 0
    for x in a:
        if isinstance(x, Range):
            c += sum(
                1 for v in x.values() if isinstance(v, (int, float, dt.date)) and not isinstance(v, bool)
            )
        elif x is not EMPTY:
            try:
                to_num(x)
                c += 1
            except XLErr:
                pass
    return c


@fn("COUNTA")
def _counta(*a):
    return sum(1 for v in _flat(a) if v is not None and v != "")


@fn("COUNTBLANK")
def _countblank(*a):
    return sum(1 for v in _flat(a) if v is None or v == "")


@fn("PRODUCT")
def _product(*a):
    return _clean(math.prod(_nums(a)))


@fn("MEDIAN")
def _median(*a):
    n = _nums(a)
    if not n:
        raise XLErr("#NUM!")
    return _clean(statistics.median(n))


@fn("STDEV", "STDEV.S")
def _stdev(*a):
    n = _nums(a)
    if len(n) < 2:
        raise XLErr("#DIV/0!")
    return statistics.stdev(n)


@fn("STDEVP", "STDEV.P")
def _stdevp(*a):
    n = _nums(a)
    if not n:
        raise XLErr("#DIV/0!")
    return statistics.pstdev(n)


@fn("VAR", "VAR.S")
def _var(*a):
    n = _nums(a)
    if len(n) < 2:
        raise XLErr("#DIV/0!")
    return statistics.variance(n)


@fn("LARGE")
def _large(r, k):
    n = sorted(_nums([r]), reverse=True)
    k = int(to_num(k))
    if k < 1 or k > len(n):
        raise XLErr("#NUM!")
    return n[k - 1]


@fn("SMALL")
def _small(r, k):
    n = sorted(_nums([r]))
    k = int(to_num(k))
    if k < 1 or k > len(n):
        raise XLErr("#NUM!")
    return n[k - 1]


def _round_half_away(x: float, d: int) -> float:
    m = 10**d
    return _clean(math.floor(abs(x) * m + 0.5 + 1e-9) / m * (1 if x >= 0 else -1))


@fn("ROUND")
def _round(x, d=0):
    return _round_half_away(to_num(x), int(to_num(d)))


@fn("ROUNDUP")
def _roundup(x, d=0):
    x, m = to_num(x), 10 ** int(to_num(d))
    return _clean(math.ceil(abs(x) * m - 1e-9) / m * (1 if x >= 0 else -1))


@fn("ROUNDDOWN", "TRUNC")
def _rounddown(x, d=0):
    x, m = to_num(x), 10 ** int(to_num(d))
    return _clean(math.floor(abs(x) * m + 1e-9) / m * (1 if x >= 0 else -1))


@fn("INT")
def _int(x):
    return math.floor(to_num(x))


@fn("ABS")
def _abs(x):
    return abs(to_num(x))


@fn("MOD")
def _mod(a, b):
    a, b = to_num(a), to_num(b)
    if b == 0:
        raise XLErr("#DIV/0!")
    return _clean(a - b * math.floor(a / b))


@fn("POWER")
def _power(a, b):
    try:
        return _clean(to_num(a) ** to_num(b))
    except (OverflowError, ZeroDivisionError, ValueError):
        raise XLErr("#NUM!") from None


@fn("SQRT")
def _sqrt(x):
    x = to_num(x)
    if x < 0:
        raise XLErr("#NUM!")
    return _clean(math.sqrt(x))


@fn("EXP")
def _exp(x):
    return math.exp(to_num(x))


@fn("LN")
def _ln(x):
    x = to_num(x)
    if x <= 0:
        raise XLErr("#NUM!")
    return math.log(x)


@fn("LOG10")
def _log10(x):
    x = to_num(x)
    if x <= 0:
        raise XLErr("#NUM!")
    return math.log10(x)


@fn("LOG")
def _log(x, b=10):
    x, b = to_num(x), to_num(b)
    if x <= 0 or b <= 0 or b == 1:
        raise XLErr("#NUM!")
    return _clean(math.log(x, b))


@fn("SIGN")
def _sign(x):
    x = to_num(x)
    return (x > 0) - (x < 0)


@fn("PI")
def _pi():
    return math.pi


@fn("CEILING", "CEILING.MATH")
def _ceiling(x, s=1):
    x, s = to_num(x), to_num(s)
    return 0 if s == 0 else _clean(math.ceil(x / s - 1e-12) * s)


@fn("FLOOR", "FLOOR.MATH")
def _floor(x, s=1):
    x, s = to_num(x), to_num(s)
    if s == 0:
        raise XLErr("#DIV/0!")
    return _clean(math.floor(x / s + 1e-12) * s)


@fn("IF", lazy=True)
def _if(ev, args, sheet):
    if not args or len(args) > 3:
        raise XLErr("#VALUE!", "IF needs 1-3 arguments")
    cond = to_bool(ev.ev(args[0], sheet))
    if cond:
        return (
            ev.ev(args[1], sheet)
            if len(args) > 1 and args[1] != ("empty",)
            else (True if len(args) < 2 else 0)
        )
    if len(args) > 2:
        v = ev.ev(args[2], sheet)
        return 0 if v is EMPTY else v
    return False


@fn("IFS", lazy=True)
def _ifs(ev, args, sheet):
    for i in range(0, len(args) - 1, 2):
        if to_bool(ev.ev(args[i], sheet)):
            return ev.ev(args[i + 1], sheet)
    raise XLErr("#N/A")


@fn("IFERROR", lazy=True)
def _iferror(ev, args, sheet):
    try:
        v = ev.ev(args[0], sheet)
        scalar(v) if not isinstance(v, Range) else None
        return v
    except XLErr as e:
        if e.code == "#CIRC!":
            raise
        return ev.ev(args[1], sheet)


@fn("IFNA", lazy=True)
def _ifna(ev, args, sheet):
    try:
        v = ev.ev(args[0], sheet)
        scalar(v) if not isinstance(v, Range) else None
        return v
    except XLErr as e:
        if e.code != "#N/A":
            raise
        return ev.ev(args[1], sheet)


@fn("AND")
def _and(*a):
    vals = [
        v
        for v in _flat(a)
        if (v is not None and not isinstance(v, str))
        or (isinstance(v, str) and v.upper() in {"TRUE", "FALSE"})
    ]
    return all(to_bool(v) for v in vals)


@fn("OR")
def _or(*a):
    vals = [
        v
        for v in _flat(a)
        if (v is not None and not isinstance(v, str))
        or (isinstance(v, str) and v.upper() in {"TRUE", "FALSE"})
    ]
    return any(to_bool(v) for v in vals)


@fn("XOR")
def _xor(*a):
    return sum(to_bool(v) for v in _flat(a) if v is not None) % 2 == 1


@fn("NOT")
def _not(x):
    return not to_bool(x)


@fn("TRUE")
def _true():
    return True


@fn("FALSE")
def _false():
    return False


@fn("NA")
def _na():
    raise XLErr("#N/A")


def _is(test):
    def f(ev, args, sheet):
        try:
            v = ev.ev(args[0], sheet)
            v = v.grid[0][0] if isinstance(v, Range) and v.grid else v
            if isinstance(v, str) and v in ERROR_VALUES:
                raise XLErr(v)
            return test(v, None)
        except XLErr as e:
            if e.code == "#CIRC!":
                raise
            return test(None, e.code)

    return f


FUNCTIONS["ISBLANK"] = _is(lambda v, e: e is None and v is None)
FUNCTIONS["ISNUMBER"] = _is(
    lambda v, e: e is None and isinstance(v, (int, float)) and not isinstance(v, bool)
)
FUNCTIONS["ISTEXT"] = _is(lambda v, e: e is None and isinstance(v, str))
FUNCTIONS["ISERROR"] = _is(lambda v, e: e is not None)
FUNCTIONS["ISERR"] = _is(lambda v, e: e is not None and e != "#N/A")
FUNCTIONS["ISNA"] = _is(lambda v, e: e == "#N/A")
FUNCTIONS["ISLOGICAL"] = _is(lambda v, e: e is None and isinstance(v, bool))


@fn("CONCAT", "CONCATENATE")
def _concat(*a):
    return "".join(to_str(v) for v in _flat(a))


@fn("TEXTJOIN")
def _textjoin(delim, ignore_empty, *a):
    vals = [to_str(v) for v in _flat(a)]
    if to_bool(ignore_empty):
        vals = [v for v in vals if v != ""]
    return to_str(delim).join(vals)


@fn("LEN")
def _len(x):
    return len(to_str(x))


@fn("UPPER")
def _upper(x):
    return to_str(x).upper()


@fn("LOWER")
def _lower(x):
    return to_str(x).lower()


@fn("PROPER")
def _proper(x):
    return to_str(x).title()


@fn("TRIM")
def _trim(x):
    return re.sub(" +", " ", to_str(x).strip())


@fn("LEFT")
def _left(x, n=1):
    return to_str(x)[: int(to_num(n))]


@fn("RIGHT")
def _right(x, n=1):
    n = int(to_num(n))
    return to_str(x)[-n:] if n else ""


@fn("MID")
def _mid(x, start, n):
    s, st = to_str(x), int(to_num(start))
    if st < 1:
        raise XLErr("#VALUE!")
    return s[st - 1 : st - 1 + int(to_num(n))]


@fn("SUBSTITUTE")
def _substitute(text, old, new, inst=EMPTY):
    t, o, n = to_str(text), to_str(old), to_str(new)
    if inst is EMPTY:
        return t.replace(o, n)
    k, idx = int(to_num(inst)), -1
    for _ in range(k):
        idx = t.find(o, idx + 1)
        if idx < 0:
            return t
    return t[:idx] + n + t[idx + len(o) :]


@fn("REPLACE")
def _replace(text, start, n, new):
    t, s = to_str(text), int(to_num(start))
    return t[: s - 1] + to_str(new) + t[s - 1 + int(to_num(n)) :]


@fn("FIND")
def _find(needle, hay, start=1):
    i = to_str(hay).find(to_str(needle), int(to_num(start)) - 1)
    if i < 0:
        raise XLErr("#VALUE!")
    return i + 1


@fn("SEARCH")
def _search(needle, hay, start=1):
    i = to_str(hay).lower().find(to_str(needle).lower(), int(to_num(start)) - 1)
    if i < 0:
        raise XLErr("#VALUE!")
    return i + 1


@fn("VALUE")
def _value(x):
    return to_num(x)


@fn("REPT")
def _rept(x, n):
    return to_str(x) * int(to_num(n))


@fn("EXACT")
def _exact(a, b):
    return to_str(a) == to_str(b)


@fn("TEXT")
def _text(v, fmt):
    v, f = scalar(v), to_str(fmt)
    if isinstance(v, (dt.date, dt.datetime)) or (
        re.search(r"[ymd]", f.lower()) and not re.search(r"[0#]", f)
    ):
        d = _date(v)
        py = f.replace("yyyy", "%Y").replace("yy", "%y").replace("mmmm", "%B").replace("mmm", "%b")
        py = py.replace("mm", "%m").replace("dd", "%d").replace("hh", "%H").replace("ss", "%S")
        return d.strftime(py)
    x = to_num(v)
    pct = f.endswith("%")
    if pct:
        x *= 100
    dec = len(f.split(".")[1].rstrip("%")) if "." in f else 0
    s = f"{x:,.{dec}f}" if "," in f else f"{x:.{dec}f}"
    return s + ("%" if pct else "")


@fn("CHOOSE")
def _choose(i, *opts):
    i = int(to_num(i))
    if i < 1 or i > len(opts):
        raise XLErr("#VALUE!")
    return opts[i - 1]


@fn("ROWS")
def _rows(r):
    return r.rows if isinstance(r, Range) else 1


@fn("COLUMNS")
def _columns(r):
    return r.cols if isinstance(r, Range) else 1


@fn("TODAY")
def _today():
    return dt.date.today()


@fn("NOW")
def _now():
    return dt.datetime.now().replace(microsecond=0)


@fn("DATE")
def _datefn(y, m, d):
    y, m, d = int(to_num(y)), int(to_num(m)), int(to_num(d))
    y += (m - 1) // 12
    m = (m - 1) % 12 + 1
    return dt.datetime(y, m, 1) + dt.timedelta(days=d - 1)


@fn("YEAR")
def _year(v):
    return _date(v).year


@fn("MONTH")
def _month(v):
    return _date(v).month


@fn("DAY")
def _day(v):
    return _date(v).day


@fn("EDATE")
def _edate(v, months):
    d, m = _date(v), int(to_num(months))
    y, mo = d.year + (d.month - 1 + m) // 12, (d.month - 1 + m) % 12 + 1
    import calendar

    return d.replace(year=y, month=mo, day=min(d.day, calendar.monthrange(y, mo)[1]))


@fn("DAYS")
def _days(end, start):
    return (_date(end) - _date(start)).days


def _pairs(ranges_criteria: list[Any]) -> list[tuple[Range, Any]]:
    out = []
    for i in range(0, len(ranges_criteria), 2):
        r = ranges_criteria[i]
        if not isinstance(r, Range):
            r = Range([[r]])
        out.append((r, make_criteria(ranges_criteria[i + 1])))
    return out


def _mask(pairs: list[tuple[Range, Any]]) -> list[bool]:
    first = pairs[0][0]
    n = first.rows * first.cols
    mask = [True] * n
    for r, crit in pairs:
        vals = list(r.values())
        if len(vals) != n:
            raise XLErr("#VALUE!", "criteria ranges must be the same size")
        for i, v in enumerate(vals):
            if mask[i] and not crit(v):
                mask[i] = False
    return mask


def _vals(r: Any, n: int) -> list[Any]:
    vals = list(r.values()) if isinstance(r, Range) else [r]
    if len(vals) != n:
        raise XLErr("#VALUE!", "ranges must be the same size")
    return vals


def _numeric(v: Any) -> bool:
    return isinstance(v, (int, float)) and not isinstance(v, bool)


@fn("SUMIF")
def _sumif(r, crit, sum_r=None):
    pairs = _pairs([r, crit])
    mask = _mask(pairs)
    vals = _vals(sum_r if sum_r is not None and sum_r is not EMPTY else r, len(mask))
    return _clean(math.fsum(v for v, m in zip(vals, mask) if m and _numeric(v)))


@fn("SUMIFS")
def _sumifs(sum_r, *rc):
    mask = _mask(_pairs(list(rc)))
    return _clean(math.fsum(v for v, m in zip(_vals(sum_r, len(mask)), mask) if m and _numeric(v)))


@fn("COUNTIF")
def _countif(r, crit):
    return sum(_mask(_pairs([r, crit])))


@fn("COUNTIFS")
def _countifs(*rc):
    return sum(_mask(_pairs(list(rc))))


@fn("AVERAGEIF")
def _averageif(r, crit, avg_r=None):
    mask = _mask(_pairs([r, crit]))
    vals = [
        v
        for v, m in zip(_vals(avg_r if avg_r not in (None, EMPTY) else r, len(mask)), mask)
        if m and _numeric(v)
    ]
    if not vals:
        raise XLErr("#DIV/0!")
    return _clean(math.fsum(vals) / len(vals))


@fn("AVERAGEIFS")
def _averageifs(avg_r, *rc):
    mask = _mask(_pairs(list(rc)))
    vals = [v for v, m in zip(_vals(avg_r, len(mask)), mask) if m and _numeric(v)]
    if not vals:
        raise XLErr("#DIV/0!")
    return _clean(math.fsum(vals) / len(vals))


@fn("MAXIFS")
def _maxifs(r, *rc):
    mask = _mask(_pairs(list(rc)))
    vals = [v for v, m in zip(_vals(r, len(mask)), mask) if m and _numeric(v)]
    return max(vals) if vals else 0


@fn("MINIFS")
def _minifs(r, *rc):
    mask = _mask(_pairs(list(rc)))
    vals = [v for v, m in zip(_vals(r, len(mask)), mask) if m and _numeric(v)]
    return min(vals) if vals else 0


@fn("SUMPRODUCT")
def _sumproduct(*arrs):
    lists = [list(a.values()) if isinstance(a, Range) else [a] for a in arrs]
    n = len(lists[0])
    if any(len(x) != n for x in lists):
        raise XLErr("#VALUE!")
    total = 0.0
    for i in range(n):
        p = 1.0
        for x in lists:
            v = x[i]
            p *= v if _numeric(v) else (int(v) if isinstance(v, bool) else 0)
        total += p
    return _clean(total)


def _match_index(needle: Any, vals: list[Any], mode: int) -> int:
    needle = scalar(needle)
    if mode == 0:
        crit = make_criteria(needle) if isinstance(needle, str) and any(c in needle for c in "*?") else None
        for i, v in enumerate(vals):
            if crit(v) if crit else (v is not None and compare(v, needle, "=")):
                return i
        raise XLErr("#N/A")
    best = -1
    for i, v in enumerate(vals):
        if v is None or _rank(v) != _rank(needle):
            continue
        if (mode == 1 and compare(v, needle, "<=")) or (mode == -1 and compare(v, needle, ">=")):
            best = i
    if best < 0:
        raise XLErr("#N/A")
    return best


@fn("MATCH")
def _match(needle, r, mode=1):
    vals = list(r.values()) if isinstance(r, Range) else [r]
    return _match_index(needle, vals, int(to_num(mode))) + 1


@fn("INDEX")
def _index(r, row=0, col=EMPTY):
    if not isinstance(r, Range):
        r = Range([[r]])
    row = int(to_num(row)) if row is not EMPTY else 0
    col = int(to_num(col)) if col is not EMPTY else 0
    if r.rows == 1 and col == 0 and row:
        row, col = 1, row
    if row < 0 or col < 0 or row > r.rows or col > r.cols:
        raise XLErr("#REF!")
    if row == 0 and col == 0:
        return r
    if row == 0:
        return Range([[g[col - 1]] for g in r.grid], r.sheet)
    if col == 0:
        return Range([list(r.grid[row - 1])], r.sheet) if r.cols > 1 else r.grid[row - 1][0]
    return r.grid[row - 1][col - 1]


@fn("VLOOKUP")
def _vlookup(needle, table, col, approx=True):
    if not isinstance(table, Range):
        raise XLErr("#VALUE!")
    c = int(to_num(col))
    if c < 1 or c > table.cols:
        raise XLErr("#REF!")
    approx = True if approx is EMPTY else to_bool(approx)
    i = _match_index(needle, [r[0] for r in table.grid], 1 if approx else 0)
    return table.grid[i][c - 1]


@fn("HLOOKUP")
def _hlookup(needle, table, row, approx=True):
    if not isinstance(table, Range):
        raise XLErr("#VALUE!")
    r = int(to_num(row))
    if r < 1 or r > table.rows:
        raise XLErr("#REF!")
    approx = True if approx is EMPTY else to_bool(approx)
    i = _match_index(needle, list(table.grid[0]), 1 if approx else 0)
    return table.grid[r - 1][i]


@fn("XLOOKUP")
def _xlookup(needle, look, ret, if_missing=EMPTY, match_mode=0):
    vals = list(look.values()) if isinstance(look, Range) else [look]
    mm = int(to_num(match_mode)) if match_mode is not EMPTY else 0
    try:
        i = _match_index(needle, vals, {0: 0, -1: 1, 1: -1}.get(mm, 0) if mm in (0, -1, 1) else 0)
    except XLErr:
        if if_missing is not EMPTY:
            return if_missing
        raise
    if not isinstance(ret, Range):
        return ret
    if ret.cols == 1 or (look.cols == 1 and ret.rows == len(vals)):
        return ret.grid[i][0] if ret.cols == 1 else Range([list(ret.grid[i])])
    return ret.grid[0][i] if ret.rows == 1 else Range([[g[i]] for g in ret.grid])
