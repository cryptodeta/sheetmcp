"""Read-only importers for formats openpyxl can't open: legacy .xls (optional xlrd), .ods and HTML tables.

Each reader returns ``{sheet_name: rows}``; results are cached by (path, mtime, size).
"""

from __future__ import annotations

import contextlib
import datetime as dt
import threading
import zipfile
from html.parser import HTMLParser
from pathlib import Path
from typing import Any
from xml.etree import ElementTree as ET

from .errors import SheetError

_cache: dict[tuple[Any, ...], dict[str, list[list[Any]]]] = {}
_lock = threading.Lock()


def read_sheets(p: Path, kind: str) -> dict[str, list[list[Any]]]:
    st = p.stat()
    key = (str(p), st.st_mtime_ns, st.st_size)
    with _lock:
        if key in _cache:
            return _cache[key]
    reader = {"xls": read_xls, "ods": read_ods, "html": read_html}[kind]
    data = reader(p)
    if not data:
        raise SheetError(f"No tables/sheets found in '{p.name}'.", "Check that the file contains data.")
    with _lock:
        if len(_cache) > 8:
            _cache.clear()
        _cache[key] = data
    return data


# ------------------------------------------------------------------ .xls (BIFF) via optional xlrd
def read_xls(p: Path) -> dict[str, list[list[Any]]]:
    try:
        import xlrd  # type: ignore[import-untyped]
    except ImportError as e:
        raise SheetError(
            f"Reading legacy .xls ('{p.name}') needs the optional 'xlrd' package.",
            "Install the extra: pip install 'sheetmcp[xls]' (or uvx --with xlrd sheetmcp ...), "
            "or re-save the file as .xlsx.",
        ) from e
    try:
        book = xlrd.open_workbook(str(p), on_demand=True)
    except Exception as e:
        raise SheetError(
            f"Could not read '{p.name}' as .xls: {e}", "It may be corrupt or encrypted; re-save as .xlsx."
        ) from e
    out: dict[str, list[list[Any]]] = {}
    try:
        for sh in book.sheets():
            rows = []
            for r in range(sh.nrows):
                row = []
                for c in range(sh.ncols):
                    cell = sh.cell(r, c)
                    t, v = cell.ctype, cell.value
                    if t == xlrd.XL_CELL_EMPTY or t == xlrd.XL_CELL_BLANK:
                        v = None
                    elif t == xlrd.XL_CELL_DATE:
                        with contextlib.suppress(Exception):
                            v = xlrd.xldate_as_datetime(v, book.datemode)
                    elif t == xlrd.XL_CELL_BOOLEAN:
                        v = bool(v)
                    elif t == xlrd.XL_CELL_ERROR:
                        v = xlrd.error_text_from_code.get(v, "#VALUE!")
                    elif t == xlrd.XL_CELL_NUMBER and float(v).is_integer():
                        v = int(v)
                    row.append(v)
                rows.append(row)
            out[sh.name] = _trim(rows)
    finally:
        book.release_resources()
    return out


# ------------------------------------------------------------------ .ods (OpenDocument) with the stdlib
_NS = {
    "table": "urn:oasis:names:tc:opendocument:xmlns:table:1.0",
    "office": "urn:oasis:names:tc:opendocument:xmlns:office:1.0",
    "text": "urn:oasis:names:tc:opendocument:xmlns:text:1.0",
}
_T = "{" + _NS["table"] + "}"
_O = "{" + _NS["office"] + "}"
_X = "{" + _NS["text"] + "}"


def _ods_value(cell: ET.Element) -> Any:
    vt = cell.get(_O + "value-type")
    if vt in ("float", "percentage", "currency"):
        v = float(cell.get(_O + "value", "0"))
        return int(v) if v.is_integer() else v
    if vt == "date":
        s = cell.get(_O + "date-value", "")
        try:
            d = dt.datetime.fromisoformat(s)
            return d
        except ValueError:
            return s
    if vt == "time":
        return cell.get(_O + "time-value")
    if vt == "boolean":
        return cell.get(_O + "boolean-value") == "true"
    parts = []
    for p in cell.iter(_X + "p"):
        parts.append("".join(p.itertext()))
    txt = "\n".join(parts)
    return txt if txt != "" else None


def read_ods(p: Path) -> dict[str, list[list[Any]]]:
    try:
        with zipfile.ZipFile(p) as z:
            xml = z.read("content.xml")
    except (zipfile.BadZipFile, KeyError) as e:
        raise SheetError(f"'{p.name}' is not a valid .ods file.", "Re-save it from LibreOffice/Excel.") from e
    root = ET.fromstring(xml)
    out: dict[str, list[list[Any]]] = {}
    for tbl in root.iter(_T + "table"):
        name = tbl.get(_T + "name") or f"Sheet{len(out) + 1}"
        rows: list[list[Any]] = []
        pending_empty_rows = 0
        for row in tbl.iter(_T + "table-row"):
            rrep = int(row.get(_T + "number-rows-repeated", "1"))
            vals: list[Any] = []
            pending_empty = 0
            for cell in row:
                if cell.tag not in (_T + "table-cell", _T + "covered-table-cell"):
                    continue
                crep = int(cell.get(_T + "number-columns-repeated", "1"))
                v = _ods_value(cell)
                if v is None:
                    pending_empty += crep
                    continue
                vals.extend([None] * pending_empty)
                pending_empty = 0
                vals.extend([v] * min(crep, 1000))
            if not vals:
                pending_empty_rows += rrep
                continue
            rows.extend([[] for _ in range(min(pending_empty_rows, 100_000))])
            pending_empty_rows = 0
            rows.extend([list(vals) for _ in range(min(rrep, 100_000))])
        out[name] = _trim(rows)
    return out


# ------------------------------------------------------------------ HTML <table>s with the stdlib
class _TableParser(HTMLParser):
    def __init__(self) -> None:
        super().__init__(convert_charrefs=True)
        self.tables: list[tuple[str, list[list[Any]]]] = []
        self._stack: list[dict[str, Any]] = []
        self._cell: list[str] | None = None
        self._span = 1
        self._caption: list[str] | None = None

    def handle_starttag(self, tag: str, attrs: list[tuple[str, str | None]]) -> None:
        a = dict(attrs)
        if tag == "table":
            self._stack.append({"name": a.get("id") or a.get("name") or "", "rows": []})
        elif not self._stack:
            return
        elif tag == "caption":
            self._caption = []
        elif tag == "tr":
            self._stack[-1]["rows"].append([])
        elif tag in ("td", "th"):
            self._cell = []
            try:
                self._span = max(1, min(int(a.get("colspan") or 1), 100))
            except ValueError:
                self._span = 1
        elif tag == "br" and self._cell is not None:
            self._cell.append("\n")

    def handle_endtag(self, tag: str) -> None:
        if not self._stack:
            return
        if tag in ("td", "th") and self._cell is not None:
            t = self._stack[-1]
            if not t["rows"]:
                t["rows"].append([])
            text = " ".join("".join(self._cell).split())
            t["rows"][-1].extend([text] + [""] * (self._span - 1))
            self._cell = None
        elif tag == "caption" and self._caption is not None:
            if not self._stack[-1]["name"]:
                self._stack[-1]["name"] = " ".join("".join(self._caption).split())[:31]
            self._caption = None
        elif tag == "table":
            t = self._stack.pop()
            rows = [r for r in t["rows"] if r]
            if rows:
                self.tables.append((t["name"], rows))

    def handle_data(self, data: str) -> None:
        if self._caption is not None:
            self._caption.append(data)
        if self._cell is not None:
            self._cell.append(data)


def read_html(p: Path) -> dict[str, list[list[Any]]]:
    raw = p.read_bytes()
    text = raw.decode("utf-8", errors="replace")
    parser = _TableParser()
    parser.feed(text)
    out: dict[str, list[list[Any]]] = {}
    for i, (name, rows) in enumerate(parser.tables, start=1):
        key = name or f"table{i}"
        while key in out:
            key += f"_{i}"
        out[key] = [[_num(v) for v in r] for r in rows]
    return out


def _num(v: str) -> Any:
    s = v.strip()
    if s == "":
        return None
    try:
        t = s.replace(",", "")
        f = float(t)
        return int(f) if f.is_integer() and "." not in t and "e" not in t.lower() else f
    except ValueError:
        return s


def _trim(rows: list[list[Any]]) -> list[list[Any]]:
    while rows and not any(v is not None and v != "" for v in rows[-1]):
        rows.pop()
    return rows
