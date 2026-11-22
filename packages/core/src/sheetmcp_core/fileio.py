"""File IO: type detection, cached workbook loading, safe atomic saves with backups, CSV grids."""

from __future__ import annotations

import contextlib
import csv
import hashlib
import io as _io
import os
import re
import shutil
import tempfile
import threading
import time
import zipfile
from collections import OrderedDict
from pathlib import Path
from typing import Any

import openpyxl
from openpyxl.workbook.workbook import Workbook

from .config import Config
from .errors import SheetError
from .util import jsonable, unique_headers

EXCEL_EXT = {".xlsx", ".xlsm", ".xltx", ".xltm"}
CSV_EXT = {".csv", ".tsv", ".tab", ".txt"}
PARQUET_EXT = {".parquet", ".pq"}
JSON_EXT = {".json", ".jsonl", ".ndjson"}
TABULAR_EXT = EXCEL_EXT | CSV_EXT | PARQUET_EXT | JSON_EXT
OUTPUT_FORMATS = {"csv", "tsv", "xlsx", "parquet", "json", "jsonl", "md"}


def file_kind(p: Path) -> str:
    ext = p.suffix.lower()
    if ext in EXCEL_EXT:
        return "excel"
    if ext in {".tsv", ".tab"}:
        return "tsv"
    if ext in CSV_EXT:
        return "csv"
    if ext in PARQUET_EXT:
        return "parquet"
    if ext in {".jsonl", ".ndjson"}:
        return "jsonl"
    if ext == ".json":
        return "json"
    if ext == ".xls":
        raise SheetError(
            f"'{p.name}' is a legacy .xls file, which is not supported.",
            "Re-save it as .xlsx (Excel/LibreOffice: Save As → .xlsx), or convert with "
            "`libreoffice --headless --convert-to xlsx file.xls`.",
        )
    if ext == ".ods":
        raise SheetError(
            f"'{p.name}' is an OpenDocument sheet (.ods), which is not supported.", "Save it as .xlsx first."
        )
    raise SheetError(
        f"Unsupported file type '{ext or p.name}'.",
        "Supported: .xlsx .xlsm .csv .tsv .txt .parquet .json .jsonl",
    )


def output_format(p: Path, fmt: str | None = None) -> str:
    if fmt:
        fmt = fmt.lower().lstrip(".")
        fmt = {"markdown": "md", "excel": "xlsx", "ndjson": "jsonl", "pq": "parquet"}.get(fmt, fmt)
    else:
        ext = p.suffix.lower().lstrip(".")
        fmt = {"markdown": "md", "xlsm": "xlsx", "ndjson": "jsonl", "pq": "parquet", "txt": "csv"}.get(
            ext, ext
        )
    if fmt not in OUTPUT_FORMATS:
        raise SheetError(f"Cannot write format '{fmt}'.", "Use one of: " + ", ".join(sorted(OUTPUT_FORMATS)))
    return fmt


# ---------------------------------------------------------------- workbook cache
_cache: OrderedDict[tuple[Any, ...], Workbook] = OrderedDict()
_lock = threading.Lock()
_CACHE_SIZE = 6


def _stamp(p: Path) -> tuple[float, int]:
    st = p.stat()
    return (st.st_mtime_ns, st.st_size)


def load_workbook(p: Path, formulas: bool = True, read_only: bool = False, fresh: bool = False) -> Workbook:
    """Load an xlsx. formulas=True gives formula strings, False gives cached values.

    Non-fresh loads are cached by (path, mtime, size, mode) and MUST NOT be mutated.
    Pass fresh=True when you intend to modify and save.
    """
    if file_kind(p) != "excel":
        raise SheetError(f"'{p.name}' is not an Excel workbook.", "This operation needs an .xlsx/.xlsm file.")
    key = (str(p), *_stamp(p), formulas, read_only)
    if not fresh:
        with _lock:
            wb = _cache.get(key)
            if wb is not None:
                _cache.move_to_end(key)
                return wb
    try:
        wb = openpyxl.load_workbook(
            p,
            data_only=not formulas,
            read_only=read_only,
            keep_vba=p.suffix.lower() in {".xlsm", ".xltm"},
            keep_links=True,
        )
    except zipfile.BadZipFile as e:
        raise SheetError(
            f"'{p.name}' is not a valid .xlsx file (corrupt, encrypted/password-protected, or another format renamed).",
            "Open it in Excel and re-save as .xlsx without a password.",
        ) from e
    except PermissionError as e:
        raise SheetError(
            f"Permission denied reading '{p.name}'.", "Close the file in Excel or fix file permissions."
        ) from e
    except Exception as e:
        raise SheetError(f"Could not open '{p.name}': {e}", "Make sure it is a real .xlsx file.") from e
    if not fresh and not read_only:
        with _lock:
            _cache[key] = wb
            while len(_cache) > _CACHE_SIZE:
                _cache.popitem(last=False)
    return wb


def get_sheet(wb: Workbook, sheet: str | int | None):
    if sheet is None or sheet == "":
        return wb.worksheets[0] if wb.worksheets else wb.active
    if isinstance(sheet, int) or (isinstance(sheet, str) and sheet.isdigit() and sheet not in wb.sheetnames):
        i = int(sheet)
        if 0 <= i < len(wb.sheetnames):
            return wb[wb.sheetnames[i]]
    if sheet in wb.sheetnames:
        return wb[sheet]
    low = {n.lower(): n for n in wb.sheetnames}
    if str(sheet).lower() in low:
        return wb[low[str(sheet).lower()]]
    raise SheetError(f"Sheet '{sheet}' does not exist.", "Available sheets: " + ", ".join(wb.sheetnames))


def has_charts(p: Path) -> bool:
    try:
        with zipfile.ZipFile(p) as z:
            return any(n.startswith("xl/charts/chart") for n in z.namelist())
    except Exception:
        return False


def xlsx_sheet_names(p: Path) -> list[str]:
    """Fast sheet listing straight from workbook.xml (no full parse)."""
    try:
        with zipfile.ZipFile(p) as z:
            xml = z.read("xl/workbook.xml").decode("utf-8", "replace")
        import html

        return [html.unescape(m) for m in re.findall(r'<(?:\w+:)?sheet\b[^>]*?\bname="([^"]*)"', xml)]
    except Exception:
        return []


_formula_cache: dict[tuple[Any, ...], bool] = {}


def has_formulas(p: Path) -> bool:
    """Cheap check (raw XML scan) whether any worksheet contains a formula."""
    key = (str(p), *_stamp(p))
    if key in _formula_cache:
        return _formula_cache[key]
    found = True
    try:
        with zipfile.ZipFile(p) as z:
            found = False
            for n in z.namelist():
                if n.startswith("xl/worksheets/") and n.endswith(".xml"):
                    with z.open(n) as fh:
                        tail = b""
                        while chunk := fh.read(1 << 20):
                            buf = tail + chunk
                            if b"<f>" in buf or b"<f " in buf or b":f>" in buf or b":f " in buf:
                                found = True
                                break
                            tail = buf[-8:]
                if found:
                    break
    except Exception:
        found = True
    _formula_cache[key] = found
    return found


# ---------------------------------------------------------------- saving
def backup_file(p: Path, cfg: Config) -> str | None:
    if not cfg.backup or not p.exists():
        return None
    bdir = p.parent / ".sheetmcp_backups"
    bdir.mkdir(exist_ok=True)
    stamp = time.strftime("%Y%m%d-%H%M%S") + f"-{time.time_ns() % 1_000_000:06d}"
    dest = bdir / f"{p.stem}.{stamp}{p.suffix}"
    shutil.copy2(p, dest)
    olds = sorted(bdir.glob(f"{p.stem}.*{p.suffix}"), key=lambda x: x.stat().st_mtime)
    for old in olds[: max(0, len(olds) - cfg.max_backups)]:
        with contextlib.suppress(OSError):
            old.unlink()
    return str(dest)


def atomic_write(p: Path, writer, cfg: Config) -> str | None:
    """Back up ``p`` (if it exists) and write it atomically via ``writer(tmp_path)``."""
    if cfg.read_only:
        from .errors import ReadOnlyError

        raise ReadOnlyError(f"write '{p.name}'")
    p.parent.mkdir(parents=True, exist_ok=True)
    backup = backup_file(p, cfg)
    fd, tmp = tempfile.mkstemp(prefix=".sheetmcp-", suffix=p.suffix, dir=p.parent)
    os.close(fd)
    try:
        writer(Path(tmp))
        os.replace(tmp, p)
    except PermissionError as e:
        raise SheetError(
            f"Permission denied writing '{p.name}'.", "If the file is open in Excel, close it and try again."
        ) from e
    finally:
        if os.path.exists(tmp):
            os.unlink(tmp)
    return backup


def save_workbook(wb: Workbook, p: Path, cfg: Config) -> str | None:
    return atomic_write(p, lambda t: wb.save(t), cfg)


# ---------------------------------------------------------------- CSV grids
def sniff_dialect(p: Path) -> tuple[str, str]:
    """Return (encoding, delimiter)."""
    raw = p.read_bytes()[:65536]
    enc = "utf-8-sig"
    try:
        raw.decode("utf-8")
    except UnicodeDecodeError:
        enc = "cp1252"
    if p.suffix.lower() in {".tsv", ".tab"}:
        return enc, "\t"
    sample = raw.decode(enc, errors="replace")
    try:
        delim = csv.Sniffer().sniff(sample, delimiters=",;\t|").delimiter
    except csv.Error:
        delim = ","
    return enc, delim


def read_csv_grid(p: Path) -> list[list[str]]:
    enc, delim = sniff_dialect(p)
    with open(p, newline="", encoding=enc, errors="replace") as f:
        return [row for row in csv.reader(f, delimiter=delim)]


def write_csv_grid(p: Path, grid: list[list[Any]], cfg: Config, delimiter: str | None = None) -> str | None:
    if delimiter is None:
        delimiter = (
            sniff_dialect(p)[1] if p.exists() else ("\t" if p.suffix.lower() in {".tsv", ".tab"} else ",")
        )

    def w(t: Path) -> None:
        with open(t, "w", newline="", encoding="utf-8") as f:
            wr = csv.writer(f, delimiter=delimiter)
            for row in grid:
                wr.writerow(["" if v is None else jsonable(v) for v in row])

    return atomic_write(p, w, cfg)


def coerce_text(v: str) -> Any:
    """CSV text -> number/bool when it clearly is one (for evaluation and typing)."""
    if not isinstance(v, str):
        return v
    s = v.strip()
    if s == "":
        return None
    try:
        if s.lstrip("-+").isdigit():
            return int(s)
        return float(s)
    except ValueError:
        pass
    if s.lower() in {"true", "false"}:
        return s.lower() == "true"
    return v


# ---------------------------------------------------------------- xlsx -> csv for DuckDB
_CONV_DIR = Path(tempfile.gettempdir()) / f"sheetmcp-cache-{os.getuid() if hasattr(os, 'getuid') else 'u'}"


def conv_dir() -> Path:
    _CONV_DIR.mkdir(parents=True, exist_ok=True)
    return _CONV_DIR


def sheet_rows(p: Path, sheet: str | int | None, header_row: int | None = None):
    """Yield (headers, rows-iterator) of a sheet's cached values, skipping leading blank rows."""
    wb = load_workbook(p, formulas=False, read_only=True, fresh=True)
    wbf = load_workbook(p, formulas=True, read_only=True, fresh=True) if has_formulas(p) else None
    try:
        ws = get_sheet(wb, sheet)
        fit = wbf[ws.title].iter_rows(values_only=True) if wbf is not None else iter(())
        evaluator = None

        def fill(row: tuple[Any, ...], frow: tuple[Any, ...] | None, rown: int) -> list[Any]:
            nonlocal evaluator
            out = list(row)
            if frow is None:
                return out
            for i, v in enumerate(out):
                f = frow[i] if i < len(frow) else None
                if v is None and isinstance(f, str) and f.startswith("="):
                    if evaluator is None:
                        from .evaluator import Evaluator

                        evaluator = Evaluator(load_workbook(p, formulas=True))
                    from openpyxl.utils.cell import get_column_letter

                    out[i] = evaluator.evaluate_cell(ws.title, f"{get_column_letter(i + 1)}{rown}")
            return out

        it = ws.iter_rows(values_only=True)
        header: list[Any] | None = None
        rown = 0
        for row in it:
            rown += 1
            frow = next(fit, None)
            if header_row and rown < header_row:
                continue
            if any(v is not None and str(v).strip() != "" for v in row):
                header = fill(row, frow, rown)
                break
        if header is None:
            return [], [], ws.title
        while header and (header[-1] is None or str(header[-1]).strip() == ""):
            header.pop()
        width = len(header)
        rows = []
        for row in it:
            rown += 1
            frow = next(fit, None)
            row = fill(row, frow, rown)
            r = list(row[:width]) + [None] * max(0, width - len(row))
            if any(v is not None for v in r):
                rows.append(r)
        return unique_headers(header), rows, ws.title
    finally:
        wb.close()
        if wbf is not None:
            wbf.close()


def excel_to_csv_cache(p: Path, sheet: str | int | None, header_row: int | None = None) -> tuple[Path, str]:
    """Materialise one sheet as a UTF-8 CSV (cached by mtime) so DuckDB can scan it fast."""
    st = p.stat()
    key = hashlib.sha1(f"{p}|{st.st_mtime_ns}|{st.st_size}|{sheet}|{header_row}".encode()).hexdigest()[:20]
    out = conv_dir() / f"{key}.csv"
    meta = conv_dir() / f"{key}.sheet"
    if out.exists() and meta.exists():
        return out, meta.read_text(encoding="utf-8")
    headers, rows, title = sheet_rows(p, sheet, header_row)
    if not headers:
        raise SheetError(f"Sheet '{title}' in '{p.name}' is empty.", "Pick another sheet (see describe).")
    buf = _io.StringIO()
    wr = csv.writer(buf)
    wr.writerow(headers)
    for r in rows:
        wr.writerow(["" if v is None else jsonable(v) for v in r])
    tmp = out.with_suffix(".tmp")
    tmp.write_text(buf.getvalue(), encoding="utf-8")
    os.replace(tmp, out)
    meta.write_text(title, encoding="utf-8")
    return out, title
