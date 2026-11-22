"""Workbook operations: describe, read, write, sheets, rows/columns."""

from __future__ import annotations

import datetime as dt
import re
from collections import Counter
from pathlib import Path
from typing import Any

import openpyxl
from openpyxl.utils.cell import get_column_letter, range_boundaries

from .base import BIG_FILE_BYTES, Base
from .errors import SheetError
from .evaluator import Evaluator
from .fileio import (
    coerce_text,
    file_kind,
    get_sheet,
    has_charts,
    load_workbook,
    read_csv_grid,
    save_workbook,
    write_csv_grid,
)
from .refshift import shift_address, shift_formula
from .util import cell_ref, col_index, is_error_value, parse_range, to_markdown

_range = range
_ISO_DATE = re.compile(r"^\d{4}-\d{2}-\d{2}([ T]\d{2}:\d{2}(:\d{2}(\.\d+)?)?)?$")


def _is_formula(v: Any) -> bool:
    return (isinstance(v, str) and v.startswith("=") and len(v) > 1) or getattr(v, "text", None) is not None


def _formula_text(v: Any) -> str:
    t = getattr(v, "text", None)
    return t if t is not None else v


def infer_type(values: list[Any]) -> dict[str, Any]:
    """Infer a column type from sample values."""
    kinds: Counter[str] = Counter()
    nulls = 0
    example = None
    for v in values:
        if v is None or (isinstance(v, str) and v.strip() == ""):
            nulls += 1
            continue
        if example is None:
            example = v
        if isinstance(v, bool):
            kinds["boolean"] += 1
        elif isinstance(v, int):
            kinds["integer"] += 1
        elif isinstance(v, float):
            kinds["number"] += 1
        elif isinstance(v, dt.datetime):
            kinds["date" if v.time() == dt.time(0) else "datetime"] += 1
        elif isinstance(v, dt.date):
            kinds["date"] += 1
        elif isinstance(v, dt.time):
            kinds["time"] += 1
        elif isinstance(v, str) and is_error_value(v):
            kinds["error"] += 1
        elif isinstance(v, str):
            c = coerce_text(v)
            if isinstance(c, bool):
                kinds["boolean"] += 1
            elif isinstance(c, int):
                kinds["integer"] += 1
            elif isinstance(c, float):
                kinds["number"] += 1
            elif _ISO_DATE.match(v.strip()):
                kinds["date"] += 1
            else:
                kinds["text"] += 1
        else:
            kinds["text"] += 1
    if not kinds:
        t = "empty"
    elif len(kinds) == 1:
        t = next(iter(kinds))
    elif set(kinds) <= {"integer", "number"}:
        t = "number"
    elif set(kinds) <= {"date", "datetime"}:
        t = "datetime"
    else:
        t = "mixed(" + ", ".join(f"{k}:{n}" for k, n in kinds.most_common()) + ")"
    return {"type": t, "nulls": nulls, "example": example}


class WorkbookOps(Base):
    # ------------------------------------------------------------------ describe
    def describe(
        self, path: str, sheet: str | None = None, preview_rows: int = 5, header_row: int | None = None
    ) -> dict:
        """Structure of a spreadsheet file: sheets, sizes, tables, names, inferred column schema and preview."""
        p = self.sb.resolve(path)
        kind = file_kind(p)
        st = p.stat()
        info: dict[str, Any] = {
            "file": self.sb.display(p),
            "kind": kind,
            "size_bytes": st.st_size,
            "modified": dt.datetime.fromtimestamp(st.st_mtime).isoformat(timespec="seconds"),
        }
        preview_rows = max(0, min(int(preview_rows), 50))
        if kind != "excel":
            from .sqlops import SqlOps

            assert isinstance(self, SqlOps)
            info.update(self.table_schema(str(p), preview_rows=preview_rows))
            return info

        big = st.st_size > BIG_FILE_BYTES
        info["streaming_mode"] = big
        wbf = load_workbook(p, formulas=True, read_only=big, fresh=big)
        wbv = load_workbook(p, formulas=False, read_only=True, fresh=True)
        try:
            names = []
            try:
                for name, dn in wbf.defined_names.items():
                    names.append({"name": name, "refers_to": dn.attr_text})
            except Exception:
                pass
            info["named_ranges"] = names
            info["has_charts"] = has_charts(p)
            sheets = []
            targets = [get_sheet(wbf, sheet).title] if sheet not in (None, "") else wbf.sheetnames[:10]
            for i, ws in enumerate(wbf.worksheets):
                s: dict[str, Any] = {
                    "index": i,
                    "name": ws.title,
                    "state": getattr(ws, "sheet_state", "visible"),
                    "dimensions": ws.calculate_dimension() if hasattr(ws, "calculate_dimension") else None,
                    "max_row": ws.max_row,
                    "max_column": ws.max_column,
                }
                if not big:
                    s["tables"] = [{"name": t.displayName, "ref": t.ref} for t in ws.tables.values()]
                    s["merged_cells"] = len(ws.merged_cells.ranges)
                    s["freeze_panes"] = ws.freeze_panes
                    s["formula_cells"] = sum(1 for row in ws.iter_rows() for c in row if _is_formula(c.value))
                    s["charts"] = len(getattr(ws, "_charts", []))
                if ws.title in targets:
                    s.update(self._sheet_schema(wbv[ws.title], preview_rows, header_row, None if big else ws))
                sheets.append(s)
            info["sheets"] = sheets
            if len(wbf.sheetnames) > 10 and sheet in (None, ""):
                info["note"] = "Schema shown for the first 10 sheets; pass sheet= for others."
        finally:
            wbv.close()
            if big:
                wbf.close()
        return info

    def _sheet_schema(self, ws: Any, preview_rows: int, header_row: int | None, wsf: Any = None) -> dict:
        rows_iter = ws.iter_rows(values_only=True)
        header: list[Any] | None = None
        header_at = 0
        for n, row in enumerate(rows_iter, start=1):
            if header_row and n < header_row:
                continue
            if any(v is not None and str(v).strip() != "" for v in row):
                header, header_at = list(row), n
                break
        if header is None:
            return {"columns": [], "preview": [], "empty": True}
        while header and header[-1] is None:
            header.pop()
        sample: list[list[Any]] = []
        for row in rows_iter:
            sample.append(list(row[: len(header)]))
            if len(sample) >= 1000:
                break
        cols = []
        for i, h in enumerate(header):
            t = infer_type([r[i] if i < len(r) else None for r in sample])
            col = {
                "column": get_column_letter(i + 1),
                "name": self._j(h) if h is not None else f"column_{i + 1}",
                "type": t["type"],
                "nulls_in_sample": t["nulls"],
                "example": self._j(t["example"]),
            }
            if wsf is not None and sample:
                fv = wsf.cell(row=header_at + 1, column=i + 1).value
                if _is_formula(fv):
                    col["formula_example"] = _formula_text(fv)
                    if t["type"] == "empty":
                        col["type"] = "formula (no cached values; use read_range to compute)"
            cols.append(col)
        return {
            "header_row": header_at,
            "columns": cols,
            "sample_rows_scanned": len(sample),
            "preview": [[self._j(v) for v in r] for r in sample[:preview_rows]],
        }

    # ------------------------------------------------------------------ read
    def read_range(
        self,
        path: str,
        sheet: str | None = None,
        range: str | None = None,
        include_formulas: bool = True,
        evaluate: str = "auto",
        offset: int = 0,
        limit: int | None = None,
        format: str = "json",
    ) -> dict:
        """Read cells. Values and formulas side by side; missing cached values are computed (evaluate='auto')."""
        p = self.sb.resolve(path)
        kind = file_kind(p)
        limit = self._limit(limit)
        offset = max(0, int(offset or 0))
        if kind in ("csv", "tsv"):
            return self._read_csv_range(p, range, offset, limit, format)
        if kind != "excel":
            raise SheetError(
                f"read_range works on .xlsx/.csv files; '{p.name}' is {kind}.",
                "Use query (SQL) or transform to read Parquet/JSON files.",
            )
        big = p.stat().st_size > BIG_FILE_BYTES
        wbv = load_workbook(p, formulas=False, read_only=big, fresh=big)
        wbf = (
            load_workbook(p, formulas=True, read_only=big, fresh=big)
            if include_formulas or evaluate != "never"
            else None
        )
        try:
            wsv = get_sheet(wbv, sheet)
            wsf = wbf[wsv.title] if wbf is not None else None
            max_row, max_col = max(wsv.max_row or 1, 1), max(wsv.max_column or 1, 1)
            if range:
                c1, r1, c2, r2 = parse_range(range, max_row, max_col)
            else:
                c1, r1, c2, r2 = 1, 1, max_col, max_row
            c2 = min(c2, c1 + 199)  # at most 200 columns per call
            start = r1 + offset
            end = min(r2, start + limit - 1)
            values: list[list[Any]] = []
            formulas: dict[str, str] = {}
            computed: list[str] = []
            evaluator = None
            if start <= end:
                vit = wsv.iter_rows(min_row=start, max_row=end, min_col=c1, max_col=c2, values_only=True)
                fit = (
                    wsf.iter_rows(min_row=start, max_row=end, min_col=c1, max_col=c2, values_only=True)
                    if wsf is not None
                    else None
                )
                for ri, vrow in enumerate(vit):
                    frow = next(fit) if fit is not None else None
                    out = []
                    for ci, v in enumerate(vrow):
                        f = frow[ci] if frow is not None and ci < len(frow) else None
                        ref = cell_ref(c1 + ci, start + ri)
                        if _is_formula(f):
                            ftxt = _formula_text(f)
                            if include_formulas:
                                formulas[ref] = ftxt
                            if (v is None and evaluate in ("auto", "always")) or evaluate == "always":
                                if evaluator is None:
                                    evaluator = Evaluator(wbf if not big else load_workbook(p, formulas=True))
                                v = evaluator.evaluate_cell(wsv.title, ref)
                                computed.append(ref)
                        out.append(self._j(v))
                    out += [None] * ((c2 - c1 + 1) - len(out))
                    values.append(out)
            res: dict[str, Any] = {
                "file": self.sb.display(p),
                "sheet": wsv.title,
                "range": f"{cell_ref(c1, start)}:{cell_ref(c2, max(start, end))}" if values else None,
                "first_row": start,
                "columns": [get_column_letter(c) for c in _range(c1, c2 + 1)],
                "values": values,
            }
            if include_formulas:
                res["formulas"] = formulas
            if computed:
                res["computed_by_evaluator"] = computed
                res["note"] = (
                    "Some formula cells had no cached value (file was saved by a tool, not Excel); "
                    "their values were computed by SheetMCP's evaluator."
                )
            if end < r2:
                res["truncated"] = True
                res["next_offset"] = offset + len(values)
                res["rows_remaining"] = r2 - end
            if format == "markdown":
                res["markdown"] = to_markdown(
                    ["#", *res["columns"]], [[start + i, *r] for i, r in enumerate(values)]
                )
                del res["values"]
            return res
        finally:
            if big:
                wbv.close()
                if wbf is not None:
                    wbf.close()

    def _read_csv_range(self, p: Path, range: str | None, offset: int, limit: int, format: str) -> dict:
        grid = read_csv_grid(p)
        nrows = len(grid)
        ncols = max((len(r) for r in grid), default=0)
        if range:
            c1, r1, c2, r2 = parse_range(range, nrows, ncols)
        else:
            c1, r1, c2, r2 = 1, 1, max(ncols, 1), max(nrows, 1)
        c2 = min(c2, c1 + 199)
        start = r1 + offset
        end = min(r2, nrows, start + limit - 1)
        values = []
        for r in _range(start, end + 1):
            row = grid[r - 1] if r - 1 < nrows else []
            values.append(
                [self._j(coerce_text(row[c - 1])) if c - 1 < len(row) else None for c in _range(c1, c2 + 1)]
            )
        cols = [get_column_letter(c) for c in _range(c1, c2 + 1)]
        res: dict[str, Any] = {
            "file": self.sb.display(p),
            "first_row": start,
            "columns": cols,
            "values": values,
            "total_rows": nrows,
        }
        if end < min(r2, nrows):
            res["truncated"] = True
            res["next_offset"] = offset + len(values)
        if format == "markdown":
            res["markdown"] = to_markdown(["#", *cols], [[start + i, *r] for i, r in enumerate(values)])
            del res["values"]
        return res

    # ------------------------------------------------------------------ write
    @staticmethod
    def _convert_input(v: Any, parse_dates: bool) -> Any:
        if isinstance(v, str) and parse_dates and _ISO_DATE.match(v.strip()):
            try:
                d = dt.datetime.fromisoformat(v.strip().replace(" ", "T"))
                return d
            except ValueError:
                return v
        return v

    def write_range(
        self,
        path: str,
        sheet: str | None = None,
        start_cell: str = "A1",
        values: list[list[Any]] | None = None,
        cells: dict[str, Any] | None = None,
        append: bool = False,
        parse_dates: bool = True,
        create: bool = True,
        dry_run: bool = False,
    ) -> dict:
        """Write a 2D block (values), individual cells ({"B2": 5}) or append rows. '=...' strings are formulas."""
        if not values and not cells:
            raise SheetError(
                "Nothing to write.", "Pass values=[[...], ...] (a 2D list) and/or cells={'A1': 1}."
            )
        if values is not None and (
            not isinstance(values, list) or any(not isinstance(r, list) for r in values)
        ):
            if isinstance(values, list) and all(not isinstance(r, list) for r in values):
                values = [values]  # a single row
            else:
                raise SheetError(
                    "values must be a 2D list (list of rows).", 'Example: [["Name", "Age"], ["Ana", 31]]'
                )
        if not dry_run:
            self._writable("write cells")
        p = self.sb.resolve(path, must_exist=False)
        exists = p.exists()
        if not exists and not create:
            raise SheetError(f"'{path}' does not exist.", "Pass create=true to create it.")
        kind = file_kind(p)
        if kind in ("csv", "tsv"):
            return self._write_csv(p, start_cell, values, cells, append, dry_run)
        if kind != "excel":
            raise SheetError(
                f"Cannot write cells into a {kind} file.", "Convert it to .xlsx or .csv first (convert tool)."
            )

        warnings: list[str] = []
        if exists:
            wb = load_workbook(p, formulas=True, fresh=True)
        else:
            wb = openpyxl.Workbook()
            if sheet:
                wb.active.title = sheet[:31]
        if sheet and sheet not in wb.sheetnames and sheet.lower() not in {n.lower() for n in wb.sheetnames}:
            if not create:
                raise SheetError(f"Sheet '{sheet}' does not exist.", "Available: " + ", ".join(wb.sheetnames))
            wb.create_sheet(sheet[:31])
            warnings.append(f"Created new sheet '{sheet[:31]}'.")
        ws = get_sheet(wb, sheet)
        changes: list[dict[str, Any]] = []

        def put(ref: str, v: Any) -> None:
            c = ws[ref]
            new = self._convert_input(v, parse_dates)
            changes.append({"cell": ref, "old": self._j(c.value), "new": self._j(new)})
            if not dry_run:
                old_fmt = c.number_format
                c.value = new
                if isinstance(new, dt.datetime) and old_fmt == "General":
                    c.number_format = "yyyy-mm-dd" if new.time() == dt.time(0) else "yyyy-mm-dd hh:mm:ss"

        if values:
            if append:
                last = ws.max_row
                if last == 1 and all(c.value is None for c in ws[1]):
                    last = 0
                r0, c0 = last + 1, col_index(re.sub(r"\d", "", start_cell or "A") or "A")
            else:
                c0, r0, _, _ = parse_range(start_cell or "A1")
            for i, row in enumerate(values):
                for j, v in enumerate(row):
                    put(cell_ref(c0 + j, r0 + i), v)
        for ref, v in (cells or {}).items():
            c, r, _, _ = parse_range(ref)
            put(cell_ref(c, r), v)

        res: dict[str, Any] = {
            "file": self.sb.display(p),
            "sheet": ws.title,
            "cells_written": len(changes),
            "dry_run": dry_run,
            "changes": changes[:50],
        }
        if len(changes) > 50:
            res["changes_truncated"] = len(changes) - 50
        if dry_run:
            res["note"] = "Dry run: nothing was saved. Call again with dry_run=false to apply."
            return res
        overwritten_formulas = [
            ch["cell"]
            for ch in changes
            if isinstance(ch["old"], str) and ch["old"].startswith("=") and ch["old"] != ch["new"]
        ]
        if overwritten_formulas:
            warnings.append(
                f"Replaced {len(overwritten_formulas)} existing formula(s): {', '.join(overwritten_formulas[:10])}"
            )
        res["backup"] = save_workbook(wb, p, self.cfg)
        if warnings:
            res["warnings"] = warnings
        return res

    def _write_csv(
        self, p: Path, start_cell: str, values: Any, cells: Any, append: bool, dry_run: bool
    ) -> dict:
        grid = read_csv_grid(p) if p.exists() else []
        changes = []

        def put(r: int, c: int, v: Any) -> None:
            while len(grid) < r:
                grid.append([])
            row = grid[r - 1]
            while len(row) < c:
                row.append("")
            changes.append(
                {"cell": cell_ref(c, r), "old": self._j(coerce_text(row[c - 1])), "new": self._j(v)}
            )
            row[c - 1] = v

        if values:
            if append:
                r0, c0 = len(grid) + 1, 1
            else:
                c0, r0, _, _ = parse_range(start_cell or "A1")
            for i, row in enumerate(values):
                for j, v in enumerate(row):
                    put(r0 + i, c0 + j, v)
        for ref, v in (cells or {}).items():
            c, r, _, _ = parse_range(ref)
            put(r, c, v)
        res: dict[str, Any] = {
            "file": self.sb.display(p),
            "cells_written": len(changes),
            "dry_run": dry_run,
            "changes": changes[:50],
        }
        if any(isinstance(ch["new"], str) and ch["new"].startswith("=") for ch in changes):
            res["warnings"] = [
                "CSV files cannot hold formulas; '=...' was stored as plain text. Use an .xlsx file for formulas."
            ]
        if not dry_run:
            res["backup"] = write_csv_grid(p, grid, self.cfg)
        return res

    # ------------------------------------------------------------------ sheets
    def manage_sheets(
        self,
        path: str,
        action: str,
        sheet: str | None = None,
        new_name: str | None = None,
        position: int | None = None,
        sheets: list[str] | None = None,
        dry_run: bool = False,
    ) -> dict:
        """create_workbook | list | add | rename | delete | copy | move | hide | unhide."""
        action = (action or "").lower().strip()
        valid = {"create_workbook", "list", "add", "rename", "delete", "copy", "move", "hide", "unhide"}
        if action not in valid:
            raise SheetError(f"Unknown action '{action}'.", "Use one of: " + ", ".join(sorted(valid)))
        if action == "list":
            p = self.sb.resolve(path)
            wb = load_workbook(p, formulas=True, read_only=True, fresh=True)
            try:
                return {"file": self.sb.display(p), "sheets": wb.sheetnames}
            finally:
                wb.close()
        if not dry_run:
            self._writable(f"{action} sheets")
        if action == "create_workbook":
            p = self.sb.resolve(path, must_exist=False)
            if p.exists():
                raise SheetError(
                    f"'{path}' already exists.", "Pick a new file name, or use action='add' to add sheets."
                )
            if file_kind(p) != "excel":
                raise SheetError("create_workbook makes .xlsx files.", "Use a path ending in .xlsx")
            wb = openpyxl.Workbook()
            names = sheets or ([sheet] if sheet else ["Sheet1"])
            wb.active.title = names[0][:31]
            for n in names[1:]:
                wb.create_sheet(n[:31])
            if not dry_run:
                save_workbook(wb, p, self.cfg)
            return {"file": self.sb.display(p), "created": True, "sheets": wb.sheetnames, "dry_run": dry_run}

        p = self.sb.resolve(path)
        wb = load_workbook(p, formulas=True, fresh=True)
        res: dict[str, Any] = {"file": self.sb.display(p), "action": action, "dry_run": dry_run}
        if action == "add":
            name = (new_name or sheet or f"Sheet{len(wb.sheetnames) + 1}")[:31]
            if name in wb.sheetnames:
                raise SheetError(f"Sheet '{name}' already exists.", "Choose another name.")
            wb.create_sheet(name, position)
            res["sheet"] = name
        else:
            ws = get_sheet(wb, sheet)
            if action == "rename":
                if not new_name:
                    raise SheetError("rename needs new_name.", "Pass new_name='Summary'.")
                if new_name in wb.sheetnames:
                    raise SheetError(f"Sheet '{new_name}' already exists.", "Choose another name.")
                old = ws.title
                ws.title = new_name[:31]
                fixed = self._rename_refs(wb, old, ws.title)
                res.update({"old_name": old, "new_name": ws.title, "formulas_updated": fixed})
            elif action == "delete":
                if len(wb.sheetnames) == 1:
                    raise SheetError(
                        "Cannot delete the only sheet in a workbook.", "Add another sheet first."
                    )
                refs = self._count_refs_to(wb, ws.title)
                wb.remove(ws)
                res["deleted"] = ws.title
                if refs:
                    res["warnings"] = [
                        f"{refs} formula(s) in other sheets referenced '{ws.title}' and will now show #REF! in Excel."
                    ]
            elif action == "copy":
                cp = wb.copy_worksheet(ws)
                if new_name:
                    cp.title = new_name[:31]
                res["copy"] = cp.title
                res["note"] = (
                    "Cell values, styles, merged cells and dimensions are copied; charts/images are not."
                )
            elif action == "move":
                if position is None:
                    raise SheetError("move needs position (0 = first).", "Pass position=0")
                cur = wb.sheetnames.index(ws.title)
                wb.move_sheet(ws, offset=int(position) - cur)
                res["position"] = wb.sheetnames.index(ws.title)
            elif action in ("hide", "unhide"):
                if action == "hide" and sum(1 for w in wb.worksheets if w.sheet_state == "visible") <= 1:
                    raise SheetError(
                        "Cannot hide the last visible sheet.", "Unhide or add another sheet first."
                    )
                ws.sheet_state = "hidden" if action == "hide" else "visible"
                res["sheet"] = ws.title
        res["sheets"] = wb.sheetnames
        if not dry_run:
            res["backup"] = save_workbook(wb, p, self.cfg)
        return res

    @staticmethod
    def _sheet_token(name: str) -> str:
        return (
            f"'{name.replace(chr(39), chr(39) * 2)}'"
            if re.search(r"[^A-Za-z0-9_]", name) or name[:1].isdigit()
            else name
        )

    def _rename_refs(self, wb: Any, old: str, new: str) -> int:
        pat = re.compile(
            r"(?<![A-Za-z0-9_.])("
            + re.escape(self._sheet_token(old))
            + r"|'"
            + re.escape(old.replace("'", "''"))
            + r"')!"
        )
        n = 0
        for ws in wb.worksheets:
            for row in ws.iter_rows():
                for c in row:
                    if isinstance(c.value, str) and c.value.startswith("=") and "!" in c.value:
                        nv = pat.sub(self._sheet_token(new) + "!", c.value)
                        if nv != c.value:
                            c.value = nv
                            n += 1
        try:
            for _, dn in wb.defined_names.items():
                if dn.attr_text:
                    dn.attr_text = pat.sub(self._sheet_token(new) + "!", dn.attr_text)
        except Exception:
            pass
        return n

    @staticmethod
    def _count_refs_to(wb: Any, title: str) -> int:
        needles = (f"{title}!", f"'{title}'!")
        n = 0
        for ws in wb.worksheets:
            if ws.title == title:
                continue
            for row in ws.iter_rows():
                for c in row:
                    if (
                        isinstance(c.value, str)
                        and c.value.startswith("=")
                        and any(x in c.value for x in needles)
                    ):
                        n += 1
        return n

    # ------------------------------------------------------------------ rows / cols
    def insert_delete(
        self,
        path: str,
        action: str,
        index: int | str,
        count: int = 1,
        sheet: str | None = None,
        dry_run: bool = False,
    ) -> dict:
        """insert_rows | delete_rows | insert_cols | delete_cols, with formula references updated across the workbook."""
        action = (action or "").lower().strip()
        if action not in {"insert_rows", "delete_rows", "insert_cols", "delete_cols"}:
            raise SheetError(
                f"Unknown action '{action}'.", "Use insert_rows, delete_rows, insert_cols or delete_cols."
            )
        count = int(count)
        if count < 1:
            raise SheetError("count must be >= 1.")
        axis = "row" if action.endswith("rows") else "col"
        if axis == "col":
            idx = col_index(index) if isinstance(index, str) and not index.isdigit() else int(index)
        else:
            idx = int(index)
        if idx < 1:
            raise SheetError("index starts at 1 (row 1 / column A).")
        if not dry_run:
            self._writable(action)
        p = self.sb.resolve(path)
        if file_kind(p) in ("csv", "tsv"):
            grid = read_csv_grid(p)
            if axis == "row":
                if action == "insert_rows":
                    grid[idx - 1 : idx - 1] = [[] for _ in range(count)]
                else:
                    del grid[idx - 1 : idx - 1 + count]
            else:
                for row in grid:
                    if action == "insert_cols":
                        while len(row) < idx - 1:
                            row.append("")
                        row[idx - 1 : idx - 1] = [""] * count
                    else:
                        del row[idx - 1 : idx - 1 + count]
            res = {
                "file": self.sb.display(p),
                "action": action,
                "index": idx,
                "count": count,
                "dry_run": dry_run,
            }
            if not dry_run:
                res["backup"] = write_csv_grid(p, grid, self.cfg)
            return res
        wb = load_workbook(p, formulas=True, fresh=True)
        ws = get_sheet(wb, sheet)
        delta = count if action.startswith("insert") else -count
        # move merged ranges ourselves (openpyxl leaves them in place)
        merged = [str(m) for m in ws.merged_cells.ranges]
        for m in merged:
            ws.unmerge_cells(m)
        {
            "insert_rows": ws.insert_rows,
            "delete_rows": ws.delete_rows,
            "insert_cols": ws.insert_cols,
            "delete_cols": ws.delete_cols,
        }[action](idx, count)
        for m in merged:
            nm = shift_address(m, axis, idx, delta)
            if nm and ":" in nm:
                ws.merge_cells(nm)
        updated = 0
        for w in wb.worksheets:
            for row in w.iter_rows():
                for c in row:
                    if isinstance(c.value, str) and c.value.startswith("="):
                        nv = shift_formula(c.value, w.title, ws.title, axis, idx, delta)
                        if nv != c.value:
                            c.value = nv
                            updated += 1
        try:
            for _, dn in wb.defined_names.items():
                if dn.attr_text:
                    dn.attr_text = shift_formula("=" + dn.attr_text, ws.title, ws.title, axis, idx, delta)[1:]
        except Exception:
            pass
        for t in ws.tables.values():
            nr = shift_address(t.ref, axis, idx, delta)
            if nr:
                t.ref = nr
        res = {
            "file": self.sb.display(p),
            "sheet": ws.title,
            "action": action,
            "index": idx,
            "count": count,
            "formulas_updated": updated,
            "dry_run": dry_run,
            "new_dimensions": ws.calculate_dimension(),
        }
        if not dry_run:
            res["backup"] = save_workbook(wb, p, self.cfg)
        return res


_ = range_boundaries  # re-exported for callers
