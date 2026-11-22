"""v0.2 power features: quick_look, SQL catalog, clean pipelines, Excel tables & names, backups/undo,
batch operations and one-call report building."""

from __future__ import annotations

import datetime as dt
import re
import shutil
from pathlib import Path
from typing import Any

import duckdb
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.table import Table, TableStyleInfo

from .base import Base
from .errors import SheetError
from .fileio import (
    LEGACY_KINDS,
    SHEETED_KINDS,
    TABULAR_EXT,
    atomic_write,
    backup_stamp,
    file_kind,
    get_sheet,
    legacy_sheet,
    load_workbook,
    save_workbook,
    sheet_names,
)
from .util import get_column_letter, parse_range, qident, qstr, safe_identifier, unique_headers

# DuckDB >= 1.3 has Python-style lambdas; the old arrow form is deprecated (an error in 2.x).
_LAMBDA_X = (
    "lambda x:" if tuple(int(p) for p in re.findall(r"\d+", duckdb.__version__)[:2]) >= (1, 3) else "x ->"
)
_BACKUP_RE = re.compile(r"^(?P<stem>.+)\.(?P<ts>\d{8}-\d{6}-\d{6})(?P<tag>-undo)?(?P<ext>\.[^.]+)$")
_TABLE_NAME = re.compile(r"^[A-Za-z_\\][A-Za-z0-9_.]{0,254}$")


class PowerOps(Base):
    # ================================================================== catalog
    def _header_of(self, p: Path, sheet: str | None) -> list[str]:
        kind = file_kind(p)
        if kind == "excel":
            wb = load_workbook(p, formulas=False, read_only=True, fresh=True)
            try:
                ws = get_sheet(wb, sheet)
                for row in ws.iter_rows(max_row=30, values_only=True):
                    if any(v is not None and str(v).strip() for v in row):
                        hdr = list(row)
                        while hdr and hdr[-1] is None:
                            hdr.pop()
                        return unique_headers(hdr)
                return []
            finally:
                wb.close()
        if kind in LEGACY_KINDS:
            grid = legacy_sheet(p, sheet)[1]
            for row in grid[:30]:
                if any(v is not None and str(v).strip() for v in row):
                    return unique_headers(list(row))
            return []
        from .sqlops import SqlOps

        assert isinstance(self, SqlOps)
        con = self._connect()
        try:
            expr, _ = self.source_sql(str(p))
            return self._columns(con, expr)
        finally:
            con.close()

    def catalog(self, folder: str | None = None, limit: int = 200) -> dict:
        """Every table name you can use in SQL (auto-registered from files in the folder) with its columns."""
        d = self.sb.resolve(folder or ".", want="dir")
        tables = []
        for f in sorted(d.iterdir()):
            if not f.is_file() or f.suffix.lower() not in TABULAR_EXT or f.name.startswith((".", "~$")):
                continue
            stem = safe_identifier(f.stem)
            try:
                kind = file_kind(f)
                if kind in SHEETED_KINDS:
                    names = sheet_names(f)
                    for i, nm in enumerate(names):
                        aliases = [f"{stem}__{safe_identifier(nm)}"] + ([stem] if i == 0 else [])
                        tables.append(
                            {
                                "table": aliases[-1],
                                "aliases": aliases,
                                "source": f"{f.name}#{nm}",
                                "columns": self._header_of(f, nm)[:60],
                            }
                        )
                else:
                    tables.append({"table": stem, "source": f.name, "columns": self._header_of(f, None)[:60]})
            except SheetError as e:
                tables.append({"table": stem, "source": f.name, "error": e.message})
            if len(tables) >= limit:
                break
        return {
            "folder": str(d),
            "tables": tables,
            "how_to": "Use these names directly in query(sql=...). Quote names that are SQL keywords or contain "
            'capitals/spaces with double quotes, e.g. SELECT "Revenue" FROM sales. Files in subfolders: '
            "FROM 'sub/file.csv'.",
        }

    # ================================================================== quick look
    def quick_look(self, path: str = ".", sheet: str | None = None) -> dict:
        """One-call overview of a folder or a file: structure, sizes, columns, stats, issues, next steps."""
        p = self.sb.resolve(path or ".", want="any")
        if p.is_dir():
            return self._look_dir(p)
        return self._look_file(p, sheet)

    def _rows_of(self, p: Path, sheet: str | None) -> int | None:
        kind = file_kind(p)
        try:
            if kind == "excel":
                wb = load_workbook(p, formulas=False, read_only=True, fresh=True)
                try:
                    ws = get_sheet(wb, sheet)
                    return max((ws.max_row or 1) - 1, 0)
                finally:
                    wb.close()
            if kind in LEGACY_KINDS:
                return max(len(legacy_sheet(p, sheet)[1]) - 1, 0)
            from .sqlops import SqlOps

            assert isinstance(self, SqlOps)
            con = self._connect()
            try:
                expr, _ = self.source_sql(str(p))
                return con.execute(f"SELECT count(*) FROM {expr}").fetchone()[0]
            finally:
                con.close()
        except (SheetError, duckdb.Error):
            return None

    def _look_dir(self, d: Path) -> dict:
        from .analysis import AnalysisOps

        assert isinstance(self, AnalysisOps)
        listing = self.list_files(str(d), recursive=True, limit=40)
        files = []
        for item in listing["files"]:
            f = self.sb.resolve(item["path"])
            entry: dict[str, Any] = {
                "path": item["path"],
                "kind": item["kind"],
                "size_kb": round(item["size_bytes"] / 1024, 1),
            }
            in_root = f.parent == self.sb.roots[0]
            stem = safe_identifier(f.stem)
            if item.get("sheets"):
                entry["sheets"] = []
                for i, nm in enumerate(item["sheets"][:10]):
                    try:
                        cols = self._header_of(f, nm)
                    except SheetError:
                        cols = []
                    sh: dict[str, Any] = {"name": nm, "rows": self._rows_of(f, nm), "columns": cols[:15]}
                    sh["sql"] = (
                        (stem if i == 0 else f"{stem}__{safe_identifier(nm)}")
                        if in_root
                        else f"'{item['path']}#{nm}'"
                    )
                    entry["sheets"].append(sh)
            else:
                try:
                    entry["columns"] = self._header_of(f, None)[:15]
                except SheetError as e:
                    entry["error"] = e.message
                entry["rows"] = self._rows_of(f, None)
                entry["sql"] = stem if in_root else f"'{item['path']}'"
            files.append(entry)
        res: dict[str, Any] = {"folder": str(d), "file_count": listing["count"], "files": files}
        if listing.get("truncated"):
            res["note"] = "Showing the first 40 files; narrow with list_files(pattern=...)."
        res["next_steps"] = [
            "quick_look(path='<file>') for one file's columns, stats and issues",
            "query(sql='SELECT ...') using the 'sql' table names above (joins across files work)",
            "query() with no sql lists every table name and its columns",
        ]
        return res

    def _look_file(self, p: Path, sheet: str | None) -> dict:
        from .formulaops import FormulaOps
        from .sqlops import SqlOps
        from .workbook import WorkbookOps

        assert isinstance(self, SqlOps) and isinstance(self, WorkbookOps) and isinstance(self, FormulaOps)
        d = self.describe(str(p), sheet, preview_rows=5)
        kind = d["kind"]
        out: dict[str, Any] = {
            "file": d["file"],
            "kind": kind,
            "size_kb": round(d["size_bytes"] / 1024, 1),
            "modified": d["modified"],
        }
        target = sheet
        if kind in SHEETED_KINDS:
            out["sheets"] = [
                {
                    k: s.get(k)
                    for k in ("name", "max_row", "max_column", "formula_cells", "tables", "charts", "state")
                    if s.get(k) not in (None, [], 0, "visible")
                }
                for s in d.get("sheets", [])
            ]
            if not target and d.get("sheets"):
                target = d["sheets"][0]["name"]
            if d.get("named_ranges"):
                out["named_ranges"] = d["named_ranges"][:20]
        try:
            prof = self.profile(str(p), target, top_values=3)
            cols = []
            for c in prof["columns"]:
                e: dict[str, Any] = {
                    "name": c["name"],
                    "type": c["type"],
                    "null_pct": c["null_pct"],
                    "distinct": c["approx_distinct"],
                }
                if "avg" in c:
                    e.update({"min": c["min"], "max": c["max"], "avg": c.get("avg")})
                elif c.get("top_values"):
                    e["top"] = [t["value"] for t in c["top_values"]]
                if c.get("unique"):
                    e["unique"] = True
                cols.append(e)
            out["rows"] = prof["rows"]
            out["columns"] = cols
            out["quality_score"] = prof["quality_score"]
            out["issues"] = [
                f"[{i['severity']}] {i['column'] or 'table'}: {i['issue']} - {i['detail']}"
                for i in prof["issues"]
            ][:12]
        except SheetError as e:
            out["profile_error"] = e.message
        if kind == "excel":
            try:
                st = self.formulas(str(p), "stats")
                if st["total_formulas"]:
                    er = self.formulas(str(p), "errors", limit=10)
                    out["formulas"] = {
                        "count": st["total_formulas"],
                        "top_functions": list(st["functions_used"])[:8],
                        "errors": er["error_count"],
                        "error_cells": [f"{x['sheet']}!{x['cell']} {x['error']}" for x in er["errors"][:5]],
                    }
            except SheetError:
                pass
        sh = (
            next((s for s in d.get("sheets", []) if s.get("name") == target), None)
            if kind in SHEETED_KINDS
            else d
        )
        if sh and sh.get("preview") is not None:
            names = [c["name"] if isinstance(c, dict) else c for c in sh.get("columns", [])]
            out["preview"] = {"columns": names, "rows": sh["preview"][:5]}
        if p.parent == self.sb.roots[0]:
            stem = safe_identifier(p.stem)
            if kind in SHEETED_KINDS and target and d.get("sheets") and target != d["sheets"][0]["name"]:
                out["sql_table"] = f"{stem}__{safe_identifier(target)}"
            else:
                out["sql_table"] = stem
        else:
            spec = self.sb.display(p) + (f"#{target}" if kind in SHEETED_KINDS and target else "")
            out["sql_table"] = f"'{spec}'"
        steps = ["query(sql=...) to answer questions", "aggregate(...) for group-by / pivot summaries"]
        if out.get("issues"):
            steps.insert(0, "clean(...) to fix the issues listed above (preview first, then output=...)")
        if out.get("formulas", {}).get("errors"):
            steps.insert(0, "formulas(action='trace', cell=...) on an error cell to find the root cause")
        out["next_steps"] = steps
        return out

    # ================================================================== clean pipelines
    def clean(
        self,
        source: str,
        operations: list[dict[str, Any]],
        sheet: str | None = None,
        output: str | None = None,
        output_sheet: str | None = None,
        limit: int | None = None,
        format: str = "json",
    ) -> dict:
        """Apply cleaning steps in order. Without output: preview; with output: save the cleaned data."""
        from .sqlops import SqlOps

        assert isinstance(self, SqlOps)
        if not operations:
            raise SheetError("operations is empty.", "Example: [{'op': 'trim'}, {'op': 'dedupe'}]")
        if len(operations) > 50:
            raise SheetError("At most 50 operations per call.")
        con = self._connect()
        try:
            expr, label = self.source_sql(source, sheet)
            con.execute(f"CREATE TEMP VIEW step0 AS SELECT row_number() OVER () AS __rn, * FROM {expr}")
            rows_before = con.execute("SELECT count(*) FROM step0").fetchone()[0]
            applied = []
            for i, op in enumerate(operations, start=1):
                info = con.execute(f"DESCRIBE step{i - 1}").fetchall()
                types = {r[0]: r[1] for r in info if r[0] != "__rn"}
                sel = self._clean_step(op, types, f"step{i - 1}")
                try:
                    con.execute(f"CREATE TEMP VIEW step{i} AS {sel}")
                    con.execute(f"SELECT * FROM step{i} LIMIT 0")
                except duckdb.Error as e:
                    raise SheetError(
                        f"Operation {i} ({op.get('op')}) failed: {str(e).splitlines()[0]}",
                        "Check column names (case-sensitive) and option values.",
                    ) from e
                applied.append(op.get("op"))
            last = f"step{len(operations)}"
            final = f"SELECT * EXCLUDE (__rn) FROM {last} ORDER BY __rn"
            rows_after = con.execute(f"SELECT count(*) FROM {last}").fetchone()[0]
            meta = {
                "source": label,
                "operations": applied,
                "rows_before": rows_before,
                "rows_after": rows_after,
            }
            if output:
                src_file, src_sheet = self._split_spec(source, sheet)
                src_path = self.sb.resolve(src_file)
                out_path = self.sb.resolve(output, must_exist=False)
                if (
                    out_path == src_path
                    and file_kind(src_path) == "excel"
                    and output_sheet in (None, "", src_sheet)
                    and (src_sheet or output_sheet is None)
                ):
                    return self._clean_in_place(con, final, src_path, src_sheet) | meta
                if out_path == src_path and file_kind(src_path) == "excel" and not output_sheet:
                    return self._clean_in_place(con, final, src_path, src_sheet) | meta
                return self._write_output(con, final, output, None, output_sheet) | meta
            return (
                self._result(con, final, limit, 0, format, {})
                | meta
                | {"note": "Preview only. Pass output='file.xlsx' (or the same file) to save."}
            )
        finally:
            con.close()

    def _clean_in_place(self, con: Any, final: str, p: Path, sheet: str | None) -> dict:
        """Write cleaned values back into the same sheet, keeping styles, tables, widths and other sheets."""
        self._writable("save cleaned data")
        cur = con.execute(final)
        cols = [d[0] for d in cur.description]
        rows = cur.fetchall()
        wb = load_workbook(p, formulas=True, fresh=True)
        ws = get_sheet(wb, sheet)
        hdr_row = 1
        for r in builtins_range(1, min(ws.max_row, 30) + 1):
            if any(c.value not in (None, "") for c in ws[r]):
                hdr_row = r
                break
        old_last_row, old_last_col = ws.max_row, ws.max_column
        formulas_lost = 0
        for row in ws.iter_rows(min_row=hdr_row, max_row=old_last_row, max_col=old_last_col):
            for c in row:
                if isinstance(c.value, str) and c.value.startswith("="):
                    formulas_lost += 1
                c.value = None
        for j, name in enumerate(cols, start=1):
            ws.cell(row=hdr_row, column=j, value=name)
        for i, r in enumerate(rows, start=hdr_row + 1):
            for j, v in enumerate(r, start=1):
                ws.cell(row=i, column=j, value=v)
                if (
                    isinstance(v, (dt.date, dt.datetime))
                    and ws.cell(row=i, column=j).number_format == "General"
                ):
                    ws.cell(row=i, column=j).number_format = (
                        "yyyy-mm-dd" if not isinstance(v, dt.datetime) else "yyyy-mm-dd hh:mm"
                    )
        last_row = hdr_row + max(len(rows), 1)
        resized = []
        for t in ws.tables.values():
            c1, r1, _, _ = parse_range(t.ref)
            if r1 == hdr_row and c1 == 1:
                t.ref = f"A{hdr_row}:{get_column_letter(len(cols))}{last_row}"
                t.tableColumns = []
                t._initialise_columns()
                for k, tc in enumerate(t.tableColumns):
                    tc.name = str(cols[k])
                if t.autoFilter is not None:
                    t.autoFilter.ref = t.ref
                resized.append(t.displayName)
        if ws.auto_filter.ref:
            ws.auto_filter.ref = f"A{hdr_row}:{get_column_letter(len(cols))}{last_row}"
        res: dict[str, Any] = {
            "output": self.sb.display(p),
            "sheet": ws.title,
            "rows_written": len(rows),
            "in_place": True,
            "backup": save_workbook(wb, p, self.cfg),
        }
        if resized:
            res["tables_resized"] = resized
        if formulas_lost:
            res["warning"] = (
                f"{formulas_lost} formula cell(s) in the cleaned range now hold their values. "
                "Undo with backups(action='undo')."
            )
        return res

    def _clean_step(self, op: dict[str, Any], types: dict[str, str], prev: str) -> str:
        name = str(op.get("op") or op.get("type") or "").lower()
        cols_all = list(types)

        def pick(default_text: bool = True) -> list[str]:
            cols = op.get("columns") or op.get("column")
            if isinstance(cols, str):
                cols = [cols]
            if not cols:
                return [c for c, t in types.items() if t == "VARCHAR"] if default_text else cols_all
            low = {c.lower(): c for c in cols_all}
            out = []
            for c in cols:
                if c not in types and c.lower() not in low:
                    raise SheetError(f"Column '{c}' not found.", "Available: " + ", ".join(cols_all[:40]))
                out.append(c if c in types else low[c.lower()])
            return out

        def as_text(c: str) -> str:
            return qident(c) if types.get(c) == "VARCHAR" else f"CAST({qident(c)} AS VARCHAR)"

        def replace_cols(exprs: dict[str, str]) -> str:
            if not exprs:
                return f"SELECT * FROM {prev}"
            return (
                f"SELECT * REPLACE ({', '.join(f'{e} AS {qident(c)}' for c, e in exprs.items())}) FROM {prev}"
            )

        if name == "trim":
            collapse = op.get("collapse_spaces", True)
            return replace_cols(
                {
                    c: (
                        f"regexp_replace(trim({as_text(c)}), '\\s+', ' ', 'g')"
                        if collapse
                        else f"trim({as_text(c)})"
                    )
                    for c in pick()
                }
            )
        if name == "case":
            mode = str(op.get("mode", "title")).lower()
            if mode not in ("upper", "lower", "title"):
                raise SheetError("case mode must be upper, lower or title.")

            def f(c: str) -> str:
                t = as_text(c)
                if mode == "upper":
                    return f"upper({t})"
                if mode == "lower":
                    return f"lower({t})"
                return (
                    f"array_to_string(list_transform(string_split(lower({t}), ' '), "
                    f"{_LAMBDA_X} upper(x[1]) || x[2:]), ' ')"
                )

            return replace_cols({c: f(c) for c in pick()})
        if name in ("replace", "find_replace"):
            find = op.get("find")
            if find is None:
                raise SheetError("replace needs find (and replace).")
            repl = str(op.get("replace", ""))
            if op.get("regex"):
                flags = "g" + ("" if op.get("case_sensitive", True) else "i")
                return replace_cols(
                    {
                        c: f"regexp_replace({as_text(c)}, {qstr(str(find))}, {qstr(repl)}, '{flags}')"
                        for c in pick()
                    }
                )
            if op.get("whole_cell"):
                return replace_cols(
                    {
                        c: f"CASE WHEN {as_text(c)} = {qstr(str(find))} THEN {qstr(repl)} ELSE {as_text(c)} END"
                        for c in pick()
                    }
                )
            return replace_cols(
                {c: f"replace({as_text(c)}, {qstr(str(find))}, {qstr(repl)})" for c in pick()}
            )
        if name == "split":
            col = pick()[0] if op.get("column") or op.get("columns") else None
            if not col:
                raise SheetError(
                    "split needs column.",
                    "Example: {'op': 'split', 'column': 'Name', 'sep': ' ', 'into': ['First', 'Last']}",
                )
            sep = str(op.get("sep", ","))
            into = op.get("into") or [f"{col}_{k}" for k in range(1, int(op.get("parts", 2)) + 1)]
            arr = (
                f"regexp_split_to_array({as_text(col)}, {qstr(sep)})"
                if op.get("regex")
                else f"string_split({as_text(col)}, {qstr(sep)})"
            )
            parts = []
            n = len(into)
            for k, nm in enumerate(into, start=1):
                if k == n and op.get("rest_in_last", True) and not op.get("regex"):
                    parts.append(f"array_to_string(({arr})[{k}:], {qstr(sep)}) AS {qident(nm)}")
                else:
                    parts.append(f"trim(({arr})[{k}]) AS {qident(nm)}")
            drop = f" EXCLUDE ({qident(col)})" if op.get("drop_original", False) else ""
            return f"SELECT *{drop}, {', '.join(parts)} FROM {prev}"
        if name == "merge":
            cols = pick(default_text=False)
            if len(cols) < 2:
                raise SheetError("merge needs at least 2 columns.")
            into = op.get("into") or "_".join(cols)
            sep = str(op.get("sep", " "))
            drop = (
                f" EXCLUDE ({', '.join(qident(c) for c in cols)})" if op.get("drop_original", False) else ""
            )
            return f"SELECT *{drop}, concat_ws({qstr(sep)}, {', '.join(as_text(c) for c in cols)}) AS {qident(into)} FROM {prev}"
        if name == "fill_down":
            cols = pick(default_text=False)
            exprs = {}
            for c in cols:
                v = f"NULLIF(trim({qident(c)}), '')" if types[c] == "VARCHAR" else qident(c)
                exprs[c] = (
                    f"last_value({v} IGNORE NULLS) OVER (ORDER BY __rn ROWS BETWEEN UNBOUNDED PRECEDING AND CURRENT ROW)"
                )
            return replace_cols(exprs)
        if name == "fill_null":
            if "value" not in op:
                raise SheetError("fill_null needs value.")
            v = op["value"]
            lit = qstr(str(v)) if isinstance(v, str) else ("NULL" if v is None else str(v))
            exprs = {}
            for c in pick(default_text=False):
                null = f"NULLIF({qident(c)}, '')" if types[c] == "VARCHAR" else qident(c)
                exprs[c] = f"coalesce({null}, TRY_CAST({lit} AS {types[c]}))"
            return replace_cols(exprs)
        if name in ("to_number", "number"):
            dec = str(op.get("decimal", "."))
            exprs = {}
            for c in pick(default_text=False):
                t = f"regexp_replace({as_text(c)}, '[^0-9eE,.+-]', '', 'g')"
                t = f"replace(replace({t}, '.', ''), ',', '.')" if dec == "," else f"replace({t}, ',', '')"
                exprs[c] = f"TRY_CAST(NULLIF({t}, '') AS DOUBLE)"
            return replace_cols(exprs)
        if name in ("to_date", "date"):
            fmt = op.get("format")
            dayfirst = bool(op.get("dayfirst", False))
            exprs = {}
            for c in pick(default_text=False):
                t = as_text(c)
                if fmt:
                    exprs[c] = f"CAST(TRY_STRPTIME({t}, {qstr(str(fmt))}) AS DATE)"
                    continue
                dm = ["%d/%m/%Y", "%m/%d/%Y"] if dayfirst else ["%m/%d/%Y", "%d/%m/%Y"]
                fmts = [
                    "%Y/%m/%d",
                    *dm,
                    "%d-%m-%Y",
                    "%d.%m.%Y",
                    "%d %b %Y",
                    "%b %d, %Y",
                    "%d-%b-%Y",
                    "%B %d, %Y",
                ]
                tries = [f"TRY_CAST({t} AS DATE)"] + [
                    f"CAST(TRY_STRPTIME({t}, {qstr(f)}) AS DATE)" for f in fmts
                ]
                exprs[c] = f"coalesce({', '.join(tries)})"
            return replace_cols(exprs)
        if name in ("drop_empty_rows", "drop_empty"):
            conds = " AND ".join(f"coalesce(CAST({qident(c)} AS VARCHAR), '') = ''" for c in cols_all)
            return f"SELECT * FROM {prev} WHERE NOT ({conds})"
        if name in ("dedupe", "remove_duplicates", "distinct"):
            cols = pick(default_text=False)
            order = "__rn DESC" if op.get("keep") == "last" else "__rn"
            part = ", ".join(qident(c) for c in cols)
            return (
                f"SELECT * FROM {prev} QUALIFY row_number() OVER (PARTITION BY {part} ORDER BY {order}) = 1"
            )
        if name in ("drop_columns", "drop"):
            cols = pick(default_text=False) if (op.get("columns") or op.get("column")) else []
            if not cols:
                raise SheetError("drop_columns needs columns.")
            return f"SELECT * EXCLUDE ({', '.join(qident(c) for c in cols)}) FROM {prev}"
        if name in ("keep_columns", "select"):
            cols = pick(default_text=False)
            return f"SELECT __rn, {', '.join(qident(c) for c in cols)} FROM {prev}"
        if name == "rename":
            mapping = op.get("mapping") or {}
            if not mapping:
                raise SheetError("rename needs mapping={'old': 'new'}.")
            parts = ["__rn"] + [f"{qident(c)} AS {qident(mapping.get(c, c))}" for c in cols_all]
            return f"SELECT {', '.join(parts)} FROM {prev}"
        if name == "filter":
            if not op.get("where"):
                raise SheetError('filter needs where, e.g. "Amount > 0".')
            from .sqlops import SqlOps

            SqlOps._check_select(f"SELECT * FROM {prev} WHERE {op['where']}")
            return f"SELECT * FROM {prev} WHERE {op['where']}"
        if name == "sort":
            ob = op.get("order_by") or op.get("by")
            if not ob:
                raise SheetError("sort needs order_by, e.g. ['Amount DESC'].")
            ob = [ob] if isinstance(ob, str) else ob
            parts = []
            for o in ob:
                m = re.match(r"^\s*(.+?)(\s+(?:asc|desc))?\s*$", o, re.I)
                col, direction = (m.group(1), m.group(2) or "") if m else (o, "")
                parts.append((qident(col) if col in types else col) + direction)
            return f"SELECT * REPLACE (row_number() OVER (ORDER BY {', '.join(parts)}) AS __rn) FROM {prev}"
        if name in ("add_column", "compute"):
            nm, ex = op.get("name"), op.get("expr") or op.get("expression")
            if not nm or not ex:
                raise SheetError(
                    "add_column needs name and expr, e.g. {'name': 'Total', 'expr': 'Qty * Price'}."
                )
            return f"SELECT *, {ex} AS {qident(nm)} FROM {prev}"
        raise SheetError(
            f"Unknown clean op '{name}'.",
            "Ops: trim, case, replace, split, merge, fill_down, fill_null, to_number, to_date, drop_empty_rows, "
            "dedupe, drop_columns, keep_columns, rename, filter, sort, add_column",
        )

    # ================================================================== Excel tables / names / rules
    def manage_objects(
        self,
        path: str,
        action: str = "list",
        sheet: str | None = None,
        name: str | None = None,
        range: str | None = None,
        refers_to: str | None = None,
        style: str = "TableStyleMedium9",
        dry_run: bool = False,
    ) -> dict:
        """list | table_create | table_delete | name_create | name_delete | validation_delete | conditional_delete."""
        action = (action or "list").lower()
        valid = {
            "list",
            "table_create",
            "table_delete",
            "name_create",
            "name_delete",
            "validation_delete",
            "conditional_delete",
        }
        if action not in valid:
            raise SheetError(f"Unknown action '{action}'.", "Use one of: " + ", ".join(sorted(valid)))
        p = self.sb.resolve(path)
        if file_kind(p) != "excel":
            raise SheetError("Tables, names and rules exist only in .xlsx files.", "Convert the file first.")
        if action == "list":
            wb = load_workbook(p, formulas=True)
            out: dict[str, Any] = {
                "file": self.sb.display(p),
                "tables": [],
                "named_ranges": [],
                "validations": [],
                "conditional_formats": [],
            }
            for ws in wb.worksheets:
                for t in ws.tables.values():
                    cols = [c.name for c in t.tableColumns]
                    out["tables"].append(
                        {
                            "name": t.displayName,
                            "sheet": ws.title,
                            "ref": t.ref,
                            "columns": cols,
                            "style": t.tableStyleInfo.name if t.tableStyleInfo else None,
                        }
                    )
                for dv in ws.data_validations.dataValidation:
                    out["validations"].append(
                        {
                            "sheet": ws.title,
                            "range": str(dv.sqref),
                            "type": dv.type,
                            "formula1": dv.formula1,
                            "formula2": dv.formula2,
                        }
                    )
                for cf in ws.conditional_formatting:
                    for r in cf.rules:
                        out["conditional_formats"].append(
                            {
                                "sheet": ws.title,
                                "range": str(cf.sqref),
                                "type": r.type,
                                "operator": r.operator,
                                "formula": r.formula,
                            }
                        )
                for n, dn in getattr(ws, "defined_names", {}).items():
                    out["named_ranges"].append({"name": n, "refers_to": dn.attr_text, "scope": ws.title})
            for n, dn in wb.defined_names.items():
                out["named_ranges"].append({"name": n, "refers_to": dn.attr_text, "scope": "workbook"})
            return out
        if not dry_run:
            self._writable(action)
        wb = load_workbook(p, formulas=True, fresh=True)
        ws = get_sheet(wb, sheet)
        res: dict[str, Any] = {"file": self.sb.display(p), "action": action, "dry_run": dry_run}
        if action == "table_create":
            if not range:
                raise SheetError("table_create needs range (including the header row), e.g. 'A1:F41'.")
            name = name or f"Table{sum(len(w.tables) for w in wb.worksheets) + 1}"
            if not _TABLE_NAME.match(name) or re.match(r"^[A-Za-z]{1,3}\d+$", name):
                raise SheetError(
                    f"'{name}' is not a valid table name.",
                    "Use letters, digits and _ (no spaces), e.g. SalesData.",
                )
            existing = {t.displayName.lower() for w in wb.worksheets for t in w.tables.values()}
            if name.lower() in existing or name in wb.defined_names:
                raise SheetError(f"A table or name '{name}' already exists.", "Pick another name.")
            c1, r1, c2, r2 = parse_range(range, ws.max_row, ws.max_column)
            if r2 <= r1:
                raise SheetError("A table needs a header row and at least one data row.")
            for t in ws.tables.values():
                tc1, tr1, tc2, tr2 = parse_range(t.ref)
                if not (c2 < tc1 or c1 > tc2 or r2 < tr1 or r1 > tr2):
                    raise SheetError(f"Range overlaps existing table '{t.displayName}' ({t.ref}).")
            fixed = []
            seen: set[str] = set()
            for c in builtins_range(c1, c2 + 1):
                cell = ws.cell(row=r1, column=c)
                h = "" if cell.value is None else str(cell.value).strip()
                if not h or h.lower() in seen or not isinstance(cell.value, str):
                    new = h or f"Column{c - c1 + 1}"
                    while new.lower() in seen:
                        new += "_2"
                    if new != cell.value:
                        fixed.append(f"{cell.coordinate}: {cell.value!r} -> {new!r}")
                        if not dry_run:
                            cell.value = new
                    h = new
                seen.add(h.lower())
            ref = f"{get_column_letter(c1)}{r1}:{get_column_letter(c2)}{r2}"
            t = Table(displayName=name, ref=ref)
            t.tableStyleInfo = TableStyleInfo(name=style, showRowStripes=True, showColumnStripes=False)
            ws.add_table(t)
            res.update(
                {
                    "table": name,
                    "sheet": ws.title,
                    "ref": ref,
                    "use_in_formulas": f"=SUM({name}[<column>]), [@<column>] inside the table",
                }
            )
            if fixed:
                res["headers_fixed"] = fixed
        elif action == "table_delete":
            found = False
            for w in wb.worksheets:
                for tname in list(w.tables.keys()):
                    if name and tname.lower() == name.lower():
                        del w.tables[tname]
                        found = True
            if not found:
                raise SheetError(f"Table '{name}' not found.", "Use action='list' to see tables.")
            res["deleted"] = name
            res["note"] = "The cells and their values are kept; only the table object was removed."
        elif action == "name_create":
            if not name:
                raise SheetError("name_create needs name.")
            if not refers_to:
                if not range:
                    raise SheetError(
                        "name_create needs refers_to ('Sheet1!$A$1:$A$9' or '=0.15') or range (+sheet)."
                    )
                c1, r1, c2, r2 = parse_range(range, ws.max_row, ws.max_column)
                title = (
                    ws.title
                    if re.fullmatch(r"[A-Za-z_]\w*", ws.title)
                    else "'" + ws.title.replace("'", "''") + "'"
                )
                refers_to = f"{title}!${get_column_letter(c1)}${r1}" + (
                    f":${get_column_letter(c2)}${r2}" if (c1, r1) != (c2, r2) else ""
                )
            refers_to = refers_to.lstrip("=")
            if not re.fullmatch(r"[A-Za-z_\\][\w.]*", name) or re.fullmatch(r"[A-Za-z]{1,3}\d+", name):
                raise SheetError(
                    f"'{name}' is not a valid name.", "Use letters, digits, _ and . (no spaces, not like A1)."
                )
            if name in wb.defined_names:
                del wb.defined_names[name]
                res["replaced"] = True
            wb.defined_names[name] = DefinedName(name, attr_text=refers_to)
            res.update({"name": name, "refers_to": refers_to})
        elif action == "name_delete":
            if not name or name not in wb.defined_names:
                raise SheetError(f"Name '{name}' not found.", "Use action='list' to see names.")
            del wb.defined_names[name]
            res["deleted"] = name
        elif action in ("validation_delete", "conditional_delete"):
            n = self._delete_rules(ws, action, range)
            res.update({"sheet": ws.title, "removed": n})
        if not dry_run:
            res["backup"] = save_workbook(wb, p, self.cfg)
        return res

    @staticmethod
    def _delete_rules(ws: Any, action: str, rng: str | None) -> int:
        from openpyxl.formatting.formatting import ConditionalFormattingList

        def overlaps(sqref: str) -> bool:
            if not rng:
                return True
            a = parse_range(rng)
            for part in str(sqref).split():
                b = parse_range(part)
                if not (a[2] < b[0] or a[0] > b[2] or a[3] < b[1] or a[1] > b[3]):
                    return True
            return False

        if action == "validation_delete":
            before = ws.data_validations.dataValidation
            keep = [dv for dv in before if not overlaps(str(dv.sqref))]
            ws.data_validations.dataValidation = keep
            return len(before) - len(keep)
        new = ConditionalFormattingList()
        removed = 0
        for cf in ws.conditional_formatting:
            if overlaps(str(cf.sqref)):
                removed += len(cf.rules)
                continue
            for r in cf.rules:
                new.add(str(cf.sqref), r)
        ws.conditional_formatting = new
        return removed

    # ================================================================== backups / undo
    def _backup_entries(self, target: Path | None, root: Path) -> list[dict[str, Any]]:
        dirs = [target.parent / ".sheetmcp_backups"] if target else list(root.rglob(".sheetmcp_backups"))
        out = []
        for bdir in dirs:
            if not bdir.is_dir():
                continue
            for f in bdir.iterdir():
                m = _BACKUP_RE.match(f.name)
                if not m:
                    continue
                orig = bdir.parent / (m.group("stem") + m.group("ext"))
                if target and orig != target:
                    continue
                ts = dt.datetime.strptime(m.group("ts")[:15], "%Y%m%d-%H%M%S")
                out.append(
                    {
                        "backup": self.sb.display(f),
                        "file": self.sb.display(orig),
                        "created": ts.isoformat(),
                        "size_bytes": f.stat().st_size,
                        "kind": "pre-undo snapshot" if m.group("tag") else "before write",
                        "_path": f,
                        "_orig": orig,
                        "_key": m.group("ts"),
                    }
                )
        out.sort(key=lambda e: e["_key"], reverse=True)
        return out

    def backups(
        self, path: str | None = None, action: str = "list", backup: str | None = None, limit: int = 50
    ) -> dict:
        """list | restore (a given backup) | undo (latest change to a file) | delete. Restores are undoable too."""
        action = (action or "list").lower()
        target = None
        if path:
            tp = self.sb.resolve(path, must_exist=False, want="any")
            target = None if tp.is_dir() else tp
            root = tp if tp.is_dir() else tp.parent
        else:
            root = self.sb.roots[0]

        def clean(e: dict[str, Any]) -> dict[str, Any]:
            return {k: v for k, v in e.items() if not k.startswith("_")}

        if action == "list":
            entries = self._backup_entries(target, root)
            return {
                "count": len(entries),
                "backups": [clean(e) for e in entries[:limit]],
                "hint": "backups(action='undo', path=<file>) restores the latest; action='restore', backup=<path> picks one.",
            }
        if action not in ("restore", "undo", "delete"):
            raise SheetError(f"Unknown action '{action}'.", "Use list, restore, undo or delete.")
        self._writable(f"{action} backups")
        if action == "undo":
            if not target:
                raise SheetError("undo needs path=<file>.")
            cands = [e for e in self._backup_entries(target, root) if e["kind"] == "before write"]
            if not cands:
                raise SheetError(
                    f"No backups found for '{path}'.", "Backups are created when SheetMCP overwrites a file."
                )
            chosen = cands[0]
        else:
            if not backup:
                raise SheetError(f"{action} needs backup=<backup path from action='list'>.")
            bp = self.sb.resolve(backup)
            m = _BACKUP_RE.match(bp.name)
            if not m or bp.parent.name != ".sheetmcp_backups":
                raise SheetError(
                    f"'{backup}' is not a SheetMCP backup file.", "Use a path from backups(action='list')."
                )
            chosen = {"_path": bp, "_orig": bp.parent.parent / (m.group("stem") + m.group("ext"))}
        bp, orig = chosen["_path"], chosen["_orig"]
        if action == "delete":
            bp.unlink()
            return {"deleted": self.sb.display(bp)}
        snapshot = None
        if orig.exists():  # keep the current state so the undo itself can be undone (restore that snapshot)
            snap = bp.parent / f"{orig.stem}.{backup_stamp()}-undo{orig.suffix}"
            shutil.copy2(orig, snap)
            snapshot = self.sb.display(snap)
        old = self.cfg.backup
        self.cfg.backup = False
        try:
            atomic_write(orig, lambda t: shutil.copy2(bp, t), self.cfg)
        finally:
            self.cfg.backup = old
        if action == "undo":
            bp.unlink()  # consumed, so the next undo goes one step further back
        return {
            "restored": self.sb.display(orig),
            "from": self.sb.display(bp) if action != "undo" else "latest backup",
            "previous_state_saved_as": snapshot,
            "hint": "To revert this undo: backups(action='restore', backup=previous_state_saved_as).",
        }

    # ================================================================== batch
    BATCHABLE = (
        "list_files",
        "describe",
        "read_range",
        "write_range",
        "manage_sheets",
        "insert_delete",
        "query",
        "transform",
        "aggregate",
        "combine",
        "convert",
        "profile",
        "formulas",
        "search",
        "diff",
        "format_range",
        "add_chart",
        "quick_look",
        "clean",
        "manage_objects",
        "backups",
        "build_report",
        "catalog",
    )

    def batch(self, operations: list[dict[str, Any]], stop_on_error: bool = True) -> dict:
        """Run many operations in one call: [{'tool': 'write_range', 'args': {...}}, ...]."""
        if not operations:
            raise SheetError("operations is empty.", "Example: [{'tool': 'format_range', 'args': {...}}]")
        if len(operations) > 100:
            raise SheetError("At most 100 operations per batch.")
        results = []
        ok = 0
        for i, op in enumerate(operations, start=1):
            tool = op.get("tool") or op.get("op")
            args = op.get("args") or {k: v for k, v in op.items() if k not in ("tool", "op")}
            if tool not in self.BATCHABLE:
                err = f"Unknown tool '{tool}'. Allowed: {', '.join(self.BATCHABLE)}"
                results.append({"step": i, "tool": tool, "ok": False, "error": err})
                if stop_on_error:
                    break
                continue
            try:
                r = getattr(self, tool)(**args)
                results.append({"step": i, "tool": tool, "ok": True, "result": r})
                ok += 1
            except SheetError as e:
                results.append({"step": i, "tool": tool, "ok": False, "error": str(e)})
                if stop_on_error:
                    break
            except TypeError as e:
                results.append({"step": i, "tool": tool, "ok": False, "error": f"Bad arguments: {e}"})
                if stop_on_error:
                    break
        res: dict[str, Any] = {"steps": len(operations), "succeeded": ok, "results": results}
        if ok < len(operations):
            res["stopped_early"] = stop_on_error and len(results) < len(operations)
        return res

    # ================================================================== report builder
    def build_report(
        self,
        source: str,
        output: str,
        group_by: list[str],
        metrics: list[str],
        sheet: str | None = None,
        where: str | None = None,
        title: str | None = None,
        chart_type: str = "column",
        include_data: bool = True,
        number_format: str = "#,##0.00",
        overwrite: bool = False,
        top: int | None = None,
    ) -> dict:
        """One call: Summary sheet (group-by + totals row + chart) and Data sheet, fully formatted."""
        from .formatops import FormatOps
        from .sqlops import SqlOps
        from .workbook import WorkbookOps

        assert isinstance(self, SqlOps) and isinstance(self, FormatOps) and isinstance(self, WorkbookOps)
        self._writable("build a report")
        if not group_by or not metrics:
            raise SheetError(
                "build_report needs group_by and metrics.",
                "Example: group_by=['Region'], metrics=['Revenue:sum', 'Units:avg', 'count']",
            )
        out = self.sb.resolve(output, must_exist=False)
        if out.suffix.lower() != ".xlsx":
            raise SheetError("Reports are .xlsx files.", "Use output='report.xlsx'.")
        if out.exists() and not overwrite:
            raise SheetError(
                f"'{output}' already exists.", "Pass overwrite=true (a backup is kept) or another name."
            )
        preview = self.aggregate(source, metrics, group_by, sheet, where=where, limit=10_000)
        cols = preview["columns"]
        n = preview["row_count"]
        if n == 0:
            raise SheetError("The summary is empty (no rows matched).", "Check where= and group_by.")
        first_metric = cols[len(group_by)] if len(cols) > len(group_by) else None
        order = [f"{qident(first_metric)} DESC"] if first_metric else None
        backup = None
        if out.exists():
            from .fileio import backup_file

            backup = backup_file(out, self.cfg)
            out.unlink()
        old = self.cfg.backup
        self.cfg.backup = False  # intermediate saves of a brand-new file need no backups
        try:
            agg = self.aggregate(
                source,
                metrics,
                group_by,
                sheet,
                where=where,
                order_by=order,
                limit=top,
                output=str(out),
                output_sheet="Summary",
            )
            if top:
                agg_rows = min(n, top)
            else:
                agg_rows = agg["rows_written"]
            data_rows = None
            if include_data:
                d = self.transform(source, sheet, where=where, output=str(out), output_sheet="Data")
                data_rows = d["rows_written"]
            wb = load_workbook(out, formulas=True, fresh=True)
            ws = wb["Summary"]
            last = agg_rows + 1
            total_row = last + 1
            ws.cell(row=total_row, column=1, value="Total")
            for j, cname in enumerate(cols, start=1):
                ws.cell(row=1, column=j, value=_pretty_header(str(cname)))
            fn_for = {
                "sum": "SUM",
                "count": "SUM",
                "avg": "AVERAGE",
                "mean": "AVERAGE",
                "min": "MIN",
                "max": "MAX",
                "count_distinct": None,
                "median": "MEDIAN",
            }
            for i, m in enumerate(metrics):
                col = len(group_by) + 1 + i
                fnn = "SUM" if m.strip().lower() in ("count", "count(*)") else None
                if ":" in m and "(" not in m:
                    fnn = fn_for.get(m.rsplit(":", 1)[1].strip().lower())
                if fnn and last >= 2:
                    L = get_column_letter(col)
                    ws.cell(row=total_row, column=col, value=f"={fnn}({L}2:{L}{last})")
            wb.save(out)
            nm_cols = len(cols)
            end_col = get_column_letter(nm_cols)
            self.format_range(
                str(out),
                f"A1:{end_col}1",
                "Summary",
                bold=True,
                fill_color="1F4E78",
                font_color="FFFFFF",
                h_align="center",
                freeze_panes="A2",
            )
            if nm_cols > len(group_by):
                self.format_range(
                    str(out),
                    f"{get_column_letter(len(group_by) + 1)}2:{end_col}{total_row}",
                    "Summary",
                    number_format=number_format,
                )
            self.format_range(
                str(out),
                f"A{total_row}:{end_col}{total_row}",
                "Summary",
                bold=True,
                border="thin",
                fill_color="DDEBF7",
            )
            self.format_range(str(out), f"A1:{end_col}{total_row}", "Summary", autofit=True)
            chart = None
            if first_metric and chart_type and chart_type != "none" and last >= 2:
                L = get_column_letter(len(group_by) + 1)
                chart = self.add_chart(
                    str(out),
                    f"{L}1:{L}{last}",
                    chart_type,
                    "Summary",
                    categories_range=f"A2:A{last}",
                    title=title or f"{_pretty_header(first_metric)} by {group_by[0]}",
                    anchor=f"{get_column_letter(nm_cols + 2)}2",
                )
            if include_data and data_rows:
                wbd = load_workbook(out, formulas=False, read_only=True, fresh=True)
                dcols = wbd["Data"].max_column
                wbd.close()
                dend = get_column_letter(max(dcols, 1))
                self.format_range(
                    str(out),
                    f"A1:{dend}1",
                    "Data",
                    bold=True,
                    fill_color="DDEBF7",
                    auto_filter=None,
                    freeze_panes="A2",
                )
                self.format_range(str(out), f"A1:{dend}{min(data_rows + 1, 5000)}", "Data", autofit=True)
                if data_rows + 1 <= 1_048_576:
                    wb2 = load_workbook(out, formulas=True, fresh=True)
                    wb2["Data"].auto_filter.ref = f"A1:{dend}{data_rows + 1}"
                    wb2.save(out)
            if title:
                wb3 = load_workbook(out, formulas=True, fresh=True)
                wb3["Summary"].sheet_properties.tabColor = "1F4E78"
                wb3.properties.title = title
                wb3.save(out)
        finally:
            self.cfg.backup = old
        res: dict[str, Any] = {
            "output": self.sb.display(out),
            "sheets": ["Summary"] + (["Data"] if include_data else []),
            "summary_rows": agg_rows,
            "summary_columns": cols,
            "data_rows": data_rows,
            "chart": chart["chart_type"] if chart else None,
            "totals_row": total_row,
        }
        if backup:
            res["backup_of_previous"] = backup
        return res


builtins_range = range


_PRETTY = {
    "sum": "Total",
    "avg": "Average",
    "mean": "Average",
    "min": "Min",
    "max": "Max",
    "median": "Median",
    "count_distinct": "Distinct",
    "stddev": "Std dev",
}


def _pretty_header(name: str) -> str:
    if name == "count":
        return "Count"
    for k in sorted(_PRETTY, key=len, reverse=True):
        if name.startswith(k + "_"):
            return f"{_PRETTY[k]} {name[len(k) + 1 :]}"
    return name
