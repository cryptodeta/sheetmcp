"""DuckDB-powered operations: SQL across many files, transform, aggregate/pivot, combine, convert, profile."""

from __future__ import annotations

import os
import re
from pathlib import Path
from typing import Any

import duckdb
import openpyxl
from openpyxl.cell import WriteOnlyCell
from openpyxl.styles import Font

from .base import Base
from .errors import SheetError
from .fileio import (
    SHEETED_KINDS,
    TABULAR_EXT,
    atomic_write,
    conv_dir,
    excel_to_csv_cache,
    file_kind,
    get_sheet,
    load_workbook,
    output_format,
    save_workbook,
    sheet_names,
    sniff_dialect,
)
from .util import jsonable, qident, qstr, safe_identifier, to_markdown

EXCEL_MAX_ROWS = 1_048_576
_FILE_LIT = re.compile(
    r"(\b(?:FROM|JOIN)\s+)'([^']+?\.(?:csv|tsv|txt|xlsx|xlsm|parquet|pq|json|jsonl|ndjson))(#[^']*)?'",
    re.IGNORECASE,
)
_AGG_FUNCS = {
    "sum": "sum",
    "avg": "avg",
    "mean": "avg",
    "average": "avg",
    "min": "min",
    "max": "max",
    "count": "count",
    "count_distinct": "count_distinct",
    "nunique": "count_distinct",
    "median": "median",
    "std": "stddev_samp",
    "stddev": "stddev_samp",
    "var": "var_samp",
    "first": "first",
    "last": "last",
    "list": "list",
    "mode": "mode",
}


class SqlOps(Base):
    # ------------------------------------------------------------------ connection & sources
    def _connect(self) -> duckdb.DuckDBPyConnection:
        con = duckdb.connect(":memory:")
        dirs = [str(r).rstrip("/\\") + os.sep for r in self.sb.roots] + [
            str(conv_dir()).rstrip("/\\") + os.sep
        ]
        con.execute("SET allowed_directories=[" + ", ".join(qstr(d) for d in dirs) + "]")
        con.execute(f"SET temp_directory={qstr(str(conv_dir() / 'spill'))}")
        con.execute("SET enable_external_access=false")
        con.execute("SET autoinstall_known_extensions=false")
        con.execute("SET autoload_known_extensions=false")
        con.execute("SET lock_configuration=true")
        return con

    def _split_spec(self, spec: Any, sheet: str | None = None) -> tuple[str, str | None]:
        if isinstance(spec, dict):
            return str(spec.get("path") or spec.get("file") or ""), spec.get("sheet") or sheet
        s = str(spec)
        if "#" in s:
            base, _, sh = s.rpartition("#")
            if Path(base).suffix.lower() in TABULAR_EXT:
                return base, sh or sheet
        return s, sheet

    def source_sql(
        self, spec: Any, sheet: str | None = None, header_row: int | None = None
    ) -> tuple[str, str]:
        """Return (DuckDB table expression, display label) for a file (optionally one sheet)."""
        path, sheet = self._split_spec(spec, sheet)
        p = self.sb.resolve(path)
        kind = file_kind(p)
        if kind in SHEETED_KINDS:
            csv_path, title = excel_to_csv_cache(p, sheet, header_row)
            return (
                f"read_csv({qstr(str(csv_path))}, header=true, auto_detect=true, sample_size=-1)",
                f"{self.sb.display(p)}#{title}",
            )
        if kind in ("csv", "tsv"):
            enc, _ = sniff_dialect(p)
            opts = ["auto_detect=true", "header=true"]
            if kind == "tsv":
                opts.append("delim='\\t'")
            if enc == "cp1252":
                opts.append("encoding='latin-1'")
            return f"read_csv({qstr(str(p))}, {', '.join(opts)})", self.sb.display(p)
        if kind == "parquet":
            return f"read_parquet({qstr(str(p))})", self.sb.display(p)
        if kind == "jsonl":
            return f"read_json_auto({qstr(str(p))}, format='newline_delimited')", self.sb.display(p)
        return f"read_json_auto({qstr(str(p))})", self.sb.display(p)

    def _register(
        self, con: duckdb.DuckDBPyConnection, sql: str, tables: Any = None, folder: str | None = None
    ) -> dict[str, str]:
        registered: dict[str, str] = {}
        used: set[str] = set()
        words = {w.lower() for w in re.findall(r"[A-Za-z_][A-Za-z0-9_]*", sql)}
        words |= {w.lower() for w in re.findall(r'"([^"]+)"', sql)}

        def add(alias: str, expr: str, label: str) -> None:
            con.execute(f"CREATE OR REPLACE VIEW {qident(alias)} AS SELECT * FROM {expr}")
            registered[alias] = label

        if isinstance(tables, dict):
            for alias, spec in tables.items():
                a = safe_identifier(alias, used) if not re.fullmatch(r"[A-Za-z_]\w*", alias) else alias
                used.add(a.lower())
                add(a, *self.source_sql(spec))
        elif isinstance(tables, (list, tuple)):
            for spec in tables:
                path, sh = self._split_spec(spec)
                a = safe_identifier(Path(path).stem + (f"_{sh}" if sh else ""), used)
                add(a, *self.source_sql(spec))
        elif isinstance(tables, str) and tables.strip():
            path, sh = self._split_spec(tables)
            add(safe_identifier(Path(path).stem + (f"_{sh}" if sh else ""), used), *self.source_sql(tables))

        scan_dir = folder if folder else (str(self.sb.roots[0]) if tables in (None, {}, []) else None)
        if scan_dir:
            d = self.sb.resolve(scan_dir, want="dir")
            for f in sorted(d.iterdir()):
                if not f.is_file() or f.suffix.lower() not in TABULAR_EXT or f.name.startswith("."):
                    continue
                stem = safe_identifier(f.stem)
                if file_kind(f) in SHEETED_KINDS:
                    # lazy: only convert sheets the SQL mentions (alias or alias__sheet)
                    cand = [w for w in words if w == stem or w.startswith(stem + "__")]
                    if not cand or all(w in registered for w in cand):
                        continue
                    try:
                        names = sheet_names(f)
                    except SheetError:
                        continue
                    for nm in names:
                        a = f"{stem}__{safe_identifier(nm)}"
                        if a in words and a not in registered:
                            add(a, *self.source_sql(str(f), nm))
                    if stem in words and stem not in registered:
                        add(stem, *self.source_sql(str(f)))
                elif stem in words and stem not in registered:
                    add(stem, *self.source_sql(str(f)))
        return registered

    def _rewrite_literals(self, sql: str) -> str:
        def rep(m: re.Match[str]) -> str:
            sheet = m.group(3)[1:] if m.group(3) else None
            expr, _ = self.source_sql(m.group(2), sheet)
            return m.group(1) + expr

        return _FILE_LIT.sub(rep, sql)

    @staticmethod
    def _check_select(sql: str) -> str:
        sql = sql.strip().rstrip(";").strip()
        if not sql:
            raise SheetError("Empty SQL.", "Example: SELECT * FROM sales LIMIT 10")
        try:
            stmts = duckdb.extract_statements(sql)
        except duckdb.Error as e:
            raise SheetError(
                f"SQL syntax error: {e}",
                "Check quotes and commas. Names with spaces/capitals, and table names that are SQL keywords "
                '(order, group, semi, user...), need "double quotes", e.g. SELECT * FROM "order".',
            ) from e
        if len(stmts) != 1:
            raise SheetError(
                "Only one SQL statement per call is allowed.", "Remove extra ';' separated statements."
            )
        if stmts[0].type != duckdb.StatementType.SELECT:
            raise SheetError(
                f"Only read-only SELECT queries are allowed (got {stmts[0].type.name}).",
                "To save results, pass output='result.xlsx' (or .csv/.parquet/.json) instead of COPY/CREATE.",
            )
        return sql

    def _db_error(self, e: Exception, registered: dict[str, str]) -> SheetError:
        msg = str(e).strip()
        hint = (
            'Use describe on the file to check column names; names with spaces/capitals need "double quotes".'
        )
        if "Table with name" in msg or ("does not exist" in msg and "Table" in msg):
            hint = (
                "Registered tables: "
                + (", ".join(registered) or "none")
                + ". Pass tables={'alias': 'file.xlsx'} or reference a file directly: FROM 'data/file.csv'. "
                "Files in the root folder are auto-registered by their name (sales.csv -> sales; "
                "book.xlsx sheet 'Q1' -> book__q1)."
            )
        elif "Conversion Error" in msg or "Could not convert" in msg:
            hint = "A column has mixed types. Use TRY_CAST(col AS DOUBLE) or compare as text."
        elif "Permission Error" in msg:
            hint = "That file is outside the allowed folders."
        return SheetError(
            f"Query failed: {msg.splitlines()[0] if msg else type(e).__name__}"
            + ("" if len(msg.splitlines()) < 2 else " " + " ".join(msg.splitlines()[1:3])),
            hint,
        )

    # ------------------------------------------------------------------ results
    def _result(
        self,
        con: duckdb.DuckDBPyConnection,
        sql: str,
        limit: int | None,
        offset: int,
        fmt: str,
        registered: dict[str, str],
    ) -> dict:
        limit = self._limit(limit)
        offset = max(0, int(offset or 0))
        try:
            cur = con.execute(f"SELECT * FROM ({sql}) AS _q LIMIT {limit + 1} OFFSET {offset}")
            cols = [d[0] for d in cur.description]
            types = [str(d[1]) for d in cur.description]
            rows = cur.fetchall()
        except duckdb.Error as e:
            raise self._db_error(e, registered) from e
        truncated = len(rows) > limit
        rows = rows[:limit]
        res: dict[str, Any] = {"columns": cols, "types": types, "row_count": len(rows)}
        data = [[self._j(v) for v in r] for r in rows]
        if fmt == "markdown":
            res["markdown"] = to_markdown(cols, data)
        elif fmt == "records":
            res["rows"] = [dict(zip(cols, r)) for r in data]
        else:
            res["rows"] = data
        if truncated:
            res["truncated"] = True
            res["next_offset"] = offset + limit
            res["hint"] = (
                "More rows available: call again with offset=next_offset, or pass output='file.xlsx' to save all rows."
            )
        if registered:
            res["tables"] = registered
        return res

    def _write_output(
        self,
        con: duckdb.DuckDBPyConnection,
        sql: str,
        output: str,
        fmt: str | None = None,
        sheet: str | None = None,
        registered: dict[str, str] | None = None,
        csv_dialect: str = "standard",
    ) -> dict:
        self._writable("save results")
        out = self.sb.resolve(output, must_exist=False)
        fmt = output_format(out, fmt)
        rows_written = 0
        try:
            n = con.execute(f"SELECT count(*) FROM ({sql}) AS _q").fetchone()[0]
        except duckdb.Error as e:
            raise self._db_error(e, registered or {}) from e
        if fmt in ("csv", "tsv", "parquet", "json", "jsonl"):
            opts = {
                "csv": "FORMAT csv, HEADER true",
                "tsv": "FORMAT csv, HEADER true, DELIMITER '\t'",
                "parquet": "FORMAT parquet, COMPRESSION zstd",
                "json": "FORMAT json, ARRAY true",
                "jsonl": "FORMAT json",
            }[fmt]
            dialect = (csv_dialect or "standard").lower().replace("-", "_")
            if dialect not in ("standard", "excel", "google_sheets"):
                raise SheetError(
                    f"Unknown csv_dialect '{csv_dialect}'.", "Use standard, excel or google_sheets."
                )

            def copy_to(t: Path) -> None:
                extra = ""
                if fmt in ("csv", "tsv") and dialect == "google_sheets":
                    extra = ", DATEFORMAT '%Y-%m-%d', TIMESTAMPFORMAT '%Y-%m-%d %H:%M:%S'"
                con.execute(f"COPY ({sql}) TO {qstr(str(t))} ({opts}{extra})")
                if fmt in ("csv", "tsv") and dialect == "excel":  # BOM so Excel detects UTF-8 (৳, é, ...)
                    data = t.read_bytes()
                    t.write_bytes(b"\xef\xbb\xbf" + data)

            backup = atomic_write(out, copy_to, self.cfg)
            rows_written = n
        elif fmt == "md":
            cur = con.execute(sql)
            cols = [d[0] for d in cur.description]
            text = to_markdown(cols, cur.fetchall()) + "\n"
            backup = atomic_write(out, lambda t: t.write_text(text, encoding="utf-8"), self.cfg)
            rows_written = n
        else:  # xlsx
            if n + 1 > EXCEL_MAX_ROWS:
                raise SheetError(
                    f"Result has {n:,} rows; Excel's limit is 1,048,576.", "Save as .parquet or .csv instead."
                )
            cur = con.execute(sql)
            cols = [d[0] for d in cur.description]
            sheet_name = (sheet or "Sheet1")[:31]
            if out.exists() and sheet:
                wb = load_workbook(out, formulas=True, fresh=True)
                idx = None
                if sheet_name in wb.sheetnames:
                    idx = wb.sheetnames.index(sheet_name)
                    wb.remove(wb[sheet_name])
                ws = wb.create_sheet(sheet_name, idx)
                ws.append(cols)
                for c in ws[1]:
                    c.font = Font(bold=True)
                while batch := cur.fetchmany(10_000):
                    for r in batch:
                        ws.append([_xl(v) for v in r])
                ws.freeze_panes = "A2"
                backup = save_workbook(wb, out, self.cfg)
            else:
                wb = openpyxl.Workbook(write_only=True)
                ws = wb.create_sheet(sheet_name)
                ws.freeze_panes = "A2"
                widths = [len(str(c)) for c in cols]
                header = []
                for c in cols:
                    cell = WriteOnlyCell(ws, value=c)
                    cell.font = Font(bold=True)
                    header.append(cell)
                buffered = cur.fetchmany(500)
                for r in buffered:
                    for i, v in enumerate(r):
                        widths[i] = max(widths[i], min(len(str(v)) if v is not None else 0, 60))
                for i, w in enumerate(widths):
                    ws.column_dimensions[openpyxl.utils.get_column_letter(i + 1)].width = w + 2
                ws.append(header)
                for r in buffered:
                    ws.append([_xl(v) for v in r])
                while batch := cur.fetchmany(10_000):
                    for r in batch:
                        ws.append([_xl(v) for v in r])
                backup = atomic_write(out, lambda t: wb.save(t), self.cfg)
            rows_written = n
        res = {"output": self.sb.display(out), "format": fmt, "rows_written": rows_written}
        if backup:
            res["backup_of_previous"] = backup
        return res

    # ------------------------------------------------------------------ public ops
    def query(
        self,
        sql: str | None = None,
        tables: Any = None,
        folder: str | None = None,
        limit: int | None = None,
        offset: int = 0,
        output: str | None = None,
        output_sheet: str | None = None,
        format: str = "json",
    ) -> dict:
        """Run a SELECT over one or many files. Files are tables (auto-registered by name or via tables=).

        Without sql, returns the catalog of table names and columns you can query.
        """
        if not sql or not str(sql).strip():
            return self.catalog(folder)  # type: ignore[attr-defined]
        con = self._connect()
        try:
            registered = self._register(con, sql, tables, folder)
            sql2 = self._check_select(self._rewrite_literals(sql))
            if output:
                return self._write_output(con, sql2, output, None, output_sheet, registered) | {
                    "tables": registered
                }
            return self._result(con, sql2, limit, offset, format, registered)
        finally:
            con.close()

    def table_schema(self, path: str, sheet: str | None = None, preview_rows: int = 5) -> dict:
        con = self._connect()
        try:
            expr, label = self.source_sql(path, sheet)
            try:
                cols = con.execute(f"DESCRIBE SELECT * FROM {expr}").fetchall()
                n = con.execute(f"SELECT count(*) FROM {expr}").fetchone()[0]
                prev = con.execute(f"SELECT * FROM {expr} LIMIT {int(preview_rows)}").fetchall()
            except duckdb.Error as e:
                raise self._db_error(e, {}) from e
            return {
                "source": label,
                "row_count": n,
                "columns": [{"name": c[0], "type": c[1]} for c in cols],
                "preview": [[self._j(v) for v in r] for r in prev],
            }
        finally:
            con.close()

    def _columns(self, con: duckdb.DuckDBPyConnection, expr: str) -> list[str]:
        return [d[0] for d in con.execute(f"SELECT * FROM {expr} LIMIT 0").description]

    @staticmethod
    def _col_or_expr(c: str, cols: list[str]) -> str:
        if c in cols:
            return qident(c)
        low = {x.lower(): x for x in cols}
        if c.lower() in low:
            return qident(low[c.lower()])
        return c  # treat as SQL expression

    def transform(
        self,
        source: str,
        sheet: str | None = None,
        where: str | None = None,
        columns: list[str] | None = None,
        rename: dict[str, str] | None = None,
        order_by: list[str] | str | None = None,
        distinct: bool = False,
        dedupe_on: list[str] | None = None,
        keep: str = "first",
        limit: int | None = None,
        offset: int = 0,
        output: str | None = None,
        output_sheet: str | None = None,
        format: str = "json",
    ) -> dict:
        """Filter / select / rename / sort / dedupe one table without writing SQL by hand."""
        con = self._connect()
        try:
            expr, label = self.source_sql(source, sheet)
            cols = self._columns(con, expr)
            rename = rename or {}
            sel = []
            for c in columns or cols:
                e = self._col_or_expr(c, cols)
                alias = rename.get(c)
                sel.append(f"{e} AS {qident(alias)}" if alias else e)
            if not columns and rename:
                sel = [f"{qident(c)} AS {qident(rename[c])}" if c in rename else qident(c) for c in cols]
            sql = f"SELECT {'DISTINCT ' if distinct else ''}{', '.join(sel)} FROM {expr} AS src"
            if where:
                sql += f" WHERE {where}"
            if dedupe_on:
                part = ", ".join(self._col_or_expr(c, cols) for c in dedupe_on)
                order = "rowid_ DESC" if keep == "last" else "rowid_"
                inner = f"SELECT *, row_number() OVER () AS rowid_ FROM {expr}"
                if where:
                    inner = f"SELECT *, row_number() OVER () AS rowid_ FROM {expr} AS src WHERE {where}"
                sql = (
                    f"SELECT {', '.join(sel)} FROM ({inner}) AS src "
                    f"QUALIFY row_number() OVER (PARTITION BY {part} ORDER BY {order}) = 1"
                )
                if not order_by:
                    order_by = ["rowid_"]
            if order_by:
                ob = [order_by] if isinstance(order_by, str) else order_by
                parts = []
                for o in ob:
                    m = re.match(r"^\s*(.+?)(\s+(?:asc|desc))?\s*$", o, re.I)
                    name, direction = (m.group(1), m.group(2) or "") if m else (o, "")
                    parts.append(self._col_or_expr(name, [*cols, "rowid_"]) + direction)
                sql += " ORDER BY " + ", ".join(parts)
            sql = self._check_select(sql)
            if output:
                return self._write_output(con, sql, output, None, output_sheet) | {"source": label}
            return self._result(con, sql, limit, offset, format, {}) | {
                "source": label,
                "sql": sql.replace(expr, "src"),
            }
        except duckdb.Error as e:
            raise self._db_error(e, {}) from e
        finally:
            con.close()

    def aggregate(
        self,
        source: str,
        metrics: list[str],
        group_by: list[str] | None = None,
        sheet: str | None = None,
        pivot_on: str | None = None,
        where: str | None = None,
        having: str | None = None,
        order_by: list[str] | str | None = None,
        limit: int | None = None,
        output: str | None = None,
        output_sheet: str | None = None,
        format: str = "json",
    ) -> dict:
        """Group-by and pivot tables. metrics like 'amount:sum', 'count', or raw SQL 'sum(qty*price) AS revenue'."""
        if not metrics:
            raise SheetError("metrics is required.", "Example: metrics=['amount:sum', 'amount:avg', 'count']")
        con = self._connect()
        try:
            expr, label = self.source_sql(source, sheet)
            cols = self._columns(con, expr)
            ms = [self._metric(m, cols) for m in metrics]
            gb = [self._col_or_expr(g, cols) for g in (group_by or [])]
            src = f"(SELECT * FROM {expr} AS s{' WHERE ' + where if where else ''})"
            if pivot_on:
                pv = self._col_or_expr(pivot_on, cols)
                # validate user fragments as a plain SELECT first
                self._check_select(
                    f"SELECT {', '.join([*gb, *ms])} FROM {src} AS t"
                    + (" GROUP BY " + ", ".join(gb) if gb else "")
                )
                vals = [
                    r[0]
                    for r in con.execute(
                        f"SELECT DISTINCT {pv} FROM {src} AS t WHERE {pv} IS NOT NULL ORDER BY 1 LIMIT 201"
                    ).fetchall()
                ]
                if len(vals) > 200:
                    raise SheetError(
                        "pivot_on column has more than 200 distinct values.",
                        "Pivot on a column with fewer categories, or filter with where=.",
                    )
                in_list = ", ".join(
                    qstr(str(v))
                    if isinstance(v, str)
                    else str(jsonable(v))
                    if isinstance(v, (int, float))
                    else qstr(str(jsonable(v)))
                    for v in vals
                )
                using = ", ".join(ms)
                sql = (
                    f"SELECT * FROM (PIVOT {src} ON {pv} IN ({in_list or 'NULL'}) USING {using}"
                    f"{' GROUP BY ' + ', '.join(gb) if gb else ''})"
                )
            else:
                sql = f"SELECT {', '.join([*gb, *ms])} FROM {src} AS t"
                if gb:
                    sql += " GROUP BY " + ", ".join(gb)
                if having:
                    sql += f" HAVING {having}"
            if order_by:
                ob = [order_by] if isinstance(order_by, str) else order_by
                sql += " ORDER BY " + ", ".join(ob)
            elif gb:
                sql += " ORDER BY " + ", ".join(gb)
            sql = self._check_select(sql) if not pivot_on else sql
            if output:
                return self._write_output(con, sql, output, None, output_sheet) | {"source": label}
            return self._result(con, sql, limit, 0, format, {}) | {"source": label}
        except duckdb.Error as e:
            raise self._db_error(e, {}) from e
        finally:
            con.close()

    def _metric(self, m: str, cols: list[str]) -> str:
        m = m.strip()
        if m.lower() in ("count", "count(*)", "rows"):
            return 'count(*) AS "count"'
        if ":" in m and "(" not in m:
            col, fn = m.rsplit(":", 1)
            fnl = fn.strip().lower()
            if fnl not in _AGG_FUNCS:
                raise SheetError(f"Unknown aggregate '{fn}'.", "Use: " + ", ".join(sorted(_AGG_FUNCS)))
            c = self._col_or_expr(col.strip(), cols)
            name = f"{fnl}_{col.strip()}"
            if fnl == "count_distinct":
                return f"count(DISTINCT {c}) AS {qident(name)}"
            return f"{_AGG_FUNCS[fnl]}({c}) AS {qident(name)}"
        return m

    def combine(
        self,
        sources: list[Any],
        mode: str = "append",
        on: list[str] | str | None = None,
        how: str = "inner",
        add_source_column: bool = True,
        limit: int | None = None,
        output: str | None = None,
        output_sheet: str | None = None,
        format: str = "json",
    ) -> dict:
        """Append (stack, aligning columns by name) or join several files/sheets."""
        if not sources or len(sources) < 2:
            raise SheetError(
                "combine needs at least 2 sources.",
                "Example: sources=['jan.csv', 'feb.csv', 'mar.xlsx#Data']",
            )
        con = self._connect()
        try:
            exprs = [self.source_sql(s) for s in sources]
            if mode == "append":
                parts = [
                    f"SELECT *{', ' + qstr(lbl) + ' AS _source' if add_source_column else ''} FROM {e}"
                    for e, lbl in exprs
                ]
                sql = " UNION ALL BY NAME ".join(parts)
            elif mode == "join":
                if not on:
                    raise SheetError("join mode needs on=[key columns].", "Example: on=['customer_id']")
                keys = [on] if isinstance(on, str) else list(on)
                how_sql = {
                    "inner": "INNER",
                    "left": "LEFT",
                    "right": "RIGHT",
                    "full": "FULL OUTER",
                    "outer": "FULL OUTER",
                    "semi": "SEMI",
                    "anti": "ANTI",
                }.get(how.lower())
                if not how_sql:
                    raise SheetError(
                        f"Unknown join type '{how}'.", "Use inner, left, right, full, semi or anti."
                    )
                using = ", ".join(qident(k) for k in keys)
                sql = f"SELECT * FROM {exprs[0][0]} AS t1"
                for i, (e, _) in enumerate(exprs[1:], start=2):
                    sql += f" {how_sql} JOIN {e} AS t{i} USING ({using})"
            else:
                raise SheetError(f"Unknown mode '{mode}'.", "Use 'append' or 'join'.")
            sql = self._check_select(sql)
            labels = [lbl for _, lbl in exprs]
            if output:
                return self._write_output(con, sql, output, None, output_sheet) | {"sources": labels}
            return self._result(con, sql, limit, 0, format, {}) | {"sources": labels}
        except duckdb.Error as e:
            raise self._db_error(e, {}) from e
        finally:
            con.close()

    def convert(
        self,
        input: str,
        output: str,
        sheet: str | None = None,
        all_sheets: bool = False,
        format: str | None = None,
        csv_dialect: str = "standard",
    ) -> dict:
        """Convert between csv/tsv/xlsx/parquet/json/jsonl/markdown (inputs also .xls/.ods/.html tables).

        all_sheets=True converts every sheet. csv_dialect: standard | excel (UTF-8 BOM) | google_sheets.
        """
        p = self.sb.resolve(input)
        con = self._connect()
        try:
            if all_sheets and file_kind(p) in SHEETED_KINDS:
                names = sheet_names(p)
                out = self.sb.resolve(output, must_exist=False)
                fmt = output_format(out, format)
                results = []
                for nm in names:
                    expr, _ = self.source_sql(str(p), nm)
                    if fmt == "xlsx":
                        results.append(self._write_output(con, f"SELECT * FROM {expr}", str(out), fmt, nm))
                    else:
                        target = out.with_name(
                            f"{out.stem}_{safe_identifier(nm)}.{out.suffix.lstrip('.') or fmt}"
                        )
                        results.append(
                            self._write_output(
                                con, f"SELECT * FROM {expr}", str(target), fmt, csv_dialect=csv_dialect
                            )
                        )
                return {"input": self.sb.display(p), "outputs": results}
            expr, label = self.source_sql(str(p), sheet)
            res = self._write_output(
                con, f"SELECT * FROM {expr}", output, format, None, csv_dialect=csv_dialect
            )
            if csv_dialect == "google_sheets":
                ncols = len(self._columns(con, expr))
                if res["rows_written"] * ncols > 10_000_000:
                    res["warning"] = (
                        "Google Sheets allows at most 10 million cells per spreadsheet; split the file."
                    )
            return res | {"input": label}
        finally:
            con.close()

    # ------------------------------------------------------------------ profiling
    def profile(
        self, source: str, sheet: str | None = None, columns: list[str] | None = None, top_values: int = 5
    ) -> dict:
        """Column statistics + data-quality issues (nulls, duplicates, numbers-as-text, whitespace, outliers...)."""
        con = self._connect()
        try:
            expr, label = self.source_sql(source, sheet)
            all_cols = self._columns(con, expr)
            cols = [self._col_or_expr(c, all_cols).strip('"') for c in columns] if columns else all_cols
            cols = cols[:100]
            try:
                total = con.execute(f"SELECT count(*) FROM {expr}").fetchone()[0]
                distinct_rows = con.execute(
                    f"SELECT count(*) FROM (SELECT DISTINCT * FROM {expr})"
                ).fetchone()[0]
                summ = con.execute(f"SUMMARIZE SELECT {', '.join(qident(c) for c in cols)} FROM {expr}")
                names = [d[0] for d in summ.description]
                summary = [dict(zip(names, r)) for r in summ.fetchall()]
            except duckdb.Error as e:
                raise self._db_error(e, {}) from e
            issues: list[dict[str, Any]] = []
            dup = total - distinct_rows
            if dup:
                issues.append(
                    {
                        "severity": "warning",
                        "column": None,
                        "issue": "duplicate_rows",
                        "detail": f"{dup} fully duplicated row(s).",
                        "fix": "transform(..., distinct=true)",
                    }
                )
            out_cols = []
            for s in summary:
                c = s["column_name"]
                qc = qident(c)
                typ = str(s["column_type"])
                nullp = float(s.get("null_percentage") or 0)
                info: dict[str, Any] = {
                    "name": c,
                    "type": typ,
                    "non_null": int(s.get("count") or 0) - round(total * nullp / 100),
                    "null_pct": round(nullp, 2),
                    "approx_distinct": s.get("approx_unique"),
                    "min": self._j(s.get("min")),
                    "max": self._j(s.get("max")),
                }
                if s.get("avg") is not None:
                    for k in ("avg", "std", "q25", "q50", "q75"):
                        v = s.get(k)
                        try:
                            info[k] = round(float(v), 6) if v is not None else None
                        except (TypeError, ValueError):
                            info[k] = self._j(v)
                if top_values and total:
                    tv = con.execute(
                        f"SELECT {qc} AS v, count(*) AS n FROM {expr} WHERE {qc} IS NOT NULL GROUP BY 1 ORDER BY 2 DESC, 1 LIMIT {int(top_values)}"
                    ).fetchall()
                    info["top_values"] = [{"value": self._j(v), "count": n} for v, n in tv]
                out_cols.append(info)
                # ---- quality checks
                if nullp >= 100:
                    issues.append(
                        {
                            "severity": "warning",
                            "column": c,
                            "issue": "empty_column",
                            "detail": "Column has no values.",
                        }
                    )
                    continue
                if nullp >= 50:
                    issues.append(
                        {
                            "severity": "info",
                            "column": c,
                            "issue": "mostly_null",
                            "detail": f"{nullp:.1f}% empty.",
                        }
                    )
                if s.get("approx_unique") == 1 and total > 1:
                    issues.append(
                        {
                            "severity": "info",
                            "column": c,
                            "issue": "constant",
                            "detail": "Same value in every row.",
                        }
                    )
                if re.fullmatch(r"column_\d+", c):
                    issues.append(
                        {
                            "severity": "info",
                            "column": c,
                            "issue": "missing_header",
                            "detail": "Header cell was blank.",
                        }
                    )
                if typ == "VARCHAR":
                    nn, numeric, ws_ = con.execute(
                        f"SELECT count({qc}), count(TRY_CAST(replace({qc}, ',', '') AS DOUBLE)), "
                        f"count(*) FILTER (WHERE {qc} <> trim({qc})) FROM {expr}"
                    ).fetchone()
                    if nn and numeric / nn >= 0.8 and numeric < nn:
                        bad = con.execute(
                            f"SELECT DISTINCT {qc} FROM {expr} WHERE {qc} IS NOT NULL AND TRY_CAST(replace({qc}, ',', '') AS DOUBLE) IS NULL LIMIT 5"
                        ).fetchall()
                        issues.append(
                            {
                                "severity": "warning",
                                "column": c,
                                "issue": "numbers_stored_as_text",
                                "detail": f"{numeric}/{nn} values are numeric but {nn - numeric} are not, e.g. "
                                + ", ".join(repr(b[0]) for b in bad),
                                "fix": "Clean the non-numeric values, or use TRY_CAST in queries.",
                            }
                        )
                    elif nn and numeric == nn and nn > 0:
                        issues.append(
                            {
                                "severity": "info",
                                "column": c,
                                "issue": "numeric_text",
                                "detail": "All values look numeric but are stored as text (e.g. thousand separators).",
                            }
                        )
                    if ws_:
                        issues.append(
                            {
                                "severity": "info",
                                "column": c,
                                "issue": "untrimmed_whitespace",
                                "detail": f"{ws_} value(s) have leading/trailing spaces.",
                                "fix": "clean(operations=[{'op': 'trim'}])",
                            }
                        )
                    if s.get("approx_unique") and s["approx_unique"] <= 5000:
                        groups = con.execute(
                            f"SELECT lower(trim({qc})) AS k, list(DISTINCT {qc} ORDER BY {qc}) AS variants FROM {expr} "
                            f"WHERE {qc} IS NOT NULL GROUP BY 1 HAVING count(DISTINCT {qc}) > 1 ORDER BY count(*) DESC LIMIT 3"
                        ).fetchall()
                        if groups:
                            issues.append(
                                {
                                    "severity": "warning",
                                    "column": c,
                                    "issue": "inconsistent_categories",
                                    "detail": "Same value written differently: "
                                    + "; ".join(" / ".join(repr(v) for v in g[1]) for g in groups),
                                    "fix": "clean(operations=[{'op': 'trim'}, {'op': 'case', 'mode': 'title'}])",
                                }
                            )
                elif s.get("q25") is not None and typ not in ("DATE", "TIMESTAMP", "TIME", "BOOLEAN"):
                    try:
                        q1, q3 = float(s["q25"]), float(s["q75"])
                        iqr = q3 - q1
                        if iqr > 0:
                            lo, hi = q1 - 1.5 * iqr, q3 + 1.5 * iqr
                            n_out = con.execute(
                                f"SELECT count(*) FROM {expr} WHERE {qc} < {lo} OR {qc} > {hi}"
                            ).fetchone()[0]
                            if n_out:
                                issues.append(
                                    {
                                        "severity": "info",
                                        "column": c,
                                        "issue": "outliers",
                                        "detail": f"{n_out} value(s) outside [{lo:.4g}, {hi:.4g}] (1.5×IQR).",
                                    }
                                )
                    except (TypeError, ValueError, duckdb.Error):
                        pass
                    try:  # robust z-score (median / MAD) anomalies with row numbers
                        an = con.execute(
                            f"WITH t AS (SELECT row_number() OVER () + 1 AS r, {qc} AS v FROM {expr}), "
                            f"m AS (SELECT median(v) AS med, mad(v) AS mad, count(v) AS n FROM t) "
                            f"SELECT r, v, abs(v - med) / (1.4826 * mad) AS z FROM t, m "
                            f"WHERE n >= 10 AND mad > 0 AND abs(v - med) / (1.4826 * mad) > 3.5 ORDER BY z DESC LIMIT 5"
                        ).fetchall()
                        if an:
                            info["anomalies"] = [
                                {"row": r, "value": self._j(v), "robust_z": round(float(z), 2)}
                                for r, v, z in an
                            ]
                            issues.append(
                                {
                                    "severity": "info",
                                    "column": c,
                                    "issue": "anomalies",
                                    "detail": "Unusual values (robust z > 3.5): "
                                    + ", ".join(f"row {r}: {self._j(v)}" for r, v, _ in an[:3]),
                                }
                            )
                    except (TypeError, ValueError, duckdb.Error):
                        pass
                if s.get("approx_unique") and total and nullp == 0:
                    exact = con.execute(f"SELECT count(DISTINCT {qc}) FROM {expr}").fetchone()[0]
                    if exact == total and total > 1:
                        info["unique"] = True
            return {
                "source": label,
                "rows": total,
                "columns_count": len(all_cols),
                "duplicate_rows": dup,
                "columns": out_cols,
                "issues": issues,
                "quality_score": max(0, 100 - sum(10 if i["severity"] == "warning" else 2 for i in issues)),
            }
        finally:
            con.close()


def _xl(v: Any) -> Any:
    if isinstance(v, (list, dict, tuple)):
        return str(jsonable(v))
    if hasattr(v, "as_tuple"):  # Decimal
        return float(v)
    if hasattr(v, "hex") and isinstance(v, (bytes, bytearray)):
        return v.hex()
    if v.__class__.__name__ == "UUID":
        return str(v)
    if hasattr(v, "tzinfo") and getattr(v, "tzinfo", None) is not None:
        return v.replace(tzinfo=None)
    return v


_ = get_sheet
