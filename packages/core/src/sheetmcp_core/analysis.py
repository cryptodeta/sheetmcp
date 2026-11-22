"""Discovery and comparison: list files, search text across files, diff sheets/files."""

from __future__ import annotations

import datetime as dt
import fnmatch
import re
from pathlib import Path
from typing import Any

import duckdb

from .base import Base
from .errors import SheetError
from .fileio import (
    TABULAR_EXT,
    coerce_text,
    file_kind,
    get_sheet,
    has_formulas,
    load_workbook,
    read_csv_grid,
    xlsx_sheet_names,
)
from .util import cell_ref, qident

SKIP_DIRS = {
    ".git",
    ".venv",
    "venv",
    "node_modules",
    "__pycache__",
    ".sheetmcp_backups",
    ".mypy_cache",
    ".ruff_cache",
}


class AnalysisOps(Base):
    def _walk(self, root: Path, recursive: bool, pattern: str | None):
        it = root.rglob("*") if recursive else root.iterdir()
        for f in it:
            if any(part in SKIP_DIRS or part.startswith(".") for part in f.relative_to(root).parts[:-1]):
                continue
            if not f.is_file() or f.name.startswith((".", "~$")):
                continue
            if f.suffix.lower() not in TABULAR_EXT:
                continue
            if pattern and not fnmatch.fnmatch(f.name.lower(), pattern.lower()):
                continue
            yield f

    def list_files(
        self, path: str = ".", recursive: bool = True, pattern: str | None = None, limit: int = 500
    ) -> dict:
        """Find spreadsheet/data files (xlsx, xlsm, csv, tsv, parquet, json) under a folder."""
        root = self.sb.resolve(path or ".", want="dir")
        files, total = [], 0
        for f in sorted(self._walk(root, recursive, pattern)):
            total += 1
            if len(files) >= limit:
                continue
            st = f.stat()
            item: dict[str, Any] = {
                "path": self.sb.display(f),
                "kind": file_kind(f),
                "size_bytes": st.st_size,
                "modified": dt.datetime.fromtimestamp(st.st_mtime).isoformat(timespec="seconds"),
            }
            if item["kind"] == "excel":
                item["sheets"] = xlsx_sheet_names(f)
            files.append(item)
        res: dict[str, Any] = {"root": str(root), "count": total, "files": files}
        if total > len(files):
            res["truncated"] = True
            res["hint"] = "Narrow down with pattern='*.xlsx' or a subfolder."
        if not total:
            res["hint"] = f"No spreadsheet files found under {root}. Allowed roots: " + ", ".join(
                map(str, self.sb.roots)
            )
        return res

    # ------------------------------------------------------------------ search
    def search(
        self,
        text: str,
        path: str = ".",
        regex: bool = False,
        case_sensitive: bool = False,
        in_formulas: bool = True,
        pattern: str | None = None,
        recursive: bool = True,
        limit: int = 100,
    ) -> dict:
        """Find cells containing text (or a regex) across every spreadsheet in a folder (or one file)."""
        if not text:
            raise SheetError("text is required.")
        flags = 0 if case_sensitive else re.IGNORECASE
        try:
            rx = re.compile(text if regex else re.escape(text), flags)
        except re.error as e:
            raise SheetError(f"Invalid regex: {e}", "Escape special characters or pass regex=false.") from e
        target = self.sb.resolve(path or ".", want="any")
        files = [target] if target.is_file() else list(self._walk(target, recursive, pattern))
        hits: list[dict[str, Any]] = []
        scanned = 0
        total = 0
        for f in sorted(files):
            scanned += 1
            kind = file_kind(f)
            try:
                if kind == "excel":
                    wbv = load_workbook(f, formulas=False, read_only=True, fresh=True)
                    wbf = (
                        load_workbook(f, formulas=True, read_only=True, fresh=True)
                        if in_formulas and has_formulas(f)
                        else None
                    )
                    try:
                        for ws in wbv.worksheets:
                            fit = wbf[ws.title].iter_rows(values_only=True) if wbf is not None else None
                            for r, row in enumerate(ws.iter_rows(values_only=True), start=1):
                                frow = next(fit, None) if fit is not None else None
                                for c, v in enumerate(row, start=1):
                                    fv = frow[c - 1] if frow is not None and c - 1 < len(frow) else None
                                    is_f = isinstance(fv, str) and fv.startswith("=")
                                    if (v is not None and rx.search(str(v))) or (is_f and rx.search(fv)):
                                        total += 1
                                        if len(hits) < limit:
                                            h = {
                                                "file": self.sb.display(f),
                                                "sheet": ws.title,
                                                "cell": cell_ref(c, r),
                                                "value": self._j(v),
                                            }
                                            if is_f:
                                                h["formula"] = fv
                                            hits.append(h)
                    finally:
                        wbv.close()
                        if wbf is not None:
                            wbf.close()
                elif kind in ("csv", "tsv"):
                    for r, row in enumerate(read_csv_grid(f), start=1):
                        for c, v in enumerate(row, start=1):
                            if v and rx.search(v):
                                total += 1
                                if len(hits) < limit:
                                    hits.append(
                                        {
                                            "file": self.sb.display(f),
                                            "cell": cell_ref(c, r),
                                            "row": r,
                                            "value": self._j(coerce_text(v)),
                                        }
                                    )
                else:
                    from .sqlops import SqlOps

                    assert isinstance(self, SqlOps)
                    con = self._connect()
                    try:
                        expr, _ = self.source_sql(str(f))
                        cur = con.execute(f"SELECT * FROM {expr}")
                        cols = [d[0] for d in cur.description]
                        r = 0
                        while batch := cur.fetchmany(5000):
                            for row in batch:
                                r += 1
                                for ci, v in enumerate(row):
                                    if v is not None and rx.search(str(v)):
                                        total += 1
                                        if len(hits) < limit:
                                            hits.append(
                                                {
                                                    "file": self.sb.display(f),
                                                    "row": r,
                                                    "column": cols[ci],
                                                    "value": self._j(v),
                                                }
                                            )
                    finally:
                        con.close()
            except SheetError:
                continue
        res: dict[str, Any] = {
            "query": text,
            "files_scanned": scanned,
            "total_matches": total,
            "matches": hits,
        }
        if total > len(hits):
            res["truncated"] = True
        return res

    # ------------------------------------------------------------------ diff
    def diff(
        self,
        file_a: str,
        file_b: str | None = None,
        sheet_a: str | None = None,
        sheet_b: str | None = None,
        key: list[str] | str | None = None,
        limit: int = 200,
    ) -> dict:
        """Compare two sheets/files. With key=[cols]: row-level diff (added/removed/changed). Else cell-by-cell."""
        file_b = file_b or file_a
        pa, pb = self.sb.resolve(file_a), self.sb.resolve(file_b)
        if key:
            return self._diff_keyed(
                pa, pb, sheet_a, sheet_b, [key] if isinstance(key, str) else list(key), limit
            )
        ka, kb = file_kind(pa), file_kind(pb)
        if ka == kb == "excel" and sheet_a is None and sheet_b is None and pa != pb:
            wa = load_workbook(pa, formulas=True)
            wb_ = load_workbook(pb, formulas=True)
            common = [s for s in wa.sheetnames if s in wb_.sheetnames]
            res: dict[str, Any] = {
                "file_a": self.sb.display(pa),
                "file_b": self.sb.display(pb),
                "sheets_only_in_a": [s for s in wa.sheetnames if s not in wb_.sheetnames],
                "sheets_only_in_b": [s for s in wb_.sheetnames if s not in wa.sheetnames],
                "sheets": [],
            }
            budget = limit
            for s in common:
                d = self._diff_cells(pa, pb, s, s, budget)
                budget = max(0, budget - len(d["differences"]))
                res["sheets"].append(d)
            res["total_differences"] = sum(d["difference_count"] for d in res["sheets"])
            res["identical"] = (
                res["total_differences"] == 0 and not res["sheets_only_in_a"] and not res["sheets_only_in_b"]
            )
            return res
        d = self._diff_cells(pa, pb, sheet_a, sheet_b, limit)
        d.update(
            {
                "file_a": self.sb.display(pa),
                "file_b": self.sb.display(pb),
                "identical": d["difference_count"] == 0,
            }
        )
        return d

    def _grid(self, p: Path, sheet: str | None) -> tuple[str, dict[tuple[int, int], tuple[Any, Any]]]:
        kind = file_kind(p)
        cells: dict[tuple[int, int], tuple[Any, Any]] = {}
        if kind == "excel":
            wbf = load_workbook(p, formulas=True)
            wbv = load_workbook(p, formulas=False)
            wsf = get_sheet(wbf, sheet)
            wsv = wbv[wsf.title]
            for rf, rv in zip(wsf.iter_rows(), wsv.iter_rows()):
                for cf, cv in zip(rf, rv):
                    if cf.value is not None or cv.value is not None:
                        f = cf.value if isinstance(cf.value, str) and cf.value.startswith("=") else None
                        cells[(cf.row, cf.column)] = (cv.value, f)
            return wsf.title, cells
        if kind in ("csv", "tsv"):
            for r, row in enumerate(read_csv_grid(p), start=1):
                for c, v in enumerate(row, start=1):
                    if v != "":
                        cells[(r, c)] = (coerce_text(v), None)
            return p.name, cells
        raise SheetError(
            f"Cell diff supports .xlsx/.csv, not {kind}.", "Pass key=[...] for a row diff of any format."
        )

    def _diff_cells(self, pa: Path, pb: Path, sa: str | None, sb_: str | None, limit: int) -> dict:
        ta, ga = self._grid(pa, sa)
        tb, gb = self._grid(pb, sb_)
        diffs = []
        count = 0
        for k in sorted(set(ga) | set(gb)):
            va, fa = ga.get(k, (None, None))
            vb, fb = gb.get(k, (None, None))
            if _eq(va, vb) and fa == fb:
                continue
            count += 1
            if len(diffs) < limit:
                d: dict[str, Any] = {"cell": cell_ref(k[1], k[0]), "a": self._j(va), "b": self._j(vb)}
                if fa or fb:
                    d["formula_a"], d["formula_b"] = fa, fb
                d["change"] = "added" if k not in ga else "removed" if k not in gb else "modified"
                diffs.append(d)
        return {
            "sheet_a": ta,
            "sheet_b": tb,
            "difference_count": count,
            "differences": diffs,
            **({"truncated": True} if count > len(diffs) else {}),
        }

    def _diff_keyed(
        self, pa: Path, pb: Path, sa: str | None, sb_: str | None, keys: list[str], limit: int
    ) -> dict:
        from .sqlops import SqlOps

        assert isinstance(self, SqlOps)
        con = self._connect()
        try:
            ea, la = self.source_sql(str(pa), sa)
            eb, lb = self.source_sql(str(pb), sb_)
            con.execute(f"CREATE VIEW a AS SELECT * FROM {ea}")
            con.execute(f"CREATE VIEW b AS SELECT * FROM {eb}")
            ca = [d[0] for d in con.execute("SELECT * FROM a LIMIT 0").description]
            cb = [d[0] for d in con.execute("SELECT * FROM b LIMIT 0").description]
            missing = [k for k in keys if k not in ca or k not in cb]
            if missing:
                raise SheetError(
                    f"Key column(s) {missing} not found in both files.",
                    f"Columns in A: {ca[:30]}; in B: {cb[:30]}",
                )
            common = [c for c in ca if c in cb and c not in keys]
            on = " AND ".join(f"a.{qident(k)} IS NOT DISTINCT FROM b.{qident(k)}" for k in keys)
            kcols = ", ".join(f"a.{qident(k)}" for k in keys)
            kcols_b = ", ".join(f"b.{qident(k)}" for k in keys)
            added = con.execute(f"SELECT b.* FROM b ANTI JOIN a ON {on} LIMIT {limit}").fetchall()
            n_added = con.execute(f"SELECT count(*) FROM b ANTI JOIN a ON {on}").fetchone()[0]
            removed = con.execute(f"SELECT a.* FROM a ANTI JOIN b ON {on} LIMIT {limit}").fetchall()
            n_removed = con.execute(f"SELECT count(*) FROM a ANTI JOIN b ON {on}").fetchone()[0]
            dup_a = con.execute(
                f"SELECT count(*) - count(DISTINCT ({', '.join(qident(k) for k in keys)})) FROM a"
            ).fetchone()[0]
            changed: list[dict[str, Any]] = []
            n_changed = 0
            if common:
                cond = " OR ".join(
                    f"a.{qident(c)}::VARCHAR IS DISTINCT FROM b.{qident(c)}::VARCHAR" for c in common
                )
                sel = ", ".join([kcols] + [f"a.{qident(c)}, b.{qident(c)}" for c in common])
                n_changed = con.execute(f"SELECT count(*) FROM a JOIN b ON {on} WHERE {cond}").fetchone()[0]
                for row in con.execute(
                    f"SELECT {sel} FROM a JOIN b ON {on} WHERE {cond} LIMIT {limit}"
                ).fetchall():
                    kv = row[: len(keys)]
                    rest = row[len(keys) :]
                    ch = {}
                    for i, c in enumerate(common):
                        x, y = rest[2 * i], rest[2 * i + 1]
                        if str(x) != str(y):
                            ch[c] = {"a": self._j(x), "b": self._j(y)}
                    changed.append({"key": dict(zip(keys, [self._j(v) for v in kv])), "changes": ch})
            _ = kcols_b
            return {
                "a": la,
                "b": lb,
                "key": keys,
                "columns_only_in_a": [c for c in ca if c not in cb],
                "columns_only_in_b": [c for c in cb if c not in ca],
                "added_count": n_added,
                "removed_count": n_removed,
                "changed_count": n_changed,
                "added": [dict(zip(cb, [self._j(v) for v in r])) for r in added],
                "removed": [dict(zip(ca, [self._j(v) for v in r])) for r in removed],
                "changed": changed,
                "identical": n_added == n_removed == n_changed == 0 and ca == cb,
                **({"warning": f"{dup_a} duplicate key(s) in A; matches may be ambiguous."} if dup_a else {}),
            }
        except duckdb.Error as e:
            raise self._db_error(e, {}) from e
        finally:
            con.close()


def _eq(a: Any, b: Any) -> bool:
    if (
        isinstance(a, (int, float))
        and isinstance(b, (int, float))
        and not isinstance(a, bool)
        and not isinstance(b, bool)
    ):
        return abs(a - b) <= 1e-9 * max(1.0, abs(a), abs(b))
    return a == b
