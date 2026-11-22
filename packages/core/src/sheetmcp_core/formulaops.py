"""Formula intelligence: list, find errors, trace precedents/dependents, evaluate, stats."""

from __future__ import annotations

import re
from collections import Counter
from typing import Any

from openpyxl.utils.cell import range_boundaries

from .base import Base
from .errors import SheetError
from .evaluator import FUNCTIONS, Evaluator, formula_refs, split_ref
from .fileio import get_sheet, load_workbook
from .util import ERROR_VALUES, cell_ref, parse_range

_FUNC_RE = re.compile(r"([A-Z][A-Z0-9.]*)\(")

ERROR_HELP = {
    "#DIV/0!": "Division by zero or by an empty cell. Wrap with IFERROR(x, 0) or check the divisor.",
    "#REF!": "A referenced cell/sheet was deleted. Rewrite the reference.",
    "#NAME?": "Unknown function or named range, or text missing quotes.",
    "#VALUE!": "Wrong type, e.g. text used in arithmetic.",
    "#N/A": "Lookup value not found. Check the key or wrap with IFNA().",
    "#NUM!": "Invalid number (e.g. SQRT of a negative).",
    "#NULL!": "Ranges that do not intersect.",
    "#CIRC!": "Circular reference: the formula depends on itself.",
}


def _formula(v: Any) -> str | None:
    t = getattr(v, "text", None)
    if t:
        return t
    if isinstance(v, str) and v.startswith("=") and len(v) > 1:
        return v
    return None


class FormulaOps(Base):
    def _iter_formulas(self, wb: Any, sheet: str | None):
        sheets = [get_sheet(wb, sheet)] if sheet else wb.worksheets
        for ws in sheets:
            for row in ws.iter_rows():
                for c in row:
                    f = _formula(c.value)
                    if f:
                        yield ws.title, c.coordinate, f

    def formulas(
        self,
        path: str,
        action: str = "list",
        sheet: str | None = None,
        cell: str | None = None,
        formula: str | None = None,
        depth: int = 3,
        offset: int = 0,
        limit: int | None = None,
    ) -> dict:
        """action: list | errors | trace | evaluate | stats."""
        action = (action or "list").lower()
        p = self.sb.resolve(path)
        wbf = load_workbook(p, formulas=True)
        wbv = load_workbook(p, formulas=False)
        if action == "list":
            return self._list(p, wbf, wbv, sheet, offset, limit)
        if action == "errors":
            return self._errors(p, wbf, wbv, sheet, limit)
        if action == "trace":
            if not cell:
                raise SheetError("trace needs cell (e.g. 'D10' or 'Summary!B2').")
            return self._trace(p, wbf, wbv, sheet, cell, depth)
        if action == "evaluate":
            return self._evaluate(p, wbf, wbv, sheet, cell, formula)
        if action == "stats":
            return self._stats(p, wbf)
        raise SheetError(f"Unknown action '{action}'.", "Use list, errors, trace, evaluate or stats.")

    def _list(self, p, wbf, wbv, sheet, offset, limit) -> dict:
        limit = self._limit(limit, 5000)
        items, total = [], 0
        for title, coord, f in self._iter_formulas(wbf, sheet):
            total += 1
            if total <= offset or len(items) >= limit:
                continue
            items.append(
                {
                    "sheet": title,
                    "cell": coord,
                    "formula": f,
                    "cached_value": self._j(wbv[title][coord].value),
                }
            )
        res: dict[str, Any] = {"file": self.sb.display(p), "total_formulas": total, "formulas": items}
        if offset + len(items) < total:
            res["truncated"] = True
            res["next_offset"] = offset + len(items)
        return res

    def _errors(self, p, wbf, wbv, sheet, limit) -> dict:
        limit = self._limit(limit, 5000)
        ev = Evaluator(wbf)
        found: list[dict[str, Any]] = []
        checked = 0
        counts: Counter[str] = Counter()
        sheets = [get_sheet(wbv, sheet)] if sheet else wbv.worksheets
        # literal error values typed into cells
        for ws in sheets:
            for row in ws.iter_rows():
                for c in row:
                    if (
                        isinstance(c.value, str)
                        and c.value in ERROR_VALUES
                        and not _formula(wbf[ws.title][c.coordinate].value)
                    ):
                        counts[c.value] += 1
                        if len(found) < limit:
                            found.append(
                                {
                                    "sheet": ws.title,
                                    "cell": c.coordinate,
                                    "error": c.value,
                                    "source": "literal",
                                    "help": ERROR_HELP.get(c.value, ""),
                                }
                            )
        for title, coord, f in self._iter_formulas(wbf, sheet):
            checked += 1
            cached = wbv[title][coord].value
            err, src = None, "cached"
            if isinstance(cached, str) and cached in ERROR_VALUES:
                err = cached
            elif "#REF!" in f:
                err, src = "#REF!", "formula_text"
            elif cached is None:
                v = ev.evaluate_cell(title, coord)
                if isinstance(v, str) and (v in ERROR_VALUES or v == "#CIRC!"):
                    err, src = v, "evaluated"
            if err:
                counts[err] += 1
                if len(found) < limit:
                    found.append(
                        {
                            "sheet": title,
                            "cell": coord,
                            "formula": f,
                            "error": err,
                            "source": src,
                            "help": ERROR_HELP.get(err, ""),
                        }
                    )
        res: dict[str, Any] = {
            "file": self.sb.display(p),
            "formulas_checked": checked,
            "error_count": sum(counts.values()),
            "by_type": dict(counts),
            "errors": found,
        }
        if ev.unsupported:
            res["note"] = "Not evaluated (unsupported functions): " + ", ".join(sorted(ev.unsupported))
        if not found:
            res["summary"] = "No formula errors found."
        return res

    def _trace(self, p, wbf, wbv, sheet, cell, depth) -> dict:
        ws = get_sheet(wbf, sheet)
        sh, coord = split_ref(cell, ws.title)
        ws = get_sheet(wbf, sh)
        c1, r1, _, _ = parse_range(coord)
        coord = cell_ref(c1, r1)
        depth = max(1, min(int(depth), 10))
        ev = Evaluator(wbf)

        def node(title: str, ref: str, d: int, seen: set[str]) -> dict[str, Any]:
            key = f"{title}!{ref}"
            raw = wbf[title][ref].value if ":" not in ref else None
            f = _formula(raw)
            n: dict[str, Any] = {"ref": key}
            if ":" in ref:
                n["range"] = True
                return n
            n["value"] = self._j(
                wbv[title][ref].value
                if wbv[title][ref].value is not None or not f
                else ev.evaluate_cell(title, ref)
            )
            if f:
                n["formula"] = f
                if d < depth and key not in seen:
                    kids = []
                    for r in formula_refs(f):
                        rs, ra = split_ref(r, title)
                        if rs not in wbf.sheetnames:
                            kids.append({"ref": r, "note": "external/unknown sheet or named range"})
                            continue
                        kids.append(node(rs, ra.upper(), d + 1, seen | {key}))
                    n["precedents"] = kids
            return n

        # dependents: scan all formulas whose refs cover the target
        index: list[tuple[str, str, list[tuple[str, int, int, int, int]]]] = []
        for title, cref, f in self._iter_formulas(wbf, None):
            bounds = []
            for r in formula_refs(f):
                rs, ra = split_ref(r, title)
                try:
                    a, b, c, d_ = range_boundaries(ra.upper())
                except ValueError:
                    continue
                bounds.append((rs.lower(), a or 1, b or 1, c or 16384, d_ or 1048576))
            index.append((title, cref, bounds))

        def dependents_of(title: str, ref: str) -> list[str]:
            cc, rr, _, _ = parse_range(ref)
            out = []
            for t, cref, bounds in index:
                for s, a, b, c, d_ in bounds:
                    if s == title.lower() and a <= cc <= c and b <= rr <= d_:
                        out.append(f"{t}!{cref}")
                        break
            return out

        deps: list[dict[str, Any]] = []
        frontier, seen = [(ws.title, coord)], {f"{ws.title}!{coord}"}
        for level in range(1, depth + 1):
            nxt = []
            for t, r in frontier:
                for dref in dependents_of(t, r):
                    if dref in seen:
                        continue
                    seen.add(dref)
                    dt_, dc = dref.split("!", 1)
                    deps.append({"ref": dref, "level": level, "formula": _formula(wbf[dt_][dc].value)})
                    nxt.append((dt_, dc))
            frontier = nxt
            if len(deps) > 500:
                break
        return {
            "file": self.sb.display(p),
            "cell": f"{ws.title}!{coord}",
            "precedents_tree": node(ws.title, coord, 0, set()),
            "dependents": deps[:500],
        }

    def _evaluate(self, p, wbf, wbv, sheet, cell, formula) -> dict:
        ws = get_sheet(wbf, sheet)
        ev = Evaluator(wbf)
        if formula:
            v = ev.evaluate_formula(formula, ws.title)
            res: dict[str, Any] = {"formula": formula, "sheet": ws.title, "value": self._j(v)}
        elif cell:
            sh, coord = split_ref(cell, ws.title)
            ws = get_sheet(wbf, sh)
            c1, r1, c2, r2 = parse_range(coord)
            if (c2 - c1 + 1) * (r2 - r1 + 1) > 2000:
                raise SheetError(
                    "Range too large to evaluate at once (max 2000 cells).", "Evaluate a smaller range."
                )
            results = []
            for r in range(r1, r2 + 1):
                for c in range(c1, c2 + 1):
                    ref = cell_ref(c, r)
                    f = _formula(ws[ref].value)
                    v = ev.evaluate_cell(ws.title, ref)
                    cached = wbv[ws.title][ref].value
                    item = {"cell": ref, "value": self._j(v)}
                    if f:
                        item["formula"] = f
                        item["cached_value"] = self._j(cached)
                        if cached is not None and not _same(cached, v):
                            item["mismatch"] = True
                    results.append(item)
            res = {"sheet": ws.title, "results": results}
        else:
            raise SheetError("evaluate needs cell='B5' (or a range) or formula='=SUM(A1:A9)'.")
        if ev.unsupported:
            res["unsupported_functions"] = sorted(ev.unsupported)
            res["note"] = (
                "Cells using unsupported functions show #NAME?; their Excel cached values (if any) are still valid."
            )
        return res

    def _stats(self, p, wbf) -> dict:
        per_sheet: Counter[str] = Counter()
        funcs: Counter[str] = Counter()
        cross = 0
        for title, _, f in self._iter_formulas(wbf, None):
            per_sheet[title] += 1
            for fn in _FUNC_RE.findall(f.upper()):
                funcs[fn.replace("_XLFN.", "")] += 1
            if "!" in f:
                cross += 1
        unsupported = sorted(fn for fn in funcs if fn not in FUNCTIONS)
        return {
            "file": self.sb.display(p),
            "total_formulas": sum(per_sheet.values()),
            "per_sheet": dict(per_sheet),
            "functions_used": dict(funcs.most_common()),
            "cross_sheet_formulas": cross,
            "evaluator_unsupported": unsupported,
        }


def _same(a: Any, b: Any) -> bool:
    if isinstance(a, (int, float)) and isinstance(b, (int, float)) and not isinstance(a, bool):
        return abs(a - b) <= 1e-9 * max(1, abs(a))
    if hasattr(a, "isoformat") and hasattr(b, "isoformat"):
        return str(a)[:10] == str(b)[:10]
    return a == b
