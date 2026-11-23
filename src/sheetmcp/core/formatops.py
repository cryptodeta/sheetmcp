"""Formatting, conditional formatting, data validation and charts (openpyxl)."""

from __future__ import annotations

from typing import Any

from openpyxl.chart import AreaChart, BarChart, LineChart, PieChart, Reference, ScatterChart, Series
from openpyxl.formatting.rule import CellIsRule, ColorScaleRule, DataBarRule, FormulaRule, IconSetRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils.cell import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

from .base import Base
from .errors import SheetError
from .fileio import file_kind, get_sheet, load_workbook, save_workbook
from .util import col_index, parse_range

COLORS = {
    "black": "000000",
    "white": "FFFFFF",
    "red": "FF0000",
    "green": "00B050",
    "blue": "0070C0",
    "yellow": "FFFF00",
    "orange": "FFC000",
    "purple": "7030A0",
    "gray": "808080",
    "grey": "808080",
    "lightgray": "D9D9D9",
    "lightgrey": "D9D9D9",
    "lightgreen": "C6EFCE",
    "lightred": "FFC7CE",
    "lightyellow": "FFEB9C",
    "lightblue": "DDEBF7",
    "darkblue": "1F4E78",
    "darkgreen": "006100",
    "darkred": "9C0006",
}
NUMBER_FORMATS = {
    "general": "General",
    "integer": "#,##0",
    "number": "#,##0.00",
    "decimal": "#,##0.00",
    "decimal2": "#,##0.00",
    "percent": "0.00%",
    "percent0": "0%",
    "currency": '"$"#,##0.00',
    "usd": '"$"#,##0.00',
    "eur": '"€"#,##0.00',
    "bdt": '"৳"#,##0.00',
    "gbp": '"£"#,##0.00',
    "inr": '"₹"#,##0.00',
    "date": "yyyy-mm-dd",
    "datetime": "yyyy-mm-dd hh:mm:ss",
    "time": "hh:mm:ss",
    "text": "@",
    "accounting": '_("$"* #,##0.00_)',
    "scientific": "0.00E+00",
}
CF_OPERATORS = {
    ">": "greaterThan",
    ">=": "greaterThanOrEqual",
    "<": "lessThan",
    "<=": "lessThanOrEqual",
    "=": "equal",
    "==": "equal",
    "!=": "notEqual",
    "<>": "notEqual",
    "between": "between",
    "not_between": "notBetween",
}


def color(c: str | None) -> str | None:
    if not c:
        return None
    s = str(c).strip().lower()
    if s in COLORS:
        return COLORS[s]
    s = s.lstrip("#").upper()
    if len(s) == 3:
        s = "".join(ch * 2 for ch in s)
    if len(s) in (6, 8) and all(ch in "0123456789ABCDEF" for ch in s):
        return s
    raise SheetError(
        f"Unknown color '{c}'.", "Use hex like '#FF0000' or a name: " + ", ".join(sorted(COLORS))
    )


class FormatOps(Base):
    def _open_xlsx(self, path: str, action: str, dry_run: bool):
        if not dry_run:
            self._writable(action)
        p = self.sb.resolve(path)
        if file_kind(p) != "excel":
            raise SheetError(
                f"{action} needs an .xlsx file; '{p.name}' is not one.",
                "CSV has no formatting. Convert it first: convert(input='x.csv', output='x.xlsx').",
            )
        return p, load_workbook(p, formulas=True, fresh=True)

    def format_range(
        self,
        path: str,
        range: str | None = None,
        sheet: str | None = None,
        bold: bool | None = None,
        italic: bool | None = None,
        underline: bool | None = None,
        font_size: float | None = None,
        font_color: str | None = None,
        font_name: str | None = None,
        fill_color: str | None = None,
        number_format: str | None = None,
        h_align: str | None = None,
        v_align: str | None = None,
        wrap_text: bool | None = None,
        border: str | None = None,
        border_color: str | None = None,
        column_width: float | dict[str, float] | None = None,
        row_height: float | None = None,
        autofit: bool = False,
        freeze_panes: str | None = None,
        merge: bool | None = None,
        auto_filter: bool | None = None,
        conditional: dict[str, Any] | None = None,
        validation: dict[str, Any] | None = None,
        dry_run: bool = False,
    ) -> dict:
        """Style cells, set widths, freeze panes, merge, filters, conditional formatting and dropdown validation."""
        p, wb = self._open_xlsx(path, "format cells", dry_run)
        ws = get_sheet(wb, sheet)
        done: list[str] = []
        if range:
            c1, r1, c2, r2 = parse_range(range, ws.max_row, ws.max_column)
            ref = f"{get_column_letter(c1)}{r1}:{get_column_letter(c2)}{r2}"
            if (c2 - c1 + 1) * (r2 - r1 + 1) > 2_000_000:
                raise SheetError("Range too large to format.", "Format a smaller range.")
            style_font = any(
                x is not None for x in (bold, italic, underline, font_size, font_color, font_name)
            )
            align = any(x is not None for x in (h_align, v_align, wrap_text))
            nf = NUMBER_FORMATS.get(number_format.lower(), number_format) if number_format else None
            fill = PatternFill("solid", fgColor=color(fill_color)) if fill_color else None
            if fill_color and fill_color.lower() in ("none", "clear"):
                fill = PatternFill(fill_type=None)
            side = None
            if border:
                if border not in ("thin", "medium", "thick", "dashed", "dotted", "double", "hair", "none"):
                    raise SheetError(
                        f"Unknown border '{border}'.",
                        "Use thin, medium, thick, dashed, dotted, double or none.",
                    )
                side = Side(style=None if border == "none" else border, color=color(border_color) or "000000")
            for row in ws.iter_rows(min_row=r1, max_row=r2, min_col=c1, max_col=c2):
                for c in row:
                    if style_font:
                        f = c.font
                        c.font = Font(
                            name=font_name or f.name,
                            size=font_size or f.size,
                            bold=f.bold if bold is None else bold,
                            italic=f.italic if italic is None else italic,
                            underline=("single" if underline else None)
                            if underline is not None
                            else f.underline,
                            color=color(font_color) if font_color else f.color,
                            strike=f.strike,
                            vertAlign=f.vertAlign,
                        )
                    if fill is not None:
                        c.fill = fill
                    if nf:
                        c.number_format = nf
                    if align:
                        a = c.alignment
                        c.alignment = Alignment(
                            horizontal=h_align or a.horizontal,
                            vertical=v_align or a.vertical,
                            wrap_text=a.wrap_text if wrap_text is None else wrap_text,
                            indent=a.indent,
                            text_rotation=a.text_rotation,
                        )
                    if side is not None:
                        c.border = Border(left=side, right=side, top=side, bottom=side)
            for flag, label in (
                (style_font, "font"),
                (fill is not None, "fill"),
                (nf, f"number_format={nf}"),
                (align, "alignment"),
                (side is not None, f"border={border}"),
            ):
                if flag:
                    done.append(label)
            if merge is True:
                ws.merge_cells(ref)
                done.append("merged")
            elif merge is False:
                for m in list(ws.merged_cells.ranges):
                    if m.coord == ref or (
                        m.min_row >= r1 and m.max_row <= r2 and m.min_col >= c1 and m.max_col <= c2
                    ):
                        ws.unmerge_cells(str(m))
                done.append("unmerged")
            if row_height:
                for r in builtins_range(r1, r2 + 1):
                    ws.row_dimensions[r].height = row_height
                done.append("row_height")
            if autofit:
                for c in builtins_range(c1, c2 + 1):
                    letter = get_column_letter(c)
                    width = 0
                    for (v,) in ws.iter_rows(
                        min_col=c, max_col=c, min_row=1, max_row=min(ws.max_row, 5000), values_only=True
                    ):
                        if v is not None:
                            s = str(v)
                            width = max(
                                width, max(len(x) for x in s.split("\n")) if not s.startswith("=") else 10
                            )
                    ws.column_dimensions[letter].width = min(max(width + 2, 8), 80)
                done.append("autofit")
            if auto_filter is True:
                ws.auto_filter.ref = ref
                done.append("auto_filter")
            elif auto_filter is False:
                ws.auto_filter.ref = None
            if conditional:
                done.append(self._conditional(ws, ref, conditional))
            if validation:
                done.append(self._validation(ws, ref, validation))
        elif any(
            x is not None for x in (bold, italic, fill_color, number_format, conditional, validation, merge)
        ):
            raise SheetError("These options need a range.", "Pass range='A1:D1' (or 'B:B' for a column).")
        if column_width is not None:
            if isinstance(column_width, dict):
                for col, w in column_width.items():
                    letter = col if not str(col).isdigit() else get_column_letter(int(col))
                    col_index(letter)
                    ws.column_dimensions[letter.upper()].width = float(w)
            elif range:
                c1, _, c2, _ = parse_range(range, ws.max_row, ws.max_column)
                for c in builtins_range(c1, c2 + 1):
                    ws.column_dimensions[get_column_letter(c)].width = float(column_width)
            else:
                raise SheetError("column_width needs a range or a dict like {'A': 20}.")
            done.append("column_width")
        if freeze_panes is not None:
            ws.freeze_panes = None if freeze_panes in ("", "none", "A1") else freeze_panes.upper()
            done.append(f"freeze_panes={freeze_panes}")
        if not done:
            raise SheetError(
                "No formatting options were given.",
                "e.g. bold=true, fill_color='lightyellow', number_format='currency'.",
            )
        res: dict[str, Any] = {
            "file": self.sb.display(p),
            "sheet": ws.title,
            "range": range,
            "applied": done,
            "dry_run": dry_run,
        }
        if not dry_run:
            res["backup"] = save_workbook(wb, p, self.cfg)
        return res

    def _conditional(self, ws: Any, ref: str, cond: dict[str, Any]) -> str:
        t = (cond.get("type") or "").lower()
        fill = PatternFill(
            "solid",
            start_color=color(cond.get("fill_color", "lightred")),
            end_color=color(cond.get("fill_color", "lightred")),
        )
        font = Font(color=color(cond["font_color"])) if cond.get("font_color") else None
        if t == "color_scale":
            if cond.get("mid_color"):
                rule = ColorScaleRule(
                    start_type="min",
                    start_color=color(cond.get("start_color", "F8696B")),
                    mid_type="percentile",
                    mid_value=50,
                    mid_color=color(cond["mid_color"]),
                    end_type="max",
                    end_color=color(cond.get("end_color", "63BE7B")),
                )
            else:
                rule = ColorScaleRule(
                    start_type="min",
                    start_color=color(cond.get("start_color", "F8696B")),
                    end_type="max",
                    end_color=color(cond.get("end_color", "63BE7B")),
                )
        elif t == "data_bar":
            rule = DataBarRule(
                start_type="min", end_type="max", color=color(cond.get("color", "638EC6")), showValue=True
            )
        elif t == "icon_set":
            rule = IconSetRule(cond.get("icon_style", "3TrafficLights1"), "percent", [0, 33, 67])
        elif t == "cell_is":
            op = CF_OPERATORS.get(str(cond.get("operator", ">")), cond.get("operator"))
            vals = cond.get("values")
            if vals is None:
                vals = [cond.get("value")]
            if not isinstance(vals, list):
                vals = [vals]
            vals = [
                f'"{v}"' if isinstance(v, str) and not str(v).startswith("=") else str(v).lstrip("=")
                for v in vals
            ]
            rule = CellIsRule(operator=op, formula=vals, fill=fill, font=font)
        elif t == "formula":
            if not cond.get("formula"):
                raise SheetError("formula conditional needs 'formula', e.g. '$C2>100'.")
            rule = FormulaRule(formula=[str(cond["formula"]).lstrip("=")], fill=fill, font=font)
        elif t in ("duplicates", "unique"):
            first = ref.split(":")[0].replace("$", "")
            rng = (
                "$"
                + "$".join(_split_cell(ref.split(":")[0]))
                + ":$"
                + "$".join(_split_cell(ref.split(":")[-1]))
            )
            op = ">1" if t == "duplicates" else "=1"
            rule = FormulaRule(formula=[f"COUNTIF({rng},{first}){op}"], fill=fill, font=font)
        else:
            raise SheetError(
                f"Unknown conditional type '{t}'.",
                "Use color_scale, data_bar, icon_set, cell_is, formula, duplicates or unique.",
            )
        ws.conditional_formatting.add(ref, rule)
        return f"conditional:{t}"

    def _validation(self, ws: Any, ref: str, v: dict[str, Any]) -> str:
        t = (v.get("type") or "list").lower()
        if t == "list":
            vals = v.get("values")
            if vals:
                formula = '"' + ",".join(str(x) for x in vals) + '"'
            elif v.get("source"):
                formula = "=" + str(v["source"]).lstrip("=")
            else:
                raise SheetError("list validation needs values=[...] or source='Sheet2!$A$1:$A$9'.")
            dv = DataValidation(type="list", formula1=formula, allow_blank=True)
        elif t in ("whole", "decimal", "date", "textLength", "text_length"):
            dv = DataValidation(
                type="textLength" if t == "text_length" else t,
                operator=v.get("operator", "between"),
                formula1=str(v.get("min", 0)),
                formula2=str(v.get("max", 10**9)) if v.get("max") is not None else None,
            )
        else:
            raise SheetError(
                f"Unknown validation type '{t}'.", "Use list, whole, decimal, date or text_length."
            )
        if v.get("error_message"):
            dv.error = v["error_message"]
            dv.showErrorMessage = True
        ws.add_data_validation(dv)
        dv.add(ref)
        return f"validation:{t}"

    def add_chart(
        self,
        path: str,
        data_range: str,
        chart_type: str = "bar",
        sheet: str | None = None,
        categories_range: str | None = None,
        title: str | None = None,
        x_title: str | None = None,
        y_title: str | None = None,
        anchor: str | None = None,
        target_sheet: str | None = None,
        width: float = 16,
        height: float = 8,
        titles_from_data: bool = True,
        stacked: bool = False,
        dry_run: bool = False,
    ) -> dict:
        """Add a native Excel chart (bar, column, line, pie, scatter, area)."""
        p, wb = self._open_xlsx(path, "add a chart", dry_run)
        ws = get_sheet(wb, sheet)
        ct = chart_type.lower()
        c1, r1, c2, r2 = parse_range(data_range, ws.max_row, ws.max_column)
        if ct in ("bar", "column", "col"):
            ch = BarChart()
            ch.type = "bar" if ct == "bar" else "col"
            if stacked:
                ch.grouping, ch.overlap = "stacked", 100
        elif ct == "line":
            ch = LineChart()
        elif ct == "pie":
            ch = PieChart()
        elif ct == "area":
            ch = AreaChart()
            if stacked:
                ch.grouping = "stacked"
        elif ct == "scatter":
            ch = ScatterChart()
            ch.style = 13
        else:
            raise SheetError(
                f"Unknown chart_type '{chart_type}'.", "Use bar, column, line, pie, scatter or area."
            )
        if ct == "scatter":
            if not categories_range:
                raise SheetError("scatter needs categories_range for the X values.")
            xc1, xr1, xc2, xr2 = parse_range(categories_range, ws.max_row, ws.max_column)
            xs = Reference(ws, min_col=xc1, min_row=xr1, max_col=xc2, max_row=xr2)
            for c in builtins_range(c1, c2 + 1):
                ys = Reference(ws, min_col=c, min_row=r1, max_row=r2)
                s = Series(ys, xs, title_from_data=titles_from_data)
                s.marker.symbol = "circle"
                s.graphicalProperties.line.noFill = True
                ch.series.append(s)
        else:
            ch.add_data(
                Reference(ws, min_col=c1, min_row=r1, max_col=c2, max_row=r2),
                titles_from_data=titles_from_data,
            )
            if categories_range:
                cc1, cr1, cc2, cr2 = parse_range(categories_range, ws.max_row, ws.max_column)
                ch.set_categories(Reference(ws, min_col=cc1, min_row=cr1, max_col=cc2, max_row=cr2))
        if title:
            ch.title = title
        if ct != "pie":
            if x_title:
                ch.x_axis.title = x_title
            if y_title:
                ch.y_axis.title = y_title
            try:  # openpyxl >= 3.1 hides axes by default in some viewers
                ch.x_axis.delete = False
                ch.y_axis.delete = False
            except AttributeError:
                pass
        ch.width, ch.height = width, height
        dest = get_sheet(wb, target_sheet) if target_sheet else ws
        anchor = anchor or f"{get_column_letter(c2 + 2)}{r1}"
        dest.add_chart(ch, anchor)
        res: dict[str, Any] = {
            "file": self.sb.display(p),
            "sheet": dest.title,
            "chart_type": ct,
            "anchor": anchor,
            "series": len(ch.series),
            "dry_run": dry_run,
        }
        if not dry_run:
            res["backup"] = save_workbook(wb, p, self.cfg)
        return res


def _split_cell(ref: str) -> tuple[str, str]:
    ref = ref.replace("$", "")
    i = next((k for k, ch in enumerate(ref) if ch.isdigit()), len(ref))
    return ref[:i], ref[i:]


builtins_range = range
