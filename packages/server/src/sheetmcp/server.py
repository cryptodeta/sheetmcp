"""MCP server exposing the sheetmcp_core engine as a compact set of tools, resources and prompts."""

from __future__ import annotations

import functools
import json
import logging
from typing import Annotated, Any, Literal
from urllib.parse import quote, unquote

from mcp.server.fastmcp import FastMCP
from mcp.server.fastmcp.exceptions import ToolError
from mcp.types import ToolAnnotations
from pydantic import Field

from sheetmcp_core import Config, SheetError, Sheets, __version__

log = logging.getLogger("sheetmcp")

INSTRUCTIONS = """SheetMCP works with Excel (.xlsx/.xlsm, plus read-only .xls/.ods/.html tables), CSV/TSV,
Parquet and JSON files. Paths are relative to the first allowed root.
Workflow tips:
1. Start with quick_look (a folder or a file): structure, columns, stats, data issues and next steps in one call.
2. read_range shows values AND formulas side by side; use offset for paging.
3. For questions use query (DuckDB SQL). query() with no sql lists every table name. Files in the root are
   tables by name: sales.csv -> sales, book.xlsx -> book (first sheet) or book__<sheet>; or FROM 'dir/f.xlsx#Sheet'.
4. clean fixes messy data (trim, dedupe, to_number, to_date, fill_down, split...); preview first, then output=.
5. build_report makes a formatted Excel report (summary + totals + chart + data) in one call.
6. Writes keep formatting/formulas and make a backup: backups(action='undo', path=...) reverts the last change.
   Use dry_run=true to preview. batch runs many steps in one call.
7. Outputs are row-limited; pass output='file.xlsx' to save full results instead of printing them."""

RO = ToolAnnotations(readOnlyHint=True, destructiveHint=False, idempotentHint=True, openWorldHint=False)
WRITE = ToolAnnotations(readOnlyHint=False, destructiveHint=False, idempotentHint=False, openWorldHint=False)
DESTRUCTIVE = ToolAnnotations(
    readOnlyHint=False, destructiveHint=True, idempotentHint=False, openWorldHint=False
)

Path_ = Annotated[str, Field(description="File path, relative to the root folder (or absolute inside it).")]
Sheet_ = Annotated[str | None, Field(description="Sheet name (or 0-based index). Default: first sheet.")]
DryRun = Annotated[bool, Field(description="Preview the change without saving.")]
Limit = Annotated[
    int | None, Field(description="Max rows to return (default from server config, usually 100).")
]
Output = Annotated[
    str | None,
    Field(
        description="Save the FULL result to this file (.xlsx/.csv/.parquet/.json/.md) instead of returning rows."
    ),
]
OutSheet = Annotated[
    str | None,
    Field(
        description="When output is an existing .xlsx: write into this sheet (replaced if it exists), keeping other sheets."
    ),
]
Fmt = Annotated[
    Literal["json", "records", "markdown"],
    Field(description="json = rows as arrays, records = list of objects, markdown = a table."),
]
Source = Annotated[
    str, Field(description="Data file; for a specific sheet use 'book.xlsx#Sheet1' or the sheet parameter.")
]


def _safe(fn):
    @functools.wraps(fn)
    def wrapper(*args: Any, **kwargs: Any) -> Any:
        try:
            return fn(*args, **kwargs)
        except SheetError as e:
            raise ToolError(str(e)) from e
        except ToolError:
            raise
        except Exception as e:  # pragma: no cover - last-resort safety net
            log.exception("unexpected error in %s", fn.__name__)
            raise ToolError(
                f"Unexpected {type(e).__name__}: {e}. If this looks like a bug, please report it."
            ) from e

    return wrapper


def build_server(config: Config | None = None, host: str = "127.0.0.1", port: int = 8000) -> FastMCP:
    cfg = config or Config.from_env()
    S = Sheets(cfg)
    mcp = FastMCP("sheetmcp", instructions=INSTRUCTIONS, host=host, port=port)

    def tool(annotations: ToolAnnotations, write: bool = False):
        def deco(fn):
            if write and cfg.read_only:
                return fn  # hide write tools entirely in read-only mode
            return mcp.tool(annotations=annotations)(_safe(fn))

        return deco

    # ================================================================= discovery
    @tool(RO)
    def server_info() -> dict:
        """Show SheetMCP's configuration: allowed folders, read-only mode, row limits, supported formats and usage tips."""
        return S.info()

    @tool(RO)
    def list_files(
        path: Annotated[str, Field(description="Folder to scan, relative to the root.")] = ".",
        recursive: Annotated[bool, Field(description="Include subfolders.")] = True,
        pattern: Annotated[
            str | None, Field(description="Filename glob, e.g. '*.xlsx' or 'sales_*'.")
        ] = None,
        limit: Annotated[int, Field(description="Max files to list.")] = 500,
    ) -> dict:
        """List spreadsheet/data files (xlsx, xlsm, xls, ods, html, csv, tsv, parquet, json/jsonl) in a folder, with
        size, date and sheet names. For columns and stats too, use quick_look."""
        return S.list_files(path, recursive, pattern, limit)

    @tool(RO)
    def describe(
        path: Path_,
        sheet: Sheet_ = None,
        preview_rows: Annotated[int, Field(description="Sample rows to include per sheet (0-50).")] = 5,
        header_row: Annotated[
            int | None, Field(description="Row number of the header if it is not the first non-empty row.")
        ] = None,
    ) -> dict:
        """Understand a file before working with it: sheets, dimensions, Excel tables, named ranges, merged cells,
        formula counts, charts, and per-column inferred types with a data preview. Works for xlsx/csv/parquet/json."""
        return S.describe(path, sheet, preview_rows, header_row)

    # ================================================================= read / write
    @tool(RO)
    def read_range(
        path: Path_,
        sheet: Sheet_ = None,
        range: Annotated[
            str | None, Field(description="A1 range like 'A1:F50', 'B:D', '5:20'. Default: the used range.")
        ] = None,
        include_formulas: Annotated[
            bool, Field(description="Also return formulas ({cell: '=...'}) next to values.")
        ] = True,
        evaluate: Annotated[
            Literal["auto", "always", "never"],
            Field(
                description="auto: compute formulas that have no cached value (e.g. files saved by scripts); always: recompute all; never: raw cached values."
            ),
        ] = "auto",
        offset: Annotated[
            int, Field(description="Skip this many rows from the start of the range (paging).")
        ] = 0,
        limit: Limit = None,
        format: Annotated[
            Literal["json", "markdown"], Field(description="json grid or a markdown table.")
        ] = "json",
    ) -> dict:
        """Read cells from an xlsx/csv sheet. Returns a grid of VALUES plus a map of FORMULAS for formula cells,
        so you see both what Excel shows and how it was computed. Paged: follow next_offset when truncated."""
        return S.read_range(path, sheet, range, include_formulas, evaluate, offset, limit, format)

    @tool(WRITE, write=True)
    def write_range(
        path: Path_,
        sheet: Sheet_ = None,
        start_cell: Annotated[str, Field(description="Top-left cell for `values`.")] = "A1",
        values: Annotated[
            list[list[Any]] | None,
            Field(
                description="2D list of rows. Strings starting with '=' become formulas, e.g. '=SUM(B2:B9)'."
            ),
        ] = None,
        cells: Annotated[
            dict[str, Any] | None, Field(description="Individual cells, e.g. {'B2': 42, 'C2': '=B2*1.15'}.")
        ] = None,
        append: Annotated[
            bool,
            Field(
                description="Append `values` as new rows after the last used row (ignores start_cell row)."
            ),
        ] = False,
        parse_dates: Annotated[
            bool, Field(description="Turn ISO strings like '2024-05-01' into real Excel dates.")
        ] = True,
        create: Annotated[bool, Field(description="Create the file/sheet if missing.")] = True,
        dry_run: DryRun = False,
        fill_formula: Annotated[
            str | None,
            Field(
                description="Formula for the FIRST cell of fill_range, e.g. '=B2*C2'; copied down/right with "
                "relative references adjusted (like dragging the fill handle)."
            ),
        ] = None,
        fill_range: Annotated[
            str | None,
            Field(
                description="Where fill_formula goes, e.g. 'D2:D100', or 'D2:D' / 'D2' = down to the last used row."
            ),
        ] = None,
    ) -> dict:
        """Write values and formulas into an xlsx (or csv), or fill a formula down a column. Only cell values change;
        existing formatting, other formulas, charts and sheets are preserved. A backup is made automatically
        (undo with backups). Appending rows to a CSV streams, so it is fast even for huge files."""
        return S.write_range(
            path,
            sheet,
            start_cell,
            values,
            cells,
            append,
            parse_dates,
            create,
            dry_run,
            fill_formula,
            fill_range,
        )

    @tool(DESTRUCTIVE, write=True)
    def manage_sheets(
        path: Path_,
        action: Annotated[
            Literal["create_workbook", "list", "add", "rename", "delete", "copy", "move", "hide", "unhide"],
            Field(description="What to do."),
        ],
        sheet: Annotated[
            str | None, Field(description="Target sheet (for rename/delete/copy/move/hide).")
        ] = None,
        new_name: Annotated[str | None, Field(description="New name (rename/add/copy).")] = None,
        position: Annotated[int | None, Field(description="0-based position (add/move).")] = None,
        sheets: Annotated[list[str] | None, Field(description="Sheet names for create_workbook.")] = None,
        dry_run: DryRun = False,
    ) -> dict:
        """Create workbooks and add/rename/delete/copy/move/hide sheets. Renaming updates formulas that reference the sheet."""
        return S.manage_sheets(path, action, sheet, new_name, position, sheets, dry_run)

    @tool(DESTRUCTIVE, write=True)
    def insert_delete(
        path: Path_,
        action: Annotated[
            Literal["insert_rows", "delete_rows", "insert_cols", "delete_cols"],
            Field(description="Operation."),
        ],
        index: Annotated[
            int | str,
            Field(description="Row number (1-based) or column letter/number where to insert/delete."),
        ],
        count: Annotated[int, Field(description="How many rows/columns.")] = 1,
        sheet: Sheet_ = None,
        dry_run: DryRun = False,
    ) -> dict:
        """Insert or delete rows/columns. Unlike plain openpyxl, formulas across ALL sheets, named ranges, tables and
        merged cells are re-pointed (references into deleted cells become #REF! like Excel)."""
        return S.insert_delete(path, action, index, count, sheet, dry_run)

    # ================================================================= SQL & data ops
    @tool(RO)
    def query(
        sql: Annotated[
            str | None,
            Field(
                description="One DuckDB SELECT (leave empty to list every table name and its columns). Tables: files in the root by name (sales.csv -> sales; "
                "book.xlsx -> book / book__sheetname), tables= aliases, or FROM 'dir/file.xlsx#Sheet'."
            ),
        ] = None,
        tables: Annotated[
            dict[str, str] | list[str] | None,
            Field(
                description="Optional aliases: {'s': 'data/sales.xlsx#2024', 'c': 'crm/customers.csv'}. A list auto-names by file."
            ),
        ] = None,
        folder: Annotated[
            str | None,
            Field(description="Auto-register every file in this folder by name (default: the root)."),
        ] = None,
        limit: Limit = None,
        offset: Annotated[int, Field(description="Row offset for paging.")] = 0,
        output: Output = None,
        output_sheet: OutSheet = None,
        format: Fmt = "json",
    ) -> dict:
        """Answer questions with SQL across one or MANY spreadsheet files at once (joins, aggregates, window
        functions, CTEs, PIVOT) using DuckDB. Read-only SELECT. Call without sql to see which tables exist.
        Use `output` (+ output_sheet) to save results to a new file or a new sheet."""
        return S.query(sql, tables, folder, limit, offset, output, output_sheet, format)

    @tool(RO)
    def transform(
        source: Source,
        sheet: Sheet_ = None,
        where: Annotated[
            str | None, Field(description="SQL filter, e.g. \"region = 'EU' AND amount > 100\".")
        ] = None,
        columns: Annotated[
            list[str] | None,
            Field(description="Columns or expressions to keep, e.g. ['name', 'price*qty AS total']."),
        ] = None,
        rename: Annotated[dict[str, str] | None, Field(description="Rename columns: {'old': 'new'}.")] = None,
        order_by: Annotated[
            list[str] | None, Field(description="Sort, e.g. ['amount DESC', 'name'].")
        ] = None,
        distinct: Annotated[bool, Field(description="Drop fully duplicated rows.")] = False,
        dedupe_on: Annotated[
            list[str] | None, Field(description="Keep one row per these key columns.")
        ] = None,
        keep: Annotated[
            Literal["first", "last"], Field(description="Which duplicate to keep with dedupe_on.")
        ] = "first",
        limit: Limit = None,
        offset: Annotated[int, Field(description="Row offset for paging.")] = 0,
        output: Output = None,
        output_sheet: OutSheet = None,
        format: Fmt = "json",
    ) -> dict:
        """Filter, select, rename, sort and de-duplicate one table without writing SQL. Save with `output`."""
        return S.transform(
            source,
            sheet,
            where,
            columns,
            rename,
            order_by,
            distinct,
            dedupe_on,
            keep,
            limit,
            offset,
            output,
            output_sheet,
            format,
        )

    @tool(RO)
    def aggregate(
        source: Source,
        metrics: Annotated[
            list[str],
            Field(
                description="e.g. ['amount:sum', 'amount:avg', 'customer:count_distinct', 'count'] "
                "or SQL like 'sum(qty*price) AS revenue'. Funcs: sum avg min max count "
                "count_distinct median std var first last."
            ),
        ],
        group_by: Annotated[list[str] | None, Field(description="Columns to group by.")] = None,
        sheet: Sheet_ = None,
        pivot_on: Annotated[
            str | None,
            Field(description="Make a pivot table: one output column per distinct value of this column."),
        ] = None,
        where: Annotated[str | None, Field(description="SQL filter applied before grouping.")] = None,
        having: Annotated[
            str | None, Field(description="SQL filter on groups, e.g. 'sum_amount > 1000'.")
        ] = None,
        order_by: Annotated[list[str] | None, Field(description="Sort results.")] = None,
        limit: Limit = None,
        output: Output = None,
        output_sheet: OutSheet = None,
        format: Fmt = "json",
    ) -> dict:
        """Group-by summaries and pivot tables (like Excel PivotTables) over xlsx/csv/parquet/json."""
        return S.aggregate(
            source,
            metrics,
            group_by,
            sheet,
            pivot_on,
            where,
            having,
            order_by,
            limit,
            output,
            output_sheet,
            format,
        )

    @tool(RO)
    def combine(
        sources: Annotated[
            list[str], Field(description="2+ files/sheets, e.g. ['jan.csv', 'feb.csv', 'q1.xlsx#March'].")
        ],
        mode: Annotated[
            Literal["append", "join"],
            Field(description="append = stack rows (columns matched by name); join = merge on keys."),
        ] = "append",
        on: Annotated[list[str] | None, Field(description="Key columns for join mode.")] = None,
        how: Annotated[
            Literal["inner", "left", "right", "full", "semi", "anti"], Field(description="Join type.")
        ] = "inner",
        add_source_column: Annotated[
            bool, Field(description="append: add a _source column with the origin file.")
        ] = True,
        limit: Limit = None,
        output: Output = None,
        output_sheet: OutSheet = None,
        format: Fmt = "json",
    ) -> dict:
        """Merge many files into one: append/concatenate (even with different column orders) or join on keys."""
        return S.combine(sources, mode, on, how, add_source_column, limit, output, output_sheet, format)

    @tool(WRITE, write=True)
    def convert(
        input: Annotated[
            str, Field(description="Source file (xlsx/xls/ods/html/csv/tsv/parquet/json/jsonl).")
        ],
        output: Annotated[
            str,
            Field(
                description="Target file; format from extension (.xlsx .csv .tsv .parquet .json .jsonl .md)."
            ),
        ],
        sheet: Sheet_ = None,
        all_sheets: Annotated[
            bool,
            Field(description="xlsx input: convert every sheet (one file per sheet, or sheets of one xlsx)."),
        ] = False,
        csv_dialect: Annotated[
            Literal["standard", "excel", "google_sheets"],
            Field(
                description="CSV output flavour: excel adds a UTF-8 BOM (accents open correctly in Excel); "
                "google_sheets uses ISO dates for clean import."
            ),
        ] = "standard",
    ) -> dict:
        """Convert between Excel, legacy .xls/.ods, HTML tables, CSV, TSV, Parquet, JSON, JSON Lines and Markdown.
        Fast (DuckDB) and streams big files. Also exports Excel- or Google-Sheets-friendly CSV."""
        return S.convert(input, output, sheet, all_sheets, None, csv_dialect)

    # ================================================================= analysis
    @tool(RO)
    def profile(
        source: Source,
        sheet: Sheet_ = None,
        columns: Annotated[
            list[str] | None, Field(description="Only these columns (default all, max 100).")
        ] = None,
        top_values: Annotated[int, Field(description="Most frequent values to show per column.")] = 5,
    ) -> dict:
        """Data profile + quality report: per-column type, nulls, distinct, min/max/mean/quartiles, top values, and
        issues like duplicate rows, numbers stored as text, stray whitespace, empty/constant columns, outliers,
        inconsistent category spellings and anomalies (with row numbers)."""
        return S.profile(source, sheet, columns, top_values)

    @tool(RO)
    def formulas(
        path: Path_,
        action: Annotated[
            Literal["list", "errors", "trace", "evaluate", "stats"],
            Field(
                description="list all formulas | errors: find #REF!/#DIV/0!/#N/A/circular | trace a cell's precedents & dependents | "
                "evaluate a cell/range or an ad-hoc formula | stats: functions used"
            ),
        ] = "list",
        sheet: Sheet_ = None,
        cell: Annotated[
            str | None, Field(description="Cell or range for trace/evaluate, e.g. 'D10' or 'Summary!B2'.")
        ] = None,
        formula: Annotated[
            str | None,
            Field(
                description="Ad-hoc formula to evaluate against the workbook, e.g. '=SUMIFS(C:C,A:A,\"EU\")'."
            ),
        ] = None,
        depth: Annotated[int, Field(description="Trace depth (1-10).")] = 3,
        offset: Annotated[int, Field(description="Paging offset for list.")] = 0,
        limit: Limit = None,
    ) -> dict:
        """Formula intelligence for xlsx: audit errors, trace dependencies, and compute formulas with the built-in
        evaluator (~100 Excel functions incl. SUMIFS, XLOOKUP, INDEX/MATCH) — no Excel needed, works on Linux/macOS/Windows."""
        return S.formulas(path, action, sheet, cell, formula, depth, offset, limit)

    @tool(RO)
    def search(
        text: Annotated[str, Field(description="Text (or regex) to find.")],
        path: Annotated[str, Field(description="Folder or single file to search.")] = ".",
        regex: Annotated[bool, Field(description="Treat text as a regular expression.")] = False,
        case_sensitive: bool = False,
        in_formulas: Annotated[bool, Field(description="Also match inside formulas.")] = True,
        pattern: Annotated[
            str | None, Field(description="Only files matching this glob, e.g. '*.xlsx'.")
        ] = None,
        limit: Annotated[int, Field(description="Max matches.")] = 100,
    ) -> dict:
        """Search for text across every sheet of every spreadsheet in a folder. Returns file, sheet, cell and value."""
        return S.search(text, path, regex, case_sensitive, in_formulas, pattern, True, limit)

    @tool(RO)
    def diff(
        file_a: Annotated[str, Field(description="First file.")],
        file_b: Annotated[
            str | None, Field(description="Second file (omit to compare two sheets of file_a).")
        ] = None,
        sheet_a: Sheet_ = None,
        sheet_b: Sheet_ = None,
        key: Annotated[
            list[str] | None,
            Field(
                description="Key column(s) for a row-level diff (added/removed/changed rows). "
                "Without a key: cell-by-cell diff incl. formulas."
            ),
        ] = None,
        limit: Annotated[int, Field(description="Max differences to list.")] = 200,
    ) -> dict:
        """Compare two versions of a sheet or file: cell-by-cell (values + formulas) or row-level by key columns."""
        return S.diff(file_a, file_b, sheet_a, sheet_b, key, limit)

    # ================================================================= formatting & charts
    @tool(WRITE, write=True)
    def format_range(
        path: Path_,
        range: Annotated[str | None, Field(description="Cells to format, e.g. 'A1:F1' or 'C:C'.")] = None,
        sheet: Sheet_ = None,
        bold: bool | None = None,
        italic: bool | None = None,
        underline: bool | None = None,
        font_size: float | None = None,
        font_color: Annotated[
            str | None, Field(description="Hex '#RRGGBB' or name (red, darkblue, ...).")
        ] = None,
        font_name: str | None = None,
        fill_color: Annotated[
            str | None,
            Field(description="Background: hex or name (lightyellow, lightgreen, ...); 'none' clears."),
        ] = None,
        number_format: Annotated[
            str | None,
            Field(
                description="Preset (currency, usd, eur, bdt, percent, integer, decimal, date, datetime, text) or Excel code like '#,##0.0'."
            ),
        ] = None,
        h_align: Annotated[
            Literal["left", "center", "right", "justify"] | None, Field(description="Horizontal alignment.")
        ] = None,
        v_align: Annotated[
            Literal["top", "center", "bottom"] | None, Field(description="Vertical alignment.")
        ] = None,
        wrap_text: bool | None = None,
        border: Annotated[
            Literal["thin", "medium", "thick", "dashed", "dotted", "double", "none"] | None,
            Field(description="Border on every cell."),
        ] = None,
        border_color: str | None = None,
        column_width: Annotated[
            float | dict[str, float] | None,
            Field(description="Width for columns in range, or {'A': 30, 'B': 12}."),
        ] = None,
        row_height: float | None = None,
        autofit: Annotated[bool, Field(description="Auto-size columns in range to their content.")] = False,
        freeze_panes: Annotated[
            str | None,
            Field(
                description="Freeze above/left of this cell, e.g. 'A2' (header row) or 'B2'; 'none' unfreezes."
            ),
        ] = None,
        merge: Annotated[bool | None, Field(description="true = merge range, false = unmerge.")] = None,
        auto_filter: Annotated[
            bool | None, Field(description="Add (true) or remove (false) filter dropdowns on the range.")
        ] = None,
        conditional: Annotated[
            dict[str, Any] | None,
            Field(
                description=(
                    "Conditional formatting: {'type':'color_scale'} | {'type':'data_bar','color':'#638EC6'} | {'type':'icon_set'} | "
                    "{'type':'cell_is','operator':'>','values':[100],'fill_color':'lightgreen'} | {'type':'formula','formula':'$D2<0','fill_color':'lightred'} | "
                    "{'type':'duplicates'}"
                )
            ),
        ] = None,
        validation: Annotated[
            dict[str, Any] | None,
            Field(
                description=(
                    "Data validation: {'type':'list','values':['Yes','No']} (dropdown) or {'type':'whole','min':0,'max':100}."
                )
            ),
        ] = None,
        dry_run: DryRun = False,
    ) -> dict:
        """Make sheets look good: fonts, fills, number formats, alignment, borders, widths, freeze panes, merge,
        filters, conditional formatting and dropdown validation. Combine several options in one call."""
        return S.format_range(
            path,
            range,
            sheet,
            bold,
            italic,
            underline,
            font_size,
            font_color,
            font_name,
            fill_color,
            number_format,
            h_align,
            v_align,
            wrap_text,
            border,
            border_color,
            column_width,
            row_height,
            autofit,
            freeze_panes,
            merge,
            auto_filter,
            conditional,
            validation,
            dry_run,
        )

    @tool(WRITE, write=True)
    def add_chart(
        path: Path_,
        data_range: Annotated[
            str, Field(description="Data incl. header row, e.g. 'B1:D13' (one series per column).")
        ],
        chart_type: Annotated[
            Literal["bar", "column", "line", "pie", "scatter", "area"], Field(description="Chart type.")
        ] = "column",
        sheet: Sheet_ = None,
        categories_range: Annotated[
            str | None, Field(description="Labels / X values without header, e.g. 'A2:A13'.")
        ] = None,
        title: str | None = None,
        x_title: str | None = None,
        y_title: str | None = None,
        anchor: Annotated[
            str | None,
            Field(description="Top-left cell where the chart is placed (default: right of the data)."),
        ] = None,
        target_sheet: Annotated[
            str | None, Field(description="Put the chart on another sheet (e.g. 'Dashboard').")
        ] = None,
        stacked: bool = False,
        dry_run: DryRun = False,
    ) -> dict:
        """Add a native, editable Excel chart (column, bar, line, pie, scatter, area) built from a cell range."""
        return S.add_chart(
            path,
            data_range,
            chart_type,
            sheet,
            categories_range,
            title,
            x_title,
            y_title,
            anchor,
            target_sheet,
            16,
            8,
            True,
            stacked,
            dry_run,
        )

    # ================================================================= power tools (v0.2)
    @tool(RO)
    def quick_look(
        path: Annotated[str, Field(description="A folder (default: the root) or a file.")] = ".",
        sheet: Sheet_ = None,
    ) -> dict:
        """The best first call. For a folder: every data file with sheets, row counts, columns and the SQL table name
        to use. For a file: sheets, column types and stats, data-quality score and issues, formula errors,
        a preview and suggested next steps. One call instead of list_files + describe + profile."""
        return S.quick_look(path, sheet)

    @tool(RO)
    def clean(
        source: Source,
        operations: Annotated[
            list[dict[str, Any]],
            Field(
                description="Steps applied in order. Each is {'op': ..., options}. Ops: "
                "trim {columns?, collapse_spaces?} | case {mode: upper|lower|title, columns?} | "
                "replace {find, replace, columns?, regex?, whole_cell?} | split {column, sep, into: [names]} | "
                "merge {columns, into, sep} | fill_down {columns} | fill_null {columns, value} | "
                "to_number {columns, decimal?: '.'|','} | to_date {columns, format?, dayfirst?} | "
                "drop_empty_rows | dedupe {columns?, keep?: first|last} | drop_columns {columns} | "
                "keep_columns {columns} | rename {mapping} | filter {where} | sort {order_by} | "
                "add_column {name, expr}. Text ops default to all text columns."
            ),
        ],
        sheet: Sheet_ = None,
        output: Annotated[
            str | None,
            Field(
                description="Save the cleaned table here. The SAME xlsx file cleans the sheet in place (styles, "
                "tables and other sheets kept; backup made). Omit to preview."
            ),
        ] = None,
        output_sheet: OutSheet = None,
        limit: Limit = None,
        format: Fmt = "json",
    ) -> dict:
        """Fix messy data without writing code: trim spaces, fix case, find/replace, split/merge columns, fill down
        blanks (merged-cell exports), numbers stored as text, mixed date formats, empty rows and duplicates.
        Previews by default; pass output to save."""
        return S.clean(source, operations, sheet, output, output_sheet, limit, format)

    @tool(WRITE, write=True)
    def manage_objects(
        path: Path_,
        action: Annotated[
            Literal[
                "list",
                "table_create",
                "table_delete",
                "name_create",
                "name_delete",
                "validation_delete",
                "conditional_delete",
            ],
            Field(
                description="list shows Excel tables, named ranges, data validations and conditional formats."
            ),
        ] = "list",
        sheet: Sheet_ = None,
        name: Annotated[
            str | None, Field(description="Table or name to create/delete, e.g. 'SalesData'.")
        ] = None,
        range: Annotated[
            str | None,
            Field(
                description="Range for table_create (with header row), name_create, or rule deletion scope."
            ),
        ] = None,
        refers_to: Annotated[
            str | None, Field(description="name_create: a reference 'Sheet1!$B$2:$B$9' or constant '=0.15'.")
        ] = None,
        style: Annotated[str, Field(description="Excel table style.")] = "TableStyleMedium9",
        dry_run: DryRun = False,
    ) -> dict:
        """Manage Excel structure objects: turn a range into a formatted Excel Table (usable as Table[Column] in
        formulas), create/delete named ranges and constants, list or remove data validation and conditional
        formatting rules."""
        return S.manage_objects(path, action, sheet, name, range, refers_to, style, dry_run)

    @tool(DESTRUCTIVE, write=True)
    def backups(
        path: Annotated[
            str | None, Field(description="File (or folder) whose backups to use. Default: all.")
        ] = None,
        action: Annotated[
            Literal["list", "undo", "restore", "delete"],
            Field(
                description="undo = revert the file's latest change; restore = put back a specific backup."
            ),
        ] = "list",
        backup: Annotated[
            str | None, Field(description="Backup path from action='list' (restore/delete).")
        ] = None,
        limit: Annotated[int, Field(description="Max backups to list.")] = 50,
    ) -> dict:
        """Undo mistakes. Every write keeps a timestamped backup; list them, undo the last change to a file, or
        restore any version. An undo itself is reversible (the replaced state is kept as a snapshot)."""
        return S.backups(path, action, backup, limit)

    @tool(DESTRUCTIVE)
    def batch(
        operations: Annotated[
            list[dict[str, Any]],
            Field(
                description="Up to 100 steps: [{'tool': 'write_range', 'args': {...}}, {'tool': "
                "'format_range', 'args': {...}}]. Any SheetMCP tool except batch/server_info."
            ),
        ],
        stop_on_error: Annotated[bool, Field(description="Stop at the first failing step.")] = True,
    ) -> dict:
        """Run many SheetMCP operations in one call (e.g. write data, format header, add chart) and get every
        result back. Saves round trips for multi-step edits."""
        if cfg.read_only:
            ro = {
                "list_files",
                "describe",
                "read_range",
                "query",
                "profile",
                "formulas",
                "search",
                "diff",
                "quick_look",
                "catalog",
                "transform",
                "aggregate",
                "combine",
                "clean",
            }
            for op in operations:
                if (op.get("tool") or op.get("op")) not in ro:
                    raise ToolError("Server is read-only; batch may only contain read tools.")
        return S.batch(operations, stop_on_error)

    @tool(WRITE, write=True)
    def build_report(
        source: Source,
        output: Annotated[str, Field(description="Report file to create, e.g. 'sales_report.xlsx'.")],
        group_by: Annotated[list[str], Field(description="Summary rows, e.g. ['Region'].")],
        metrics: Annotated[
            list[str],
            Field(description="e.g. ['Revenue:sum', 'Units:avg', 'count'] (same syntax as aggregate)."),
        ],
        sheet: Sheet_ = None,
        where: Annotated[str | None, Field(description='Optional SQL filter, e.g. "Year = 2024".')] = None,
        title: Annotated[str | None, Field(description="Report / chart title.")] = None,
        chart_type: Annotated[
            Literal["column", "bar", "line", "pie", "area", "none"], Field(description="Summary chart.")
        ] = "column",
        include_data: Annotated[
            bool, Field(description="Add a 'Data' sheet with the (filtered) rows.")
        ] = True,
        number_format: Annotated[
            str,
            Field(description="Number format for metrics, e.g. '#,##0.00' or '[$৳-845]#,##0.00' for BDT."),
        ] = "#,##0.00",
        top: Annotated[int | None, Field(description="Only the top N groups (by the first metric).")] = None,
        overwrite: Annotated[
            bool, Field(description="Replace an existing report (a backup is kept).")
        ] = False,
    ) -> dict:
        """Create a finished Excel report in one call: a Summary sheet (grouped metrics sorted by the first metric,
        a live-formula totals row, styled header, number formats, frozen header, chart) plus a filterable Data
        sheet."""
        return S.build_report(
            source,
            output,
            group_by,
            metrics,
            sheet,
            where,
            title,
            chart_type,
            include_data,
            number_format,
            overwrite,
            top,
        )

    # ================================================================= resources
    @mcp.resource(
        "sheet://info",
        name="server-info",
        mime_type="application/json",
        description="SheetMCP configuration and tips.",
    )
    def info_resource() -> str:
        return json.dumps(S.info(), indent=2)

    @mcp.resource(
        "sheet://files",
        name="files",
        mime_type="application/json",
        description="All spreadsheet/data files under the allowed root(s).",
    )
    def files_resource() -> str:
        return json.dumps(S.list_files(".", True, None, 500), indent=2)

    @mcp.resource(
        "sheet://describe/{path}",
        name="describe-file",
        mime_type="application/json",
        description="Structure + schema of a file. URL-encode '/' in the path as %2F.",
    )
    def describe_resource(path: str) -> str:
        try:
            return json.dumps(S.describe(unquote(path)), indent=2, default=str)
        except SheetError as e:
            return json.dumps({"error": str(e)})

    @mcp.resource(
        "sheet://preview/{path}/{sheet}",
        name="sheet-preview",
        mime_type="text/markdown",
        description="First rows of a sheet as a markdown table. URL-encode '/' in path as %2F.",
    )
    def preview_resource(path: str, sheet: str) -> str:
        try:
            r = S.read_range(unquote(path), unquote(sheet), None, False, "auto", 0, 50, "markdown")
            return r.get("markdown", "")
        except SheetError as e:
            return f"Error: {e}"

    # concrete per-file resources (so clients can browse/attach files directly)
    try:
        for f in S.list_files(".", True, None, 100)["files"]:
            rel = f["path"]
            uri = "sheet://describe/" + quote(rel, safe="")

            def make(r: str):
                return lambda: json.dumps(S.describe(r), indent=2, default=str)

            mcp.resource(uri, name=rel, mime_type="application/json", description=f"{f['kind']} file {rel}")(
                make(rel)
            )
    except Exception as e:  # pragma: no cover
        log.warning("could not list files for resources: %s", e)

    # ================================================================= prompts
    @mcp.prompt(
        name="analyze_spreadsheet",
        description="Explore a spreadsheet and summarise what it contains and what stands out.",
    )
    def analyze_spreadsheet(path: str, question: str = "") -> str:
        return (
            f"Analyze the file '{path}'.\n"
            "1. Call quick_look to learn its sheets, columns, types, statistics and data-quality issues.\n"
            "2. If needed, call profile on other sheets.\n"
            "3. Use query (SQL) or aggregate to compute the key totals, trends and top/bottom items.\n"
            "4. If it has formulas, run formulas(action='errors').\n"
            "Then write a concise report: what the data is, key numbers, notable patterns, data problems, and next steps."
            + (f"\nFocus on this question: {question}" if question else "")
        )

    @mcp.prompt(name="clean_data", description="Find and fix data quality problems, saving a cleaned copy.")
    def clean_data(path: str, output: str = "") -> str:
        out = output or "a new file next to the original with a _clean suffix"
        return (
            f"Clean the data in '{path}'. First run quick_look and list every issue found. Propose fixes (trim "
            "whitespace, remove duplicates, fix numbers stored as text, standardise dates/categories, fill down blanks). "
            f"Preview them with the clean tool, and after I agree save the result to {out} (never overwrite the original)."
        )

    @mcp.prompt(name="audit_formulas", description="Audit an Excel model's formulas for errors and risks.")
    def audit_formulas(path: str) -> str:
        return (
            f"Audit the formulas in '{path}'. Use formulas(action='stats'), then formulas(action='errors'). For each error, "
            "trace the cell (action='trace') to find the root cause and suggest a corrected formula. Also flag hard-coded "
            "numbers inside formulas, inconsistent formulas in a column, and functions the evaluator could not compute."
        )

    @mcp.prompt(name="compare_files", description="Explain the differences between two versions of a file.")
    def compare_files(file_a: str, file_b: str, key: str = "") -> str:
        k = f" using key column(s) {key}" if key else ""
        return (
            f"Compare '{file_a}' with '{file_b}'{k} using the diff tool. Summarise what was added, removed and changed, "
            "highlight the most significant changes, and note any formula changes."
        )

    @mcp.prompt(
        name="build_report",
        description="Turn raw data into a formatted Excel report with a summary and chart.",
    )
    def report_prompt(source: str, output: str = "report.xlsx", goal: str = "") -> str:
        return (
            f"Build an Excel report from '{source}' into '{output}'.\n"
            "1. quick_look the data to pick sensible group-by columns and metrics. 2. Call build_report "
            f"(output='{output}'). 3. If more is needed, add sheets with aggregate(output=..., output_sheet=...), "
            "format_range and add_chart. 4. Describe what the report shows."
            + (f"\nGoal: {goal}" if goal else "")
        )

    return mcp


__all__ = ["__version__", "build_server"]
