"""The public facade: one object exposing every operation."""

from __future__ import annotations

from typing import Any

from . import __version__
from .analysis import AnalysisOps
from .formatops import FormatOps
from .formulaops import FormulaOps
from .power import PowerOps
from .sqlops import SqlOps
from .workbook import WorkbookOps


class Sheets(WorkbookOps, SqlOps, FormulaOps, FormatOps, AnalysisOps, PowerOps):
    """All spreadsheet operations, sandboxed to ``config.roots``.

    >>> s = Sheets(Config(roots=["./data"]))
    >>> s.describe("sales.xlsx")
    """

    def info(self) -> dict[str, Any]:
        import duckdb
        import openpyxl

        return {
            "name": "SheetMCP",
            "version": __version__,
            "roots": [str(r) for r in self.sb.roots],
            "read_only": self.cfg.read_only,
            "backups": self.cfg.backup,
            "max_rows_per_response": self.cfg.max_rows,
            "engines": {"openpyxl": openpyxl.__version__, "duckdb": duckdb.__version__},
            "supported_inputs": [
                "xlsx",
                "xlsm",
                "csv",
                "tsv",
                "txt",
                "parquet",
                "json",
                "jsonl",
                "xls (extra)",
                "ods",
                "html",
            ],
            "supported_outputs": ["xlsx", "csv", "tsv", "parquet", "json", "jsonl", "md"],
            "tips": [
                "Relative paths resolve against the first root.",
                "In SQL, files in the root folder are tables by name: sales.csv -> sales, book.xlsx sheet 'Q1' -> book__q1.",
                "You can also query a file directly: SELECT * FROM 'data/file.xlsx#Sheet1'.",
                "Every overwrite makes a backup in .sheetmcp_backups/ next to the file; backups(action='undo') reverts.",
                "Start with quick_look(path='.') for an overview of a folder or file.",
            ],
        }
