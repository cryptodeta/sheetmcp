# Architecture

```
            ┌──────────────── AI app (Claude, Cursor, VS Code, Windsurf, …) ───────────────┐
            │                         MCP (stdio or streamable HTTP)                        │
            └───────────────────────────────────────┬───────────────────────────────────────┘
                                                    │
packages/server  sheetmcp ─ FastMCP: 18 tools, resources, prompts, error → ToolError mapping
                                                    │  (thin: no logic, just schemas + calls)
packages/core    sheetmcp_core.Sheets  (facade made of mixins)
                 ├─ workbook.py   describe / read_range / write_range / manage_sheets / insert_delete
                 ├─ sqlops.py     query / transform / aggregate / combine / convert / profile (DuckDB)
                 ├─ formulaops.py list / errors / trace / evaluate / stats
                 ├─ evaluator.py  dependency-free Excel formula evaluator (~100 functions)
                 ├─ refshift.py   re-points references after row/column insert/delete
                 ├─ formatops.py  styles, conditional formatting, validation, charts
                 ├─ analysis.py   list_files / search / diff
                 ├─ fileio.py     type detection, workbook LRU cache, atomic save + backups, xlsx→csv cache
                 ├─ paths.py      sandbox (allowed roots, did-you-mean)
                 └─ config.py     env/CLI configuration
packages/cli     sheetcli ─ argparse wrapper over the same facade
```

## Key design decisions

* **Core has no MCP dependency.** The engine is a plain library (`openpyxl`, `duckdb`), so it is testable,
  reusable, and the server stays a thin adapter.
* **Formulas + values.** xlsx files store formulas and the last computed value separately. We load both
  (`data_only=False/True`) and return them side by side. When a value is missing (files written by scripts
  never get computed), the built-in evaluator computes it; Excel's cached values win when present.
* **SQL over xlsx.** DuckDB reads CSV/Parquet/JSON natively. For xlsx, a sheet is streamed once with openpyxl in
  read-only mode into a UTF-8 CSV cache keyed by path + mtime + size, then scanned by DuckDB. Repeat queries are fast.
  Excel tables referenced in SQL are registered lazily: only the files and sheets your SQL mentions get loaded.
* **Sandboxed DuckDB.** Every connection sets `allowed_directories`, disables external access and extension
  autoloading, and locks the configuration. Only single `SELECT` statements are accepted.
* **Safe writes.** Writes load a fresh workbook, change only `cell.value`, then save atomically
  (temp file + `os.replace`) after copying the original to `.sheetmcp_backups/`.
* **Context protection.** Every list-returning operation is limited (default 100 rows, configurable) and paged
  with `next_offset`. Long strings are truncated. `output=` writes full results to disk instead.
* **Performance.** LRU cache of parsed workbooks (keyed by mtime), raw-XML scan to skip formula parsing on
  formula-free files, fast sheet listing from `workbook.xml`, read-only streaming for files > 8 MB.
