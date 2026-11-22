# Changelog

All notable changes to this project are documented here. Format: [Keep a Changelog](https://keepachangelog.com/),
versioning: [SemVer](https://semver.org/).

## [0.2.0] - 2026-11-22

### Added
- **Easy setup**: `sheetmcp install` writes/merges MCP configs for Claude Desktop (Windows/macOS/Linux), Cursor,
  VS Code, Windsurf and Claude Code (with backups, `--dry-run`, `--print`, `--read-only`); `sheetmcp uninstall`;
  `sheetmcp doctor` (versions, sandbox self-test, folders, temp dir, configured clients, optional extras).
- 6 new tools (24 total):
  - `quick_look`: one-call overview of a folder or file (columns, stats, quality issues, formula errors, SQL names, next steps).
  - `clean`: cleaning pipeline (trim, case, replace/regex, split, merge, fill_down, fill_null, to_number,
    to_date with mixed formats, drop_empty_rows, dedupe, drop/keep/rename columns, filter, sort, add_column);
    preview, save, or clean a sheet in place keeping styles/tables.
  - `manage_objects`: Excel Tables, named ranges/constants, list/remove data validation and conditional formats.
  - `backups`: list / undo / restore / delete; undo is reversible.
  - `batch`: up to 100 operations in one call (read-only servers allow read tools only).
  - `build_report`: Summary (grouped metrics, live totals row, styling, chart) + Data sheet in one call.
- `query` without `sql` returns the catalog of every table name and its columns.
- Read support for `.ods` and HTML tables (stdlib parsers) and legacy `.xls` (optional `xlrd`, extra `[xls]`):
  describe, read, query, convert, search, diff, profile.
- `write_range`: `fill_formula` + `fill_range` (fill-down with relative references); streaming CSV append.
- `insert_delete` now shifts conditional formatting, data validation and auto-filter ranges.
- `profile`: inconsistent category spellings and robust z-score anomalies with row numbers.
- `convert`: `csv_dialect` = `excel` (UTF-8 BOM) or `google_sheets` (ISO dates).
- Formula evaluator: structured references (`Table[Col]`, `[@Col]`, `[#All]`, `[#Totals]`, `[[A]:[C]]`), array math,
  LET, SWITCH, FILTER, UNIQUE, SORT, SORTBY, SEQUENCE, TRANSPOSE, VSTACK, HSTACK, TAKE, DROP, CHOOSECOLS,
  CHOOSEROWS, TEXTBEFORE, TEXTAFTER, TEXTSPLIT, CHAR/UNICHAR, CODE/UNICODE, CLEAN, NUMBERVALUE, FIXED, T, N,
  PMT, FV, PV, NPER, RATE, NPV, IRR, EOMONTH, WEEKDAY, WEEKNUM, ISOWEEKNUM, NETWORKDAYS, WORKDAY, DATEDIF,
  HOUR, MINUTE, SECOND, TIME, DATEVALUE (143 functions total).
- `sheetcli look | clean | report | backups`.
- Sample `.ods` and HTML files (`make_samples.py`), 88 new tests (200 total).

### Changed
- Tool descriptions rewritten in plain language; server instructions point to `quick_look` first.
- Backup file names use real microseconds and are strictly ordered.
- English-only documentation.

### Fixed
- Windows: the xlsx/ods/html → CSV cache was written with doubled line endings (`\r\r\n`), breaking SQL on
  Excel files (failing Windows CI since 0.1.0). Cache keys are versioned so old cache files are ignored.
- Short defined names such as `Ids` or `Tax` were evaluated as whole-column references.

## [0.1.0] - 2026-11-22

### Added
- Monorepo with `sheetmcp-core` (engine), `sheetmcp` (MCP server) and `sheetmcp-cli` (`sheetcli`).
- 18 MCP tools: `server_info`, `list_files`, `describe`, `read_range`, `write_range`, `manage_sheets`,
  `insert_delete`, `query`, `transform`, `aggregate`, `combine`, `convert`, `profile`, `formulas`, `search`,
  `diff`, `format_range`, `add_chart`.
- Resources (`sheet://files`, `sheet://describe/{path}`, `sheet://preview/{path}/{sheet}`, per-file) and prompts
  (`analyze_spreadsheet`, `clean_data`, `audit_formulas`, `compare_files`, `build_report`).
- Built-in formula evaluator (~100 functions), formula error audit, precedent/dependent tracing.
- DuckDB SQL across xlsx/csv/tsv/parquet/json with auto-registration and sandboxing.
- Reference-aware row/column insert/delete and sheet rename.
- Sandbox roots, read-only mode, dry-run, automatic backups, atomic saves.
- stdio, streamable HTTP and SSE transports; Dockerfile; GitHub Actions CI; 112 tests.
