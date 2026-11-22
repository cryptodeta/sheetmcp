# Changelog

All notable changes to this project are documented here. Format: [Keep a Changelog](https://keepachangelog.com/),
versioning: [SemVer](https://semver.org/).

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
