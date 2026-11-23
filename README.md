# SheetMCP

[![PyPI](https://img.shields.io/pypi/v/sheetmcp?logo=pypi&logoColor=white)](https://pypi.org/project/sheetmcp/)
[![Python](https://img.shields.io/pypi/pyversions/sheetmcp)](https://pypi.org/project/sheetmcp/)
[![CI](https://github.com/cryptodeta/sheetmcp/actions/workflows/ci.yml/badge.svg)](https://github.com/cryptodeta/sheetmcp/actions/workflows/ci.yml)
[![Release](https://github.com/cryptodeta/sheetmcp/actions/workflows/release.yml/badge.svg)](https://github.com/cryptodeta/sheetmcp/actions/workflows/release.yml)
[![GitHub release](https://img.shields.io/github/v/release/cryptodeta/sheetmcp)](https://github.com/cryptodeta/sheetmcp/releases)
[![Docker](https://img.shields.io/badge/docker-ghcr.io%2Fcryptodeta%2Fsheetmcp-2496ED?logo=docker&logoColor=white)](https://github.com/cryptodeta/sheetmcp/pkgs/container/sheetmcp)
[![License: MIT](https://img.shields.io/badge/license-MIT-blue)](LICENSE)

**The Excel / CSV MCP server that actually works.** Read formulas *and* values, run SQL across a whole
folder of spreadsheets, clean messy data, write safely without breaking formatting (with undo), audit
formulas, and build formatted reports with charts in one call. Pure Python, runs on macOS / Linux / Windows,
**no Excel, no API keys, no cloud.**

```bash
pip install sheetmcp                             # or run it without installing: uvx sheetmcp
sheetmcp install ~/Documents/spreadsheets        # adds SheetMCP to Claude Desktop / Cursor / Windsurf ...
sheetmcp doctor                                  # checks that everything works
```

Then ask your AI: *"Take a quick look at my spreadsheets and tell me what's there."*

---

## Install

Requires Python 3.10 or newer. Published on [PyPI](https://pypi.org/project/sheetmcp/) and as a Docker image on
[ghcr.io](https://github.com/cryptodeta/sheetmcp/pkgs/container/sheetmcp); every [GitHub Release](https://github.com/cryptodeta/sheetmcp/releases)
also has the wheel and sdist attached. One package gives you both commands: `sheetmcp` (the MCP server) and
`sheetcli` (the terminal tool).

```bash
# Run without installing (needs uv: https://docs.astral.sh/uv/)
uvx sheetmcp --version

# Install with pip
pip install sheetmcp

# Or as an isolated app with pipx
pipx install sheetmcp

# Optional extra: legacy .xls support (xlrd)
pip install "sheetmcp[xls]"            # uvx --from "sheetmcp[xls]" sheetmcp / pipx install "sheetmcp[xls]"

# Docker (stdio; mount the folder with your spreadsheets at /data)
docker run -i --rm -v "$PWD:/data" ghcr.io/cryptodeta/sheetmcp
# Docker over HTTP: http://localhost:8000/mcp
docker run -d -p 8000:8000 -v "$PWD:/data" ghcr.io/cryptodeta/sheetmcp --transport http --host 0.0.0.0
```

Add it to your AI app (or run `sheetmcp install <folder>` to do it for you; more clients in
[Manual configuration](#manual-configuration)):

**Claude Desktop** (`claude_desktop_config.json`) and **Cursor** (`~/.cursor/mcp.json`):

```json
{
  "mcpServers": {
    "sheets": { "command": "uvx", "args": ["sheetmcp", "/absolute/path/to/spreadsheets"] }
  }
}
```

**VS Code** (`.vscode/mcp.json`):

```json
{
  "servers": {
    "sheets": { "type": "stdio", "command": "uvx", "args": ["sheetmcp", "${workspaceFolder}"] }
  }
}
```

> **Note:** as of 0.3.0, `sheetmcp-core` and `sheetmcp-cli` are merged into the single `sheetmcp` package.
> The engine is now `sheetmcp.core` (was `sheetmcp_core`) and `sheetcli` ships with `sheetmcp`. Uninstall the
> old packages (`pip uninstall sheetmcp-core sheetmcp-cli`) and change `from sheetmcp_core import ...` to
> `from sheetmcp.core import ...`.

## Why SheetMCP?

| Common complaint about Excel MCPs | SheetMCP |
| --- | --- |
| "It only reads values, not formulas" | `read_range` returns values **and** formulas side by side |
| "Formulas show empty because the file was made by a script" | Built-in evaluator (140+ Excel functions incl. dynamic arrays, LET, financial, `Table[Column]` refs) fills missing values |
| "Can't work across many CSV/xlsx files" | `query` = DuckDB SQL over every file in a folder, joins included |
| "Only works on Windows (COM)" | Pure Python: openpyxl + DuckDB, works everywhere |
| "60 tools, the model gets confused" | 24 focused tools with plain-language descriptions; `quick_look` is the one obvious first call |
| "Setup is JSON-file fiddling" | `sheetmcp install` writes the config for Claude Desktop, Cursor, VS Code, Windsurf and Claude Code |
| "My data is messy" | `clean`: trim, case, split/merge, fill down, numbers-as-text, mixed date formats, dedupe, in one call |
| "The AI broke my file" | Every write is backed up; `backups(action="undo")` reverts it, and undo is itself undoable |
| "Can't open old .xls / LibreOffice files" | `.xls` (optional extra), `.ods` and HTML tables are readable and queryable |
| "Writes broke my formatting / formulas" | Only cell values change, formatting/charts/tables preserved, auto-backup, `dry_run` |
| "Inserting a row broke all formulas" | `insert_delete` re-points formulas, names, tables, conditional formats and validations, like Excel |
| "Huge output blew up the context" | Row limits, paging (`next_offset`), `output=` saves full results to a file |
| "Cryptic errors" | Every error says what went wrong **and how to fix it** (with did-you-mean) |
| "Needs an API key / account" | None. Local files only, sandboxed to folders you choose |

## 1-minute quickstart

Install it first (see [Install](#install)): `pip install sheetmcp`, or prefix the commands below with `uvx`.
From a clone instead: `pip install .` (with [uv](https://docs.astral.sh/uv/): `uv sync`, then prefix commands
with `uv run`).

### Easiest: `sheetmcp install`

```bash
sheetmcp install ~/Documents/spreadsheets                 # every AI app found on this computer
sheetmcp install . --client cursor --client vscode        # pick apps (claude-desktop, cursor, vscode, windsurf, claude-code, all)
sheetmcp install ~/data --read-only                       # no write tools
sheetmcp install ~/data --dry-run                         # show what would change
sheetmcp install ~/data --print                           # just print the JSON snippet
sheetmcp uninstall                                        # remove it again
sheetmcp doctor                                           # versions, sandbox self-test, folders, configured apps
```

`install` merges into existing config files (other servers are kept, the old file is backed up), uses the
current Python interpreter, and works on Windows, macOS and Linux. VS Code and Claude Code configs are written
to the project folder (`.vscode/mcp.json`, `.mcp.json`). Restart / reload the app afterwards.

Try it without an AI app:

```bash
sheetcli --root ./examples/data look
sheetcli --root ./examples/data query "SELECT Region, sum(Revenue) FROM sales GROUP BY 1" --md
```

### Manual configuration

The snippets below use `uvx sheetmcp`, which downloads and runs the PyPI package. With a plain
`pip install` or `pipx install`, use `"command": "sheetmcp"` (or the full path of your Python with `"args": ["-m", "sheetmcp", "<folder>"]`). The default folder is the current
directory when none is given.

### Claude Desktop

`~/Library/Application Support/Claude/claude_desktop_config.json` (macOS) or `%APPDATA%\Claude\claude_desktop_config.json` (Windows):

```json
{
  "mcpServers": {
    "sheets": {
      "command": "uvx",
      "args": ["sheetmcp", "/Users/you/Documents/spreadsheets"]
    }
  }
}
```

### Cursor

`~/.cursor/mcp.json` (global) or `.cursor/mcp.json` (project):

```json
{
  "mcpServers": {
    "sheets": {
      "command": "uvx",
      "args": ["sheetmcp", "${workspaceFolder}"]
    }
  }
}
```

### VS Code (Copilot agent mode)

`.vscode/mcp.json`:

```json
{
  "servers": {
    "sheets": {
      "type": "stdio",
      "command": "uvx",
      "args": ["sheetmcp", "${workspaceFolder}"]
    }
  }
}
```

### Windsurf

`~/.codeium/windsurf/mcp_config.json`:

```json
{
  "mcpServers": {
    "sheets": {
      "command": "uvx",
      "args": ["sheetmcp", "/path/to/spreadsheets"]
    }
  }
}
```

### Claude Code

```bash
claude mcp add sheets -- uvx sheetmcp "$PWD"
```

### Generic HTTP (any client that supports Streamable HTTP)

```bash
uvx sheetmcp /data --transport http --host 127.0.0.1 --port 8000
# endpoint: http://127.0.0.1:8000/mcp
```

```json
{ "mcpServers": { "sheets": { "url": "http://127.0.0.1:8000/mcp" } } }
```

### Docker

```bash
docker build -t sheetmcp .
docker run --rm -i -v "$PWD/data:/data" sheetmcp                       # stdio
docker run --rm -p 8000:8000 -v "$PWD/data:/data" sheetmcp --transport http --host 0.0.0.0
```

More snippets (read-only mode, multiple folders, env vars) in [`examples/configs/`](examples/configs).

## Tools (24)

**Start here**

| Tool | What it does |
| --- | --- |
| `quick_look` | One call: a folder's files with sheets, row counts, columns and SQL table names, or one file's column stats, quality score, issues, formula errors, preview and suggested next steps |
| `server_info` | Allowed folders, read-only mode, limits, tips |
| `list_files` | Find xlsx/xlsm/xls/ods/html/csv/tsv/parquet/json files (with sheet names) |
| `describe` | Sheets, sizes, tables, named ranges, merged cells, formula counts, charts, inferred column types, preview |

**Read & write**

| Tool | What it does |
| --- | --- |
| `read_range` | Values **and** formulas side by side, auto-computes uncached formulas, paging, JSON or markdown |
| `write_range` | Write values/formulas (2D block, single cells, append rows), fill a formula down a column (`fill_formula`); preserves formatting; backup; dry-run; streaming CSV append |
| `manage_sheets` | Create workbook; add / rename (updates references) / delete / copy / move / hide sheets |
| `insert_delete` | Insert/delete rows/columns with formulas, named ranges, tables, merged cells, conditional formats, validations and filters updated |
| `manage_objects` | Excel Tables (create/delete, usable as `Table[Column]`), named ranges and constants, list/remove data validation and conditional formatting |
| `backups` | List backups, `undo` the last change to a file, `restore` any version, delete |
| `batch` | Run up to 100 operations in one call |

**Analyse & transform**

| Tool | What it does |
| --- | --- |
| `query` | DuckDB SQL across many files at once: joins, aggregates, window functions, CTEs; no `sql` = list every table; save with `output` (new file or new sheet) |
| `transform` | Filter / select / rename / sort / distinct / dedupe without writing SQL |
| `aggregate` | Group-by summaries and pivot tables |
| `combine` | Append (aligns columns by name) or join several files/sheets |
| `clean` | Cleaning pipeline: trim, case, find/replace (regex), split, merge, fill down, fill blanks, to_number, to_date (mixed formats), drop empty rows, dedupe, drop/keep/rename columns, filter, sort, computed columns; preview, save, or clean a sheet in place |
| `convert` | xlsx / xls / ods / html ⇄ csv ⇄ tsv ⇄ parquet ⇄ json ⇄ jsonl → markdown; all sheets at once; Excel- and Google-Sheets-friendly CSV |
| `profile` | Column stats + data-quality issues (duplicates, numbers-as-text, whitespace, outliers, inconsistent category spellings, anomalies with row numbers, empty/constant columns) |
| `formulas` | `list`, `errors` (#REF!, #DIV/0!, #N/A, circular), `trace` precedents/dependents, `evaluate`, `stats` |
| `search` | Find text/regex in every cell (and formula) of every file in a folder |
| `diff` | Compare files/sheets cell-by-cell (incl. formulas) or row-level by key columns |

**Format & report**

| Tool | What it does |
| --- | --- |
| `format_range` | Fonts, fills, number formats (incl. ৳ BDT, $, €), borders, widths/autofit, freeze panes, merge, filters, conditional formatting, dropdown validation |
| `add_chart` | Native Excel charts: column, bar, line, pie, scatter, area |
| `build_report` | One call: Summary sheet (grouped metrics, live-formula totals row, styled header, number format, chart) + filterable Data sheet |

Plus **resources** (`sheet://files`, `sheet://describe/{path}`, `sheet://preview/{path}/{sheet}`, one per file)
and **prompts** (`analyze_spreadsheet`, `clean_data`, `audit_formulas`, `compare_files`, `build_report`).
Full reference: [docs/tools.md](docs/tools.md).

### SQL across files

Files in the root folder are tables automatically:

```sql
-- sales.xlsx (first sheet) ⨝ customers.csv ⨝ targets.parquet
SELECT c.Country, s.Region, sum(s.Revenue) AS revenue, any_value(t.Target) AS target
FROM sales s
JOIN customers c ON c.Customer = s.Customer
JOIN targets t USING (Region)
GROUP BY ALL ORDER BY revenue DESC
```

* `book.xlsx` → `book` (first sheet), `book__<sheet>` for other sheets (lower-case, non-alphanumerics → `_`).
* Any path: `SELECT * FROM 'reports/2024/q1.xlsx#Summary'`.
* Explicit aliases: `tables={"s": "data/sales.xlsx#2024"}`.
* Don't know the names? Call `query` with no SQL, or `quick_look`, to list every table and its columns.
* `.xls`, `.ods` and HTML files work too: `budget.ods` → `budget` / `budget__<sheet>`, each HTML `<table>` is a sheet.
* Only `SELECT` is allowed; DuckDB file access is locked to your folders.

### Cleaning data

```json
{"source": "raw/customers.csv",
 "operations": [{"op": "trim"}, {"op": "dedupe"},
                {"op": "to_number", "columns": ["CreditLimit"]},
                {"op": "to_date", "columns": ["Joined"], "dayfirst": true},
                {"op": "fill_down", "columns": ["Region"]}],
 "output": "clean/customers.xlsx"}
```

Without `output` you get a preview. Passing the same `.xlsx` as `output` cleans that sheet in place (styles,
tables and other sheets are kept, and a backup is made).

### Formulas

The evaluator computes formulas that have no cached value (files written by scripts) with 140+ functions:
math/stats, lookup (XLOOKUP, INDEX/MATCH...), text (TEXTBEFORE/AFTER, TEXTSPLIT...), dates (EOMONTH,
NETWORKDAYS, WORKDAY, DATEDIF...), financial (PMT, FV, PV, NPER, RATE, NPV, IRR), logic (LET, SWITCH, IFS),
dynamic arrays (FILTER, UNIQUE, SORT, SORTBY, SEQUENCE, VSTACK, HSTACK, TAKE, DROP, CHOOSECOLS...), array
math like `SUM(B2:B9*C2:C9)`, and structured references (`Sales[Qty]`, `[@Price]`, `Sales[#Totals]`).

## Configuration

| CLI flag | Env var | Default | Meaning |
| --- | --- | --- | --- |
| `ROOT ...` / `--root` | `SHEETMCP_ROOTS` (`:`-separated, `;` on Windows) | current dir | Folders the server may touch (sandbox) |
| `--read-only` | `SHEETMCP_READ_ONLY=1` | off | Hide/disable all write tools |
| `--no-backup` | `SHEETMCP_BACKUP=0` | backups on | Copy files to `.sheetmcp_backups/` before overwriting |
| | `SHEETMCP_MAX_BACKUPS` | 10 | Backups kept per file |
| `--max-rows` | `SHEETMCP_MAX_ROWS` | 100 | Default rows per response |
| | `SHEETMCP_MAX_CELL_CHARS` | 500 | Long text truncated in responses |
| `--transport` | `SHEETMCP_TRANSPORT` | `stdio` | `stdio`, `http` (streamable HTTP), `sse` |
| `--host` / `--port` | `SHEETMCP_HOST` / `SHEETMCP_PORT` | 127.0.0.1 / 8000 | HTTP bind address |
| `--log-level` | `SHEETMCP_LOG_LEVEL` | WARNING | Logs go to stderr |

## Safety

* **Sandbox**: every path is resolved (symlinks too) and must live inside an allowed root. DuckDB runs with
  `enable_external_access=false` + `allowed_directories` + locked configuration, so SQL cannot read other files,
  install extensions or reach the network.
* **Read-only mode** hides every write tool.
* **Dry run** on every write tool shows exactly which cells would change.
* **Automatic backups** (`.sheetmcp_backups/<name>.<timestamp>.xlsx`) and atomic saves (write temp → rename),
  so a crash never leaves a half-written file. `backups(action="undo")` reverts the latest change.

## Python library & CLI

```python
from sheetmcp.core import Sheets, Config

s = Sheets(Config(roots=["./data"]))
s.read_range("sales.xlsx", "Summary", "A1:C6")  # values + formulas
s.query("SELECT * FROM sales WHERE Units > 30", output="big_orders.xlsx")
s.formulas("model.xlsx", "errors")
s.quick_look(".")  # overview of the folder
s.clean("raw.csv", [{"op": "trim"}, {"op": "dedupe"}], output="clean.xlsx")
s.build_report("sales.xlsx", "report.xlsx", group_by=["Region"], metrics=["Revenue:sum", "count"])
s.backups("sales.xlsx", "undo")
```

```bash
sheetcli ls data/
sheetcli describe data/sales.xlsx
sheetcli formulas data/sales.xlsx errors
sheetcli diff v1.xlsx v2.xlsx
sheetcli look data/
sheetcli clean data/raw.csv trim dedupe to_number:Amount -o data/clean.xlsx
sheetcli report data/sales.xlsx report.xlsx -g Region -m Revenue:sum -m count
sheetcli backups data/sales.xlsx --undo
```

## Performance (laptop-class VM, Python 3.13)

| Operation | Time |
| --- | --- |
| GROUP BY over a 500k-row CSV (9 MB) | ~0.1–0.5 s |
| First SQL over a 100k-row xlsx (converted once, cached by mtime) | ~2.5 s |
| Repeat SQL over the same xlsx | ~0.5 s |
| Profile 500k-row CSV | ~1 s |
| CSV → Parquet (500k rows) | ~0.2 s |

xlsx files above 8 MB are streamed (openpyxl read-only mode); appending rows to a CSV streams without
rewriting the file. Core dependencies: **openpyxl + duckdb** only (`xlrd` optional for `.xls`).

## Repository layout

```
src/sheetmcp/          MCP server (stdio + streamable HTTP), tools/resources/prompts, install/doctor
src/sheetmcp/core/     the engine, importable as `sheetmcp.core` (no MCP dependency)
src/sheetmcp/cli/      `sheetcli` terminal tool
examples/         sample data generator, client configs, Python usage
docs/             tool reference, architecture
tests/            pytest suite (fixtures generated on the fly)
```

## Development

```bash
uv sync                 # create .venv with sheetmcp (editable) + dev tools
uv run pytest           # 200 tests
uv run ruff check . && uv run ruff format --check .
uv run sheetmcp examples/data                     # run the server
npx @modelcontextprotocol/inspector uv run sheetmcp examples/data   # inspect interactively
```

See [CONTRIBUTING.md](CONTRIBUTING.md). License: [MIT](LICENSE).

## Releasing

[`VERSION.md`](VERSION.md) holds the release version on its first line. To ship a release, add the notes to
`CHANGELOG.md`, bump that line and push to `main`. The `Release` workflow tags `vX.Y.Z`, creates the GitHub
Release (wheel and sdist attached) and publishes to PyPI and ghcr.io only when the new version is greater
than the latest tag; an equal or lower version is a green no-op. Package versions are synced from
`VERSION.md` automatically (`python scripts/version.py sync` does the same locally).

### PyPI trusted publishing (no tokens)

PyPI releases use [trusted publishing](https://docs.pypi.org/trusted-publishers/) (OIDC), so there is no
API token secret. On pypi.org open the `sheetmcp` project, then **Manage** -> **Settings** -> **Publishing**
-> **Add a new publisher** -> **GitHub**, with these values:

| Field | Value |
| --- | --- |
| Owner | `cryptodeta` |
| Repository name | `sheetmcp` |
| Workflow name | `release.yml` |
| Environment name | `pypi` |

The `pypi` job runs in the `pypi` GitHub environment, which already exists in this repository.

Bumping `VERSION.md` on `main` then publishes the `sheetmcp` wheel and sdist. To publish a version whose tag already exists
(for example the current one, or after a failed run), run the `Release` workflow manually with **republish**
checked (`gh workflow run release.yml -f republish=true`); it uploads the files attached to that GitHub Release
and never creates a second tag or release. Until the publisher is configured, the PyPI job ends green with a
warning. ghcr.io uses the built-in `GITHUB_TOKEN` and needs no setup.

## Known limitations

* `.xls`, `.ods` and HTML are **read-only** inputs (convert to `.xlsx` to edit). `.xls` needs the optional `xlrd`
  package (`pip install xlrd`, or the `xls` extra). Password-protected files can't be opened.
* openpyxl preserves most content but drops some exotic objects on save (shapes, form controls, some chart styling,
  slicers, external-data connections). Pivot tables are kept but not refreshed.
* The formula evaluator covers 140+ functions; others (e.g. LAMBDA/MAP, cube and web functions) return `#NAME?`
  and are reported. Excel's own cached values are always preferred when present. External workbook links are
  not evaluated.
* Cleaning a sheet in place replaces formula cells inside the cleaned table with their values (reported in the
  response; undo with `backups`).
* `sheetmcp install` writes configs that point at the Python interpreter it runs from; re-run it if you move or
  recreate that environment.
