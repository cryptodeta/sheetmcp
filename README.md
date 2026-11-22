# SheetMCP

**The Excel / CSV MCP server that actually works.** Read formulas *and* values, run SQL across a whole
folder of spreadsheets, write safely without breaking formatting, audit formulas, build formatted reports
with charts. Pure Python, runs on macOS / Linux / Windows, **no Excel, no API keys, no cloud.**

```bash
uvx sheetmcp ~/Documents/spreadsheets
```

---

## Why SheetMCP?

| Common complaint about Excel MCPs | SheetMCP |
| --- | --- |
| "It only reads values, not formulas" | `read_range` returns values **and** formulas side by side |
| "Formulas show empty because the file was made by a script" | Built-in evaluator (~100 Excel functions) fills missing values |
| "Can't work across many CSV/xlsx files" | `query` = DuckDB SQL over every file in a folder, joins included |
| "Only works on Windows (COM)" | Pure Python: openpyxl + DuckDB, works everywhere |
| "60 tools, the model gets confused" | 18 focused tools with clear schemas |
| "Writes broke my formatting / formulas" | Only cell values change, formatting/charts/tables preserved, auto-backup, `dry_run` |
| "Inserting a row broke all formulas" | `insert_delete` re-points formulas across all sheets, like Excel |
| "Huge output blew up the context" | Row limits, paging (`next_offset`), `output=` saves full results to a file |
| "Cryptic errors" | Every error says what went wrong **and how to fix it** (with did-you-mean) |
| "Needs an API key / account" | None. Local files only, sandboxed to folders you choose |

## 1-minute quickstart

1. Install [uv](https://docs.astral.sh/uv/) (`curl -LsSf https://astral.sh/uv/install.sh | sh`, or `pip install uv`).
2. Add SheetMCP to your AI app (pick one below). Replace the folder with yours.
3. Ask: *"List my spreadsheets and summarise sales by region"*.

Try it without an AI app:

```bash
uvx --from sheetmcp-cli sheetcli --root ./examples/data query "SELECT Region, sum(Revenue) FROM sales GROUP BY 1" --md
```

> Not on PyPI yet? Install from the source tree / zip:
> `uv tool install ./packages/server` (or `pip install ./packages/core ./packages/server`), then use `sheetmcp` instead of `uvx sheetmcp`.

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

## Tools (18)

| Tool | What it does |
| --- | --- |
| `server_info` | Allowed folders, read-only mode, limits, tips |
| `list_files` | Find xlsx/xlsm/csv/tsv/parquet/json files (with sheet names) |
| `describe` | Sheets, sizes, tables, named ranges, merged cells, formula counts, charts, inferred column types, preview |
| `read_range` | Values **and** formulas side by side, auto-computes uncached formulas, paging, JSON or markdown |
| `write_range` | Write values/formulas (2D block, single cells, or append rows); preserves formatting; backup; dry-run |
| `manage_sheets` | Create workbook; add / rename (updates references) / delete / copy / move / hide sheets |
| `insert_delete` | Insert/delete rows/columns with formula, named range, table and merged-cell references updated |
| `query` | DuckDB SQL across many files at once: joins, aggregates, window functions, CTEs; save with `output` |
| `transform` | Filter / select / rename / sort / distinct / dedupe without writing SQL |
| `aggregate` | Group-by summaries and pivot tables |
| `combine` | Append (aligns columns by name) or join several files/sheets |
| `convert` | xlsx ⇄ csv ⇄ tsv ⇄ parquet ⇄ json ⇄ jsonl → markdown; all sheets at once |
| `profile` | Column stats + data-quality issues (duplicates, numbers-as-text, whitespace, outliers, empty/constant columns) |
| `formulas` | `list`, `errors` (#REF!, #DIV/0!, #N/A, circular), `trace` precedents/dependents, `evaluate`, `stats` |
| `search` | Find text/regex in every cell (and formula) of every file in a folder |
| `diff` | Compare files/sheets cell-by-cell (incl. formulas) or row-level by key columns |
| `format_range` | Fonts, fills, number formats (incl. ৳ BDT, $, €), borders, widths/autofit, freeze panes, merge, filters, conditional formatting, dropdown validation |
| `add_chart` | Native Excel charts: column, bar, line, pie, scatter, area |

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
* Only `SELECT` is allowed; DuckDB file access is locked to your folders.

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
  so a crash never leaves a half-written file.

## Python library & CLI

```python
from sheetmcp_core import Sheets, Config

s = Sheets(Config(roots=["./data"]))
s.read_range("sales.xlsx", "Summary", "A1:C6")  # values + formulas
s.query("SELECT * FROM sales WHERE Units > 30", output="big_orders.xlsx")
s.formulas("model.xlsx", "errors")
```

```bash
sheetcli ls data/
sheetcli describe data/sales.xlsx
sheetcli formulas data/sales.xlsx errors
sheetcli diff v1.xlsx v2.xlsx
```

## Performance (laptop-class VM, Python 3.13)

| Operation | Time |
| --- | --- |
| GROUP BY over a 500k-row CSV (9 MB) | ~0.1–0.5 s |
| First SQL over a 100k-row xlsx (converted once, cached by mtime) | ~2.5 s |
| Repeat SQL over the same xlsx | ~0.5 s |
| Profile 500k-row CSV | ~1 s |
| CSV → Parquet (500k rows) | ~0.2 s |

xlsx files above 8 MB are streamed (openpyxl read-only mode). Core dependencies: **openpyxl + duckdb** only.

## Repository layout

```
packages/core     sheetmcp-core: the engine (no MCP dependency)
packages/server   sheetmcp: MCP server (stdio + streamable HTTP), tools/resources/prompts
packages/cli      sheetmcp-cli: `sheetcli` terminal tool
examples/         sample data generator, client configs, Python usage
docs/             tool reference, architecture
tests/            pytest suite (fixtures generated on the fly)
```

## Development

```bash
uv sync                 # create .venv with all workspace packages + dev tools
uv run pytest           # 112 tests
uv run ruff check . && uv run ruff format --check .
uv run sheetmcp examples/data                     # run the server
npx @modelcontextprotocol/inspector uv run sheetmcp examples/data   # inspect interactively
```

See [CONTRIBUTING.md](CONTRIBUTING.md). License: [MIT](LICENSE).

## Known limitations

* Legacy `.xls` and `.ods` are not supported (clear error with a conversion hint). Password-protected files can't be opened.
* openpyxl preserves most content but drops some exotic objects on save (shapes, form controls, some chart styling,
  slicers, external-data connections). Pivot tables are kept but not refreshed.
* The formula evaluator covers ~100 common functions; others (e.g. dynamic arrays like FILTER/UNIQUE, financial
  functions) return `#NAME?` and are reported. Excel's own cached values are always preferred when present.
* `insert_delete` updates formulas, named ranges, tables and merged cells; conditional-formatting and data-validation
  ranges are not shifted.
* Structured references (`Table1[Column]`) are read and preserved but not evaluated.
