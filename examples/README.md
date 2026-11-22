# Examples

| File | What |
| --- | --- |
| `make_samples.py` | Generates `data/` (sales.xlsx with formulas/table/named range, customers.csv with deliberate quality issues, targets.parquet, monthly order CSVs, events.json). Tests use it too. |
| `data/` | Pre-generated sample files, ready to point the server at: `uv run sheetmcp examples/data` |
| `python_usage.py` | Use the engine as a Python library |
| `mcp_client_demo.py` | Connect to the server over stdio with the official MCP client |
| `configs/` | Copy-paste configs: Claude Desktop, Cursor, VS Code, Windsurf, HTTP, Docker, from-source |

Prompts to try once connected:

* "Describe sales.xlsx and tell me which region is furthest from its target (targets.parquet)."
* "Find every formula error in sales.xlsx, explain the cause and fix it (show me a dry run first)."
* "Combine all orders_*.csv into orders_q1.xlsx, then add a bold header, autofit columns and a column chart."
* "Profile customers.csv and save a cleaned copy without duplicates and with trimmed names."
