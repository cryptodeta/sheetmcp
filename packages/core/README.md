# sheetmcp-core

Pure-Python spreadsheet engine used by [SheetMCP](../../README.md): read/write xlsx with formulas
(openpyxl), SQL over CSV/xlsx/Parquet/JSON (DuckDB), formula tracing & evaluation, profiling, diffing.
It has no MCP dependency, so you can use it from any Python program:

```python
from sheetmcp_core import Sheets, Config

s = Sheets(Config(roots=["./data"]))
print(s.query("SELECT region, sum(amount) FROM sales GROUP BY 1", tables={"sales": "sales.xlsx"}))
```
