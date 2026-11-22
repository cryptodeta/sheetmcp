# sheetmcp-cli

`sheetcli` runs the same engine as the MCP server from your terminal:

```bash
sheetcli ls ./data
sheetcli query "SELECT * FROM sales LIMIT 5" -t sales=data/sales.xlsx
sheetcli convert data/sales.xlsx data/sales.parquet
```
