"""Use the engine directly from Python (no MCP needed).

uv run python examples/python_usage.py
"""

from pathlib import Path

from sheetmcp_core import Config, Sheets

data = Path(__file__).parent / "data"
if not (data / "sales.xlsx").exists():
    import make_samples  # type: ignore[import-not-found]

    make_samples.make(data)

s = Sheets(Config(roots=[str(data)]))

print(
    s.query(
        "SELECT Region, round(sum(Revenue), 2) AS revenue FROM sales GROUP BY 1 ORDER BY 2 DESC",
        format="markdown",
    )["markdown"]
)
print(s.read_range("sales.xlsx", "Summary", "A1:C6", format="markdown")["markdown"])
print("Formula errors:", [(e["cell"], e["error"]) for e in s.formulas("sales.xlsx", "errors")["errors"]])
print("Quality issues:", [(i["issue"], i["column"]) for i in s.profile("customers.csv")["issues"]])

# 0.2 power features --------------------------------------------------------
look = s.quick_look(".")
print("Tables:", [f.get("sql") or [sh["sql"] for sh in f.get("sheets", [])] for f in look["files"]])
cleaned = s.clean(
    "customers.csv", [{"op": "trim"}, {"op": "dedupe"}, {"op": "to_number", "columns": ["CreditLimit"]}]
)
print("Rows before/after cleaning:", cleaned["rows_before"], cleaned["rows_after"])
print("Catalog:", [t["table"] for t in s.query()["tables"]])
