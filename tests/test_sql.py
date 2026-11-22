import json
from pathlib import Path

import openpyxl
import pytest

from sheetmcp_core import SheetError


def test_query_auto_registers_files(sheets):
    r = sheets.query("SELECT Region, round(sum(Revenue), 2) AS rev FROM sales GROUP BY 1 ORDER BY rev DESC")
    assert r["columns"] == ["Region", "rev"] and r["row_count"] == 4
    assert r["tables"] == {"sales": "sales.xlsx#Sales"}


def test_query_join_across_formats(sheets):
    r = sheets.query(
        "SELECT s.Region, sum(s.Revenue) AS rev, any_value(t.Target) AS target "
        "FROM sales s JOIN targets t USING (Region) JOIN customers c ON c.Customer = s.Customer GROUP BY 1",
        format="records",
    )
    assert {row["Region"] for row in r["rows"]} == {"North", "South", "East", "West"}
    assert set(r["tables"]) == {"sales", "targets", "customers"}


def test_query_explicit_tables_and_sheet_alias(sheets):
    r = sheets.query("SELECT count(*) AS n FROM s", tables={"s": "sales.xlsx#Summary"})
    assert r["rows"] == [[5]]
    r = sheets.query("SELECT count(*) AS n FROM sales__summary")
    assert r["rows"] == [[5]]


def test_query_file_literal(sheets):
    r = sheets.query("SELECT count(*) FROM 'orders_2024-02.csv'")
    assert r["rows"] == [[5]]


def test_query_paging_and_markdown(sheets):
    r = sheets.query("SELECT * FROM sales", limit=7)
    assert r["row_count"] == 7 and r["truncated"] and r["next_offset"] == 7
    md = sheets.query("SELECT Region FROM sales LIMIT 2", format="markdown")["markdown"]
    assert md.startswith("| Region |")


@pytest.mark.parametrize(
    "sql",
    [
        "DROP TABLE x",
        "COPY (SELECT 1) TO 'x.csv'",
        "SELECT 1; SELECT 2",
        "CREATE TABLE t AS SELECT 1",
        "INSTALL httpfs",
    ],
)
def test_query_rejects_non_select(sheets, sql):
    with pytest.raises(SheetError):
        sheets.query(sql)


def test_query_sandbox(sheets, tmp_path: Path):
    outside = tmp_path / "secret.csv"
    outside.write_text("a\n1\n")
    with pytest.raises(SheetError) as e:
        sheets.query(f"SELECT * FROM read_csv('{outside}')")
    assert "outside" in str(e.value).lower() or "permission" in str(e.value).lower()
    with pytest.raises(SheetError):
        sheets.query("SELECT * FROM read_text('/etc/passwd')")


def test_query_unknown_table_hint(sheets):
    with pytest.raises(SheetError) as e:
        sheets.query("SELECT * FROM nosuchtable")
    assert "Hint" in str(e.value)


def test_query_output_xlsx_and_sheet(sheets, data_dir: Path):
    r = sheets.query("SELECT * FROM customers", output="out/report.xlsx")
    assert r["rows_written"] == 13
    sheets.query(
        "SELECT Region, sum(Revenue) AS rev FROM sales GROUP BY 1",
        output="out/report.xlsx",
        output_sheet="Summary",
    )
    wb = openpyxl.load_workbook(data_dir / "out" / "report.xlsx")
    assert wb.sheetnames == ["Sheet1", "Summary"]
    assert wb["Sheet1"]["A1"].font.bold and wb["Summary"].max_row == 5


def test_transform(sheets):
    r = sheets.transform(
        "customers.csv", where="Country = 'Germany'", columns=["Customer", "Name"], rename={"Name": "company"}
    )
    assert r["columns"] == ["Customer", "company"] and r["row_count"] == 3
    d = sheets.transform("customers.csv", distinct=True)
    assert d["row_count"] == 12
    k = sheets.transform("customers.csv", dedupe_on=["Customer"], order_by=["Customer DESC"])
    assert k["row_count"] == 12 and k["rows"][0][0] == "C111"


def test_aggregate_and_pivot(sheets):
    r = sheets.aggregate("sales.xlsx", ["Revenue:sum", "Units:avg", "count"], group_by=["Region"])
    assert r["columns"] == ["Region", "sum_Revenue", "avg_Units", "count"]
    p = sheets.aggregate("sales.xlsx", ["sum(Units)"], group_by=["Region"], pivot_on="Product")
    assert {"Region", "Widget", "Gadget", "Gizmo", "Doohickey"} <= set(p["columns"])


def test_combine_append_and_join(sheets):
    r = sheets.combine(["orders_2024-01.csv", "orders_2024-02.csv", "orders_2024-03.csv"])
    assert r["row_count"] == 15 and "_source" in r["columns"]
    j = sheets.combine(["orders_2024-01.csv", "customers.csv"], mode="join", on=["customer"], how="left")
    assert j["row_count"] >= 5


def test_combine_join_column_case(sheets):
    with pytest.raises(SheetError):
        sheets.combine(["orders_2024-01.csv", "customers.csv"], mode="join", on=["nope"])


@pytest.mark.parametrize("ext", ["csv", "tsv", "parquet", "json", "jsonl", "md", "xlsx"])
def test_convert_formats(sheets, data_dir: Path, ext):
    r = sheets.convert("sales.xlsx", f"conv/sales.{ext}")
    assert r["rows_written"] == 40
    out = data_dir / "conv" / f"sales.{ext}"
    assert out.exists() and out.stat().st_size > 0
    if ext == "json":
        assert len(json.loads(out.read_text())) == 40
    if ext != "md":
        assert sheets.query(f"SELECT count(*) FROM 'conv/sales.{ext}'")["rows"] == [[40]]


def test_convert_all_sheets(sheets, data_dir: Path):
    r = sheets.convert("sales.xlsx", "split/sales.csv", all_sheets=True)
    assert len(r["outputs"]) == 2
    assert (data_dir / "split" / "sales_summary.csv").exists()


def test_profile_finds_quality_issues(sheets):
    p = sheets.profile("customers.csv")
    kinds = {(i["issue"], i["column"]) for i in p["issues"]}
    assert ("duplicate_rows", None) in kinds
    assert ("numbers_stored_as_text", "CreditLimit") in kinds
    assert ("untrimmed_whitespace", "Name") in kinds
    assert p["rows"] == 13 and p["quality_score"] < 100
    num = sheets.profile("sales.xlsx")
    rev = next(c for c in num["columns"] if c["name"] == "Revenue")
    assert rev["avg"] > 0 and "q50" in rev
