from pathlib import Path

import openpyxl
import pytest

from sheetmcp.core import ReadOnlyError, SheetError


def test_describe_xlsx(sheets):
    d = sheets.describe("sales.xlsx")
    assert [s["name"] for s in d["sheets"]] == ["Sales", "Summary"]
    sales = d["sheets"][0]
    assert sales["max_row"] == 41 and sales["formula_cells"] == 40
    assert sales["tables"][0]["name"] == "SalesTable"
    cols = {c["name"]: c for c in sales["columns"]}
    assert cols["Units"]["type"] == "integer"
    assert cols["Date"]["type"] == "date"
    assert cols["Revenue"]["formula_example"] == "=D2*E2"
    assert d["named_ranges"][0]["name"] == "TotalRevenue"


def test_describe_csv_and_parquet(sheets):
    d = sheets.describe("customers.csv")
    assert d["row_count"] == 13
    assert [c["name"] for c in d["columns"]][:2] == ["Customer", "Name"]
    assert sheets.describe("targets.parquet")["row_count"] == 4


def test_read_values_and_formulas(sheets):
    r = sheets.read_range("sales.xlsx", "Summary", "A1:C6")
    assert r["formulas"]["B6"] == "=SUM(B2:B5)"
    total = r["values"][5][1]
    assert total == pytest.approx(sum(row[1] for row in r["values"][1:5]))
    assert "B6" in r["computed_by_evaluator"]


def test_read_paging(sheets):
    r = sheets.read_range("sales.xlsx", "Sales", limit=10)
    assert len(r["values"]) == 10 and r["truncated"] and r["next_offset"] == 10
    r2 = sheets.read_range("sales.xlsx", "Sales", offset=r["next_offset"], limit=100)
    assert r2["first_row"] == 11 and len(r2["values"]) == 31 and "truncated" not in r2


def test_read_markdown_and_csv(sheets):
    r = sheets.read_range("customers.csv", range="A1:C3", format="markdown")
    assert "| 1 | Customer | Name | Country |" in r["markdown"]


def test_write_preserves_formatting_and_backs_up(sheets, data_dir: Path):
    res = sheets.write_range("sales.xlsx", "Sales", cells={"D2": 99, "H1": "Note", "H2": "=F2*2"})
    assert res["cells_written"] == 3 and res["backup"]
    assert Path(res["backup"]).exists()
    wb = openpyxl.load_workbook(data_dir / "sales.xlsx")
    ws = wb["Sales"]
    assert ws["D2"].value == 99
    assert ws["A1"].font.bold is True  # header formatting kept
    assert ws["F2"].value == "=D2*E2"  # other formulas kept
    assert ws["H2"].value == "=F2*2"
    assert ws.freeze_panes == "A2"
    assert "SalesTable" in ws.tables


def test_write_dry_run_does_not_save(sheets, data_dir: Path):
    before = (data_dir / "sales.xlsx").read_bytes()
    res = sheets.write_range("sales.xlsx", "Sales", "A2", [[1, 2]], dry_run=True)
    assert res["dry_run"] and res["changes"][0]["old"] == "2024-01-01"
    assert (data_dir / "sales.xlsx").read_bytes() == before


def test_write_new_file_append_and_dates(sheets, data_dir: Path):
    sheets.write_range("new.xlsx", "Log", values=[["when", "what"], ["2024-05-01", "x"]])
    sheets.write_range("new.xlsx", "Log", values=[["2024-05-02", "y"]], append=True)
    ws = openpyxl.load_workbook(data_dir / "new.xlsx")["Log"]
    assert ws.max_row == 3 and ws["A3"].value.year == 2024 and ws["A3"].number_format == "yyyy-mm-dd"


def test_write_csv(sheets, data_dir: Path):
    sheets.write_range("orders_2024-01.csv", values=[["X-1", "C999", 1.5]], append=True)
    assert "X-1,C999,1.5" in (data_dir / "orders_2024-01.csv").read_text()


def test_read_only_blocks_writes(ro_sheets):
    with pytest.raises(ReadOnlyError):
        ro_sheets.write_range("sales.xlsx", cells={"A1": 1})
    # dry runs are allowed in read-only mode
    assert ro_sheets.write_range("sales.xlsx", cells={"A1": 1}, dry_run=True)["dry_run"]


def test_manage_sheets_rename_updates_formulas(sheets, data_dir: Path):
    res = sheets.manage_sheets("sales.xlsx", "rename", sheet="Sales", new_name="Sales Data")
    assert res["formulas_updated"] >= 4
    ws = openpyxl.load_workbook(data_dir / "sales.xlsx")["Summary"]
    assert ws["B2"].value == "=SUMIF('Sales Data'!B:B,A2,'Sales Data'!F:F)"


def test_manage_sheets_lifecycle(sheets):
    sheets.manage_sheets("book.xlsx", "create_workbook", sheets=["A", "B"])
    assert sheets.manage_sheets("book.xlsx", "add", new_name="C")["sheets"] == ["A", "B", "C"]
    assert sheets.manage_sheets("book.xlsx", "move", sheet="C", position=0)["sheets"][0] == "C"
    assert sheets.manage_sheets("book.xlsx", "copy", sheet="A", new_name="A2")["copy"] == "A2"
    assert "B" not in sheets.manage_sheets("book.xlsx", "delete", sheet="B")["sheets"]
    with pytest.raises(SheetError):
        sheets.manage_sheets("book.xlsx", "create_workbook")


def test_insert_rows_shifts_formulas(sheets, data_dir: Path):
    res = sheets.insert_delete("sales.xlsx", "insert_rows", 3, 2, sheet="Sales")
    assert res["formulas_updated"] > 0
    wb = openpyxl.load_workbook(data_dir / "sales.xlsx")
    assert wb["Sales"]["F5"].value == "=D5*E5"
    assert wb["Summary"]["E2"].value == "=AVERAGE(Sales!D2:D43)"
    assert wb["Sales"].tables["SalesTable"].ref == "A1:G43"


def test_delete_rows_creates_ref_errors(sheets, data_dir: Path):
    sheets.write_range("m.xlsx", values=[[1], [2], ["=A1+A2"], ["=SUM(A1:A2)"]])
    sheets.insert_delete("m.xlsx", "delete_rows", 1, 1)
    ws = openpyxl.load_workbook(data_dir / "m.xlsx").active
    assert ws["A2"].value == "=#REF!+A1"
    assert ws["A3"].value == "=SUM(A1:A1)"


def test_insert_cols_by_letter(sheets, data_dir: Path):
    sheets.insert_delete("sales.xlsx", "insert_cols", "B", 1, sheet="Sales")
    ws = openpyxl.load_workbook(data_dir / "sales.xlsx")["Sales"]
    assert ws["G2"].value == "=E2*F2"


def test_streaming_mode_for_big_files(sheets, monkeypatch):
    import sheetmcp.core.workbook as wbmod

    monkeypatch.setattr(wbmod, "BIG_FILE_BYTES", 0)  # treat every file as "big"
    d = sheets.describe("sales.xlsx")
    assert d["streaming_mode"] is True and d["sheets"][0]["max_row"] == 41
    r = sheets.read_range("sales.xlsx", "Summary", "A1:C6")
    assert r["formulas"]["B6"] == "=SUM(B2:B5)" and r["values"][5][1] > 0
