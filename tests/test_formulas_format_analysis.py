from pathlib import Path

import openpyxl
import pytest

from sheetmcp.core import SheetError


def test_formula_list_and_stats(sheets):
    r = sheets.formulas("sales.xlsx", "list", limit=5)
    assert r["total_formulas"] == 53 and len(r["formulas"]) == 5 and r["truncated"]
    s = sheets.formulas("sales.xlsx", "stats")
    assert s["functions_used"]["SUMIF"] == 4 and s["evaluator_unsupported"] == []


def test_formula_errors(sheets):
    r = sheets.formulas("sales.xlsx", "errors")
    cells = {(e["sheet"], e["cell"]): e for e in r["errors"]}
    assert cells[("Summary", "E6")]["error"] == "#DIV/0!"
    assert "IFERROR" in cells[("Summary", "E6")]["help"]


def test_formula_trace(sheets):
    r = sheets.formulas("sales.xlsx", "trace", sheet="Summary", cell="B6", depth=2)
    tree = r["precedents_tree"]
    assert tree["formula"] == "=SUM(B2:B5)"
    deps = {d["ref"] for d in r["dependents"]}
    assert "Summary!C2" in deps
    r2 = sheets.formulas("sales.xlsx", "trace", cell="Sales!D2")
    assert "Sales!F2" in {d["ref"] for d in r2["dependents"]}


def test_formula_evaluate(sheets):
    r = sheets.formulas("sales.xlsx", "evaluate", sheet="Sales", formula='=SUMIFS(F:F,B:B,"North")')
    by_region = sheets.query("SELECT sum(Revenue) FROM sales WHERE Region='North'")["rows"][0][0]
    assert r["value"] == pytest.approx(by_region)
    c = sheets.formulas("sales.xlsx", "evaluate", sheet="Summary", cell="B2:B6")
    assert len(c["results"]) == 5


def test_format_range(sheets, data_dir: Path):
    r = sheets.format_range(
        "sales.xlsx",
        "A1:G1",
        "Sales",
        bold=True,
        fill_color="lightyellow",
        font_color="darkblue",
        h_align="center",
        border="thin",
        autofit=True,
        freeze_panes="B2",
    )
    assert "font" in r["applied"] and "autofit" in r["applied"]
    sheets.format_range(
        "sales.xlsx",
        "F2:F41",
        "Sales",
        number_format="bdt",
        conditional={"type": "cell_is", "operator": ">", "values": [500], "fill_color": "lightgreen"},
    )
    sheets.format_range(
        "sales.xlsx", "B2:B41", "Sales", validation={"type": "list", "values": ["North", "South"]}
    )
    sheets.format_range("sales.xlsx", "D2:D41", "Sales", conditional={"type": "data_bar"})
    ws = openpyxl.load_workbook(data_dir / "sales.xlsx")["Sales"]
    assert ws["A1"].font.bold and ws["A1"].fill.fgColor.rgb.endswith("FFEB9C")
    assert ws["F2"].number_format.startswith('"৳"')
    assert ws.freeze_panes == "B2"
    assert len(ws.conditional_formatting) == 2
    assert ws.data_validations.dataValidation


def test_format_errors(sheets):
    with pytest.raises(SheetError):
        sheets.format_range("sales.xlsx", "A1", fill_color="notacolor")
    with pytest.raises(SheetError):
        sheets.format_range("customers.csv", "A1", bold=True)


def test_add_chart_and_charts_survive_writes(sheets, data_dir: Path):
    sheets.add_chart(
        "sales.xlsx", "B1:B5", "column", sheet="Summary", categories_range="A2:A5", title="Revenue"
    )
    sheets.add_chart("sales.xlsx", "B1:B5", "pie", sheet="Summary", categories_range="A2:A5", anchor="H20")
    sheets.write_range("sales.xlsx", "Summary", cells={"A8": "note"})
    wb = openpyxl.load_workbook(data_dir / "sales.xlsx")
    assert len(wb["Summary"]._charts) == 2


def test_list_files(sheets):
    r = sheets.list_files()
    names = {f["path"] for f in r["files"]}
    assert {"sales.xlsx", "customers.csv", "targets.parquet", "events.json"} <= names
    assert next(f for f in r["files"] if f["path"] == "sales.xlsx")["sheets"] == ["Sales", "Summary"]
    assert sheets.list_files(pattern="orders_*")["count"] == 3


def test_search(sheets):
    r = sheets.search("initech")
    assert any(m["file"] == "customers.csv" for m in r["matches"])
    f = sheets.search("VLOOKUP")
    assert any(m.get("formula", "").startswith("=VLOOKUP") for m in f["matches"])
    rx = sheets.search(r"^C10[0-2]$", regex=True, pattern="*.csv")
    assert rx["total_matches"] >= 3


def test_diff_cells(sheets, data_dir: Path):
    import shutil

    shutil.copy(data_dir / "sales.xlsx", data_dir / "sales_v2.xlsx")
    sheets.write_range("sales_v2.xlsx", "Sales", cells={"D2": 1000, "F3": "=D3*E3*2"})
    sheets.manage_sheets("sales_v2.xlsx", "add", new_name="Extra")
    d = sheets.diff("sales.xlsx", "sales_v2.xlsx")
    assert d["sheets_only_in_b"] == ["Extra"] and not d["identical"]
    diffs = {x["cell"]: x for x in d["sheets"][0]["differences"]}
    assert diffs["D2"]["b"] == 1000 and diffs["F3"]["formula_b"] == "=D3*E3*2"
    assert sheets.diff("sales.xlsx", "sales.xlsx", "Sales", "Sales")["identical"]


def test_diff_keyed(sheets, data_dir: Path):
    (data_dir / "a.csv").write_text("id,name,amt\n1,x,10\n2,y,20\n3,z,30\n")
    (data_dir / "b.csv").write_text("id,name,amt\n1,x,10\n2,y,25\n4,w,40\n")
    d = sheets.diff("a.csv", "b.csv", key=["id"])
    assert (d["added_count"], d["removed_count"], d["changed_count"]) == (1, 1, 1)
    assert d["changed"][0]["changes"] == {"amt": {"a": 20, "b": 25}}


def test_sandbox_paths(sheets):
    with pytest.raises(SheetError) as e:
        sheets.describe("../../etc/passwd")
    assert "outside" in str(e.value)
    with pytest.raises(SheetError) as e:
        sheets.describe("sale.xlsx")
    assert "sales.xlsx" in str(e.value)  # did-you-mean
    with pytest.raises(SheetError):
        sheets.describe("old.xls") if False else sheets.read_range("sales.xlsx", "NoSheet")


def test_legacy_xls_message(sheets, data_dir: Path):
    (data_dir / "old.xls").write_bytes(b"\xd0\xcf\x11\xe0")
    with pytest.raises(SheetError) as e:
        sheets.describe("old.xls")
    assert ".xlsx" in str(e.value)
