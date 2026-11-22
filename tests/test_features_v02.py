"""Tests for the 0.2 features: legacy formats, quick_look, clean, objects, backups, batch, reports and more."""

from __future__ import annotations

import json
from pathlib import Path

import make_samples  # registered by conftest
import openpyxl
import pytest

from sheetmcp_core import Config, SheetError, Sheets


@pytest.fixture()
def extra_dir(data_dir: Path) -> Path:
    make_samples.make_extras(data_dir)
    return data_dir


# ------------------------------------------------------------------ legacy formats
def test_ods_read_query_convert(extra_dir: Path, sheets: Sheets):
    d = sheets.describe("budget.ods")
    assert d["kind"] == "ods" and [s["name"] for s in d["sheets"]] == ["Budget", "Notes"]
    r = sheets.query("SELECT sum(Planned) AS p, count(Actual) AS a FROM budget")
    assert r["rows"] == [[1750, 2]]
    rr = sheets.read_range("budget.ods", "Notes")
    assert rr["values"][1][0] == "Imported from LibreOffice"
    out = sheets.convert("budget.ods", "budget.xlsx")
    assert out["rows_written"] == 3
    with pytest.raises(SheetError, match=r"read-only|convert"):
        sheets.write_range("budget.ods", cells={"A1": 1})


def test_html_tables(extra_dir: Path, sheets: Sheets):
    names = [s["name"] for s in sheets.describe("prices.html")["sheets"]]
    assert names == ["prices", "fx"]
    r = sheets.query("SELECT Item, Price FROM prices ORDER BY Item")
    assert ["Widget", 12.5] in r["rows"]
    assert sheets.query("SELECT Rate FROM prices__fx")["rows"] == [[117.5]]


def test_xls_read(data_dir: Path, sheets: Sheets):
    xlwt = pytest.importorskip("xlwt")
    pytest.importorskip("xlrd")
    wb = xlwt.Workbook()
    ws = wb.add_sheet("Old")
    for c, h in enumerate(["Item", "Qty"]):
        ws.write(0, c, h)
    ws.write(1, 0, "a")
    ws.write(1, 1, 3)
    ws.write(2, 0, "b")
    ws.write(2, 1, 4.5)
    wb.save(str(data_dir / "legacy.xls"))
    assert sheets.query("SELECT sum(Qty) FROM legacy")["rows"] == [[7.5]]
    assert sheets.list_files(pattern="*.xls")["files"][0]["sheets"] == ["Old"]


def test_xls_without_xlrd_gives_hint(data_dir: Path, sheets: Sheets, monkeypatch):
    import builtins

    (data_dir / "fake.xls").write_bytes(b"\xd0\xcf\x11\xe0" + b"\0" * 100)
    real = builtins.__import__

    def no_xlrd(name, *a, **k):
        if name == "xlrd":
            raise ImportError("no xlrd")
        return real(name, *a, **k)

    monkeypatch.setattr(builtins, "__import__", no_xlrd)
    with pytest.raises(SheetError, match=r"sheetmcp\[xls\]"):
        sheets.describe("fake.xls")


# ------------------------------------------------------------------ write improvements
def test_fill_formula(sheets: Sheets, data_dir: Path):
    r = sheets.write_range("sales.xlsx", "Sales", fill_formula="=D2*E2*2", fill_range="H2:H")
    assert r["cells_written"] == 40
    ws = openpyxl.load_workbook(data_dir / "sales.xlsx")["Sales"]
    assert ws["H2"].value == "=D2*E2*2" and ws["H41"].value == "=D41*E41*2"
    v = sheets.read_range("sales.xlsx", "Sales", "H2:H2")["values"][0][0]
    assert v == ws["D2"].value * ws["E2"].value * 2
    with pytest.raises(SheetError):
        sheets.write_range("customers.csv", fill_formula="=A1", fill_range="F2")


def test_csv_streaming_append(sheets: Sheets, data_dir: Path):
    before = (data_dir / "customers.csv").read_text(encoding="utf-8")
    r = sheets.write_range("customers.csv", values=[["C999", "New Co", "Nepal", "SMB", 10]], append=True)
    assert r.get("streamed") or r["cells_written"] == 5
    after = (data_dir / "customers.csv").read_text(encoding="utf-8")
    assert after.startswith(before.rstrip("\r\n")) and after.rstrip().endswith("C999,New Co,Nepal,SMB,10")
    assert sheets.query("SELECT count(*) FROM customers")["rows"] == [[14]]


def test_insert_shifts_conditional_formats_and_validation(sheets: Sheets, data_dir: Path):
    sheets.format_range(
        "sales.xlsx", "F2:F41", "Sales", conditional={"type": "cell_is", "operator": ">", "values": [500]}
    )
    sheets.format_range(
        "sales.xlsx", "B2:B41", "Sales", validation={"type": "list", "values": ["North", "South"]}
    )
    r = sheets.insert_delete("sales.xlsx", "insert_rows", 5, 3, "Sales")
    assert r["conditional_formats_shifted"] >= 1 and r["validations_shifted"] >= 1
    ws = openpyxl.load_workbook(data_dir / "sales.xlsx")["Sales"]
    assert [str(cf.sqref) for cf in ws.conditional_formatting] == ["F2:F44"]
    assert str(ws.data_validations.dataValidation[0].sqref) == "B2:B44"
    sheets.insert_delete("sales.xlsx", "delete_cols", "A", 1, "Sales")
    ws = openpyxl.load_workbook(data_dir / "sales.xlsx")["Sales"]
    assert [str(cf.sqref) for cf in ws.conditional_formatting] == ["E2:E44"]
    assert str(ws.data_validations.dataValidation[0].sqref) == "A2:A44"


# ------------------------------------------------------------------ profile / convert
def test_profile_anomalies_and_inconsistent_categories(data_dir: Path, sheets: Sheets):
    rows = ["id,city,amount"] + [
        f"{i},{c},{a}"
        for i, (c, a) in enumerate(
            [
                ("Dhaka", 10),
                ("dhaka ", 11),
                ("DHAKA", 9),
                ("Paris", 10),
                ("Paris", 12),
                ("Rome", 11),
                ("Rome", 10),
                ("Oslo", 9),
                ("Oslo", 10),
                ("Lima", 11),
                ("Lima", 10),
                ("Kyiv", 5000),
            ],
            start=1,
        )
    ]
    (data_dir / "m.csv").write_text("\n".join(rows) + "\n", encoding="utf-8")
    p = sheets.profile("m.csv")
    kinds = {(i["column"], i["issue"]) for i in p["issues"]}
    assert ("city", "inconsistent_categories") in kinds
    assert ("amount", "anomalies") in kinds
    amount = next(c for c in p["columns"] if c["name"] == "amount")
    assert amount["anomalies"][0]["row"] == 13 and amount["anomalies"][0]["value"] == 5000


def test_csv_dialects(sheets: Sheets, data_dir: Path):
    sheets.convert("sales.xlsx", "excel.csv", csv_dialect="excel")
    assert (data_dir / "excel.csv").read_bytes().startswith(b"\xef\xbb\xbf")
    sheets.convert("sales.xlsx", "gs.csv", csv_dialect="google_sheets")
    first = (data_dir / "gs.csv").read_text(encoding="utf-8").splitlines()[1]
    assert first.startswith("2024-01-01")
    plain = (data_dir / "excel.csv").read_bytes()[3:]
    assert plain.splitlines()[0].startswith(b"Date,Region")
    with pytest.raises(SheetError):
        sheets.convert("sales.xlsx", "bad.csv", csv_dialect="nope")


# ------------------------------------------------------------------ quick_look / catalog
def test_quick_look_folder_and_file(extra_dir: Path, sheets: Sheets):
    q = sheets.quick_look(".")
    by = {f["path"]: f for f in q["files"]}
    assert by["customers.csv"]["sql"] == "customers" and by["customers.csv"]["rows"] == 13
    assert (
        by["sales.xlsx"]["sheets"][0]["sql"] == "sales"
        and by["sales.xlsx"]["sheets"][1]["sql"] == "sales__summary"
    )
    assert "budget.ods" in by and q["next_steps"]
    f = sheets.quick_look("customers.csv")
    assert f["rows"] == 13 and f["sql_table"] == "customers"
    assert any("duplicate_rows" in i for i in f["issues"])
    assert f["next_steps"][0].startswith("clean")
    x = sheets.quick_look("sales.xlsx")
    assert x["formulas"]["count"] > 0 and x["preview"]["columns"][0] == "Date"


def test_catalog_via_query_without_sql(data_dir: Path, sheets: Sheets):
    c = sheets.query()
    names = {t["table"] for t in c["tables"]}
    assert {"customers", "sales", "sales__summary", "targets", "events"} <= names
    sales = next(t for t in c["tables"] if t["table"] == "sales")
    assert "Revenue" in sales["columns"]


# ------------------------------------------------------------------ clean
def test_clean_preview_and_save(data_dir: Path, sheets: Sheets):
    ops = [
        {"op": "trim"},
        {"op": "dedupe"},
        {"op": "to_number", "columns": ["CreditLimit"]},
        {"op": "fill_null", "columns": ["CreditLimit"], "value": 0},
        {"op": "case", "mode": "upper", "columns": ["Segment"]},
        {"op": "filter", "where": "CreditLimit >= 0"},
        {"op": "sort", "order_by": ["CreditLimit DESC"]},
    ]
    p = sheets.clean("customers.csv", ops, limit=3)
    assert p["rows_before"] == 13 and p["rows_after"] == 12 and "note" in p
    assert p["rows"][0][4] == 12000 and p["rows"][0][3] in ("SMB", "ENTERPRISE")
    assert not (data_dir / "c.xlsx").exists()
    s = sheets.clean("customers.csv", ops, output="c.xlsx")
    assert s["rows_written"] == 12
    assert sheets.query("SELECT count(*) FROM 'c.xlsx' WHERE Name <> trim(Name)")["rows"] == [[0]]


def test_clean_text_ops(data_dir: Path, sheets: Sheets):
    (data_dir / "people.csv").write_text(
        "name,city,joined,amount\n"
        'john   smith,dhaka,15/01/2024,"1,200"\n'
        "jane doe,,03/02/2024,$300\n"
        ",,,\n"
        "ann lee,,2024-03-05,n/a\n",
        encoding="utf-8",
    )
    ops = [
        {"op": "drop_empty_rows"},
        {"op": "trim"},
        {"op": "case", "mode": "title", "columns": ["name"]},
        {"op": "split", "column": "name", "sep": " ", "into": ["first", "last"]},
        {"op": "fill_down", "columns": ["city"]},
        {"op": "to_date", "columns": ["joined"], "dayfirst": True},
        {"op": "to_number", "columns": ["amount"]},
        {
            "op": "replace",
            "find": "^Dh",
            "replace": "DH",
            "columns": ["city"],
            "regex": True,
            "case_sensitive": False,
        },
        {"op": "merge", "columns": ["first", "city"], "into": "tag", "sep": "@"},
        {"op": "add_column", "name": "double", "expr": "amount * 2"},
        {"op": "rename", "mapping": {"double": "amount_x2"}},
        {"op": "drop_columns", "columns": ["last"]},
    ]
    r = sheets.clean("people.csv", ops, format="records")
    rows = r["rows"]
    assert len(rows) == 3
    assert rows[0]["name"] == "John Smith" and rows[0]["first"] == "John"
    assert [x["city"] for x in rows] == ["DHaka"] * 3
    assert str(rows[0]["joined"]) == "2024-01-15" and str(rows[1]["joined"]) == "2024-02-03"
    assert str(rows[2]["joined"]) == "2024-03-05"
    assert rows[0]["amount"] == 1200 and rows[1]["amount"] == 300 and rows[2]["amount"] is None
    assert rows[0]["tag"] == "John@DHaka" and rows[0]["amount_x2"] == 2400 and "last" not in rows[0]


def test_clean_errors(sheets: Sheets):
    with pytest.raises(SheetError, match="Unknown clean op"):
        sheets.clean("customers.csv", [{"op": "explode"}])
    with pytest.raises(SheetError, match="not found"):
        sheets.clean("customers.csv", [{"op": "trim", "columns": ["Nope"]}])
    with pytest.raises(SheetError):
        sheets.clean("customers.csv", [{"op": "filter", "where": "1=1; DROP TABLE x"}])
    with pytest.raises(SheetError):
        sheets.clean("customers.csv", [])


def test_clean_in_place_keeps_workbook(sheets: Sheets, data_dir: Path):
    sheets.write_range("sales.xlsx", "Sales", cells={"B3": None, "B4": " "})
    r = sheets.clean("sales.xlsx#Sales", [{"op": "fill_down", "columns": ["Region"]}], output="sales.xlsx")
    assert r["in_place"] and r["backup"] and "SalesTable" in r.get("tables_resized", [])
    wb = openpyxl.load_workbook(data_dir / "sales.xlsx")
    ws = wb["Sales"]
    assert ws["B3"].value == ws["B2"].value and ws["B4"].value == ws["B2"].value
    assert "Summary" in wb.sheetnames and "SalesTable" in ws.tables
    assert ws["A1"].font.bold  # header style kept


def test_clean_read_only(ro_sheets: Sheets):
    assert ro_sheets.clean("customers.csv", [{"op": "trim"}], limit=1)["rows_after"] == 13
    with pytest.raises(SheetError, match="read-only"):
        ro_sheets.clean("customers.csv", [{"op": "trim"}], output="x.csv")


# ------------------------------------------------------------------ Excel objects
def test_manage_objects_tables_and_names(sheets: Sheets, data_dir: Path):
    sheets.clean("customers.csv", [{"op": "to_number", "columns": ["CreditLimit"]}], output="c.xlsx")
    t = sheets.manage_objects("c.xlsx", "table_create", range="A1:E14", name="Cust")
    assert t["ref"] == "A1:E14"
    with pytest.raises(SheetError, match="already exists"):
        sheets.manage_objects("c.xlsx", "table_create", range="G1:H3", name="cust")
    with pytest.raises(SheetError, match="overlaps"):
        sheets.manage_objects("c.xlsx", "table_create", range="B2:C5", name="Other")
    with pytest.raises(SheetError, match="valid table name"):
        sheets.manage_objects("c.xlsx", "table_create", range="G1:H3", name="A1")
    sheets.manage_objects("c.xlsx", "name_create", name="Rate", refers_to="=0.1")
    sheets.manage_objects("c.xlsx", "name_create", name="Ids", range="A2:A14")
    sheets.write_range("c.xlsx", cells={"G1": "=SUM(Cust[CreditLimit])*Rate", "G2": "=COUNTA(Ids)"})
    vals = sheets.read_range("c.xlsx", range="G1:G2")["values"]
    assert vals[0][0] == pytest.approx(7500) and vals[1][0] == 13
    lst = sheets.manage_objects("c.xlsx")
    assert lst["tables"][0]["name"] == "Cust"
    assert {n["name"] for n in lst["named_ranges"]} == {"Rate", "Ids"}
    sheets.manage_objects("c.xlsx", "name_delete", name="Rate")
    sheets.manage_objects("c.xlsx", "table_delete", name="Cust")
    lst = sheets.manage_objects("c.xlsx")
    assert not lst["tables"] and [n["name"] for n in lst["named_ranges"]] == ["Ids"]


def test_manage_objects_rules_and_dry_run(sheets: Sheets, data_dir: Path):
    sheets.format_range("sales.xlsx", "F2:F41", "Sales", conditional={"type": "data_bar"})
    sheets.format_range("sales.xlsx", "B2:B41", "Sales", validation={"type": "list", "values": ["North"]})
    lst = sheets.manage_objects("sales.xlsx")
    assert lst["validations"] and lst["conditional_formats"]
    d = sheets.manage_objects("sales.xlsx", "conditional_delete", "Sales", range="A1:A5", dry_run=True)
    assert d["removed"] == 0
    assert sheets.manage_objects("sales.xlsx", "conditional_delete", "Sales")["removed"] == 1
    assert sheets.manage_objects("sales.xlsx", "validation_delete", "Sales", range="B10")["removed"] == 1
    lst = sheets.manage_objects("sales.xlsx")
    assert not lst["validations"] and not lst["conditional_formats"]
    with pytest.raises(SheetError):
        sheets.manage_objects("customers.csv")


# ------------------------------------------------------------------ backups / undo
def test_backups_undo_restore(sheets: Sheets, data_dir: Path):
    sheets.write_range("sales.xlsx", "Sales", cells={"B2": "one"})
    sheets.write_range("sales.xlsx", "Sales", cells={"B2": "two"})
    sheets.write_range("sales.xlsx", "Sales", cells={"B2": "three"})

    def b2() -> str:
        return openpyxl.load_workbook(data_dir / "sales.xlsx")["Sales"]["B2"].value

    lst = sheets.backups("sales.xlsx")
    assert lst["count"] == 3
    u = sheets.backups("sales.xlsx", "undo")
    assert b2() == "two"
    sheets.backups("sales.xlsx", "undo")
    assert b2() == "one"
    sheets.backups(action="restore", backup=u["previous_state_saved_as"])
    assert b2() == "three"
    first = sheets.backups("sales.xlsx")["backups"]
    assert any(b["kind"] == "pre-undo snapshot" for b in first)
    victim = first[-1]["backup"]
    sheets.backups(action="delete", backup=victim)
    assert not (data_dir / victim).exists()
    with pytest.raises(SheetError, match="not a SheetMCP backup"):
        sheets.backups(action="restore", backup="customers.csv")
    with pytest.raises(SheetError, match="No backups"):
        sheets.backups("customers.csv", "undo")


def test_backups_all_and_read_only(sheets: Sheets, ro_sheets: Sheets):
    sheets.write_range("customers.csv", cells={"A2": "X"})
    assert sheets.backups()["count"] == 1
    assert ro_sheets.backups()["count"] == 1
    with pytest.raises(SheetError, match="read-only"):
        ro_sheets.backups("customers.csv", "undo")


# ------------------------------------------------------------------ batch
def test_batch(sheets: Sheets, data_dir: Path):
    r = sheets.batch(
        [
            {"tool": "write_range", "args": {"path": "new.xlsx", "values": [["a", "b"], [1, 2]]}},
            {"tool": "format_range", "args": {"path": "new.xlsx", "range": "A1:B1", "bold": True}},
            {"tool": "nope", "args": {}},
            {"tool": "read_range", "args": {"path": "new.xlsx"}},
        ]
    )
    assert r["succeeded"] == 2 and r["stopped_early"] and len(r["results"]) == 3
    r = sheets.batch(
        [
            {"tool": "nope"},
            {"tool": "describe", "args": {"path": "new.xlsx", "bogus": 1}},
            {"tool": "query", "args": {"sql": "SELECT 1 AS x"}},
        ],
        stop_on_error=False,
    )
    assert r["succeeded"] == 1 and "Bad arguments" in r["results"][1]["error"]
    with pytest.raises(SheetError):
        sheets.batch([])


# ------------------------------------------------------------------ report builder
def test_build_report(sheets: Sheets, data_dir: Path):
    r = sheets.build_report(
        "sales.xlsx",
        "report.xlsx",
        ["Region"],
        ["Revenue:sum", "Units:avg", "count"],
        title="Sales by region",
        number_format="[$৳-845]#,##0.00",
    )
    assert r["sheets"] == ["Summary", "Data"] and r["summary_rows"] == 4 and r["data_rows"] == 40
    wb = openpyxl.load_workbook(data_dir / "report.xlsx")
    ws = wb["Summary"]
    assert [c.value for c in ws[1]][:4] == ["Region", "Total Revenue", "Average Units", "Count"]
    assert (
        ws["A6"].value == "Total" and ws["B6"].value == "=SUM(B2:B5)" and ws["C6"].value == "=AVERAGE(C2:C5)"
    )
    assert ws["B2"].value >= ws["B3"].value  # sorted by first metric
    assert ws["B2"].number_format == "[$৳-845]#,##0.00" and ws["A1"].font.bold and ws.freeze_panes == "A2"
    assert len(ws._charts) == 1 and wb["Data"].auto_filter.ref == "A1:G41"
    total = sheets.read_range("report.xlsx", "Summary", "D6")["values"][0][0]
    assert total == 40
    assert (
        not list((data_dir / ".sheetmcp_backups").glob("report.*"))
        if (data_dir / ".sheetmcp_backups").exists()
        else True
    )
    with pytest.raises(SheetError, match="already exists"):
        sheets.build_report("sales.xlsx", "report.xlsx", ["Region"], ["count"])
    r2 = sheets.build_report(
        "sales.xlsx",
        "report.xlsx",
        ["Product"],
        ["Revenue:sum"],
        where="Units > 5",
        top=2,
        include_data=False,
        chart_type="none",
        overwrite=True,
    )
    assert r2["summary_rows"] == 2 and r2["sheets"] == ["Summary"] and r2["backup_of_previous"]
    with pytest.raises(SheetError):
        sheets.build_report("sales.xlsx", "r.csv", ["Region"], ["count"])


# ------------------------------------------------------------------ CLI additions
def test_sheetcli_new_commands(data_dir: Path, capsys, monkeypatch):
    from sheetmcp_cli.main import main

    monkeypatch.chdir(data_dir)
    assert main(["look"]) == 0
    assert "customers" in capsys.readouterr().out
    assert main(["clean", "customers.csv", "trim", "dedupe", "to_number:CreditLimit", "-o", "c.csv"]) == 0
    assert json.loads(capsys.readouterr().out)["rows_written"] == 12
    assert main(["report", "sales.xlsx", "rep.xlsx", "-g", "Region", "-m", "Revenue:sum"]) == 0
    capsys.readouterr()
    assert main(["write", "customers.csv", "A2=Z"]) == 0
    capsys.readouterr()
    assert main(["backups", "customers.csv", "--undo"]) == 0
    assert "restored" in capsys.readouterr().out
    assert (data_dir / "customers.csv").read_text(encoding="utf-8").splitlines()[1].startswith("C100")


def test_info_mentions_new_formats(sheets: Sheets):
    info = Sheets(Config(roots=[str(sheets.sb.roots[0])])).info()
    assert info["version"] == "0.2.0" and "ods" in info["supported_inputs"]


def test_csv_cache_has_clean_line_endings(sheets: Sheets, data_dir: Path):
    from sheetmcp_core.fileio import excel_to_csv_cache

    out, _ = excel_to_csv_cache(data_dir / "sales.xlsx", "Sales")
    raw = out.read_bytes()
    assert b"\r\r\n" not in raw and raw.count(b"\r\n") == 41
