"""Generate realistic sample files used by the examples and the test-suite.

uv run python examples/make_samples.py [target_dir]
"""

from __future__ import annotations

import csv
import datetime as dt
import json
import random
import sys
import zipfile
from pathlib import Path

import openpyxl
from openpyxl.styles import Font, PatternFill
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.table import Table

REGIONS = ["North", "South", "East", "West"]
PRODUCTS = [("Widget", 12.5), ("Gadget", 30.0), ("Doohickey", 7.25), ("Gizmo", 55.0)]


def make(target: Path) -> Path:
    target.mkdir(parents=True, exist_ok=True)
    rnd = random.Random(42)

    # ---- sales.xlsx: data sheet with formulas, a summary sheet, a table, a named range
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Sales"
    ws.append(["Date", "Region", "Product", "Units", "Price", "Revenue", "Customer"])
    for c in ws[1]:
        c.font = Font(bold=True)
        c.fill = PatternFill("solid", fgColor="DDEBF7")
    start = dt.date(2024, 1, 1)
    for i in range(2, 42):
        prod, price = PRODUCTS[(i * 7) % 4]
        ws.append(
            [
                start + dt.timedelta(days=(i - 2) * 9),
                REGIONS[i % 4],
                prod,
                rnd.randint(1, 40),
                price,
                f"=D{i}*E{i}",
                f"C{100 + (i % 12)}",
            ]
        )
        ws[f"A{i}"].number_format = "yyyy-mm-dd"
        ws[f"F{i}"].number_format = '"$"#,##0.00'
    ws.freeze_panes = "A2"
    tab = Table(displayName="SalesTable", ref="A1:G41")
    ws.add_table(tab)
    s = wb.create_sheet("Summary")
    s.append(["Region", "Revenue", "Share"])
    for j, r in enumerate(REGIONS, start=2):
        s.append([r, f"=SUMIF(Sales!B:B,A{j},Sales!F:F)", f"=B{j}/$B$6"])
    s.append(["Total", "=SUM(B2:B5)", "=SUM(C2:C5)"])
    s["E1"] = "Avg units"
    s["E2"] = "=AVERAGE(Sales!D2:D41)"
    s["E3"] = "Lookup Gizmo price"
    s["E4"] = '=VLOOKUP("Gizmo",Sales!C2:E41,3,FALSE)'
    s["E5"] = "Broken"
    s["E6"] = "=B2/0"
    wb.defined_names["TotalRevenue"] = DefinedName("TotalRevenue", attr_text="Summary!$B$6")
    wb.save(target / "sales.xlsx")

    # ---- customers.csv (with a few quality problems on purpose)
    with open(target / "customers.csv", "w", newline="", encoding="utf-8") as f:
        w = csv.writer(f)
        w.writerow(["Customer", "Name", "Country", "Segment", "CreditLimit"])
        names = [
            "Acme Corp",
            "Globex",
            "Initech",
            "Umbrella",
            "Hooli",
            "Stark Ind",
            "Wayne Ent",
            "Wonka",
            "Cyberdyne",
            "Soylent",
            "Tyrell",
            "Aperture",
        ]
        countries = ["Bangladesh", "USA", "Germany", "India", "Japan", "Brazil"]
        for i, n in enumerate(names):
            limit = str(1000 * (i + 1)) if i != 5 else "N/A"
            w.writerow(
                [
                    f"C{100 + i}",
                    n if i != 3 else f" {n} ",
                    countries[i % 6],
                    ["SMB", "Enterprise"][i % 2],
                    limit,
                ]
            )
        w.writerow(["C102", "Initech", "Germany", "SMB", "3000"])  # duplicate row

    # ---- targets.parquet via DuckDB (no pyarrow needed)
    import duckdb

    con = duckdb.connect()
    con.execute(
        "COPY (SELECT * FROM (VALUES ('North', 9000), ('South', 7000), ('East', 8000), ('West', 6500)) t(Region, Target)) "
        f"TO '{(target / 'targets.parquet').as_posix()}' (FORMAT parquet)"
    )
    con.close()

    # ---- monthly CSVs to combine
    for m in ("2024-01", "2024-02", "2024-03"):
        with open(target / f"orders_{m}.csv", "w", newline="", encoding="utf-8") as f:
            w = csv.writer(f)
            w.writerow(["order_id", "customer", "amount"])
            for k in range(5):
                w.writerow([f"{m}-{k}", f"C{100 + k}", round(rnd.uniform(10, 500), 2)])

    # ---- events.json
    (target / "events.json").write_text(
        json.dumps(
            [
                {"event": "signup", "customer": "C100", "ts": "2024-01-03T10:00:00"},
                {"event": "purchase", "customer": "C101", "ts": "2024-01-05T12:30:00"},
                {"event": "churn", "customer": "C104", "ts": "2024-02-11T08:15:00"},
            ],
            indent=2,
        ),
        encoding="utf-8",
    )
    return target


ODS_CONTENT = """<?xml version="1.0" encoding="UTF-8"?>
<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
 xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0"
 xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" office:version="1.2">
<office:body><office:spreadsheet>{tables}</office:spreadsheet></office:body></office:document-content>"""


def _ods_cell(v: object) -> str:
    if v is None:
        return "<table:table-cell/>"
    if isinstance(v, dt.date):
        return (
            f'<table:table-cell office:value-type="date" office:date-value="{v.isoformat()}">'
            f"<text:p>{v.isoformat()}</text:p></table:table-cell>"
        )
    if isinstance(v, (int, float)):
        return f'<table:table-cell office:value-type="float" office:value="{v}"><text:p>{v}</text:p></table:table-cell>'
    t = str(v).replace("&", "&amp;").replace("<", "&lt;")
    return f'<table:table-cell office:value-type="string"><text:p>{t}</text:p></table:table-cell>'


def write_ods(path: Path, sheets: dict[str, list[list[object]]]) -> Path:
    """Tiny dependency-free .ods writer (enough for samples and tests)."""
    tables = ""
    for name, rows in sheets.items():
        body = "".join(
            "<table:table-row>" + "".join(_ods_cell(v) for v in r) + "</table:table-row>" for r in rows
        )
        tables += f'<table:table table:name="{name}">{body}</table:table>'
    manifest = (
        '<?xml version="1.0" encoding="UTF-8"?><manifest:manifest '
        'xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0" manifest:version="1.2">'
        '<manifest:file-entry manifest:full-path="/" '
        'manifest:media-type="application/vnd.oasis.opendocument.spreadsheet"/>'
        '<manifest:file-entry manifest:full-path="content.xml" manifest:media-type="text/xml"/>'
        "</manifest:manifest>"
    )
    with zipfile.ZipFile(path, "w") as z:
        z.writestr(zipfile.ZipInfo("mimetype"), "application/vnd.oasis.opendocument.spreadsheet")
        z.writestr("META-INF/manifest.xml", manifest)
        z.writestr("content.xml", ODS_CONTENT.format(tables=tables), zipfile.ZIP_DEFLATED)
    return path


def make_extras(target: Path) -> Path:
    """Extra formats: an OpenDocument spreadsheet and an HTML page with tables."""
    target.mkdir(parents=True, exist_ok=True)
    write_ods(
        target / "budget.ods",
        {
            "Budget": [
                ["Category", "Planned", "Actual", "Month"],
                ["Rent", 1200, 1200, dt.date(2024, 1, 1)],
                ["Food", 400, 455.5, dt.date(2024, 1, 1)],
                ["Travel", 150, None, dt.date(2024, 1, 1)],
            ],
            "Notes": [["Note"], ["Imported from LibreOffice"]],
        },
    )
    (target / "prices.html").write_text(
        """<html><body><h1>Price list</h1>
<table id="prices"><caption>Prices</caption>
<tr><th>Item</th><th>Price</th><th>Stock</th></tr>
<tr><td>Widget</td><td>12.50</td><td>100</td></tr>
<tr><td>Gadget</td><td>30</td><td>7</td></tr>
<tr><td colspan="2">Gizmo bundle</td><td>3</td></tr>
</table>
<table id="fx"><tr><th>Currency</th><th>Rate</th></tr><tr><td>BDT</td><td>117.5</td></tr></table>
</body></html>""",
        encoding="utf-8",
    )
    return target


if __name__ == "__main__":
    out = make(Path(sys.argv[1]) if len(sys.argv) > 1 else Path(__file__).parent / "data")
    make_extras(out)
    print(f"sample files written to {out}")
