import openpyxl
import pytest

from sheetmcp_core import Evaluator


@pytest.fixture()
def ev():
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Data"
    rows = [
        ["Name", "Region", "Amount", "Qty"],
        ["a", "EU", 10, 1],
        ["b", "US", 20, 2],
        ["c", "EU", 30, 3],
        ["d", "APAC", 40, 4],
        ["e", "EU", None, 5],
    ]
    for r in rows:
        ws.append(r)
    o = wb.create_sheet("Other Sheet")
    o["A1"] = 100
    o["A2"] = "=Data!C2*2"
    o["B1"] = "=B1+1"  # circular
    from openpyxl.workbook.defined_name import DefinedName

    wb.defined_names["Rate"] = DefinedName("Rate", attr_text="'Other Sheet'!$A$1")
    return Evaluator(wb)


@pytest.mark.parametrize(
    ("formula", "expected"),
    [
        ("=1+2*3", 7),
        ("=(1+2)*3", 9),
        ("=2^3^2", 64),
        ("=-2^2", 4),
        ("=10/4", 2.5),
        ("=50%", 0.5),
        ('="a"&"b"&1', "ab1"),
        ("=1=1", True),
        ('="A"="a"', True),
        ("=3>2", True),
        ("=SUM(C2:C6)", 100),
        ("=AVERAGE(C2:C6)", 25),
        ("=COUNT(C2:C6)", 4),
        ("=COUNTA(A2:A6)", 5),
        ("=MAX(C2:C6)-MIN(C2:C6)", 30),
        ("=ROUND(2.345,2)", 2.35),
        ("=ROUND(-2.5,0)", -3),
        ("=INT(-1.5)", -2),
        ("=MOD(-3,2)", 1),
        ('=SUMIF(B2:B6,"EU",C2:C6)', 40),
        ('=SUMIFS(C2:C6,B2:B6,"EU",D2:D6,">1")', 30),
        ('=COUNTIF(B2:B6,"E*")', 3),
        ('=COUNTIFS(B2:B6,"<>EU")', 2),
        ('=AVERAGEIF(B2:B6,"EU",C2:C6)', 20),
        ('=VLOOKUP("c",A2:D6,3,FALSE)', 30),
        ('=INDEX(C2:C6,MATCH("d",A2:A6,0))', 40),
        ('=XLOOKUP("b",A2:A6,C2:C6)', 20),
        ('=XLOOKUP("zz",A2:A6,C2:C6,"none")', "none"),
        ('=IFERROR(1/0,"x")', "x"),
        ('=IF(C2>5,"big","small")', "big"),
        ("=IF(FALSE,1)", False),
        ("=AND(TRUE,C2>5)", True),
        ("=OR(FALSE,FALSE)", False),
        ('=UPPER(LEFT("hello",2))', "HE"),
        ('=MID("abcdef",2,3)', "bcd"),
        ('=TRIM("  a   b ")', "a b"),
        ('=LEN("abc")', 3),
        ('=SUBSTITUTE("a-b-c","-","+")', "a+b+c"),
        ('=TEXTJOIN(",",TRUE,A2:A4)', "a,b,c"),
        ("=SUMPRODUCT(C2:C5,D2:D5)", 300),
        ("='Other Sheet'!A1+1", 101),
        ("=Rate*2", 200),
        ("=YEAR(DATE(2024,14,1))", 2025),
        ("=ISNUMBER(C2)", True),
        ("=ISBLANK(C6)", True),
        ('=TEXT(0.256,"0.0%")', "25.6%"),
        ("=MEDIAN(1,3,2,10)", 2.5),
        ("=LARGE(C2:C6,2)", 30),
    ],
)
def test_formulas(ev, formula, expected):
    assert ev.evaluate_formula(formula, "Data") == expected


def test_errors(ev):
    assert ev.evaluate_formula("=1/0", "Data") == "#DIV/0!"
    assert ev.evaluate_formula('=VLOOKUP("zz",A2:D6,2,FALSE)', "Data") == "#N/A"
    assert ev.evaluate_formula('=1+"x"', "Data") == "#VALUE!"
    assert ev.evaluate_formula("=NOSUCHFUNC(1)", "Data") == "#NAME?"
    assert "NOSUCHFUNC" in ev.unsupported
    assert ev.evaluate_formula("=Missing!A1", "Data") == "#REF!"


def test_cell_eval_and_circular(ev):
    assert ev.evaluate_cell("Other Sheet", "A2") == 20
    assert ev.evaluate_cell("Other Sheet", "B1") == "#CIRC!"
