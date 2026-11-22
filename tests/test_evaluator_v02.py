"""Formula functions and structured references added in 0.2."""

import datetime as dt

import openpyxl
import pytest
from openpyxl.worksheet.table import Table

from sheetmcp_core import Evaluator


@pytest.fixture()
def ev():
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Data"
    for r in [
        ["Name", "Region", "Amount", "Qty"],
        ["a", "EU", 10, 1],
        ["b", "US", 20, 2],
        ["c", "EU", 30, 3],
        ["d", "APAC", 40, 4],
        ["e", "EU", None, 5],
    ]:
        ws.append(r)
    ws.add_table(Table(displayName="T", ref="A1:D6"))
    ws["E2"] = "=[@Amount]*[@Qty]"
    ws["E3"] = "=T[[#This Row],[Amount]]+1"
    ws["F1"] = dt.date(2024, 1, 31)
    return Evaluator(wb)


@pytest.mark.parametrize(
    ("formula", "expected"),
    [
        ("=SUM(T[Amount])", 100),
        ("=COUNTA(T[#Headers])", 4),
        ("=ROWS(T[#Data])", 5),
        ("=ROWS(T[#All])", 6),
        ("=SUM(T[[Amount]:[Qty]])", 115),
        ("=LET(x,2,y,3,x*y+1)", 7),
        ('=SWITCH(2,1,"a",2,"b","z")', "b"),
        ('=SWITCH(9,1,"a","z")', "z"),
        ('=SUM(FILTER(C2:C5,B2:B5="EU"))', 40),
        ("=ROWS(UNIQUE(B2:B6))", 3),
        ("=INDEX(SORT(C2:C5,1,-1),1,1)", 40),
        ("=INDEX(SORTBY(A2:A5,C2:C5,-1),1,1)", "d"),
        ("=SUM(SEQUENCE(4))", 10),
        ("=INDEX(SEQUENCE(2,3,10,5),2,3)", 35),
        ("=COLUMNS(TRANSPOSE(A1:D1))", 1),
        ("=ROWS(VSTACK(A1:B2,A3:B4))", 4),
        ("=COLUMNS(HSTACK(A1:A2,B1:C2))", 3),
        ("=INDEX(TAKE(C2:C5,-1),1,1)", 40),
        ("=ROWS(DROP(C2:C5,1))", 3),
        ("=INDEX(CHOOSECOLS(A2:D2,3),1,1)", 10),
        ("=INDEX(CHOOSEROWS(A2:A5,-1),1,1)", "d"),
        ('=TEXTBEFORE("a-b-c","-")', "a"),
        ('=TEXTAFTER("a-b-c","-",-1)', "c"),
        ('=COLUMNS(TEXTSPLIT("a,b,c",","))', 3),
        ("=CHAR(65)", "A"),
        ('=CODE("A")', 65),
        ("=UNICHAR(2547)", "৳"),
        ('=CLEAN("a"&CHAR(7)&"b")', "ab"),
        ('=NUMBERVALUE("1.234,5",",",".")', 1234.5),
        ("=FIXED(1234.567,1)", "1,234.6"),
        ("=T(5)", ""),
        ('=N("x")', 0),
        ("=ROUND(PMT(0.05/12,360,200000),2)", -1073.64),
        ("=ROUND(FV(0.05,10,-100),2)", 1257.79),
        ("=ROUND(PV(0.05,10,-100),2)", 772.17),
        ("=ROUND(NPER(0.01,-100,1000),4)", 10.5886),
        ("=ROUND(RATE(10,-100,800),6)", 0.042775),
        ("=ROUND(NPV(0.1,100,100),4)", 173.5537),
        ("=ROUND(IRR({-100,30,40,50}),5)", 0.08896),
        ("=EOMONTH(F1,1)", dt.datetime(2024, 2, 29, 0, 0)),
        ("=WEEKDAY(DATE(2024,1,1))", 2),
        ("=WEEKNUM(DATE(2024,3,15))", 11),
        ("=ISOWEEKNUM(DATE(2024,1,1))", 1),
        ("=NETWORKDAYS(DATE(2024,1,1),DATE(2024,1,31))", 23),
        ("=WORKDAY(DATE(2024,1,5),1)", dt.datetime(2024, 1, 8, 0, 0)),
        ('=DATEDIF(DATE(2020,1,15),DATE(2024,1,14),"Y")', 3),
        ("=HOUR(TIME(13,45,10))", 13),
        ("=MINUTE(0.75+1/1440)", 1),
        ("=SECOND(TIME(1,2,3))", 3),
        ('=DATEVALUE("2024-03-01")', dt.datetime(2024, 3, 1, 0, 0)),
        ("=SUM(C2:C4*D2:D4)", 140),
        ('=SUM(IF(B2:B6="EU",D2:D6,0))', 9),
        ("=SUM(-C2:C3)", -30),
    ],
)
def test_new_functions(ev, formula, expected):
    assert ev.evaluate_formula(formula, "Data") == expected


def test_this_row_references(ev):
    assert ev.evaluate_cell("Data", "E2") == 10
    assert ev.evaluate_cell("Data", "E3") == 21


def test_structured_ref_errors(ev):
    assert ev.evaluate_formula("=SUM(T[Nope])", "Data") == "#REF!"
    assert ev.evaluate_formula("=SUM(Missing[Amount])", "Data") in ("#REF!", "#NAME?")
    assert ev.evaluate_formula("=[1]Sheet1!A1", "Data") in ("#REF!", "#NAME?")


def test_let_scoping_does_not_leak(ev):
    assert ev.evaluate_formula("=LET(Amount,5,Amount*2)", "Data") == 10
    assert ev.evaluate_formula("=LET(x,1,x)+LET(x,2,x)", "Data") == 3


def test_short_names_are_not_columns():
    from openpyxl.workbook.defined_name import DefinedName

    wb = openpyxl.Workbook()
    ws = wb.active
    for i in range(1, 4):
        ws.cell(row=i, column=1, value=i)
    wb.defined_names["Ids"] = DefinedName("Ids", attr_text="Sheet!$A$1:$A$3")
    wb.defined_names["Tax"] = DefinedName("Tax", attr_text="0.5")
    e = Evaluator(wb)
    assert e.evaluate_formula("=SUM(Ids)*Tax", "Sheet") == 3
    assert e.evaluate_formula("=SUM(A:A)", "Sheet") == 6
    assert e.evaluate_formula("=SUM(1:2)", "Sheet") == 3
