"""
Files openpyxl writes with formulas whose newer functions lack their prefix, for `writeKept`'s `prefixes: 'repair'`
(`prefixes.ts`): `python3 bench/coverage/keep/prefixes.py` writes them into `test/excel/keep/prefixes-made/`, where they
stay, as the cases are read from them. openpyxl 3.1 takes a formula verbatim: `ws["C1"] = "=XLOOKUP(...)"` is stored
`<f>XLOOKUP(...)</f>`, which Excel reads as a call of a name. Two files are changed after openpyxl, as other generators
write them: `shared` has its column of formulas made one shared formula, as Excel saves a formula filled down, and
`pivot` is openpyxl's save of the file Excel made with a chart and a pivot table (`keep/made/pivot.xlsx`), openpyxl
keeping the pivot table and dropping the chart, with a chart of openpyxl's own.
"""
import os
import re
import shutil
import sys
import tempfile
import zipfile

from openpyxl import Workbook, load_workbook
from openpyxl.chart import BarChart, Reference
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import PatternFill
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.formula import ArrayFormula
from openpyxl.worksheet.table import Table, TableColumn, TableFormula

ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", "..", ".."))
OUT = os.path.join(ROOT, "test", "excel", "keep", "prefixes-made")


def book():
    wb = Workbook()
    ws = wb.active
    ws.title = "S"
    return wb, ws


def lookup():
    wb, ws = book()
    ws["A1"], ws["A2"], ws["B1"], ws["B2"] = 1, 2, 10, 20
    ws["C1"] = "=XLOOKUP(2,A1:A2,B1:B2)"
    ws["C2"] = '=IFERROR(XLOOKUP(2,A1:A2,B1:B2),"fallback")'
    ws["C3"] = "=xlookup(1,A1:A2,B1:B2)"
    ws["D1"] = "=_xlfn._xlws.SORT(A1:A2)"
    chart = BarChart()
    chart.add_data(Reference(ws, min_col=2, min_row=1, max_row=2))
    chart.set_categories(Reference(ws, min_col=1, min_row=1, max_row=2))
    ws.add_chart(chart, "F2")
    return wb


def filter_():
    wb, ws = book()
    for r in range(1, 5):
        ws.cell(r, 1, r)
        ws.cell(r, 2, r * 10)
    ws["C1"] = "=FILTER(B1:B4,A1:A4>2)"
    ws["C2"] = "=SUM(FILTER(B1:B4,A1:A4>2))"
    ws["C3"] = "=UNIQUE(A1:A4)"
    ws["C4"] = '=IFERROR(FILTER(B1:B4,A1:A4>9),"none")'
    return wb


def sort():
    wb, ws = book()
    for r, v in enumerate([3, 1, 2], 1):
        ws.cell(r, 1, v)
    ws["C2"] = "=SORT(A1:A3)"
    ws["C3"] = "=SORT(A1:A3,1,-1)"
    ws["C4"] = "=SUM(SORTBY(A1:A3,A1:A3))"
    return wb


def shared():
    wb, ws = book()
    for r in range(1, 5):
        ws.cell(r, 1, r)
        ws.cell(r, 2, r * 10)
        ws.cell(r, 4, 5 - r)
        ws.cell(r, 5, "=XLOOKUP(D%d,$A$1:$A$4,$B$1:$B$4)" % r)
        ws.cell(r, 6, '=IFERROR(XLOOKUP(D%d+1,$A$1:$A$4,$B$1:$B$4),"none")' % r)
    return wb


def share(path):
    """The formulas of columns E and F as one shared formula each, its master the first cell."""
    def sheet(xml):
        for col, si in (("E", 0), ("F", 1)):
            def one(m):
                r, text = int(m.group(1)), m.group(2)
                if r == 1:
                    return '<c r="%s1"><f t="shared" ref="%s1:%s4" si="%d">%s</f>' % (col, col, col, si, text)
                return '<c r="%s%d"><f t="shared" si="%d"/>' % (col, r, si)
            xml = re.sub(r'<c r="%s(\d+)"><f>([^<]*)</f>' % col, one, xml)
        return xml
    rewrite(path, {"xl/worksheets/sheet1.xml": sheet})


def array():
    wb, ws = book()
    for r, v in enumerate([3, 1, 2], 1):
        ws.cell(r, 1, v)
    ws["C1"] = ArrayFormula("C1:C3", "=SORT(A1:A3)")
    ws["E1"] = ArrayFormula("E1:E3", "=SEQUENCE(3)*10")
    ws["G1"] = ArrayFormula("G1", "=SUM(XLOOKUP({1;2},A1:A3,A1:A3))")
    return wb


def name():
    wb, ws = book()
    ws["A1"], ws["A2"], ws["B1"], ws["B2"] = 1, 2, 10, 20
    wb.defined_names["Pick"] = DefinedName("Pick", attr_text="XLOOKUP(2,S!$A$1:$A$2,S!$B$1:$B$2)")
    ws.defined_names["Local"] = DefinedName("Local", attr_text="_xlfn.IFNA(XLOOKUP(9,S!$A$1:$A$2,S!$B$1:$B$2),0)")
    ws["C1"] = "=Pick"
    ws["C2"] = "=Pick*2"
    ws["C3"] = "=Local"
    return wb


def mixed():
    wb, ws = book()
    ws["A1"], ws["A2"], ws["B1"], ws["B2"] = 1, 2, 10, 20
    ws["C1"] = "=_xlfn.IFNA(XLOOKUP(5,A1:A2,B1:B2),DAYS(10,1))"
    ws["C2"] = "=LET(x,XLOOKUP(2,A1:A2,B1:B2),x*2)"
    ws["C3"] = "=SUM(MAKEARRAY(2,2,LAMBDA(r,c,r+c)))"
    ws["C4"] = "=LET(f,LAMBDA(x,[y],IF(ISOMITTED(y),x,x+y)),f(1)+f(1,2))"
    ws["C5"] = "=SUM(A1:A2)+_xlfn.DAYS(3,1)"
    return wb


def rules():
    wb, ws = book()
    ws["A1"], ws["A2"], ws["B1"], ws["B2"] = 1, 2, 10, 20
    fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
    ws.conditional_formatting.add("A1:A2", FormulaRule(formula=["XLOOKUP(A1,$A$1:$A$2,$B$1:$B$2)>15"], fill=fill))
    dv = DataValidation(type="custom", formula1="XLOOKUP(B1,$B$1:$B$2,$A$1:$A$2)>0")
    dv.add("B1:B2")
    ws.add_data_validation(dv)
    ws["C1"] = "=A1"
    return wb


def table():
    wb, ws = book()
    ws.append(["a", "b"])
    ws.append([1, None])
    ws.append([2, None])
    ws["E1"], ws["E2"], ws["F1"], ws["F2"] = 1, 2, 10, 20
    formula = "XLOOKUP(T[[#This Row],[a]],$E$1:$E$2,$F$1:$F$2)"
    ws["B2"] = "=" + formula
    ws["B3"] = "=" + formula
    t = Table(displayName="T", ref="A1:B3")
    t._initialise_columns()
    t.tableColumns[0].name, t.tableColumns[1].name = "a", "b"
    t.tableColumns[1].calculatedColumnFormula = TableFormula(attr_text=formula)
    ws.add_table(t)
    return wb


def pivot():
    wb = load_workbook(os.path.join(ROOT, "test", "excel", "keep", "made", "pivot.xlsx"))
    ws = wb.worksheets[0]
    ws["G1"] = '=XLOOKUP("x",A2:A21,B2:B21,"none")'
    ws["G2"] = "=XLOOKUP(A2,A2:A21,B2:B21)"
    chart = BarChart()
    chart.add_data(Reference(ws, min_col=2, min_row=1, max_row=5), titles_from_data=True)
    ws.add_chart(chart, "I2")
    return wb


def rewrite(path, parts):
    """The archive with the parts named rewritten by their functions of the text."""
    fd, tmp = tempfile.mkstemp(suffix=".xlsx")
    os.close(fd)
    with zipfile.ZipFile(path) as src, zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as out:
        for item in src.infolist():
            data = src.read(item.filename)
            if item.filename in parts:
                data = parts[item.filename](data.decode("utf-8")).encode("utf-8")
            out.writestr(item, data)
    shutil.move(tmp, path)


BOOKS = {"lookup": lookup, "filter": filter_, "sort": sort, "shared": shared, "array": array, "name": name, "mixed": mixed, "rules": rules, "table": table, "pivot": pivot}
AFTER = {"shared": share}

if __name__ == "__main__":
    os.makedirs(OUT, exist_ok=True)
    only = sys.argv[1:]
    for id, make in BOOKS.items():
        if only and id not in only:
            continue
        path = os.path.join(OUT, id + ".xlsx")
        make().save(path)
        if id in AFTER:
            AFTER[id](path)
        print(path)
