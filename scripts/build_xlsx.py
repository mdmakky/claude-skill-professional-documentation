#!/usr/bin/env python3
"""Render a JSON workbook spec to a formatted XLSX (live formulas, styled inputs, frozen headers, filters, charts).
Usage: build_xlsx.py spec.json out.xlsx   (refuses to overwrite; use --force)
Spec: {title?, sheets: [Sheet...]}
Sheet kinds:
 {"name","kind":"inputs","title?","rows":[{"label","name?","value","fmt?","tag":"verified|assumption","source?","note?"}]}
     -> cols Label|Value|Unit/format|Status|Source|Note. Value cells are styled as inputs; "name" makes a workbook
        defined name (letters/digits/_) usable in formulas on other sheets, e.g. =Ad_Spend*CTR.
 {"name","kind":"table","title?","columns":[{"h","fmt?","width?","total?":"sum|avg"}],"rows":[[cell,...]],
     "filter?":true,"cond?":[{"range","op":"lt|gt|eq|between","value","value2?","color":"red|green|amber"}],
     "chart?":{"type":"bar|line|pie","title","cats":"A5:A9","vals":["C4:C9"],"anchor":"H3","x_title?","y_title?","from_rows?":true (series in rows; first cell = series name)}}
Cell: number | text | "=FORMULA" (formula) | {"v":..,"input":true} (input styling). `{r}` in a formula = current row number.
fmt: text int num2 usd usd0 pct pct2 date."""
import json, os, sys
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.formatting.rule import CellIsRule
from openpyxl.chart import BarChart, LineChart, PieChart, Reference
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.utils import get_column_letter as L

FMT = {"text": "@", "int": "#,##0", "num2": "#,##0.00", "usd": '"$"#,##0.00', "usd0": '"$"#,##0', "pct": "0.0%", "pct2": "0.00%", "date": "yyyy-mm-dd"}
UNITS = {"text": "Text", "int": "Count", "num2": "Number", "usd": "USD", "usd0": "USD", "pct": "Percent", "pct2": "Percent", "date": "Date"}
NAVY, ZEBRA = "1F3864", "F2F5FA"
FONT = "Calibri"
thin = Side(style="thin", color="BFC5D2"); BORDER = Border(left=thin, right=thin, top=thin, bottom=thin)
INPUT_FILL = PatternFill("solid", fgColor="FFF8DC"); INPUT_FONT = Font(name=FONT, size=10, color="0000FF")
CALC_FONT = Font(name=FONT, size=10, color="000000")
HDR_FILL = PatternFill("solid", fgColor=NAVY); HDR_FONT = Font(name=FONT, size=10, bold=True, color="FFFFFF")
COLORS = {"red": ("FFC7CE", "9C0006"), "green": ("C6EFCE", "006100"), "amber": ("FFEB9C", "9C5700")}

def put(ws, r, c, v, fmt=None, input_=False, row=None):
    if isinstance(v, dict): input_ = v.get("input", input_); v = v["v"]
    if isinstance(v, str) and v.startswith("="): v = v.replace("{r}", str(row or r))
    cell = ws.cell(r, c, v); cell.border = BORDER
    cell.font = INPUT_FONT if input_ else CALC_FONT
    if input_: cell.fill = INPUT_FILL
    if fmt: cell.number_format = FMT.get(fmt, fmt)
    if isinstance(v, str) and not v.startswith("="): cell.alignment = Alignment(wrap_text=True, vertical="top")
    return cell

def header(ws, r, heads):
    for i, h in enumerate(heads, 1):
        c = ws.cell(r, i, h); c.fill, c.font, c.border = HDR_FILL, HDR_FONT, BORDER
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

def title(ws, sheet, wb_title):
    ws["A1"] = sheet.get("title") or wb_title or sheet["name"]; ws["A1"].font = Font(name=FONT, size=14, bold=True, color=NAVY)
    ws["A2"] = sheet.get("subtitle", "Blue text on cream = input. Black = calculated."); ws["A2"].font = Font(name=FONT, size=9, italic=True, color="595959")
    ws.sheet_view.showGridLines = False

def build(spec, out):
    wb = Workbook(); wb.remove(wb.active)
    for s in spec["sheets"]:
        ws = wb.create_sheet(s["name"][:31]); title(ws, s, spec.get("title"))
        if s["kind"] == "inputs":
            header(ws, 4, ["Input", "Value", "Unit / format", "Status", "Source", "Note"])
            for i, row in enumerate(s["rows"], 5):
                put(ws, i, 1, row["label"])
                put(ws, i, 2, row["value"], row.get("fmt"), input_=True)
                put(ws, i, 3, UNITS.get(row.get("fmt", "text"), row.get("fmt"))); put(ws, i, 4, row.get("tag", "assumption"))
                put(ws, i, 5, row.get("source", "")); put(ws, i, 6, row.get("note", ""))
                if row.get("name"):
                    wb.defined_names[row["name"]] = DefinedName(row["name"], attr_text=f"'{ws.title}'!$B${i}")
            for col, w in zip("ABCDEF", (34, 16, 14, 13, 38, 44)): ws.column_dimensions[col].width = w
            ws.freeze_panes = "A5"
            last = 4 + len(s["rows"])
            ws.conditional_formatting.add(f"D5:D{last}", CellIsRule(operator="equal", formula=['"assumption"'], fill=PatternFill("solid", bgColor=COLORS["amber"][0]), font=Font(color=COLORS["amber"][1])))
            ws.conditional_formatting.add(f"D5:D{last}", CellIsRule(operator="equal", formula=['"verified"'], fill=PatternFill("solid", bgColor=COLORS["green"][0]), font=Font(color=COLORS["green"][1])))
        elif s["kind"] == "table":
            cols = s["columns"]; header(ws, 4, [c["h"] for c in cols]); ws.row_dimensions[4].height = 30
            for ri, row in enumerate(s["rows"], 5):
                for ci, v in enumerate(row, 1):
                    c = put(ws, ri, ci, v, cols[ci - 1].get("fmt"), row=ri)
                    if (ri - 5) % 2 and c.fill.fill_type is None: c.fill = PatternFill("solid", fgColor=ZEBRA)
            last = 4 + len(s["rows"])
            if any(c.get("total") for c in cols):
                tr = last + 1
                put(ws, tr, 1, "Total").font = Font(name=FONT, size=10, bold=True)
                for ci, c in enumerate(cols, 1):
                    if c.get("total"):
                        fn = "SUM" if c["total"] == "sum" else "AVERAGE"
                        cell = put(ws, tr, ci, f"={fn}({L(ci)}5:{L(ci)}{last})", c.get("fmt")); cell.font = Font(name=FONT, size=10, bold=True)
            for ci, c in enumerate(cols, 1): ws.column_dimensions[L(ci)].width = c.get("width", 16)
            ws.freeze_panes = "B5" if len(cols) > 4 else "A5"
            if s.get("filter"): ws.auto_filter.ref = f"A4:{L(len(cols))}{last}"
            for cf in s.get("cond", []):
                bg, fg = COLORS[cf["color"]]; f = [str(cf["value"]) if not isinstance(cf["value"], str) else cf["value"]]
                if cf["op"] == "between": f.append(str(cf["value2"]))
                op = {"lt": "lessThan", "gt": "greaterThan", "eq": "equal", "between": "between"}[cf["op"]]
                ws.conditional_formatting.add(cf["range"], CellIsRule(operator=op, formula=f, fill=PatternFill("solid", bgColor=bg), font=Font(color=fg)))
            ch = s.get("chart")
            if ch:
                C = {"bar": BarChart, "line": LineChart, "pie": PieChart}[ch["type"]]()
                C.title = ch["title"]; C.height, C.width = 8.5, 16
                for v in ch["vals"]: C.add_data(Reference(ws, range_string=f"'{ws.title}'!{v}"), titles_from_data=True, from_rows=bool(ch.get("from_rows")))
                C.set_categories(Reference(ws, range_string=f"'{ws.title}'!{ch['cats']}"))
                if ch["type"] != "pie":
                    C.x_axis.title, C.y_axis.title = ch.get("x_title"), ch.get("y_title")
                    C.x_axis.delete = C.y_axis.delete = False
                ws.add_chart(C, ch.get("anchor", f"{L(len(cols) + 2)}4"))
        else:
            sys.exit(f"unknown kind {s['kind']}")
    wb.properties.title = spec.get("title") or ""
    wb.save(out)

if __name__ == "__main__":
    a = [x for x in sys.argv[1:] if x != "--force"]
    if len(a) != 2: sys.exit(__doc__)
    if os.path.exists(a[1]) and "--force" not in sys.argv: sys.exit(f"refusing to overwrite {a[1]}")
    build(json.load(open(a[0])), a[1]); print("wrote", a[1])
