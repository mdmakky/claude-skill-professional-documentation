#!/usr/bin/env python3
"""Create NATIVE Google Docs / Sheets from the same JSON specs used by build_doc.py / build_xlsx.py.
Usage:
  publish_google.py auth                      # one-time browser login, saves token
  publish_google.py doc   spec.json           # new Google Doc  -> prints URL
  publish_google.py sheet book.json           # new Google Sheet -> prints URL
Always creates a NEW file (never edits or overwrites). Needs ~/.config/professional-documentation/credentials.json
(OAuth Desktop client). Token: same dir, token.json (mode 600). Scopes: documents, spreadsheets.
Docs API limits: no page-number field and no table of contents can be inserted via API (they are skipped)."""
import json, os, re, sys

CFG = os.path.expanduser("~/.config/professional-documentation")
CRED, TOKEN = f"{CFG}/credentials.json", f"{CFG}/token.json"
SCOPES = ["https://www.googleapis.com/auth/documents", "https://www.googleapis.com/auth/spreadsheets"]
NAVY, GREY, ZEBRA = "1F3864", "595959", "F2F5FA"

def rgb(h): return {"color": {"rgbColor": {"red": int(h[0:2], 16) / 255, "green": int(h[2:4], 16) / 255, "blue": int(h[4:6], 16) / 255}}}
def rgbc(h): return rgb(h)["color"]["rgbColor"]

def u16(t): return len(t.encode("utf-16-le")) // 2  # Docs API indexes count UTF-16 code units (emoji = 2)

def creds(login=False):
    from google.oauth2.credentials import Credentials
    from google.auth.transport.requests import Request
    from google_auth_oauthlib.flow import InstalledAppFlow
    c = Credentials.from_authorized_user_file(TOKEN, SCOPES) if os.path.exists(TOKEN) else None
    if c and c.valid: return c
    if c and c.expired and c.refresh_token:
        c.refresh(Request())
    else:
        if not login: sys.exit("Not logged in (or token expired). Run: publish_google.py auth")
        c = InstalledAppFlow.from_client_secrets_file(CRED, SCOPES).run_local_server(port=0, open_browser=True)
    with open(TOKEN, "w") as f: f.write(c.to_json())
    os.chmod(TOKEN, 0o600)
    return c

def svc(name, ver):
    from googleapiclient.discovery import build
    return build(name, ver, credentials=creds(), cache_discovery=False)

def runs(text):  # "a **b** c" -> ("a b c", [(start, end)] bold ranges relative)
    out, bold, pos = "", [], 0
    for i, t in enumerate(re.split(r"\*\*(.+?)\*\*", text)):
        if i % 2: bold.append((pos, pos + len(t)))
        out += t; pos += len(t)
    return out, bold

# ------------------------------- DOCS -------------------------------
def make_doc(spec):
    docs = svc("docs", "v1")
    doc = docs.documents().create(body={"title": spec["title"]}).execute()
    did = doc["documentId"]
    def send(reqs):
        if reqs: docs.documents().batchUpdate(documentId=did, body={"requests": reqs}).execute()
    def get(): return docs.documents().get(documentId=did).execute()
    def end(): return get()["body"]["content"][-1]["endIndex"] - 1

    state = {"idx": 1, "reqs": []}
    def para(text, style="NORMAL_TEXT", bold_all=False, size=None, color=None, shade=None, bar=False, bullet=None, space_after=6):
        t, bolds = runs(text); s = state["idx"]; e = s + u16(t) + 1
        r = state["reqs"]
        r.append({"insertText": {"location": {"index": s}, "text": t + "\n"}})
        r.append({"updateParagraphStyle": {"range": {"startIndex": s, "endIndex": e}, "fields": "namedStyleType,spaceBelow,spaceAbove" + (",shading" if shade else "") + (",borderLeft" if bar else ""),
                  "paragraphStyle": {"namedStyleType": style, "spaceBelow": {"magnitude": space_after, "unit": "PT"}, "spaceAbove": {"magnitude": 0, "unit": "PT"},
                                     **({"shading": {"backgroundColor": rgb(shade)}} if shade else {}),
                                     **({"borderLeft": {"color": rgb(NAVY), "width": {"magnitude": 3, "unit": "PT"}, "padding": {"magnitude": 6, "unit": "PT"}, "dashStyle": "SOLID"}} if bar else {})}}})
        ts = {}
        if bold_all: ts["bold"] = True
        if size: ts["fontSize"] = {"magnitude": size, "unit": "PT"}
        if color: ts["foregroundColor"] = rgb(color)
        if ts:
            r.append({"updateTextStyle": {"range": {"startIndex": s, "endIndex": e - 1}, "textStyle": ts, "fields": ",".join(ts)}})
        for a, b in bolds:
            r.append({"updateTextStyle": {"range": {"startIndex": s + u16(t[:a]), "endIndex": s + u16(t[:b])}, "textStyle": {"bold": True}, "fields": "bold"}})
        if bullet:
            r.append({"createParagraphBullets": {"range": {"startIndex": s, "endIndex": e - 1}, "bulletPreset": bullet}})
        state["idx"] = e

    def flush():
        send(state["reqs"]); state["reqs"] = []; state["idx"] = end()

    def table(headers, rows, aligns=None, caption=None):
        if caption: para(caption, bold_all=True, size=9.5, color=GREY, space_after=3)
        flush()
        nr, nc = len(rows) + 1, len(headers)
        send([{"insertTable": {"rows": nr, "columns": nc, "location": {"index": state["idx"]}}}])
        tbl = [e for e in get()["body"]["content"] if "table" in e][-1]
        cells = [(ri, ci, tbl["table"]["tableRows"][ri]["tableCells"][ci]["content"][0]["startIndex"]) for ri in range(nr) for ci in range(nc)]
        data = [headers] + rows
        send([{"insertText": {"location": {"index": idx}, "text": runs(str(data[ri][ci]))[0]}} for ri, ci, idx in reversed(cells)])
        tbl = [e for e in get()["body"]["content"] if "table" in e][-1]
        t0, reqs = tbl["startIndex"], []
        for ri in range(nr):
            for ci in range(nc):
                c = tbl["table"]["tableRows"][ri]["tableCells"][ci]; p = c["content"][0]
                rng = {"startIndex": p["startIndex"], "endIndex": p["endIndex"] - 1}
                if rng["endIndex"] <= rng["startIndex"]: continue
                ts = {"fontSize": {"magnitude": 9, "unit": "PT"}}
                if ri == 0: ts.update(bold=True, foregroundColor=rgb("FFFFFF"))
                reqs.append({"updateTextStyle": {"range": rng, "textStyle": ts, "fields": ",".join(ts)}})
                reqs.append({"updateParagraphStyle": {"range": rng, "paragraphStyle": {"spaceBelow": {"magnitude": 2, "unit": "PT"}, **({"alignment": "END"} if aligns and aligns[ci] == "r" else {})},
                              "fields": "spaceBelow" + (",alignment" if aligns and aligns[ci] == "r" else "")}})
                bg = NAVY if ri == 0 else (ZEBRA if ri % 2 == 0 else None)
                if bg:
                    reqs.append({"updateTableCellStyle": {"tableRange": {"tableCellLocation": {"tableStartLocation": {"index": t0}, "rowIndex": ri, "columnIndex": ci}, "rowSpan": 1, "columnSpan": 1},
                                 "tableCellStyle": {"backgroundColor": rgb(bg)}, "fields": "backgroundColor"}})
        reqs.append({"pinTableHeaderRows": {"tableStartLocation": {"index": t0}, "pinnedHeaderRowsCount": 1}})
        send(reqs)
        state["idx"] = end()
        para("", space_after=4)

    # title block
    para(spec["title"], style="TITLE", bold_all=True, size=26, color=NAVY)
    if spec.get("subtitle"): para(spec["subtitle"], style="SUBTITLE", size=14, color=GREY)
    meta = " | ".join(x for x in (spec.get("doc_type"), spec.get("author"), spec.get("date"), spec.get("classification")) if x)
    if meta: para(meta, size=9.5, color=GREY)
    for b in spec["blocks"]:
        k = next(iter(b))
        if k in ("h1", "h2", "h3"):
            para(b[k], style={"h1": "HEADING_1", "h2": "HEADING_2", "h3": "HEADING_3"}[k], bold_all=True, size={"h1": 16, "h2": 13, "h3": 11}[k], color=NAVY)
        elif k == "p": para(b[k])
        elif k in ("bullets", "numbered"):
            for it in b[k]:
                para(it, bullet="BULLET_DISC_CIRCLE_SQUARE" if k == "bullets" else "NUMBERED_DECIMAL_ALPHA_ROMAN", space_after=2)
        elif k == "callout": para(f"**{b.get('label', 'Key point')}:** {b[k]}", shade=ZEBRA, bar=True)
        elif k == "table":
            T = b[k]; flush(); table(T["headers"], T["rows"], T.get("align"), T.get("caption"))
        elif k == "sources":
            para("Sources and collection dates", style="HEADING_2", bold_all=True, size=13, color=NAVY); flush()
            table(["Claim / data point", "Source", "Collected"], [[s["claim"], s["url"], s["date"]] for s in b[k]])
        elif k == "pagebreak":
            flush(); send([{"insertPageBreak": {"location": {"index": state["idx"]}}}]); state["idx"] = end()
        else: sys.exit(f"unknown block {k}")
    flush()
    # global font + header/footer
    fin = [{"updateTextStyle": {"range": {"startIndex": 1, "endIndex": end()}, "textStyle": {"weightedFontFamily": {"fontFamily": "Calibri"}}, "fields": "weightedFontFamily"}}]
    send(fin)
    res = docs.documents().batchUpdate(documentId=did, body={"requests": [{"createHeader": {"type": "DEFAULT"}}]}).execute()
    hid = res["replies"][0]["createHeader"]["headerId"]
    send([{"insertText": {"location": {"segmentId": hid, "index": 0}, "text": spec["title"]}},
          {"updateTextStyle": {"range": {"segmentId": hid, "startIndex": 0, "endIndex": len(spec["title"])}, "textStyle": {"fontSize": {"magnitude": 8.5, "unit": "PT"}, "foregroundColor": rgb(GREY)}, "fields": "fontSize,foregroundColor"}}])
    return f"https://docs.google.com/document/d/{did}/edit"

# ------------------------------- SHEETS -------------------------------
FMT = {"text": None, "int": "#,##0", "num2": "#,##0.00", "usd": '"$"#,##0.00', "usd0": '"$"#,##0', "pct": "0.0%", "pct2": "0.00%", "date": "yyyy-mm-dd"}
UNITS = {"text": "Text", "int": "Count", "num2": "Number", "usd": "USD", "usd0": "USD", "pct": "Percent", "pct2": "Percent", "date": "Date"}
COND = {"red": ("FFC7CE", "9C0006"), "green": ("C6EFCE", "006100"), "amber": ("FFEB9C", "9C5700")}
A1 = re.compile(r"([A-Z]+)(\d+)")

def col_idx(s):
    n = 0
    for ch in s: n = n * 26 + ord(ch) - 64
    return n - 1

def rng_of(sid, a1):
    a, b = a1.split(":"); (ca, ra), (cb, rb) = A1.match(a).groups(), A1.match(b).groups()
    return {"sheetId": sid, "startRowIndex": int(ra) - 1, "endRowIndex": int(rb), "startColumnIndex": col_idx(ca), "endColumnIndex": col_idx(cb) + 1}

def col_letter(i):
    s = ""
    i += 1
    while i: i, r = divmod(i - 1, 26); s = chr(65 + r) + s
    return s

def make_sheet(spec):
    sh = svc("sheets", "v4")
    book = sh.spreadsheets().create(body={"properties": {"title": spec.get("title") or "Workbook"}, "sheets": [{"properties": {"title": s["name"][:31], "gridProperties": {"hideGridlines": True}}} for s in spec["sheets"]]}).execute()
    sid = book["spreadsheetId"]; ids = {s["properties"]["title"]: s["properties"]["sheetId"] for s in book["sheets"]}
    values, reqs = [], []
    def cellfmt(sheet_id, r0, r1, c0, c1, fmt=None, input_=False, bold=False, fg=None, bg=None, wrap=False, halign=None):
        cf, f = {}, []
        if fmt and FMT.get(fmt, fmt): cf["numberFormat"] = {"type": "PERCENT" if "pct" in fmt else ("DATE" if fmt == "date" else "NUMBER"), "pattern": FMT.get(fmt, fmt)}; f.append("numberFormat")
        tf = {"fontFamily": "Calibri", "fontSize": 10}
        if input_: tf["foregroundColor"] = rgbc("0000FF")
        if fg: tf["foregroundColor"] = rgbc(fg)
        if bold: tf["bold"] = True
        cf["textFormat"] = tf; f.append("textFormat")
        if input_ or bg: cf["backgroundColor"] = rgbc("FFF8DC" if input_ else bg); f.append("backgroundColor")
        if wrap: cf["wrapStrategy"] = "WRAP"; f.append("wrapStrategy")
        if halign: cf["horizontalAlignment"] = halign; f.append("horizontalAlignment")
        cf["borders"] = {k: {"style": "SOLID", "color": rgbc("BFC5D2")} for k in ("top", "bottom", "left", "right")}; f.append("borders")
        reqs.append({"repeatCell": {"range": {"sheetId": sheet_id, "startRowIndex": r0, "endRowIndex": r1, "startColumnIndex": c0, "endColumnIndex": c1}, "cell": {"userEnteredFormat": cf}, "fields": "userEnteredFormat(" + ",".join(f) + ")"}})
    def title(s, sheet_id):
        values.append({"range": f"'{s['name'][:31]}'!A1:A2", "values": [[s.get("title") or spec.get("title") or s["name"]], [s.get("subtitle", "Blue text on cream = input. Black = calculated.")]]})
        reqs.append({"repeatCell": {"range": {"sheetId": sheet_id, "startRowIndex": 0, "endRowIndex": 1, "startColumnIndex": 0, "endColumnIndex": 1}, "cell": {"userEnteredFormat": {"textFormat": {"fontFamily": "Calibri", "fontSize": 14, "bold": True, "foregroundColor": rgbc(NAVY)}}}, "fields": "userEnteredFormat.textFormat"}})
        reqs.append({"repeatCell": {"range": {"sheetId": sheet_id, "startRowIndex": 1, "endRowIndex": 2, "startColumnIndex": 0, "endColumnIndex": 1}, "cell": {"userEnteredFormat": {"textFormat": {"fontFamily": "Calibri", "fontSize": 9, "italic": True, "foregroundColor": rgbc(GREY)}}}, "fields": "userEnteredFormat.textFormat"}})
    def widths(sheet_id, ws):
        for i, w in enumerate(ws): reqs.append({"updateDimensionProperties": {"range": {"sheetId": sheet_id, "dimension": "COLUMNS", "startIndex": i, "endIndex": i + 1}, "properties": {"pixelSize": int(w * 7 + 10)}, "fields": "pixelSize"}})
    def header(sheet_id, n):
        cellfmt(sheet_id, 3, 4, 0, n, bold=True, fg="FFFFFF", bg=NAVY, halign="CENTER", wrap=True)
    for s in spec["sheets"]:
        name = s["name"][:31]; sheet_id = ids[name]; title(s, sheet_id); q = f"'{name}'!"
        if s["kind"] == "inputs":
            rows = s["rows"]; n = len(rows)
            values.append({"range": q + f"A4:F{4 + n}", "values": [["Input", "Value", "Unit / format", "Status", "Source", "Note"]] + [[r["label"], r["value"], UNITS.get(r.get("fmt", "text")), r.get("tag", "assumption"), r.get("source", ""), r.get("note", "")] for r in rows]})
            header(sheet_id, 6); cellfmt(sheet_id, 4, 4 + n, 0, 1); cellfmt(sheet_id, 4, 4 + n, 2, 6, wrap=True)
            for i, r in enumerate(rows):
                cellfmt(sheet_id, 4 + i, 5 + i, 1, 2, fmt=r.get("fmt"), input_=True)
                if r.get("name"): reqs.append({"addNamedRange": {"namedRange": {"name": r["name"], "range": {"sheetId": sheet_id, "startRowIndex": 4 + i, "endRowIndex": 5 + i, "startColumnIndex": 1, "endColumnIndex": 2}}}})
            for word, (bg, fg) in (("assumption", COND["amber"]), ("verified", COND["green"])):
                reqs.append({"addConditionalFormatRule": {"index": 0, "rule": {"ranges": [rng_of(sheet_id, f"D5:D{4 + n}")], "booleanRule": {"condition": {"type": "TEXT_EQ", "values": [{"userEnteredValue": word}]}, "format": {"backgroundColor": rgbc(bg), "textFormat": {"foregroundColor": rgbc(fg)}}}}}})
            widths(sheet_id, (34, 16, 14, 13, 38, 44)); reqs.append({"updateSheetProperties": {"properties": {"sheetId": sheet_id, "gridProperties": {"frozenRowCount": 4, "hideGridlines": True}}, "fields": "gridProperties.frozenRowCount,gridProperties.hideGridlines"}})
        elif s["kind"] == "table":
            cols = s["columns"]; nc = len(cols); rows = s["rows"]; last = 4 + len(rows); out = [[c["h"] for c in cols]]; inputs = []
            for ri, row in enumerate(rows, 5):
                line = []
                for ci, v in enumerate(row):
                    if isinstance(v, dict):
                        if v.get("input"): inputs.append((ri - 1, ci))
                        v = v["v"]
                    if isinstance(v, str) and v.startswith("="): v = v.replace("{r}", str(ri))
                    line.append(v)
                out.append(line)
            if any(c.get("total") for c in cols):
                out.append(["Total"] + [(f"={'SUM' if c['total'] == 'sum' else 'AVERAGE'}({col_letter(i)}5:{col_letter(i)}{last})" if c.get("total") else "") for i, c in enumerate(cols)][1:])
            values.append({"range": q + f"A4:{col_letter(nc - 1)}{3 + len(out)}", "values": out})
            header(sheet_id, nc)
            for ci, c in enumerate(cols):
                cellfmt(sheet_id, 4, 3 + len(out), ci, ci + 1, fmt=c.get("fmt"), wrap=isinstance(rows[0][ci], str) and not str(rows[0][ci]).startswith("=") if rows else False)
            if len(out) > len(rows) + 1: cellfmt(sheet_id, 3 + len(out) - 1, 3 + len(out), 0, nc, bold=True)
            for r, c in inputs: cellfmt(sheet_id, r, r + 1, c, c + 1, fmt=cols[c].get("fmt"), input_=True)
            widths(sheet_id, [c.get("width", 16) for c in cols])
            reqs.append({"updateSheetProperties": {"properties": {"sheetId": sheet_id, "gridProperties": {"frozenRowCount": 4, "hideGridlines": True}}, "fields": "gridProperties.frozenRowCount,gridProperties.hideGridlines"}})
            if s.get("filter"): reqs.append({"setBasicFilter": {"filter": {"range": rng_of(sheet_id, f"A4:{col_letter(nc - 1)}{last}")}}})
            for cf in s.get("cond", []):
                bg, fg = COND[cf["color"]]; t = {"lt": "NUMBER_LESS", "gt": "NUMBER_GREATER", "eq": "NUMBER_EQ", "between": "NUMBER_BETWEEN"}[cf["op"]]
                vals = [{"userEnteredValue": str(cf["value"])}] + ([{"userEnteredValue": str(cf["value2"])}] if cf["op"] == "between" else [])
                reqs.append({"addConditionalFormatRule": {"index": 0, "rule": {"ranges": [rng_of(sheet_id, cf["range"])], "booleanRule": {"condition": {"type": t, "values": vals}, "format": {"backgroundColor": rgbc(bg), "textFormat": {"foregroundColor": rgbc(fg)}}}}}})
            ch = s.get("chart")
            if ch:
                dom = rng_of(sheet_id, ch["cats"]); series = []
                for v in ch["vals"]:
                    r = rng_of(sheet_id, v)
                    if ch.get("from_rows"):  # first cell = series name
                        r = {**r, "startColumnIndex": r["startColumnIndex"] + 1}
                    series.append({"series": {"sourceRange": {"sources": [r]}}, "targetAxis": "LEFT_AXIS"})
                anchor = A1.match(ch.get("anchor", "H4")).groups()
                spec_c = {"title": ch["title"], "basicChart": {"chartType": {"bar": "COLUMN", "line": "LINE", "pie": "COLUMN"}[ch["type"]], "legendPosition": "BOTTOM_LEGEND" if len(series) > 1 else "NO_LEGEND",
                          "axis": [{"position": "BOTTOM_AXIS", "title": ch.get("x_title", "")}, {"position": "LEFT_AXIS", "title": ch.get("y_title", "")}],
                          "domains": [{"domain": {"sourceRange": {"sources": [dom]}}}], "series": series, "headerCount": 0}}
                reqs.append({"addChart": {"chart": {"spec": spec_c, "position": {"overlayPosition": {"anchorCell": {"sheetId": sheet_id, "rowIndex": int(anchor[1]) - 1, "columnIndex": col_idx(anchor[0])}, "widthPixels": 560, "heightPixels": 300}}}}})
        else: sys.exit(f"unknown kind {s['kind']}")
    # named ranges first so formulas resolve; values after
    nr = [r for r in reqs if "addNamedRange" in r]; rest = [r for r in reqs if "addNamedRange" not in r]
    if nr: sh.spreadsheets().batchUpdate(spreadsheetId=sid, body={"requests": nr}).execute()
    sh.spreadsheets().values().batchUpdate(spreadsheetId=sid, body={"valueInputOption": "USER_ENTERED", "data": values}).execute()
    sh.spreadsheets().batchUpdate(spreadsheetId=sid, body={"requests": rest}).execute()
    return f"https://docs.google.com/spreadsheets/d/{sid}/edit", sid

if __name__ == "__main__":
    a = sys.argv[1:]
    if a == ["auth"]:
        if not os.path.exists(CRED): sys.exit(f"missing {CRED}")
        creds(login=True); print("auth ok, token saved"); sys.exit(0)
    if len(a) != 2 or a[0] not in ("doc", "sheet"): sys.exit(__doc__)
    spec = json.load(open(a[1]))
    print(make_doc(spec) if a[0] == "doc" else make_sheet(spec)[0])
