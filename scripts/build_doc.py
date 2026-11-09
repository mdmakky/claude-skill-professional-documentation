#!/usr/bin/env python3
"""Render a JSON document spec to DOCX or PDF with one consistent style.
Usage: build_doc.py spec.json out.docx|out.pdf   (refuses to overwrite; use --force)
Run with the skill venv: ~/.local/share/professional-documentation-venv/bin/python3
Spec: {title, subtitle?, author?, date?, doc_type?, classification?, toc?: bool,
       blocks: [{h1|h2|h3: str} | {p: str} | {bullets: [..]} | {numbered: [..]} |
                {table: {headers, rows, widths?, caption?, align?: ["l","r",..]}} |
                {callout: str, label?: str} | {sources: [{claim,url,date}]} | {pagebreak: true}]}
Inline: **bold** only. Fonts: Calibri/Helvetica, 10.5pt body, navy accent."""
import json, os, re, subprocess, sys

NAVY, GREY, ZEBRA = "1F3864", "595959", "F2F5FA"

def load(path):
    spec = json.load(open(path))
    for k in ("title", "blocks"):
        if k not in spec: sys.exit(f"spec missing '{k}'")
    return spec

def parts(text):  # split **bold** runs -> [(text, bold)]
    out = []
    for i, t in enumerate(re.split(r"\*\*(.+?)\*\*", text)):
        if t: out.append((t, i % 2 == 1))
    return out

def new_numbering(d):
    nums = d.part.numbering_part.numbering_definitions._numbering
    base = d.styles["List Number"].element.pPr.numPr.numId.val
    num = nums.add_num(nums.num_having_numId(base).abstractNumId.val)
    num.add_lvlOverride(ilvl=0).add_startOverride(1)
    return num.numId

# ---------------- DOCX ----------------
def build_docx(spec, out):
    from docx import Document
    from docx.shared import Pt, RGBColor, Inches
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.enum.table import WD_TABLE_ALIGNMENT
    from docx.oxml.ns import qn
    from docx.oxml import OxmlElement

    d = Document()
    sec = d.sections[0]
    sec.left_margin = sec.right_margin = Inches(1); sec.top_margin = sec.bottom_margin = Inches(0.9)
    rgb = lambda h: RGBColor.from_string(h)
    st = d.styles["Normal"]; st.font.name = "Calibri"; st.font.size = Pt(10.5)
    st.element.rPr.rFonts.set(qn("w:eastAsia"), "Calibri")
    st.paragraph_format.space_after = Pt(6); st.paragraph_format.line_spacing = 1.15
    for name, size, before in (("Heading 1", 16, 18), ("Heading 2", 13, 12), ("Heading 3", 11, 10)):
        h = d.styles[name]; h.font.name = "Calibri"; h.font.size = Pt(size); h.font.bold = True
        h.font.color.rgb = rgb(NAVY); h.paragraph_format.space_before = Pt(before)
        h.paragraph_format.space_after = Pt(4); h.paragraph_format.keep_with_next = True
        h.element.rPr.rFonts.set(qn("w:eastAsia"), "Calibri")

    def field(par, instr):
        r = par.add_run()
        for t, txt in (("begin", None), (None, instr), ("end", None)):
            if t:
                e = OxmlElement("w:fldChar"); e.set(qn("w:fldCharType"), t)
            else:
                e = OxmlElement("w:instrText"); e.set(qn("xml:space"), "preserve"); e.text = txt
            r._r.append(e)
        return r

    def shade(cell, hexcolor):
        p = cell._tc.get_or_add_tcPr(); s = OxmlElement("w:shd")
        s.set(qn("w:val"), "clear"); s.set(qn("w:color"), "auto"); s.set(qn("w:fill"), hexcolor); p.append(s)

    def runs(par, text, size=None, color=None, bold_all=False):
        for t, b in parts(text):
            r = par.add_run(t); r.bold = b or bold_all
            if size: r.font.size = Pt(size)
            if color: r.font.color.rgb = rgb(color)

    # title block
    p = d.add_paragraph(); p.paragraph_format.space_before = Pt(60)
    runs(p, spec["title"], 26, NAVY, True)
    if spec.get("subtitle"):
        runs(d.add_paragraph(), spec["subtitle"], 14, GREY)
    meta = " | ".join(x for x in (spec.get("doc_type"), spec.get("author"), spec.get("date"), spec.get("classification")) if x)
    if meta: runs(d.add_paragraph(), meta, 9.5, GREY)
    if spec.get("toc"):
        runs(d.add_paragraph(), "Contents", 13, NAVY, True)
        field(d.add_paragraph(), 'TOC \\o "1-2" \\h \\z \\u')
        d.add_paragraph().add_run("(Right-click, Update Field to refresh the table of contents.)").font.size = Pt(8)
    d.add_page_break()

    # header/footer
    hp = sec.header.paragraphs[0]; hp.text = ""; runs(hp, spec["title"], 8.5, GREY)
    fp = sec.footer.paragraphs[0]; fp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    runs(fp, (spec.get("classification") or "") + ("   |   " if spec.get("classification") else "") + "Page ", 8.5, GREY)
    for instr in ("PAGE",):
        r = field(fp, instr); r.font.size = Pt(8.5)
    runs(fp, " of ", 8.5, GREY); field(fp, "NUMPAGES").font.size = Pt(8.5)

    for b in spec["blocks"]:
        k = next(iter(b))
        if k in ("h1", "h2", "h3"):
            d.add_heading(b[k], int(k[1]))
        elif k == "p":
            runs(d.add_paragraph(), b[k])
        elif k in ("bullets", "numbered"):
            nid = new_numbering(d) if k == "numbered" else None
            for it in b[k]:
                par = d.add_paragraph(style="List Bullet" if k == "bullets" else "List Number")
                if nid:  # own numbering instance, so every numbered list restarts at 1
                    np_ = par._p.get_or_add_pPr().get_or_add_numPr(); np_.get_or_add_ilvl().val = 0; np_.get_or_add_numId().val = nid
                runs(par, it)
        elif k == "callout":
            t = d.add_table(rows=1, cols=1); c = t.rows[0].cells[0]; shade(c, ZEBRA)
            c.paragraphs[0].text = ""
            runs(c.paragraphs[0], (b.get("label", "Key point") + ": "), 10.5, NAVY, True); runs(c.paragraphs[0], b[k])
            d.add_paragraph()
        elif k == "table":
            T = b[k]; hd, rows = T["headers"], T["rows"]
            al = T.get("align") or ["l"] * len(hd)
            if T.get("caption"):
                cp = d.add_paragraph(); cp.paragraph_format.keep_with_next = True; runs(cp, T["caption"], 9.5, GREY, True)
            t = d.add_table(rows=1, cols=len(hd)); t.style = "Table Grid"; t.alignment = WD_TABLE_ALIGNMENT.CENTER
            for i, h in enumerate(hd):
                c = t.rows[0].cells[i]; c.text = ""; shade(c, NAVY); runs(c.paragraphs[0], h, 9.5, "FFFFFF", True)
                if al[i] == "r": c.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.RIGHT
            trPr = t.rows[0]._tr.get_or_add_trPr(); e = OxmlElement("w:tblHeader"); e.set(qn("w:val"), "true"); trPr.append(e)
            for ri, row in enumerate(rows):
                cells = t.add_row().cells
                for i, v in enumerate(row):
                    cells[i].text = ""; runs(cells[i].paragraphs[0], str(v), 9.5)
                    cells[i].paragraphs[0].paragraph_format.space_after = Pt(2)
                    if al[i] == "r": cells[i].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.RIGHT
                    if ri % 2: shade(cells[i], ZEBRA)
            if T.get("widths"):
                for row in t.rows:
                    for i, w in enumerate(T["widths"]): row.cells[i].width = Inches(w)
            d.add_paragraph()
        elif k == "sources":
            d.add_heading("Sources and collection dates", 2)
            hd = ["Claim / data point", "Source", "Collected"]
            t = d.add_table(rows=1, cols=3); t.style = "Table Grid"
            for i, h in enumerate(hd):
                c = t.rows[0].cells[i]; c.text = ""; shade(c, NAVY); runs(c.paragraphs[0], h, 9.5, "FFFFFF", True)
            for s in b[k]:
                cells = t.add_row().cells
                for i, v in enumerate((s["claim"], s["url"], s["date"])):
                    cells[i].text = ""; runs(cells[i].paragraphs[0], v, 9)
        elif k == "pagebreak":
            d.add_page_break()
        else:
            sys.exit(f"unknown block '{k}'")
    d.core_properties.title = spec["title"]; d.core_properties.author = spec.get("author") or ""
    d.save(out)

# ---------------- PDF ----------------
def pdf_fonts(spec):
    """Helvetica for Latin text; for other scripts (e.g. Bengali) register a system font that covers them."""
    odd = sorted({ord(c) for c in json.dumps(spec, ensure_ascii=False) if ord(c) > 0x24F})
    if not odd: return None
    from reportlab.pdfbase import pdfmetrics
    from reportlab.pdfbase.ttfonts import TTFont
    pat = ":charset=" + " ".join(f"{c:x}" for c in [0x41, 0x61, 0x31] + odd[:60])
    try: path = subprocess.run(["fc-match", "-f", "%{file}", pat], capture_output=True, text=True, timeout=10).stdout.strip()
    except (OSError, subprocess.SubprocessError): path = ""
    if not path.lower().endswith((".ttf", ".otf")):
        print("warning: no system font found for non-Latin text; PDF may show boxes. Use DOCX or install Noto fonts.", file=sys.stderr)
        return None
    pdfmetrics.registerFont(TTFont("DocFont", path))
    pdfmetrics.registerFontFamily("DocFont", normal="DocFont", bold="DocFont", italic="DocFont", boldItalic="DocFont")
    return "DocFont"

def build_pdf(spec, out):
    from reportlab.lib.pagesizes import letter
    from reportlab.lib.styles import ParagraphStyle
    from reportlab.lib import colors
    from reportlab.lib.units import inch
    from reportlab.platypus import (BaseDocTemplate, PageTemplate, Frame, Paragraph, Spacer, Table,
                                    TableStyle, PageBreak, ListFlowable, ListItem, KeepTogether)
    from xml.sax.saxutils import escape
    F, FB, alt = "Helvetica", "Helvetica-Bold", pdf_fonts(spec)  # alt: font for non-Latin runs
    navy, grey, zebra = (colors.HexColor("#" + c) for c in (NAVY, GREY, ZEBRA))
    S = lambda n, **k: ParagraphStyle(n, fontName=k.pop("fontName", F), **k)
    body = S("body", fontSize=10, leading=14, spaceAfter=6)
    sty = {"h1": S("h1", fontName=FB, fontSize=16, leading=20, textColor=navy, spaceBefore=16, spaceAfter=6, keepWithNext=1),
           "h2": S("h2", fontName=FB, fontSize=13, leading=16, textColor=navy, spaceBefore=12, spaceAfter=4, keepWithNext=1),
           "h3": S("h3", fontName=FB, fontSize=11, leading=14, textColor=navy, spaceBefore=8, spaceAfter=3, keepWithNext=1)}
    cell = S("cell", fontSize=8.5, leading=11)
    cellh = S("cellh", fontName=FB, fontSize=8.5, leading=11, textColor=colors.white)
    def mk(text):
        t = re.sub(r"\*\*(.+?)\*\*", r"<b>\1</b>", escape(text))
        return re.sub(r"([^\x00-\u024f]+)", rf'<font name="{alt}">\1</font>', t) if alt else t
    def footer(c, doc):
        c.saveState(); c.setFont(F, 8); c.setFillColor(grey)
        c.drawString(inch, letter[1] - 0.6 * inch, spec["title"])
        c.drawCentredString(letter[0] / 2, 0.55 * inch, (spec.get("classification", "") + "   |   " if spec.get("classification") else "") + f"Page {doc.page}")
        c.restoreState()
    doc = BaseDocTemplate(out, pagesize=letter, title=spec["title"], author=spec.get("author") or "",
                          leftMargin=inch, rightMargin=inch, topMargin=0.9 * inch, bottomMargin=0.9 * inch)
    doc.addPageTemplates([PageTemplate(id="p", frames=[Frame(doc.leftMargin, doc.bottomMargin, doc.width, doc.height)], onPage=footer)])
    f = [Spacer(1, 1.6 * inch), Paragraph(mk(spec["title"]), S("t", fontName=FB, fontSize=26, leading=31, textColor=navy, spaceAfter=10))]
    if spec.get("subtitle"): f.append(Paragraph(mk(spec["subtitle"]), S("st", fontSize=14, leading=18, textColor=grey, spaceAfter=8)))
    meta = " | ".join(x for x in (spec.get("doc_type"), spec.get("author"), spec.get("date"), spec.get("classification")) if x)
    if meta: f.append(Paragraph(escape(meta), S("m", fontSize=9.5, textColor=grey)))
    f.append(PageBreak())
    def table(hd, rows, widths, al, cap=None):
        data = [[Paragraph(mk(h), cellh) for h in hd]] + [[Paragraph(mk(str(v)), S("c%d" % i, fontSize=8.5, leading=11, alignment=2 if al[i] == "r" else 0)) for i, v in enumerate(r)] for r in rows]
        cw = [w * inch for w in widths] if widths else [doc.width / len(hd)] * len(hd)
        t = Table(data, colWidths=cw, repeatRows=1)
        ts = [("BACKGROUND", (0, 0), (-1, 0), navy), ("GRID", (0, 0), (-1, -1), 0.4, colors.HexColor("#BFC5D2")),
              ("VALIGN", (0, 0), (-1, -1), "TOP"), ("TOPPADDING", (0, 0), (-1, -1), 3), ("BOTTOMPADDING", (0, 0), (-1, -1), 3)]
        ts += [("BACKGROUND", (0, i), (-1, i), zebra) for i in range(2, len(data), 2)]
        t.setStyle(TableStyle(ts))
        items = ([Paragraph(mk(cap), S("cap", fontName=FB, fontSize=9, textColor=grey, spaceAfter=3, keepWithNext=1))] if cap else [])
        return items + [t, Spacer(1, 10)]
    for b in spec["blocks"]:
        k = next(iter(b))
        if k in sty: f.append(Paragraph(mk(b[k]), sty[k]))
        elif k == "p": f.append(Paragraph(mk(b[k]), body))
        elif k in ("bullets", "numbered"):
            f.append(ListFlowable([ListItem(Paragraph(mk(x), body)) for x in b[k]],
                                  bulletType="bullet" if k == "bullets" else "1", leftIndent=18, bulletFontSize=10, start="•" if k == "bullets" else None))
        elif k == "callout":
            t = Table([[Paragraph(f"<b>{escape(b.get('label', 'Key point'))}:</b> {mk(b[k])}", body)]], colWidths=[doc.width])
            t.setStyle(TableStyle([("BACKGROUND", (0, 0), (-1, -1), zebra), ("LINEBEFORE", (0, 0), (0, -1), 3, navy), ("LEFTPADDING", (0, 0), (-1, -1), 10)]))
            f += [t, Spacer(1, 8)]
        elif k == "table":
            T = b[k]; f += table(T["headers"], T["rows"], T.get("widths"), T.get("align") or ["l"] * len(T["headers"]), T.get("caption"))
        elif k == "sources":
            f.append(Paragraph("Sources and collection dates", sty["h2"]))
            f += table(["Claim / data point", "Source", "Collected"], [[s["claim"], s["url"], s["date"]] for s in b[k]], [2.4, 3.1, 1.0], ["l"] * 3)
        elif k == "pagebreak": f.append(PageBreak())
        else: sys.exit(f"unknown block '{k}'")
    doc.build(f)

def main():
    a = [x for x in sys.argv[1:] if x != "--force"]
    if len(a) != 2: sys.exit(__doc__)
    spec, out = load(a[0]), a[1]
    if os.path.exists(out) and "--force" not in sys.argv: sys.exit(f"refusing to overwrite {out} (use --force only with user permission)")
    ext = os.path.splitext(out)[1].lower()
    {".docx": build_docx, ".pdf": build_pdf}.get(ext, lambda *_: sys.exit("output must be .docx or .pdf"))(spec, out)
    print("wrote", out)

if __name__ == "__main__":
    main()
