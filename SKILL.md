---
name: professional-documentation
description: "Create polished, business-ready documents and spreadsheets: executive/market research/client reports, proposals, PRDs, SOPs, meeting minutes, technical docs, project plans, strategy docs, and financial/KPI/competitor/budget workbooks. Outputs DOCX, PDF, XLSX (native Google Docs/Sheets only when API access is verified). Use whenever the user asks for a report, document, SOP, proposal, spreadsheet, workbook, financial model, or PDF, or wants a document to look professional."
---

# Professional Documentation

Produce documents a client or executive can use unedited: clear hierarchy, sourced facts, consistent formatting, verified math.

## Always
1. Identify audience, purpose, decision/action, and output format. Ask only if any is genuinely unclear.
2. Read `references/writing-standards.md` and `references/format-standards.md` before drafting. Run `references/qa-checklist.md` before delivering.
3. Facts: never invent statistics, prices, references, URLs or financial figures. Tag each as Verified / Estimate / Assumption with source and collection date. Unknown = "Not available" plus where to get it. For e-commerce research also follow `ecommerce-research-standards`; for profit math use `ecommerce-profit-analysis`.
4. Never overwrite an existing file: write to a new path (the scripts refuse unless `--force`, which needs user permission).
5. Final message: file path(s), what was checked, anything unverified. No claims of Google Docs/Sheets creation unless verified (see `references/google-workspace.md`).

## Choose the path
| Need | Do |
|---|---|
| DOCX or PDF document | Copy a template from `templates/` (or compose a spec from `templates/outlines.md`), fill it, run `build_doc.py` |
| XLSX workbook | Write a workbook spec (see `templates/financial-model.xlsx.json`), run `build_xlsx.py`, then `check_xlsx.py` |
| Google Doc/Sheet | `publish_google.py doc` or `publish_google.py sheet` with a spec (verified working 2026-11-09; needs token, see `references/google-workspace.md`); if auth fails, build DOCX/XLSX and say so |

## Commands
Python env (python-docx, openpyxl, reportlab) is a venv; create if missing:
`python3 -m venv ~/.local/share/professional-documentation-venv && ~/.local/share/professional-documentation-venv/bin/pip install python-docx openpyxl reportlab`
```
PY=~/.local/share/professional-documentation-venv/bin/python3
S=~/.claude/skills/professional-documentation
$PY $S/scripts/build_doc.py spec.json out.docx      # or out.pdf (same spec renders both)
$PY $S/scripts/build_xlsx.py book.json out.xlsx
$PY $S/scripts/publish_google.py doc spec.json      # native Google Doc, new file only; `sheet book.json` for Sheets
$PY $S/scripts/check_xlsx.py out.xlsx               # LibreOffice recalculation; exit 1 on any formula error
soffice --headless --convert-to pdf --outdir DIR out.docx   # render check
```
Doc spec blocks: `h1 h2 h3 p bullets numbered table callout sources pagebreak`; `**bold**` inline. Workbook spec: see docstring in `scripts/build_xlsx.py` (input sheets with named cells, formula tables, totals, conditional formats, charts).

## Document rules (summary)
- Executive summary first for >3 pages; findings as headings; tables for comparisons; caption every table/chart.
- Consistent terms, no filler openers/closers, no emojis, max 3 hues, body 10.5 pt, header/footer with page numbers.

## Spreadsheet rules (summary)
- Sheets: Summary, Inputs, Calculations, Raw Data, Sources. Inputs blue-on-cream, calculations are formulas, named inputs, consistent formats, frozen header, filters, conditional formatting, labelled charts. Run `check_xlsx.py` and fix every error.

## Spec for data not yet available
Use bracketed placeholders (`[Source needed]`, "Not yet collected") instead of invented values. Leave templates' placeholders visible rather than fabricating.
