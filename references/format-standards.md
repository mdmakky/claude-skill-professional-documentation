# Format standards

## Documents (DOCX / PDF / Google Docs)
- Page: US Letter (use A4 if the audience is outside North America), 1 in side margins.
- Type: Calibri (DOCX/Docs) or Helvetica (PDF). Body 10.5 pt, line spacing 1.15, 6 pt after paragraphs. Headings 16 / 13 / 11 pt bold, navy `#1F3864`. Title page 26 pt. No heading above 26 pt.
- Palette: navy `#1F3864` primary, grey `#595959` secondary, light `#F2F5FA` for banding and callouts. Red/amber/green only for status values. Max 3 hues per document.
- Header: document title (small, grey). Footer: classification (if any) and "Page X of Y".
- Tables: header row navy with white bold text and repeated on each page; zebra rows; numbers right-aligned; 8.5-9.5 pt; max ~6 columns portrait. Wider data goes to an appendix or a spreadsheet.
- Keep headings with their next paragraph; no orphan headings at the page bottom.
- Table of contents when the document exceeds ~8 pages.

## Spreadsheets (XLSX / Google Sheets)
- Sheet order: Summary (or Dashboard) -> Inputs -> Calculations -> Raw Data -> Sources/Notes.
- Inputs live only on the Inputs sheet: blue text on cream fill. Everything else is formulas (black). Name key inputs (defined names) so formulas read `=Ad_Spend*CTR`.
- No hard-coded numbers inside formulas, except structural constants (12 months, 1).
- Formats: currency `$#,##0.00`, percent `0.0%`, integers `#,##0`, dates `yyyy-mm-dd`. One format per column.
- Row 1 title, row 2 legend, row 4 headers; freeze panes below headers; filters on data tables; column widths fit content; no merged cells in data ranges.
- Conditional formatting for thresholds (negative profit red, below-target amber). Status column shows verified vs assumption.
- Charts: one message per chart, titled with the takeaway or subject, axis titles, units, legend only for multiple series, no 3D.
- Always run `scripts/check_xlsx.py` before delivery; zero formula errors.

## Google Docs / Sheets
- Only create native Google files when an authenticated Google Docs/Sheets API or connector is available and verified in the session (test with a read call first). Otherwise produce DOCX/XLSX and say so plainly; DOCX/XLSX import into Google Drive without losing structure.
- Never overwrite or edit an existing file without explicit permission; create a new file with a versioned name.
