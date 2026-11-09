# professional-documentation

A [Claude Code](https://claude.com/claude-code) Agent Skill that makes Claude produce business-ready documents and spreadsheets instead of generic AI text: reports, proposals, PRDs, SOPs, minutes, technical docs, and financial/KPI workbooks.

One JSON spec renders to **DOCX**, **PDF**, **XLSX**, or a **native Google Doc / Google Sheet**. The skill also carries the writing, formatting, sourcing and QA rules Claude follows when it uses it.

## What you get

| Output | Tool | Notes |
|---|---|---|
| DOCX, PDF | `scripts/build_doc.py` | Same spec renders both. Title page, header/footer with page numbers, zebra tables with repeating header row, callouts, numbered lists that restart, source table. Non-Latin text (e.g. Bengali) works in PDF through a system font. |
| XLSX | `scripts/build_xlsx.py` | Live formulas, named inputs (blue on cream), frozen headers, filters, conditional formatting, totals, charts. |
| Formula QA | `scripts/check_xlsx.py` | Recalculates in LibreOffice, fails on `#REF!`, `#DIV/0!`, `#NAME?`, etc. |
| Google Doc / Sheet | `scripts/publish_google.py` | Native files through the Docs and Sheets APIs. Always a new file, never edits existing ones. |

Rules baked into `SKILL.md` and `references/`: never invent figures or sources, tag every number Verified / Estimate / Assumption with source and date, give confidence levels, never overwrite existing files, run the QA checklist before delivery.

## Requirements

- Claude Code, Python 3.9+, `bash`
- LibreOffice (`soffice`), optional, needed for `check_xlsx.py` and render checks
- A Google Cloud OAuth client, optional, only for Google Docs/Sheets

## Install

```bash
git clone https://github.com/mdmakky/claude-skill-professional-documentation ~/.claude/skills/professional-documentation
cd ~/.claude/skills/professional-documentation
./install.sh            # DOCX / PDF / XLSX
./install.sh --google   # also the Google API client
./tests/run.sh          # offline smoke tests
```

Claude Code loads skills from `~/.claude/skills/<name>/`, so the folder name must be `professional-documentation`. Restart Claude Code, then ask for a document in plain language ("write a market research report on X as a Google Doc") or invoke `/professional-documentation`.

## Usage

```bash
PY=~/.local/share/professional-documentation-venv/bin/python3
S=~/.claude/skills/professional-documentation/scripts

$PY $S/build_doc.py   spec.json report.docx      # or report.pdf
$PY $S/build_xlsx.py  book.json  model.xlsx
$PY $S/check_xlsx.py  model.xlsx                 # exit 1 on any formula error
$PY $S/publish_google.py doc   spec.json         # prints the Google Doc URL
$PY $S/publish_google.py sheet book.json         # prints the Google Sheet URL
```

Builders refuse to overwrite an existing file (`--force` overrides). Start from `templates/` (business report, market research report, SOP, financial model) or `examples/`.

### Document spec

```json
{"title": "...", "subtitle": "...", "doc_type": "Executive Report", "author": "...", "date": "2026-01-31",
 "classification": "Internal", "toc": false,
 "blocks": [
   {"h1": "Executive summary"},
   {"callout": "One-sentence conclusion.", "label": "Recommendation"},
   {"p": "Text with **bold** inline."},
   {"bullets": ["a", "b"]}, {"numbered": ["step 1", "step 2"]},
   {"table": {"caption": "Table 1. ...", "headers": ["Metric", "Value"], "rows": [["A", "1"]], "align": ["l", "r"], "widths": [3, 2]}},
   {"sources": [{"claim": "...", "url": "https://...", "date": "2026-01-31"}]},
   {"pagebreak": true}]}
```

Blocks: `h1 h2 h3 p bullets numbered table callout sources pagebreak`.

### Workbook spec

Two sheet kinds. `inputs` rows have `label, name, value, fmt, tag (verified|assumption), source, note`; `name` becomes a workbook-wide named cell, usable as `=Ad_Spend*CTR`. `table` sheets have `columns` (`h, fmt, width, total`), `rows` (numbers, text, `"=FORMULA"`, or `{"v":..,"input":true}`; `{r}` = current row), optional `filter`, `cond` (conditional formats) and `chart`. Formats: `text int num2 usd usd0 pct pct2 date`. Full reference: docstring in `scripts/build_xlsx.py`, working example: `templates/financial-model.xlsx.json`.

## Google Docs and Sheets

Native Google files need a one-time OAuth setup (about 5 minutes). No Drive scope is requested; the tool can only create files.

1. In [Google Cloud Console](https://console.cloud.google.com) create a project and enable **Google Docs API** and **Google Sheets API**.
2. Configure the OAuth consent screen (External), and add your Gmail under **Test users**. Without this, login fails with "Access blocked".
3. Create credentials: **OAuth client ID**, type **Desktop app**, download the JSON.
4. Save it as `~/.config/professional-documentation/credentials.json` and `chmod 600` it.
5. Run `publish_google.py auth` once, approve in the browser ("Advanced, Go to app" if Google warns the app is unverified; it is your own app).

The token is stored next to the credentials (mode 600) and is never committed (`.gitignore`). While the app is in Testing mode, Google expires the token after about 7 days; run `auth` again.
Limits of the Docs API: it cannot insert page-number fields or a table of contents, so native Docs have a header but no page numbers. DOCX/PDF output has both.

## Layout

```
SKILL.md                  activation instructions Claude reads
references/               writing, format and QA standards, Google notes
templates/                spec skeletons for common documents; outlines.md for 11 more types
examples/                 a complete sample report spec
scripts/                  build_doc, build_xlsx, check_xlsx, publish_google
tests/                    offline smoke tests (./tests/run.sh)
```

## Limitations

- PDF/DOCX use fixed styling (navy palette, Calibri/Helvetica); customize `NAVY`/fonts in the scripts.
- Inline formatting supports `**bold**` only.
- The sample in `examples/` was built from search-result summaries (see its own Scope section) and is a demonstration of format, not vetted market data.
- Google output was verified by reading the created files back through the API, not by visual inspection.

## Author

[Md. Arafatuzzaman](https://github.com/mdmakky) ([@mdmakky](https://github.com/mdmakky)). Built with Claude Code.

## License

MIT. See `LICENSE`.
