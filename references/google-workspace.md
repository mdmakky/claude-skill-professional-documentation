# Google Docs / Sheets integration

Status: **verified 2026-11-09.** `scripts/publish_google.py` created a native Doc and a native Sheet via the Docs and Sheets APIs; content read back through the API (headings, 4 tables, named ranges, chart, formulas with 0 errors).

Setup (done once per machine): OAuth Desktop client JSON at `~/.config/professional-documentation/credentials.json`; token at `token.json` (both mode 600, never print or commit). Re-auth: `publish_google.py auth`. The Google Cloud app is in Testing mode, so the token expires after ~7 days; rerun `auth` then. Only Gmail accounts listed as Test users can log in.

Rules: new file every time, never edit existing files without permission. Report the URL and what was verified.
Known Docs API limits: no page-number field, no table of contents. (Stray empty heading paragraphs after tables were fixed 2026-11-09.)
If auth is missing or fails, fall back to DOCX/XLSX and say so plainly.
