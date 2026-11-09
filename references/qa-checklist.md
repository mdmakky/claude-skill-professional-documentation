# QA checklist (run before delivery; report results)

1. **Language**: spell-check, grammar, consistent tense and terminology, no banned patterns (see writing-standards.md).
2. **Facts**: each number traces to a source row or is labelled assumption; dates and units present; no invented references.
3. **Consistency**: figures in summary == body == tables == charts; names, dates, versions identical everywhere.
4. **Tables / math**: recompute totals and percentages; for sheets run `check_xlsx.py` (0 errors) and spot-check 3 formulas by hand.
5. **Structure**: heading hierarchy has no skips; summary present when needed; sources section present; page numbers and header/footer present.
6. **Render**: convert to PDF (`soffice --headless --convert-to pdf`) and inspect page count and layout (tables not clipped, no orphan headings). View pages as images if available.
7. **Safety**: no secrets, tokens or personal data unless required; output path did not overwrite an existing file.
Report: what was checked, what failed and was fixed, what remains unverified.
