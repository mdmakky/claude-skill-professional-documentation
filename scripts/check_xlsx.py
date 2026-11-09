#!/usr/bin/env python3
"""Recalculate an XLSX headlessly with LibreOffice and report formula errors.
Usage: check_xlsx.py book.xlsx   -> exit 1 if any #REF!/#DIV/0!/#NAME?/#VALUE!/#N/A/#NUM!/Err: or uncached formula.
Also lists sheets with formulas vs. hard-coded numbers so reviewers can spot missing formulas."""
import os, shutil, subprocess, sys, tempfile
from openpyxl import load_workbook
ERR = ("#REF!", "#DIV/0!", "#NAME?", "#VALUE!", "#N/A", "#NUM!", "#NULL!", "Err:")

def main(path):
    tmp = tempfile.mkdtemp()
    subprocess.run(["soffice", f"-env:UserInstallation=file://{tmp}/profile", "--headless", "--convert-to", "xlsx", "--outdir", tmp, path], check=True, capture_output=True, timeout=120)
    calc = load_workbook(os.path.join(tmp, os.path.basename(path)), data_only=True)
    src = load_workbook(path)
    bad, nform = [], 0
    for ws in src:
        f = h = 0
        for row in ws.iter_rows():
            for c in row:
                v = c.value
                if isinstance(v, str) and v.startswith("="):
                    f += 1; r = calc[ws.title][c.coordinate].value
                    if r is None or (isinstance(r, str) and r.startswith(ERR)): bad.append(f"{ws.title}!{c.coordinate} {v} -> {r!r}")
                elif isinstance(v, (int, float)) and c.row > 4: h += 1
        nform += f; print(f"{ws.title}: {f} formulas, {h} hard-coded numbers")
    shutil.rmtree(tmp, ignore_errors=True)
    for b in bad: print("ERROR", b)
    print(f"{nform} formulas checked, {len(bad)} errors"); sys.exit(1 if bad else 0)

if __name__ == "__main__":
    if len(sys.argv) != 2: sys.exit(__doc__)
    main(sys.argv[1])
