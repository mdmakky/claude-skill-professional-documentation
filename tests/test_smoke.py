import json, os, shutil, subprocess, sys, tempfile, unittest
ROOT = sys.argv.pop(1) if len(sys.argv) > 1 else os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
PY = sys.executable
run = lambda *a: subprocess.run([PY, *a], capture_output=True, text=True)

class Smoke(unittest.TestCase):
    def setUp(self): self.t = tempfile.mkdtemp()
    def tearDown(self): shutil.rmtree(self.t, ignore_errors=True)

    def test_templates_build_docx_and_pdf(self):
        for name in ("business-report", "market-research-report", "sop"):
            for ext in ("docx", "pdf"):
                out = f"{self.t}/{name}.{ext}"
                r = run(f"{ROOT}/scripts/build_doc.py", f"{ROOT}/templates/{name}.json", out)
                self.assertEqual(r.returncode, 0, r.stderr); self.assertGreater(os.path.getsize(out), 1000)

    def test_refuses_overwrite(self):
        out = f"{self.t}/a.docx"; spec = f"{ROOT}/templates/sop.json"
        self.assertEqual(run(f"{ROOT}/scripts/build_doc.py", spec, out).returncode, 0)
        self.assertNotEqual(run(f"{ROOT}/scripts/build_doc.py", spec, out).returncode, 0)

    def test_numbered_lists_restart(self):
        from docx import Document
        spec = {"title": "T", "blocks": [{"numbered": ["a", "b"]}, {"h1": "x"}, {"numbered": ["c"]}]}
        json.dump(spec, open(f"{self.t}/s.json", "w"))
        run(f"{ROOT}/scripts/build_doc.py", f"{self.t}/s.json", f"{self.t}/s.docx")
        ids = [p._p.pPr.numPr.numId.val for p in Document(f"{self.t}/s.docx").paragraphs if p.style.name == "List Number"]
        self.assertEqual(len(set(ids)), 2); self.assertEqual(ids[0], ids[1])

    def test_unicode_pdf(self):
        json.dump({"title": "T", "blocks": [{"p": "Bengali: বাংলা"}]}, open(f"{self.t}/u.json", "w"), ensure_ascii=False)
        self.assertEqual(run(f"{ROOT}/scripts/build_doc.py", f"{self.t}/u.json", f"{self.t}/u.pdf").returncode, 0)

    @unittest.skipUnless(shutil.which("soffice"), "LibreOffice not installed")
    def test_workbook_formulas_have_no_errors(self):
        out = f"{self.t}/m.xlsx"
        self.assertEqual(run(f"{ROOT}/scripts/build_xlsx.py", f"{ROOT}/templates/financial-model.xlsx.json", out).returncode, 0)
        r = run(f"{ROOT}/scripts/check_xlsx.py", out); self.assertEqual(r.returncode, 0, r.stdout + r.stderr)

if __name__ == "__main__":
    unittest.main(verbosity=2)
