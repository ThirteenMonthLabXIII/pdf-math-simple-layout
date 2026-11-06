"""MinerU handoff test with generated synthetic assets (no private documents)."""
import json
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path

import pymupdf
from docx import Document

ROOT = Path(__file__).resolve().parents[1]
CLI = ROOT / "scripts" / "mineru_handoff.py"


class MinerUHandoffTest(unittest.TestCase):
    def test_requires_real_source_bound_review(self):
        with tempfile.TemporaryDirectory() as tmp:
            folder = Path(tmp)
            pdf_path = folder / "sample.pdf"
            pdf = pymupdf.open()
            p = pdf.new_page()
            p.insert_text((80, 80), "Equation sample")
            pdf.save(str(pdf_path))
            pdf.close()

            docx_path = folder / "mineru.docx"
            d = Document()
            d.add_paragraph("Synthetic sample question 1: x + y = 3")
            d.add_picture(self._mock_image(folder), width=__import__("docx").shared.Inches(1))
            d.save(str(docx_path))
            workspace = folder / "output"

            proc = subprocess.run([sys.executable, str(CLI), "prepare", "--pdf",
                str(pdf_path), "--mineru-docx", str(docx_path),
                "--workspace", str(workspace)], capture_output=True, text=True)
            self.assertEqual(proc.returncode, 0, proc.stderr)
            self.assertTrue((workspace / "mineru_original.docx").exists())
            self.assertTrue((workspace / "working.docx").exists())
            self.assertEqual(len(list((workspace / "source_pages").glob("*.png"))), 1)

            result = subprocess.run([sys.executable, str(CLI), "audit",
                "--workspace", str(workspace), "--require-accepted"],
                capture_output=True, text=True)
            self.assertEqual(result.returncode, 2)
            report = json.loads((workspace / "mineru_audit.json").read_text("utf-8"))
            self.assertEqual(report["status"], "needs_review")
            self.assertTrue(report["checks"]["source_unchanged"])
            self.assertFalse(report["checks"]["all_pdf_pages_reviewed"])

    @staticmethod
    def _mock_image(folder):
        from PIL import Image
        path = folder / "diagram.png"
        Image.new("RGB", (120, 80), "white").save(str(path))
        return str(path)


if __name__ == "__main__":
    unittest.main()
