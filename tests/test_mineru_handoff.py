"""MinerU handoff test with generated synthetic assets (no private documents)."""
import json
import hashlib
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path

import pymupdf
from docx import Document
from docx.oxml import parse_xml
from docx.shared import Twips

ROOT = Path(__file__).resolve().parents[1]
CLI = ROOT / "scripts" / "mineru_handoff.py"
MATH_XML = (
    '<m:oMath xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">'
    '<m:r><m:t>x</m:t></m:r></m:oMath>'
)


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
            d.add_paragraph("1．Synthetic sample question: x + y = 3")
            formula = d.add_paragraph()
            formula.paragraph_format.left_indent = Twips(1687)
            formula._p.append(parse_xml(MATH_XML))
            answer = d.add_paragraph("答案：（    ）")
            exact_answer_text = answer.text
            d.add_picture(self._mock_image(folder), width=__import__("docx").shared.Inches(1))
            d.save(str(docx_path))
            raw_source_hash = hashlib.sha256(docx_path.read_bytes()).hexdigest()
            workspace = folder / "output"

            proc = subprocess.run([sys.executable, str(CLI), "prepare", "--pdf",
                str(pdf_path), "--mineru-docx", str(docx_path),
                "--workspace", str(workspace)], capture_output=True, text=True)
            self.assertEqual(proc.returncode, 0, proc.stderr)
            prepare_report = json.loads(proc.stdout)
            self.assertEqual(prepare_report["status"], "ready_for_codex_formatting")
            self.assertEqual(prepare_report["image_sizing"]["schema"], "image-size-policy/1")
            self.assertEqual(prepare_report["image_sizing"]["resized_count"], 0)
            question_plan = json.loads((workspace / "question_layout_plan.json").read_text(encoding="utf-8"))
            self.assertEqual(question_plan["preflight_status"], "ready_for_codex_formatting")
            self.assertEqual(question_plan["word_pagination"]["status"], "accepted")
            self.assertEqual(question_plan["word_pagination"]["split_question_groups"], [])
            self.assertTrue((workspace / "mineru_original.docx").exists())
            self.assertTrue((workspace / "working.docx").exists())
            self.assertEqual(len(list((workspace / "source_pages").glob("*.png"))), 1)

            raw_doc = Document(workspace / "mineru_original.docx")
            working_doc = Document(workspace / "working.docx")
            self.assertEqual(hashlib.sha256((workspace / "mineru_original.docx").read_bytes()).hexdigest(), raw_source_hash)
            self.assertEqual(raw_doc.paragraphs[1].paragraph_format.left_indent.twips, 1687)
            self.assertIsNone(working_doc.paragraphs[1].paragraph_format.left_indent)
            self.assertEqual(working_doc.paragraphs[1].style.name, "SIJI Formula")
            self.assertEqual(working_doc.paragraphs[1].style.paragraph_format.left_indent.twips, 0)
            self.assertIn("oMath", working_doc.paragraphs[1]._p.xml)
            self.assertEqual(working_doc.paragraphs[2].text, exact_answer_text)
            style_policy = json.loads((workspace / "style_policy.json").read_text("utf-8"))
            self.assertFalse(style_policy["text_mutated"])

            result = subprocess.run([sys.executable, str(CLI), "audit",
                "--workspace", str(workspace), "--require-accepted"],
                capture_output=True, text=True)
            self.assertEqual(result.returncode, 2)
            report = json.loads((workspace / "mineru_audit.json").read_text("utf-8"))
            self.assertEqual(report["status"], "needs_review")
            self.assertTrue(report["checks"]["source_unchanged"])
            self.assertFalse(report["checks"]["all_pdf_pages_reviewed"])
            self.assertFalse(report["checks"]["docx_layout_review_complete"])

            candidate_hash = hashlib.sha256((workspace / "working.docx").read_bytes()).hexdigest()
            mineru_review = json.loads((workspace / "mineru_review.json").read_text("utf-8"))
            mineru_review["candidate_sha256"] = candidate_hash
            mineru_review["word_preview_checked"] = True
            for page in mineru_review["pages"]:
                page["content_checked"] = True
                page["formulas_checked"] = True
                page["figures_checked"] = True
            (workspace / "mineru_review.json").write_text(json.dumps(mineru_review), encoding="utf-8")

            layout_review = json.loads((workspace / "layout_review.json").read_text("utf-8"))
            layout_review["candidate_sha256"] = candidate_hash
            layout_review["checked"] = True
            layout_review["reviewer"] = "Tester"
            findings = json.loads((workspace / "docx_layout_findings.json").read_text("utf-8"))["findings"]
            layout_review["findings"] = [
                {
                    "id": finding["id"],
                    "disposition": "intentional",
                    "reason": "Synthetic source PDF confirms the answer blank width; the original four spaces are preserved.",
                    "source_pdf_checked": True,
                }
                for finding in findings
            ]
            (workspace / "layout_review.json").write_text(json.dumps(layout_review), encoding="utf-8")

            accepted = subprocess.run([sys.executable, str(CLI), "audit",
                "--workspace", str(workspace), "--require-accepted", "--word-pages"], capture_output=True, text=True)
            self.assertEqual(accepted.returncode, 0, accepted.stdout + accepted.stderr)
            final_report = json.loads((workspace / "mineru_audit.json").read_text("utf-8"))
            self.assertEqual(final_report["status"], "accepted")

    @staticmethod
    def _mock_image(folder):
        from PIL import Image
        path = folder / "diagram.png"
        Image.new("RGB", (120, 80), "white").save(str(path))
        return str(path)


if __name__ == "__main__":
    unittest.main()
