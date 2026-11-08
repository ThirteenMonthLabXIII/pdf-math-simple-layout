"""Regression tests for Word paragraph layout anomalies."""
import json
import sys
import tempfile
import unittest
from pathlib import Path

from docx import Document
from docx.oxml import parse_xml
from docx.shared import Twips

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "scripts"))
from docx_layout_check import evaluate_layout_review, inspect_docx, make_review_template


MATH_XML = (
    '<m:oMath xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">'
    '<m:r><m:t>x</m:t></m:r></m:oMath>'
)


class DocxLayoutCheckTest(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.root = Path(self.temp.name)

    def tearDown(self):
        self.temp.cleanup()

    def test_flags_1687_twip_formula_indent_and_four_space_answer_blank(self):
        path = self.root / "suspect.docx"
        doc = Document()

        formula = doc.add_paragraph("formula x")
        formula.paragraph_format.left_indent = Twips(1687)
        formula._p.append(parse_xml(MATH_XML))

        # A modest 600-twip nested-question indent is intentional and must remain.
        nested = doc.add_paragraph("（2）nested subquestion")
        nested.paragraph_format.left_indent = Twips(600)

        doc.add_paragraph("答案：（    ）")
        doc.save(path)

        report = inspect_docx(path)
        codes = [item["code"] for item in report["findings"]]
        self.assertIn("formula_paragraph_indent", codes)
        self.assertIn("answer_blank_spacing", codes)
        self.assertNotIn("paragraph_indent", codes)
        formula_finding = next(x for x in report["findings"] if x["code"] == "formula_paragraph_indent")
        self.assertEqual(formula_finding["value_twips"], 1687)
        answer_finding = next(x for x in report["findings"] if x["code"] == "answer_blank_spacing")
        self.assertEqual(answer_finding["space_count"], 4)

    def test_fixed_formula_indent_and_two_space_answer_have_no_findings(self):
        path = self.root / "fixed.docx"
        doc = Document()
        formula = doc.add_paragraph("formula x")
        formula._p.append(parse_xml(MATH_XML))
        doc.add_paragraph("答案：（  ）")
        doc.save(path)
        self.assertEqual(inspect_docx(path)["findings"], [])

    def test_review_requires_current_hash_and_source_pdf_confirmation(self):
        path = self.root / "suspect.docx"
        review_path = self.root / "layout_review.json"
        doc = Document()
        formula = doc.add_paragraph("formula x")
        formula.paragraph_format.left_indent = Twips(1687)
        formula._p.append(parse_xml(MATH_XML))
        doc.save(path)

        template = make_review_template(path, review_path)
        self.assertEqual(len(template["findings"]), 1)
        review = json.loads(review_path.read_text("utf-8"))
        review["checked"] = True
        review["reviewer"] = "Codex regression test"
        review["findings"][0].update({
            "disposition": "intentional",
            "reason": "Synthetic test: the reviewer explicitly compared this placement with the source PDF.",
            "source_pdf_checked": False,
        })
        review_path.write_text(json.dumps(review), encoding="utf-8")
        report = inspect_docx(path)
        self.assertFalse(evaluate_layout_review(path, review_path, report)["complete"])
        review["findings"][0]["source_pdf_checked"] = True
        review_path.write_text(json.dumps(review), encoding="utf-8")
        self.assertTrue(evaluate_layout_review(path, review_path, report)["complete"])

        # A later DOCX edit invalidates the review even if the JSON was not changed.
        doc = Document(path)
        doc.add_paragraph("new content")
        doc.save(path)
        self.assertFalse(evaluate_layout_review(path, review_path)["complete"])


if __name__ == "__main__":
    unittest.main()
