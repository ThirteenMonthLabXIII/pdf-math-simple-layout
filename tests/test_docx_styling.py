"""Preventive Word style contract tests."""
import hashlib
import json
import sys
import tempfile
import unittest
from types import SimpleNamespace
from pathlib import Path

from docx import Document
from docx.oxml import parse_xml
from docx.shared import Twips

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "scripts"))
from docx_styling import apply_docx_styles, default_paragraph_policy, extract_paragraph_policy
from layout_cli import cmd_profile


MATH_XML = (
    '<m:oMath xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">'
    '<m:r><m:t>x</m:t></m:r></m:oMath>'
)


def sha(path):
    return hashlib.sha256(Path(path).read_bytes()).hexdigest()


class PreventiveDocxStylingTest(unittest.TestCase):
    def test_preformat_removes_bad_formula_indent_and_preserves_valid_content(self):
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "working.docx"
            doc = Document()

            formula = doc.add_paragraph()
            formula.paragraph_format.left_indent = Twips(1687)
            formula._p.append(parse_xml(MATH_XML))

            subquestion = doc.add_paragraph("（1）子题")
            subquestion.paragraph_format.left_indent = Twips(600)

            answer = doc.add_paragraph("答案：（    ）")
            exact_answer_text = answer.text
            doc.save(path)
            source_hash = sha(path)

            report = apply_docx_styles(path, {})
            styled = Document(path)

            self.assertEqual(sha(path), report["styled_docx_sha256"])
            self.assertNotEqual(source_hash, sha(path))
            self.assertIsNone(styled.paragraphs[0].paragraph_format.left_indent)
            self.assertEqual(styled.paragraphs[0].style.name, "SIJI Formula")
            self.assertEqual(styled.paragraphs[0].style.paragraph_format.left_indent.twips, 0)
            self.assertIsNone(styled.paragraphs[1].paragraph_format.left_indent)
            self.assertEqual(styled.paragraphs[1].style.name, "SIJI Subquestion")
            self.assertEqual(styled.paragraphs[1].style.paragraph_format.left_indent.twips, 600)
            self.assertEqual(styled.paragraphs[2].text, exact_answer_text)
            self.assertEqual(styled.paragraphs[2].text.count(" "), 4)
            self.assertIn("oMath", styled.paragraphs[0]._p.xml)
            self.assertFalse(report["text_mutated"])

    def test_sample_policy_learns_600_twip_subquestion_but_rejects_1687_formula(self):
        with tempfile.TemporaryDirectory() as tmp:
            sample_path = Path(tmp) / "teacher-sample.docx"
            doc = Document()
            formula = doc.add_paragraph()
            formula.paragraph_format.left_indent = Twips(1687)
            formula._p.append(parse_xml(MATH_XML))
            subquestion = doc.add_paragraph("（1）子题")
            subquestion.paragraph_format.left_indent = Twips(600)
            doc.save(sample_path)

            policy, warnings = extract_paragraph_policy(sample_path, default_paragraph_policy())
            self.assertEqual(policy["formula"]["left_indent_twips"], 0)
            self.assertEqual(policy["subquestion"]["left_indent_twips"], 600)
            self.assertTrue(any("over 1000 twips" in warning for warning in warnings))

    def test_safe_default_formula_policy_is_zero_indent(self):
        policy = default_paragraph_policy()
        self.assertEqual(policy["formula"]["left_indent_twips"], 0)
        self.assertEqual(policy["formula"]["right_indent_twips"], 0)
        self.assertEqual(policy["formula"]["first_line_indent_twips"], 0)

    def test_profile_command_exports_safe_style_contract(self):
        with tempfile.TemporaryDirectory() as tmp:
            sample_path = Path(tmp) / "teacher-sample.docx"
            profile_path = Path(tmp) / "style_profile.json"
            doc = Document()
            doc.add_paragraph("第 1 讲：样例")
            formula = doc.add_paragraph()
            formula.paragraph_format.left_indent = Twips(1687)
            formula._p.append(parse_xml(MATH_XML))
            subquestion = doc.add_paragraph("（1）子题")
            subquestion.paragraph_format.left_indent = Twips(600)
            doc.save(sample_path)

            cmd_profile(SimpleNamespace(sample=sample_path, output=profile_path))
            profile = json.loads(profile_path.read_text("utf-8"))
            self.assertEqual(profile["paragraph_policy"]["formula"]["left_indent_twips"], 0)
            self.assertEqual(profile["paragraph_policy"]["subquestion"]["left_indent_twips"], 600)
            self.assertTrue(any("over 1000 twips" in warning for warning in profile["paragraph_policy_warnings"]))


if __name__ == "__main__":
    unittest.main()
