"""Actual Word pagination, preservation, and failure-closed tests for Enter mode."""
import hashlib
import json
import base64
import sys
import tempfile
import unittest
from argparse import Namespace
from contextlib import redirect_stdout
from io import StringIO
from pathlib import Path

import pymupdf
from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "scripts"))

from docx_styling import apply_docx_styles
from enter_pagination import SPACER_STYLE, find_forbidden_controls
from question_layout import apply_question_layout, check_question_layout, check_word_pagination


def digest(path):
    return hashlib.sha256(path.read_bytes()).hexdigest()


class EnterPaginationTest(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.root = Path(self.tmp.name)
        self.docx = self.root / "questions.docx"

    def tearDown(self):
        self.tmp.cleanup()

    def make_questions(self):
        doc = Document()
        png = self.root / "figure.png"
        png.write_bytes(base64.b64decode(
            "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVQIHWP4z8DwHwAFgAI/ScLttAAAAABJRU5ErkJggg=="))
        doc.styles["Normal"].paragraph_format.keep_with_next = True
        doc.add_paragraph("第1讲：方程")
        for number in range(1, 7):
            doc.add_paragraph(f"{number}．某工程问题请列出方程并求出结果。")
            if number == 2:
                doc.add_paragraph().add_run().add_picture(str(png), width=Cm(0.5))
                math_p = doc.add_paragraph()
                math = OxmlElement("m:oMath")
                math_run = OxmlElement("m:r")
                math_text = OxmlElement("m:t")
                math_text.text = "x+1"
                math_run.append(math_text)
                math.append(math_run)
                math_p._p.append(math)
            doc.add_paragraph("（1）计算甲乙的工作效率。")
            if number == 2:
                table = doc.add_table(rows=2, cols=2)
                table.cell(0, 0).text = "甲"
                table.cell(0, 1).text = "乙"
                table.cell(1, 0).text = "2"
                table.cell(1, 1).text = "3"
            doc.add_paragraph("（2）求出两人合作所需的时间。")
        doc.add_paragraph()
        doc.add_paragraph()
        doc.save(self.docx)
        apply_docx_styles(self.docx, {})

    @unittest.skipUnless(sys.platform == "win32", "requires Microsoft Word on Windows")
    def test_word_moves_complete_groups_and_preserves_answer_space_on_repeat(self):
        self.make_questions()
        original_text = [p.text for p in Document(self.docx).paragraphs if p.text]
        first = apply_question_layout(self.docx, {}, "enter")
        self.assertGreater(first["pagination"]["inserted_page_spacers"], 0)
        self.assertEqual(first["question_group_count"], 6)
        self.assertEqual(first["word_pagination"]["status"], "accepted")
        self.assertGreater(first["word_pagination"]["word_page_count"], 1)
        doc = Document(self.docx)
        self.assertEqual(original_text, [p.text for p in doc.paragraphs if p.text])
        self.assertEqual(doc.tables[0].cell(1, 1).text, "3")
        self.assertEqual(len(doc.inline_shapes), 1)
        self.assertEqual("".join(n.text or "" for n in doc.element.iter(qn("m:t"))), "x+1")
        self.assertEqual(find_forbidden_controls(doc), [])
        spacers = [p for p in doc.paragraphs if p.style.name == SPACER_STYLE]
        self.assertGreater(len(spacers), 0)
        self.assertTrue(all(p.text == "" and all(c.tag == qn("w:pPr") for c in p._p) for p in spacers))
        before = check_question_layout(self.docx, {}, "enter")
        self.assertEqual(before["status"], "ready_for_word_pagination_check")
        second = apply_question_layout(self.docx, {}, "enter")
        after = check_question_layout(self.docx, {}, "enter")
        self.assertEqual(second["pagination"]["replaced_prior_page_spacers"], len(spacers))
        self.assertEqual(before["existing_empty_paragraphs"], after["existing_empty_paragraphs"])
        self.assertEqual(second["word_pagination"]["status"], "accepted")

    @unittest.skipUnless(sys.platform == "win32", "requires Microsoft Word on Windows")
    def test_oversized_question_does_not_change_original(self):
        doc = Document()
        doc.add_paragraph("1．这是一道完整的大题。")
        for number in range(80):
            doc.add_paragraph(f"题目条件第{number}行，内容属于同一道题。")
        doc.save(self.docx)
        apply_docx_styles(self.docx, {})
        before = digest(self.docx)
        with self.assertRaisesRegex(RuntimeError, "question_too_tall"):
            apply_question_layout(self.docx, {}, "enter")
        self.assertEqual(digest(self.docx), before)

    def test_manual_line_break_rejected_without_changing_source(self):
        doc = Document()
        para = doc.add_paragraph("1．题干第一行")
        para.add_run().add_break()
        para.add_run("题干第二行")
        doc.save(self.docx)
        before = digest(self.docx)
        with self.assertRaisesRegex(RuntimeError, "unsupported_line_or_column_break"):
            apply_question_layout(self.docx, {}, "enter")
        self.assertEqual(digest(self.docx), before)

    def test_unsupported_body_content_is_not_silently_skipped(self):
        doc = Document()
        doc.add_paragraph("1．题目。")
        embedded = OxmlElement("w:altChunk")
        doc.element.body.insert(1, embedded)
        doc.save(self.docx)
        before = digest(self.docx)
        with self.assertRaisesRegex(RuntimeError, "unsupported_top_level_content"):
            apply_question_layout(self.docx, {}, "enter")
        self.assertEqual(digest(self.docx), before)

    @unittest.skipUnless(sys.platform == "win32", "requires Microsoft Word on Windows")
    def test_word_detects_one_long_paragraph_crossing_pages(self):
        doc = Document()
        doc.add_paragraph("1．" + "这一段内容属于同一道题。" * 500)
        doc.add_paragraph()
        doc.save(self.docx)
        apply_docx_styles(self.docx, {})
        result = check_word_pagination(self.docx, "enter")
        self.assertEqual(result["status"], "needs_layout_work")
        self.assertTrue(any(f["reason"] == "question_spans_pages" for f in result["split_question_groups"]))

    @unittest.skipUnless(sys.platform == "win32", "requires Microsoft Word on Windows")
    def test_mineru_prepare_and_audit_retain_enter_mode(self):
        import mineru_handoff
        self.make_questions()
        pdf = self.root / "source.pdf"
        source = pymupdf.open()
        source.new_page().insert_text((72, 72), "Question source")
        source.save(str(pdf))
        source.close()
        workspace = self.root / "job"
        with redirect_stdout(StringIO()):
            mineru_handoff.prepare(Namespace(pdf=pdf, mineru_docx=self.docx, workspace=workspace,
                                             profile=None, pagination_mode="enter"))
            mineru_handoff.audit(Namespace(workspace=workspace, docx=None, word_pages=True,
                                           require_accepted=False))
        manifest = json.loads((workspace / "mineru_manifest.json").read_text("utf-8"))
        audit = json.loads((workspace / "mineru_audit.json").read_text("utf-8"))
        self.assertEqual(manifest["pagination_mode"], "enter")
        self.assertEqual(manifest["question_layout_preflight_status"], "ready_for_codex_formatting")
        self.assertEqual(digest(self.docx), digest(workspace / "mineru_original.docx"))
        self.assertTrue(audit["checks"]["question_layout_complete"])
        self.assertTrue(audit["checks"]["question_groups_same_page"])

    @unittest.skipUnless(sys.platform == "win32", "requires Microsoft Word on Windows")
    def test_fallback_build_and_audit_retain_enter_mode(self):
        import layout_cli
        pdf = self.root / "source.pdf"
        source = pymupdf.open()
        source.new_page().insert_text((72, 72), "Question source")
        source.save(str(pdf))
        source.close()
        workspace = self.root / "fallback"
        with redirect_stdout(StringIO()):
            layout_cli.cmd_prepare(Namespace(pdf=pdf, output=workspace, dpi=100, overwrite=False))
            data = layout_cli.load(workspace / "layout.json")
            data["pages"][0]["blocks"] = [
                {"type": "question", "text": "1．计算结果。"},
                {"type": "subquestion", "text": "（1）求 x 的值。"},
            ]
            data["pages"][0]["verified"] = True
            layout_cli.save(workspace / "layout.json", data)
            layout_cli.cmd_build(Namespace(workspace=workspace, profile=None, output=None,
                                           allow_incomplete=False, overwrite=False, pagination_mode="enter"))
            layout_cli.cmd_audit(Namespace(workspace=workspace, docx=None, word_pages=True))
        report = layout_cli.load(workspace / "build_report.json")
        audit = layout_cli.load(workspace / "audit_report.json")
        self.assertEqual(report["pagination_mode"], "enter")
        self.assertEqual(report["question_layout_preflight_status"], "ready_for_content_review")
        self.assertTrue(audit["checks"]["question_layout_complete"])
        self.assertTrue(audit["checks"]["question_groups_same_page"])


if __name__ == "__main__":
    unittest.main()
