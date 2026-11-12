"""Preventive question layout tests: one-page groups and Enter-only writing space."""
import sys
import tempfile
import unittest
from pathlib import Path

from docx import Document
from docx.oxml.ns import qn

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "scripts"))

from docx_styling import apply_docx_styles
from question_layout import (
    apply_question_layout,
    check_question_layout,
    check_word_pagination,
)


class QuestionLayoutTest(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.path = Path(self.tmp.name) / "questions.docx"

    def tearDown(self):
        self.tmp.cleanup()

    def _make_sample(self):
        doc = Document()
        doc.add_paragraph("第1讲：方程组")
        doc.add_paragraph("【例题1】应用题")
        doc.add_paragraph("1．某商店购买商品，列方程求解。")
        doc.add_paragraph("（1）计算甲乙两种商品的数量。")
        doc.add_paragraph("（2）证明该方案满足题目条件。")
        doc.add_paragraph("2．填空题：求出x的值。")
        doc.save(self.path)
        apply_docx_styles(self.path, {})

    def test_each_subquestion_gets_its_own_enter_space_and_is_idempotent(self):
        self._make_sample()

        first = apply_question_layout(self.path, {})
        self.assertEqual(first["question_group_count"], 2)
        self.assertEqual(
            [row["required_empty_paragraphs"] for row in first["answer_space_segments"]],
            [4, 6, 2],
        )
        self.assertEqual(first["inserted_empty_paragraphs"], 12)

        checked = check_question_layout(self.path, {})
        self.assertEqual(checked["status"], "ready_for_word_pagination_check")
        self.assertEqual(checked["findings"], [])
        self.assertEqual(checked["required_empty_paragraphs"], 12)
        self.assertEqual(checked["existing_empty_paragraphs"], 12)

        doc = Document(self.path)
        blanks = [p for p in doc.paragraphs if p.style.name == "SIJI AnswerBlank"]
        self.assertEqual(len(blanks), 12)
        for paragraph in blanks:
            self.assertEqual(paragraph.text, "")
            self.assertIsNone(paragraph._p.pPr.find(qn("w:pBdr")))
            for tag in ("tabs", "ind", "spacing", "shd", "framePr", "jc"):
                self.assertIsNone(paragraph._p.pPr.find(qn("w:" + tag)))
            self.assertEqual(paragraph.style.paragraph_format.space_before.pt, 0)
            self.assertEqual(paragraph.style.paragraph_format.space_after.pt, 0)
            self.assertAlmostEqual(paragraph.style.paragraph_format.line_spacing, 1.15)
            self.assertFalse(any(node.tag == qn("w:br") for node in paragraph._p.iter()))
            self.assertFalse(any(node.tag == qn("w:tab") for node in paragraph._p.iter()))
            self.assertFalse(any(node.tag == qn("w:t") for node in paragraph._p.iter()))

        starts = [p for p in doc.paragraphs if p.style.name == "SIJI Question Start"]
        self.assertEqual(len(starts), 2)
        for paragraph in starts:
            self.assertEqual(paragraph.paragraph_format.left_indent, None)
            self.assertEqual(paragraph.paragraph_format.first_line_indent, None)
            self.assertEqual(paragraph.style.paragraph_format.left_indent.twips, 0)
            self.assertEqual(paragraph.style.paragraph_format.first_line_indent.twips, 0)

        second = apply_question_layout(self.path, {})
        self.assertEqual(second["inserted_empty_paragraphs"], 0)
        self.assertEqual(
            len([p for p in Document(self.path).paragraphs if p.style.name == "SIJI AnswerBlank"]),
            12,
        )

    def test_check_detects_missing_answer_space(self):
        self._make_sample()
        apply_question_layout(self.path, {})
        doc = Document(self.path)
        blank = next(p for p in doc.paragraphs if p.style.name == "SIJI AnswerBlank")
        blank._element.getparent().remove(blank._element)
        doc.save(self.path)

        checked = check_question_layout(self.path, {})
        self.assertEqual(checked["status"], "needs_layout_work")
        self.assertTrue(any(item["code"] == "insufficient_answer_space" for item in checked["findings"]))

    def test_manual_page_break_wrapper_is_removed(self):
        doc = Document()
        doc.add_paragraph("1．求解方程。")
        doc.add_page_break()
        doc.add_paragraph("（1）计算结果。")
        doc.save(self.path)
        apply_docx_styles(self.path, {})

        report = apply_question_layout(self.path, {})
        self.assertEqual(report["removed_source_page_controls"]["manual_page_breaks"], 1)
        self.assertEqual(report["removed_source_page_controls"]["empty_page_break_paragraphs"], 1)
        checked = check_question_layout(self.path, {})
        self.assertEqual(checked["manual_page_breaks"], 0)
        self.assertFalse(any(item["code"] == "manual_page_break_present" for item in checked["findings"]))

    def test_parenthesis_numbered_question_is_detected(self):
        doc = Document()
        doc.add_paragraph("1) 求解方程。")
        doc.save(self.path)
        apply_docx_styles(self.path, {})
        report = apply_question_layout(self.path, {})
        self.assertEqual(report["question_group_count"], 1)
        checked = check_question_layout(self.path, {})
        self.assertEqual(checked["status"], "ready_for_word_pagination_check")

    def test_unrecognized_question_structure_is_not_silently_accepted(self):
        doc = Document()
        doc.add_paragraph("这是一段普通说明文字，没有可识别题号。")
        doc.save(self.path)
        apply_docx_styles(self.path, {})

        report = apply_question_layout(self.path, {})
        self.assertEqual(report["question_group_count"], 0)
        checked = check_question_layout(self.path, {})
        self.assertEqual(checked["status"], "needs_layout_work")
        self.assertTrue(any(item["code"] == "no_question_groups_detected" for item in checked["findings"]))

    def test_word_pagination_moves_last_question_as_a_whole_to_next_page(self):
        try:
            import win32com.client  # noqa: F401
        except ImportError:
            self.skipTest("Microsoft Word COM is unavailable")

        doc = Document()
        for index in range(65):
            doc.add_paragraph(f"前置说明第{index + 1}行：用于占据本页剩余空间。")
        doc.add_paragraph("【例题1】应用题")
        doc.add_paragraph("1．某工程需要计算工作总量并列方程求解。")
        doc.add_paragraph("（1）计算甲乙两人的工作效率。")
        doc.add_paragraph("（2）求出两人合作所需的时间。")
        doc.save(self.path)
        apply_docx_styles(self.path, {})
        apply_question_layout(self.path, {})

        result = check_word_pagination(self.path)
        self.assertEqual(result["status"], "accepted", result)
        self.assertEqual(result["question_group_count"], 1)
        self.assertEqual(result["split_question_groups"], [])
        self.assertGreaterEqual(result["word_page_count"], 2)


if __name__ == "__main__":
    unittest.main()
