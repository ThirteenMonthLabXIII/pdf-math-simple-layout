"""Preserve and classify Word list labels before conversion to text."""
import copy
import sys
import tempfile
import unittest
from pathlib import Path

from docx import Document
from docx.oxml.ns import qn
from lxml import etree

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "scripts"))

from docx_styling import apply_docx_styles, classify_paragraph
from question_layout import (apply_question_layout, build_question_groups,
                             _normalize_question_numbering_indent, _numbering_format)


def bracket_numbering(document):
    numbering = document.part.numbering_part.element
    num_id = document.styles["List Number"]._element.pPr.numPr.numId.val
    num = numbering.xpath(f'./w:num[@w:numId="{num_id}"]')[0]
    abstract_id = num.abstractNumId.val
    original = numbering.xpath(f'./w:abstractNum[@w:abstractNumId="{abstract_id}"]')[0]
    clone = copy.deepcopy(original)
    next_id = max(int(n.get(qn("w:abstractNumId"))) for n in numbering.findall(qn("w:abstractNum"))) + 1
    clone.set(qn("w:abstractNumId"), str(next_id))
    level = next(n for n in clone if n.tag == qn("w:lvl"))
    level.find(qn("w:lvlText")).set(qn("w:val"), "(%1)")
    numbering.append(clone)
    return numbering.add_num(next_id).numId


def set_list(paragraph, num_id):
    paragraph._p.get_or_add_pPr().get_or_add_numPr().get_or_add_numId().val = num_id


class WordListRecognitionTest(unittest.TestCase):
    def test_math_topic_heading_does_not_turn_a_sentence_into_a_heading(self):
        doc = Document()
        self.assertEqual(classify_paragraph(doc.add_paragraph("[例题1]")), "heading")
        self.assertEqual(classify_paragraph(doc.add_paragraph("二元一次方程组")), "heading")
        self.assertEqual(classify_paragraph(doc.add_paragraph("【代入消元法】")), "heading")
        self.assertEqual(classify_paragraph(doc.add_paragraph("二元一次方程组的解是什么？")), "body")

    def test_inherited_and_parenthesized_lists_remain_one_question_group(self):
        with tempfile.TemporaryDirectory() as folder:
            path = Path(folder) / "numbered.docx"
            doc = Document()
            sub_num = bracket_numbering(doc)
            doc.add_paragraph("计算第一个问题。", style="List Number")
            sub1 = doc.add_paragraph("求第一个结果。")
            sub2 = doc.add_paragraph("求第二个结果。")
            set_list(sub1, sub_num)
            set_list(sub2, sub_num)
            doc.add_paragraph("计算另一个问题。", style="List Number")
            expected_numbering = etree.tostring(doc.part.numbering_part.element)
            doc.save(path)

            apply_docx_styles(path, {})
            styled = Document(path)
            self.assertEqual([_numbering_format(p) for p in styled.paragraphs], ["decimal"] * 4)
            _, groups = build_question_groups(styled)
            self.assertEqual(len(groups), 2)
            self.assertEqual(len(groups[0]["subquestion_starts"]), 2)

            report = apply_question_layout(path, {})
            final = Document(path)
            self.assertEqual(report["question_group_count"], 2)
            self.assertEqual(etree.tostring(final.part.numbering_part.element), expected_numbering)
            self.assertTrue(all(_numbering_format(p) == "decimal" for p in final.paragraphs if p.text))

    def test_question_indent_does_not_mutate_shared_list_definition(self):
        doc = Document()
        first = doc.add_paragraph("第一题。", style="List Number")
        other = doc.add_paragraph("共享编号定义的其他段落。", style="List Number")
        before = etree.tostring(doc.part.numbering_part.element)
        _normalize_question_numbering_indent(first)
        self.assertEqual(etree.tostring(doc.part.numbering_part.element), before)
        self.assertEqual(first.paragraph_format.left_indent.twips, 0)
        self.assertIsNone(other.paragraph_format.left_indent)


if __name__ == "__main__":
    unittest.main()
