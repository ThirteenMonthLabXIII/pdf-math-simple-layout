"""No OCR service/Word installation required to test deterministic core."""
import json
import sys
import tempfile
import unittest
from pathlib import Path

ROOT=Path(__file__).resolve().parents[1]
sys.path.insert(0,str(ROOT/'scripts'))
import layout_cli as c


class WorkflowTest(unittest.TestCase):
    def setUp(self):
        self.t=tempfile.TemporaryDirectory()
        self.root=Path(self.t.name)
        self.pdf=self.root/'source.pdf'
        doc=c.fitz.open();page=doc.new_page(width=595,height=842)
        page.insert_text((55,60),"Sample",fontsize=12)
        page.draw_line((80,145),(180,145),color=(0,0,0),width=2)
        doc.save(str(self.pdf));doc.close()
        self.ws=self.root/'ws'
        c.cmd_prepare(type('A',(),dict(pdf=self.pdf,output=self.ws,dpi=100,overwrite=False)))
        self.layout=c.load(self.ws/'layout.json')
        self.layout['pages'][0]['blocks']=[
            {'type':'title','text':'第 1 讲：测试'},
            {'type':'question','text':'1．已知 x = 2，求 y。'},
            {'type':'subquestion','text':'（1）计算 y 的值'},
            {'type':'table','rows':[['列1','列2'],['A','B']]},
            {'type':'image','bbox':[70,120,190,155],'caption':'第1题图'},
        ]
        self.layout['pages'][0]['verified']=True
        c.save(self.ws/'layout.json',self.layout)

    def tearDown(self):self.t.cleanup()

    def test_profile_and_build_and_audit(self):
        output=self.ws/'candidate.docx'
        c.cmd_build(type('A',(),dict(workspace=self.ws,profile=None,output=None,allow_incomplete=False,overwrite=False)))
        self.assertTrue(output.exists())
        d=c.Document(str(output))
        self.assertTrue(any('第 1 讲' in p.text for p in d.paragraphs))
        self.assertEqual(len(d.tables),1)
        self.assertEqual(len(d.inline_shapes),1)
        self.assertEqual(d.paragraphs[0].style.name,'SIJI Title')
        self.assertEqual(d.paragraphs[1].style.name,'SIJI Question')
        self.assertEqual(d.paragraphs[2].style.name,'SIJI Subquestion')
        self.assertEqual(d.paragraphs[2].style.paragraph_format.left_indent.twips,600)
        figure_paragraph=next(p for p in d.paragraphs if any(node.tag.endswith('drawing') for node in p._p.iter()))
        self.assertEqual(figure_paragraph.style.name,'SIJI Figure')
        self.assertEqual(figure_paragraph.style.paragraph_format.alignment,c.WD_ALIGN_PARAGRAPH.CENTER)
        caption_paragraph=next(p for p in d.paragraphs if p.text=='第1题图')
        self.assertEqual(caption_paragraph.style.name,'SIJI Caption')
        build_report=c.load(self.ws/'build_report.json')
        self.assertEqual(build_report['style_policy']['schema'],'docx-style-policy/1')
        c.cmd_audit(type('A',(),dict(workspace=self.ws,docx=None)))
        report=c.load(self.ws/'audit_report.json')
        self.assertEqual(report['status'],'needs_review')
        self.assertFalse(report['checks']['visual_review_complete'])
        review=c.load(self.ws/'visual_review.json')
        review['reviewer']='Tester'
        review['candidate_sha256']=c.sha256(output)
        for row in review['pages']:
            for key in ('source_checked','output_checked','text_math_checked','figures_checked'):
                row[key]=True
        c.save(self.ws/'visual_review.json',review)
        c.cmd_audit(type('A',(),dict(workspace=self.ws,docx=None)))
        report=c.load(self.ws/'audit_report.json')
        self.assertEqual(report['status'],'needs_review')
        self.assertFalse(report['checks']['docx_layout_review_complete'])
        layout_review=c.load(self.ws/'layout_review.json')
        layout_review['candidate_sha256']=c.sha256(output)
        layout_review['checked']=True
        layout_review['reviewer']='Tester'
        layout_review['findings']=[]
        c.save(self.ws/'layout_review.json',layout_review)
        c.cmd_audit(type('A',(),dict(workspace=self.ws,docx=None)))
        self.assertEqual(c.load(self.ws/'audit_report.json')['status'],'accepted')

    def test_preflight_does_not_modify_source(self):
        self.assertEqual(c.sha256(self.pdf),self.layout['source_sha256'])
        self.assertEqual(len(self.layout['pages']),1)
        self.assertTrue((self.ws/'pages/page-001.png').exists())
        self.assertEqual(c.classify('【例题 2】测试'),'heading')
        self.assertEqual(c.classify('（2）计算'),'subquestion')
        c.cmd_build(type('A',(),dict(workspace=self.ws,profile=None,output=None,allow_incomplete=False,overwrite=False)))
        with self.assertRaises(FileExistsError):
            c.cmd_build(type('A',(),dict(workspace=self.ws,profile=None,output=None,allow_incomplete=False,overwrite=False)))

    def test_invalid_figure_cannot_erase_content(self):
        self.layout['pages'][0]['blocks'][4]['bbox']=[-2,0,800,900]
        c.save(self.ws/'layout.json',self.layout)
        with self.assertRaises(ValueError):
            c.cmd_build(type('A',(),dict(workspace=self.ws,profile=None,output=None,allow_incomplete=False,overwrite=False)))
        self.assertFalse((self.ws/'candidate.docx').exists())


if __name__=='__main__':unittest.main()
