import sys
import tempfile
import unittest
from pathlib import Path

from PIL import Image
from docx import Document
from docx.shared import Cm

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "scripts"))
from image_sizing import plan_image_widths, apply_docx_image_size_policy


class ImageSizingTest(unittest.TestCase):
    def test_per_lecture_median_and_document_fallback(self):
        plan = plan_image_widths([
            {"key": "a", "group": "lecture:1", "width_cm": 4.0},
            {"key": "b", "group": "lecture:1", "width_cm": 4.2},
            {"key": "c", "group": "lecture:1", "width_cm": 10.0},
            {"key": "d", "group": "lecture:2", "width_cm": 12.0},
        ])
        self.assertAlmostEqual(plan["items"]["c"]["final_width_cm"], 4.1, places=2)
        self.assertAlmostEqual(plan["items"]["d"]["final_width_cm"], 4.1, places=2)
        self.assertFalse(plan["items"]["a"]["resized"])
        self.assertEqual(plan["resized_count"], 2)
        self.assertEqual(plan["groups"]["lecture:2"]["reference_source"], "document_median_fallback")

    def test_never_enlarges_and_uses_fixed_fallback_only_without_references(self):
        plan = plan_image_widths([
            {"key": "a", "group": "lecture:1", "width_cm": 12.0},
            {"key": "b", "group": "lecture:2", "width_cm": 8.0},
        ])
        self.assertEqual(plan["items"]["a"]["final_width_cm"], 4.05)
        self.assertEqual(plan["items"]["b"]["final_width_cm"], 4.05)
        self.assertEqual(plan["resized_count"], 2)

    def test_docx_application_preserves_aspect_ratio_and_content(self):
        with tempfile.TemporaryDirectory() as tmp:
            folder = Path(tmp)
            image_path = folder / "diagram.png"
            Image.new("RGB", (200, 100), "white").save(image_path)

            docx_path = folder / "working.docx"
            doc = Document()
            doc.add_paragraph("第1讲：图片尺寸测试")
            doc.add_picture(str(image_path), width=Cm(4.0))
            doc.add_picture(str(image_path), width=Cm(4.2))
            doc.add_picture(str(image_path), width=Cm(10.0))
            doc.add_paragraph("第2讲：没有小图参照")
            doc.add_picture(str(image_path), width=Cm(12.0))
            doc.add_paragraph("保留原文")
            doc.save(docx_path)

            before = Document(docx_path)
            before_text = [p.text for p in before.paragraphs]
            before_sizes = [(int(s.width), int(s.height)) for s in before.inline_shapes]
            before_ratios = [w / h for w, h in before_sizes]

            report = apply_docx_image_size_policy(docx_path)

            after = Document(docx_path)
            after_sizes = [(int(s.width), int(s.height)) for s in after.inline_shapes]
            after_ratios = [w / h for w, h in after_sizes]

            self.assertEqual(report["resized_count"], 2)
            self.assertEqual(before_text, [p.text for p in after.paragraphs])
            self.assertAlmostEqual(after_sizes[0][0] / 360000, 4.0, places=2)
            self.assertAlmostEqual(after_sizes[1][0] / 360000, 4.2, places=2)
            self.assertAlmostEqual(after_sizes[2][0] / 360000, 4.1, places=2)
            self.assertAlmostEqual(after_sizes[3][0] / 360000, 4.1, places=2)
            for old, new in zip(before_ratios, after_ratios):
                self.assertAlmostEqual(old, new, places=4)


if __name__ == "__main__":
    unittest.main()
