# PDF Math Simple Layout

A Codex-friendly skill to turn scanned mathematics handouts into editable Microsoft Word documents.

**Preferred:** MinerU → editable DOCX → Word formatting and proofreading → render → per-page review.  
**Fallback:** local RapidOCR → editable DOCX using structured layout JSON.

Features:
- Preserve source PDFs and raw MinerU DOCX exports as immutable originals.
- Keep mathematical text editable; preserve geometry diagrams as high-resolution crops.
- Retain the original question order, mathematical notation, figures, and tables.
- Apply formatting from a reference Word document without copying its question content.
- Refuse to mark a document accepted until source pages, rendered Word pages and final hashes are checked.
- Scan DOCX paragraph XML for abnormal math indents, excessive paragraph spacing, and mismatched answer-blank spacing; bind the review to the final DOCX hash.

## Install (Windows; Python 3.11+)

~~~powershell
git clone https://github.com/ThirteenMonthLabXIII/pdf-math-simple-layout.git
cd pdf-math-simple-layout
py -3 -m venv .venv
.\.venv\Scripts\python.exe -m pip install -r requirements.txt
.\.venv\Scripts\python.exe -m unittest discover -s tests -v
~~~

## MinerU-first example

First export the **entire** source PDF from MinerU as an editable DOCX.

~~~powershell
$py = ".\.venv\Scripts\python.exe"
& $py scripts\mineru_handoff.py prepare --pdf "input.pdf" --mineru-docx "mineru.docx" --workspace "job1"
# Correct and format job1\working.docx in Microsoft Word.
& $py scripts\layout_cli.py render --docx "job1\working.docx" --output "job1\word_preview"
& $py scripts\docx_layout_check.py check --docx "job1\working.docx" --output "job1\docx_layout_findings.json"
& $py scripts\docx_layout_check.py template --docx "job1\working.docx" --output "job1\layout_review.json"
# Compare all findings with the source PDF; fix or explicitly justify any retained item.
# Complete layout_review.json and mineru_review.json against the final DOCX hash.
& $py scripts\mineru_handoff.py audit --workspace "job1" --require-accepted
~~~

See [SKILL.md](SKILL.md), [MinerU-first instructions](docs/MINERU_FIRST.md), [installation](docs/INSTALL.md), [layout JSON fallback](docs/WORKFLOW.md), and [style sample use](docs/REFERENCE.md).

Microsoft Word and Windows are required for real Word preview rendering. MinerU is a separate project, not bundled here. No scanned documents, customer files, model weights, or licensed samples are included.

---

## 中文

适用于**数学扫描PDF转可编辑Word简排**。默认使用 MinerU 导出 DOCX，再由 Codex 在原生 Word 文件上校对公式、题号、表格与几何图形，按照单独提供的Word样例排版，最后逐页对照原PDF验收。新增Word段落版式扫描，检查公式异常缩进、段前段后距和答案括号空格，并将审核绑定最终DOCX哈希。MinerU不可用时才运行RapidOCR备用识别，避免重复转换。

MIT License.
