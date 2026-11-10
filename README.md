# PDF Math Simple Layout

A Codex-friendly skill to turn scanned mathematics handouts into editable Microsoft Word documents.

**Preferred:** MinerU export → preventive Word style contract → content correction → real Word render → per-page review.
**Fallback:** local RapidOCR → editable DOCX → the same preventive style contract.

Features:
- Preserve source PDFs and raw MinerU DOCX exports as immutable originals.
- Preformat the MinerU working copy before Codex edits it; do not inherit unsafe source paragraph indents.
- Apply named paragraph styles for body text, questions, subquestions, standalone formulas, answer blanks, headings, titles and captions.
- Keep formula paragraph indents safe by default; learn question hierarchy from a teacher sample while rejecting extreme sample indents.
- Preserve answer-parenthesis text and exact spaces; never globally normalize source whitespace.
- Apply the same style contract automatically to RapidOCR-built DOCX candidates.
- Keep mathematical text editable; preserve geometry diagrams as high-resolution crops.
- Preserve every source question and figure, and plan writing space from the start; follow [the exercise workflow](docs/EXERCISE_WORKFLOW.md).
- Refuse to mark a document accepted until source pages, rendered Word pages and final hashes are checked.
- Retain a final DOCX layout scanner as a last-resort regression gate, not as the primary correction method.

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

# If a teacher sample exists, learn its safe paragraph styles first:
& $py scripts\layout_cli.py profile --sample "teacher-sample.docx" --output "style_profile.json"
& $py scripts\mineru_handoff.py prepare --pdf "input.pdf" --mineru-docx "mineru.docx" --workspace "job1" --profile "style_profile.json"

# If there is no teacher sample, use built-in safe defaults:
# & $py scripts\mineru_handoff.py prepare --pdf "input.pdf" --mineru-docx "mineru.docx" --workspace "job1"

# Correct content using the prepared SIJI paragraph styles.
& $py scripts\layout_cli.py render --docx "job1\working.docx" --output "job1\word_preview"

# Last-resort regression gate and source-bound acceptance.
& $py scripts\docx_layout_check.py check --docx "job1\working.docx" --output "job1\docx_layout_findings.json"
& $py scripts\docx_layout_check.py template --docx "job1\working.docx" --output "job1\layout_review.json"
& $py scripts\mineru_handoff.py audit --workspace "job1" --require-accepted
~~~

See [SKILL.md](SKILL.md), [MinerU-first instructions](docs/MINERU_FIRST.md), [installation](docs/INSTALL.md), [layout JSON fallback](docs/WORKFLOW.md), and [style sample use](docs/REFERENCE.md).

Microsoft Word and Windows are required for real Word preview rendering. MinerU is a separate project, not bundled here. No scanned documents, customer files, model weights, or licensed samples are included.

---

## 中文

适用于**数学扫描 PDF 转可编辑 Word 简排**。默认使用 MinerU 导出 DOCX，prepare 阶段先给工作副本套用受控段落样式，再由 Codex 校对内容。公式段落预设零缩进，题目层级从老师样例提取，异常样例缩进不会被继承；答案括号和空格原样保留。RapidOCR 构建也自动套用同一套样式。最终扫描只负责兜底，不承担主要纠错工作。

MIT License.
