# PDF Math Simple Layout

A Codex-friendly skill to turn scanned mathematics handouts into editable Microsoft Word documents.

**Preferred:** MinerU export → preventive Word styles + question/answer-space plan → content correction → real Word pagination check → per-page review.
**Fallback:** local RapidOCR → editable DOCX → the same question/answer-space planner and style contract.

Features:
- Preserve source PDFs and raw MinerU DOCX exports as immutable originals.
- Preformat the MinerU working copy before Codex edits it; do not inherit unsafe source paragraph indents.
- Keep every complete question, its subquestions, formulas and attached figures on one page; move the whole question when the remaining page space is insufficient.
- Allocate answer space for every independently solvable question/subquestion using real empty Enter paragraphs; larger questions receive more space.
- Start each page's first question number flush-left at the top of the text area; never fake answer space with borders, tabs, manual line breaks or paragraph spacing.
- Apply named paragraph styles for body text, questions, subquestions, standalone formulas, answer blanks, headings, titles and captions.
- Keep formula paragraph indents safe by default; learn question hierarchy from a teacher sample while rejecting extreme sample indents.
- Preserve answer-parenthesis text and exact spaces; never globally normalize source whitespace.
- Apply the same style contract and real Word pagination preflight automatically to RapidOCR-built DOCX candidates.
- Use scripts/question_layout.py to allocate per-subquestion writing space, enforce one-page question groups, and reject non-Enter blank-line tricks.
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
$profile = "references\style_profile.json"

# If a teacher sample exists, learn its safe styles to a temporary profile:
if (Test-Path "teacher-sample.docx") {
    $profile = Join-Path $env:TEMP "pdf-math-style-profile.json"
    & $py scripts\layout_cli.py profile --sample "teacher-sample.docx" --output $profile
}

& $py scripts\mineru_handoff.py prepare --pdf "input.pdf" --mineru-docx "mineru.docx" --workspace "job1" --profile $profile

# Only edit when prepare reports ready_for_codex_formatting.
# After any question/subquestion content change, re-run the planner:
& $py scripts\question_layout.py apply --docx "job1\working.docx" --profile $profile
& $py scripts\question_layout.py check --docx "job1\working.docx" --profile $profile --word-pages

# Render and run the regression gate/source-bound acceptance.
& $py scripts\layout_cli.py render --docx "job1\working.docx" --output "job1\word_preview"
& $py scripts\docx_layout_check.py check --docx "job1\working.docx" --output "job1\docx_layout_findings.json"
& $py scripts\docx_layout_check.py template --docx "job1\working.docx" --output "job1\layout_review.json"
& $py scripts\mineru_handoff.py audit --workspace "job1" --require-accepted --word-pages
~~~

See [SKILL.md](SKILL.md), [MinerU-first instructions](docs/MINERU_FIRST.md), [installation](docs/INSTALL.md), [layout JSON fallback](docs/WORKFLOW.md), and [style sample use](docs/REFERENCE.md).

Microsoft Word and Windows are required for real Word preview rendering. MinerU is a separate project, not bundled here. No scanned documents, customer files, model weights, or licensed samples are included.

---

## 中文

适用于**数学扫描 PDF 转可编辑 Word 简排**。默认使用 MinerU 导出 DOCX，prepare 阶段先给工作副本套用受控段落样式，并在 Codex 编辑前识别完整题组、按小题数量和复杂度插入 Enter 空段落、建立整题同页分页链。每页题号顶格开始；页尾空间不足时整题移到下一页。RapidOCR build 使用同一规划器。最终检查只负责确认规则被保持，不承担事后补救。

MIT License.
