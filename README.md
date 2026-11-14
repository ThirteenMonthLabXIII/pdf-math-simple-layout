# PDF Math Simple Layout

## 必读：GPT 与 Codex 同权限，只有 MinerU 默认优先级不同

**这是一条执行顺序规则，不是权限规则。GPT 和 Codex 都能完整使用已登录网页、本机 Token、MinerU API、转换脚本及后续排版工具；Codex 使用 API 不需要重新取得用户授权。** 以当前任务明确要求与实际效率为准。

| 执行者 | 默认先用 | 切换与直接选择 |
|---|---|---|
| **GPT** | **本机 MinerU API（Token 自动从本机文件读取）** | 可以直接使用已登录的 MinerU 网页 |
| **Codex** | **当前 Codex 对话侧边栏中已登录的 MinerU 网页（不需要 Token）** | 网页排队久、60 秒无进展、前端卡住、会话未登录或侧栏不可用时立即改用本机 API；需要时也可直接用 API |

**切换参考：** 网页预计排队超过 **120 秒**、连续 **60 秒无进展** 或界面卡住，**不要继续等网页**，调用本机 API。这些数值是效率参考，不是允许使用 API 的前置许可。即使未达到阈值，也可按实际任务选择更快的通道。

**统一 API 入口：** `scripts\mineru_entry.py`。它转发到原有 MinerU API 实现，不复制第二套代码；默认后端在 AxMath Skill，也可用 `MINERU_API_SCRIPT` 指定新位置。Token 由原 API 从本机读取，不放入文档、命令行或 Git。

**共同出口：** 无论谁使用网页还是 API，取得完整的 MinerU DOCX 后都进入 `scripts\mineru_handoff.py prepare`，首次制作即设置图片尺寸、题目布局、答题空间、Word 分页。

**新对话阅读顺序：** [START_HERE.md](START_HERE.md) → [SKILL.md](SKILL.md) → [练习工作流](docs/EXERCISE_WORKFLOW.md) → 按需读取 [Codex 网页入口](docs/CODEX_DIRECT.md) 或 [本机 API 命令](docs/MINERU_API.md)。不要跳过入口自行猜路由。

A Codex-friendly skill to turn scanned mathematics handouts into editable Microsoft Word documents.

> **New conversation? Start with [START_HERE.md](START_HERE.md).**
> Read the route-specific instructions there before touching a file. Do not hunt through old work folders or recreate the workflow.

**Preferred:** MinerU export → preventive Word styles + question/answer-space plan → content correction → real Word pagination check → per-page review.
**MinerU access:** GPT and Codex have equal permission; choose the default order shown above, and switch freely as required by the task.
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
- Set figure display sizes automatically during the first MinerU prepare / RapidOCR build, before question layout and pagination. The shared rule uses per-lecture small-image medians, a document-wide fallback and a 4.05 cm final fallback while preserving aspect ratios.
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

First export the **entire** source PDF from MinerU as an editable DOCX. Both actors may choose web or API; the default order and fast switching rules are at the top of this README. Both routes feed the same prepare command below.

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

See [SKILL.md](SKILL.md), [MinerU-first instructions](docs/MINERU_FIRST.md), [Codex browser/API routing](docs/CODEX_DIRECT.md), [local MinerU API fallback](docs/MINERU_API.md), [installation](docs/INSTALL.md), [layout JSON fallback](docs/WORKFLOW.md), and [style sample use](docs/REFERENCE.md).

Microsoft Word and Windows are required for real Word preview rendering. MinerU is a separate project, not bundled here. No scanned documents, customer files, model weights, or licensed samples are included.

---

## 中文

适用于**数学扫描 PDF 转可编辑 Word 简排**。新对话统一从根目录 START_HERE.md 进入，再按任务路线读取对应工作流。默认使用 MinerU 导出 DOCX，prepare 阶段先自动确定图片显示尺寸，再套用受控段落样式、识别完整题组、按小题数量和复杂度插入 Enter 空段落、建立整题同页分页链。每页题号顶格开始；页尾空间不足时整题移到下一页。RapidOCR build 使用同一套图片尺寸和题目规划规则。图片尺寸不留到终稿阶段返工。

MIT License.
