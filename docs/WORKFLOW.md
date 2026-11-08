# layout.json 编辑说明

`prepare` 自动产生 `layout.json`，是所有任务的单一结构真源；`source_sha256` 与 PDF 来源绑定。识别后每页 `blocks` 是数组，按原稿内容顺序排列，不要通过全局重排题号来掩盖缺失。

## 块格式

```json
{
  "page": 2,
  "verified": false,
  "blocks": [
    {"type": "heading", "text": "【例题 1】平行与相似"},
    {"type": "question", "text": "1．如图，在△ABC中，……"},
    {"type": "image", "bbox": [60, 125, 290, 230], "caption": ""},
    {"type": "subquestion", "text": "（1）求证：……"},
    {"type": "table", "rows": [["项目", "数据"], ["甲", "3"]]}
  ]
}
```

- 文本块：`title`、`heading`、`question`、`subquestion`、`paragraph`、`math`。文本含真实原文，Unicode可编辑。
- `image`：`bbox=[左,上,右,下]`，PDF点为单位；`dpi` 默认260，可在块中设300+。从源PDF图像直接采样；图中文字跟图片一起保留。
- `table`：Word原生表格，行列可以不等长但必须非空。
- `page_break`：仅原文有明确分节或老师要求时使用，不因原PDF翻页就强加分页符。
- `verified`：只在**逐页对照源稿并逐题核实**后标真。

## 容易犯错的地方

1. OCR会把 `∠` 识别成 `<`，把 `△` 识别成 `A`，把 `∥` 识别成数字11，把 `′`/`−`/`~` 混淆。每页核对数学符号和图内小字。
2. 题号“1、/1．”、分项“（1）”、括号和分式位置按老师样例和原稿保留；不能因为系统识别文本段落错误就合并两道题。
3. 原PDF的页脚页码如“71、72、73…”可能是扫描版教材页码，可不保留到Word正文，但必须确认不是题目编号；不要删除章节标题或图注。
4. 一道题若题干文字环绕图形，裁剪框不能混入要转换的正文；图形必须在对应题文附近。
5. OCR低置信度不是唯一危险。高置信度也可能错；公式的等号、分数、特殊符号必须对照图片。
6. Word页码和PDF页码可能不同；逐页审查按题号/短语/图片锚点追踪，不能靠同页码判错或判对。

## 验收字段

`visual_review.json` 包含 `candidate_sha256` 和每一页四个布尔字段。只有实际核对后才填真：
- `source_checked`：检查原始扫描页和所有内容。
- `output_checked`：实际看过Word候选渲染，并能定位全部对应题目。
- `text_math_checked`：文字、数字、公式、选项逐项核对。
- `figures_checked`：图形及图注对应、清晰、无截断。

Word内容有修改后必须清除旧的批准记录/更新哈希重新看图，不得复制旧记录冒充验收。


## Preventive Word style contract (applies before content editing)

The DOCX must be preformatted BEFORE Codex begins editing. Do not rely on a final scan to repair a bad layout.

- MinerU route: mineru_handoff.py prepare copies the raw export to mineru_original.docx, then applies docx_styling.apply_docx_styles to working.docx using references/style_profile.json or a profile extracted from the teacher sample.
- RapidOCR route: layout_cli.py build applies the same style contract to the generated candidate before it is opened for review.
- The contract creates SIJI Body, SIJI Title, SIJI Heading, SIJI Question, SIJI Subquestion, SIJI Formula, SIJI AnswerBlank, SIJI Caption, and SIJI Figure paragraph styles. Each style explicitly owns left/right/first-line indent, paragraph spacing, line spacing, and alignment.
- Formula paragraphs start with zero left/right/first-line indent. The source DOCX's direct paragraph indents are cleared in the working copy before editing; a 1687-twip source formula indent therefore never becomes the working formula style.
- Subquestion indentation is a controlled hierarchy value (600 twips by default). If a teacher sample is provided, the profile extractor learns its median per-kind paragraph settings. Any sample paragraph carrying an indent above 1000 twips is excluded from learning, so an anomalous formula does not contaminate the style profile.
- Codex must reuse the named styles for every inserted paragraph. Never set paragraph_format.left_indent, right_indent, first_line_indent, space_before, space_after, or line_spacing ad hoc. Never paste paragraph formatting from the raw MinerU export.
- Answer blank text is content, not layout decoration. Preserve the exact parentheses and spaces from the source PDF. Do not run global whitespace normalization, strip() on source text, or replace repeated spaces. The style pass changes paragraph/run formatting only and leaves text and OfficeMath XML unchanged.
- Word gray dots, return arrows and tab arrows are nonprinting formatting marks. They are not characters in the document text; do not delete or insert text to address their on-screen display.

The final layout scanner remains a last-resort regression gate. It should find no introduced anomaly because the working document was built from the style contract first.

## Word paragraph-layout regression check (MinerU and RapidOCR)

After the final DOCX save and before audit, run:

~~~powershell
python scripts/docx_layout_check.py check --docx "job1/working.docx" --output "job1/docx_layout_findings.json"
python scripts/docx_layout_check.py template --docx "job1/working.docx" --output "job1/layout_review.json"
~~~

The scanner flags direct paragraph properties and inherited paragraph-style properties:
- w:ind left/right/first-line/hanging values above 1000 twips (50 pt); formula paragraphs containing m:oMath / m:oMathPara get a dedicated finding. The regression case that motivated this check had a 1687-twip left indent that pushed a formula rightward.
- w:spacing before/after values above 720 twips (36 pt).
- Three or more consecutive spaces inside answer parentheses.

Compare every finding with the source PDF and neighbouring questions. A 600-twip nested-question indent may be correct; do not normalize every paragraph or delete blank paragraphs in bulk. Gray dots, return arrows and tab arrows in Word are normally nonprinting formatting marks; hide them or inspect the exported PDF before treating them as real characters.

After fixing issues, save and rerun both commands. layout_review.json must match the current DOCX SHA-256, have checked=true and a reviewer. Any remaining intentional finding must include disposition=intentional, a reason, and source_pdf_checked=true. Both audit routes require this review; any later DOCX edit invalidates the review hash.
