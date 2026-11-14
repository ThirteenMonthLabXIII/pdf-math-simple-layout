---
name: pdf-math-simple-layout
description: Convert mathematics PDF handouts into editable Word. Preformat DOCX with a controlled paragraph-style contract before Codex edits content; prefer MinerU and use RapidOCR only as fallback.
---

# Math PDF to Word: simple layout skill

## START HERE — new conversation entry

For every new PDF-to-Word task, use this repository as the single source of truth:

1. Read this complete SKILL.md.
2. Read the complete docs/EXERCISE_WORKFLOW.md. It contains the mandatory question, answer-space, image-size and pagination rules.
3. Choose only the matching route:
   - MinerU conversion: read docs/MINERU_FIRST.md.
   - When Codex is doing the conversion: also read docs/CODEX_DIRECT.md.
   - Read docs/MINERU_API.md whenever using the API (GPT and Codex both permitted).
   - RapidOCR fallback / layout.json: read docs/WORKFLOW.md.
4. If a teacher formatting sample is supplied, also read docs/REFERENCE.md. Read docs/INSTALL.md only for installation work.
5. Follow the existing scripts. Do not search unrelated work folders, revive old workflow copies, or create a second processing pipeline.

**Fixed production order:** preserve source → convert once with MinerU → set image sizes automatically → apply paragraph styles and question/answer-space plan → run the real Word pagination preflight → edit content → rerun the question planner only if question structure changed → perform the required source-bound acceptance once.

Both production routes apply the image-size rule automatically before question layout and pagination. Do not resize images later as a separate repair pass. See the image-size rule below and docs/EXERCISE_WORKFLOW.md.

## MUST READ: GPT and Codex have equal MinerU permissions

**Both GPT and Codex have full access to the local MinerU API and authenticated web route, including use of the existing local Token. The only distinction is the default order of use, not who is allowed to use the API. No additional Codex authorization is required.**

- **GPT default:** local MinerU Precision API; the logged-in web route is also allowed.
- **Codex default:** MinerU in the already-logged-in browser in the current Codex conversation's own sidebar, without an API Token.
- **Switch immediately** when the web UI estimates more than 120 seconds of queueing, has no progress for 60 seconds, hangs, lacks login, or the sidebar is unavailable. Codex may also call the local API directly at the outset when the task or available environment makes it preferable. These thresholds guide efficiency; they are not permission gates.
- **Single stable API entry:** scripts/mineru_entry.py delegates to the existing shared MinerU API implementation. Both GPT and Codex can use it directly; MINERU_API_SCRIPT may override the implementation path without changing the workflow. It never prints or duplicates the local Token.
- **One downstream pipeline:** complete MinerU DOCX → scripts/mineru_handoff.py prepare → image sizing → styles → answer space → question pagination. Neither actor gets a separate formatting pipeline.

When uncertain, follow the decision table at the **top of README.md**, then START_HERE.md. Read docs/CODEX_DIRECT.md for sidebar steps and docs/MINERU_API.md for API commands. Do not invent new routes or permission restrictions.

## First: preserve the source and account for every question

Read [docs/EXERCISE_WORKFLOW.md](docs/EXERCISE_WORKFLOW.md) before preparing or editing a practice handout. Unless the user explicitly requests a subset, preserve the source's complete content and order: explanations, examples, every numbered question and full stem, conditions, choices, subquestions, formulas, tables, answer blanks, diagrams, illustrations and captions. A contents page, example heading or generated knowledge summary does not replace the actual question. Keep summaries already present in the source; do not add a new assistant-written summary or answers unless requested.

Before editing, make a source-page inventory that records each page's sections, examples, question numbers and stem anchors, choices/subquestions, tables and figures, including cross-page items and genuinely blank pages. Map each source page to its corresponding final content. Use MinerU as a conversion aid, then check the output against the original PDF; OCR output alone does not establish completeness. Correct conversion omissions and recognition errors against the source. Preserve unclear or inconsistent source content and flag it in the report rather than guessing.

For this user's local workflow, when no destination is specified and it exists, use `C:\Users\Lenovo\WPSDrive\1754907989\WPS企业云盘\深圳大学\我的企业文档\University\兼职2`. Preserve the original file and name. Use the same base name for the DOCX and report, and `_Word版.pdf` for the Word-exported PDF so it cannot overwrite the source. Current user instructions take priority.

## Practice-page layout: plan complete questions before editing

These are hard requirements applied during prepare/build, before Codex edits the document. A final audit only verifies them; it is not the first time the rules are applied.

1. **One complete question stays on one page.** A question block includes its heading, number, complete stem, conditions, every subquestion, formulas, choices, tables, related figures and captions, plus all answer space. No part may spill onto the next page.
2. **Reserve answer space for every solvable item.** Each independently solvable subquestion gets its own writing area. For a big question, total answer space is the sum of its subquestion allowances, so more subquestions mean more room. If a question has no subquestions, allocate space based on the whole question's complexity.
3. **Answer space is made only with Enter.** Each writing line is a genuinely empty Word paragraph. Do not create it with Shift+Enter, manual line breaks, underscores, borders, tables, tabs, text boxes, shapes, or paragraph-before/after spacing. Empty paragraphs contain no text, tabs, breaks, drawings or formulas.
4. **Every page starts at the top of the text area.** The first content line and its question number/label start flush left: zero left indent, zero first-line indent, and no leading spaces or tabs. A heading that belongs to the question travels with it.
5. **If the last question cannot fit with its full answer area, move the whole question block to the next page.** Never leave its stem on the previous page and its subquestions or answer area on the next.
6. **A question that is taller than a full page is still not allowed to split.** Its figures already receive their final display dimensions during the first build, before pagination. Do not start a later image-resize repair cycle. Keep at least one real empty paragraph for every independent subquestion. If it still cannot fit, stop and report that question as blocked; do not deliver a split question.

Default answer-space budget (one count = one empty Enter paragraph):
- Simple choice/fill-in/short response: 2.
- Standard question: 3.
- Calculation or equation-solving: 4.
- Long application question, table/diagram reasoning: 5.
- Proof, geometry proof or multi-step derivation: 6.
- Each subquestion is estimated independently; the parent question's total is the sum of its subquestion budgets. These defaults are configurable in references/style_profile.json.

The planner adds an invisible keep-together / keep-with-next chain to each complete question block so Word moves the question and its answer paragraphs together. Those properties control pagination only; they do not create writing space. Never use manual page breaks or pageBreakBefore to fake the layout.

Put a question's figure after its complete stem and subquestions, adjacent to that question, with its original caption centered. Keep lecture illustrations with their explanatory text. Preserve all tables and images; do not replace a figure with a description. OfficeMath is the default editable formula format. Do not use an AxMath workflow unless the current user explicitly requests AxMath and identifies the applicable workflow.

### Image size is decided during the first build

Both supported routes call scripts/image_sizing.py before question grouping, answer-space planning and Word pagination. This is a production step, not a final audit or repair stage.

- Within each lecture, existing images with a displayed width above 0 and up to 6.0 cm form the reference set. Their median width is the lecture's reference size.
- Images wider than 6.0 cm in that lecture are reduced to that median width. Images at or below 6.0 cm keep their current size; they are never enlarged or normalized.
- If a lecture has no reference image, use the median of eligible small images across the document. If the document has no eligible reference image, use the 4.05 cm fallback.
- Preserve each image's aspect ratio. Only Word display dimensions change; embedded image bytes and source resolution remain untouched.
- MinerU prepare and RapidOCR build use the same policy. Defaults live in references/style_profile.json and can be overridden through image_size_policy in a supplied profile.
- If a figure is added or replaced after preparation, apply the same sizing function once to the updated working DOCX before rerunning question layout. Text-only edits do not trigger image resizing.
- Do not add a separate image-size audit, blanket resize command, or repeated resize/check loop after pagination. The existing final page review checks source-content correctness and actual pagination; it is not where image-sizing rules are first discovered.

## Core rule: prevent layout defects before editing

The working DOCX must be preformatted before Codex starts changing content. Do not create a bad layout and rely on the final audit to repair it.

The pipeline has three responsibilities:
1. **Preventive question planning:** identify question blocks, allocate answer paragraphs, and bind each full question to one page before content editing.
2. **Preventive style formatting:** create safe paragraph styles and make the intended hierarchy and flush-left page starts explicit.
3. **Regression gates:** verify the answer-space counts and real Word pagination after edits. These checks confirm the rules were maintained; they do not replace the pre-edit planning step.

## Default: MinerU-first

Read docs/MINERU_FIRST.md. Export the complete source PDF to editable DOCX using MinerU. Keep the original PDF, raw MinerU DOCX, and teacher style sample immutable. Set the Python executable once before any optional sample-profile step: `$py = 'D:\我的python\python3.12\python.exe'`.

If a teacher/reference Word sample exists, extract a controlled style profile first:

~~~powershell
$py = 'D:\我的python\python3.12\python.exe'
$profile = "references\style_profile.json"
if (Test-Path "teacher-sample.docx") {
    $profile = Join-Path $env:TEMP "pdf-math-style-profile.json"
    & $py scripts\layout_cli.py profile --sample "teacher-sample.docx" --output $profile
}
~~~

Then prepare the handout using that profile:

~~~powershell
& $py scripts\mineru_handoff.py prepare --pdf "input.pdf" --mineru-docx "mineru.docx" --workspace "job1" --profile $profile
~~~

When no teacher sample exists, keep $profile pointed at references/style_profile.json; this same selected profile must be passed to prepare and every later question-layout command.

The prepare command:
- preserves the raw MinerU export as mineru_original.docx;
- creates working.docx as a separate copy;
- applies the SIJI paragraph-style contract;
- runs scripts/question_layout.py through the shared planner before Codex opens the file;
- inserts real empty answer paragraphs after every question/subquestion, with more paragraphs for more complex work;
- applies the page-safe question chain and flush-left question-start style;
- records the question groups, required/existing/inserted answer paragraphs and hashes in question_layout_plan.json and style_policy.json;
- runs a structural answer-space check and a real Microsoft Word pagination preflight before it reports the file ready for editing.

Only preflight_status=ready_for_codex_formatting allows Codex to begin content editing. If prepare reports needs_question_mapping, needs_question_layout, needs_question_page_fit or needs_word_pagination_check, stop before editing. Map missed questions, adjust the oversized question block or restore the Word pagination environment, rerun prepare/planning, and require a clean preflight. Never silently skip a question or defer this to the final audit.

Do not bypass this step or start editing the raw MinerU export.

## Preventive paragraph-style contract

The preformatter creates these named Word paragraph styles:

| Style | Use |
|---|---|
| SIJI Body | Ordinary explanatory text |
| SIJI Title | Lecture/title lines |
| SIJI Heading | Section headings |
| SIJI Question Start | First paragraph/number of a complete question block; always flush left |
| SIJI Question | Other numbered question stems |
| SIJI Subquestion | Numbered subquestions such as （1） |
| SIJI Formula | Standalone OfficeMath formula paragraphs |
| SIJI AnswerBlank | Answer-only blank lines |
| SIJI Caption | Figure captions |
| SIJI Figure | Centered figure/image paragraphs |

Each style explicitly controls left/right/first-line indentation, before/after spacing, line spacing, and alignment. The source document's direct paragraph-layout overrides are cleared in the working copy before the controlled style is applied.

### Formula paragraphs

- Standalone OfficeMath paragraphs use SIJI Formula.
- Formula style has zero left, right, and first-line indentation by default.
- Never copy paragraph formatting from a MinerU formula paragraph into a new formula paragraph.
- Never set formula paragraph indentation manually. If the teacher sample has an abnormal formula indent over 1000 twips, the sample extractor rejects that paragraph rather than learning the bad value.
- Inline formulas remain inside their question/body paragraph; do not split them into a new paragraph unless the source layout requires it.

### Question hierarchy

- The first paragraph of every complete question block uses SIJI Question Start: zero left indent and zero first-line indent, so a question number at the top of a page is flush left.
- Other numbered question stems use SIJI Question; child items use SIJI Subquestion.
- Subquestions may retain their normal hierarchy indent inside a question. If a subquestion is itself the first item of a new question block, it is promoted to SIJI Question Start and flush left.
- The default subquestion indent is 600 twips. A provided teacher sample can override normal subquestion styling, but it cannot override the flush-left Question Start style.
- A sample paragraph with any indentation over 1000 twips is excluded from style learning. This prevents an anomalous source formula from contaminating the whole style profile.
- Do not flatten all paragraph indents to zero. Use the prepared hierarchy styles.

### Answer blanks and source text

- Answer parentheses and their spaces already present in the source are content; preserve them exactly.
- New working space is separate from source answer blanks. The planner creates it as genuinely empty Word paragraphs (w:p) with the SIJI AnswerBlank style. One empty paragraph represents one normal writing line.
- SIJI AnswerBlank has zero before/after spacing and normal body line spacing. It must not use borders, underlines, tabs, underscores, tables, shapes or manual line breaks.
- Never run global whitespace normalization, repeated-space collapsing, global strip(), or find/replace on source answer blanks.
- The style preformatter does not rewrite existing paragraph text or OfficeMath XML. The question planner only inserts new empty paragraphs and pagination metadata.
- When editing a paragraph containing OfficeMath, do not replace the entire paragraph text with paragraph.text = ...; that can discard embedded equation nodes. Edit only the relevant text runs and preserve the math XML.

### Codex editing contract

After prepare, Codex must reuse the named styles for every inserted paragraph. When using python-docx, assign the matching style and leave paragraph_format indentation/spacing properties alone. When using Word, select the matching named style from the style list. Do not hand-build answer-space formatting or manually add page breaks. If content editing adds/removes questions or subquestions, rerun the question planner so the new structure receives its answer paragraphs and one-page chain.

Do not:
- set paragraph_format.left_indent, right_indent, first_line_indent, space_before, space_after, or line_spacing ad hoc;
- paste paragraph formatting from mineru_original.docx;
- use Format Painter from a suspicious formula paragraph;
- globally normalize spaces or paragraph text;
- create new formula paragraphs with Normal style and manual indentation;
- create answer space using Shift+Enter, underscores, borders, tables, tabs, or paragraph spacing;
- split a question across pages or add manual page breaks to force a visual result.

After all content edits, run the question planner before rendering:

~~~powershell
& $py scripts\question_layout.py apply --docx "job1\working.docx" --profile $profile
& $py scripts\question_layout.py check --docx "job1\working.docx" --profile $profile
~~~

This is the planned layout pass for any new/changed question content, not a repair-after-failure step. Correct OCR/formula content by comparing the PDF original. Retain numbering, subquestions, data, tables, captions, and geometry diagrams in the original order. When geometry crops are needed, take them from the source PDF at high resolution and keep them adjacent to their corresponding question. Never infer unclear mathematics from context.

A formatting reference DOCX is for styles only, not content. Do not copy its question text into the handout.

## RapidOCR fallback

Only if MinerU is unavailable or its DOCX export fails. Do not repeat OCR after a valid MinerU output.

~~~powershell
& $py scripts\layout_cli.py prepare --pdf "input.pdf" --output "fallback_job"
& $py scripts\layout_cli.py recognize --workspace "fallback_job"
# Correct blocks and high-resolution figure bbox regions in fallback_job/layout.json.
& $py scripts\layout_cli.py build --workspace "fallback_job" --profile "references/style_profile.json"
~~~

The build command applies the same SIJI paragraph-style contract, the same question-layout planner, and a real Word pagination preflight before review. Its build_report.json records the preflight status and any split question groups. Source PDF page boundaries are inventory markers, not Word page breaks; the build does not insert a page break at every source page. Do not proceed to content review until the preflight reports ready_for_content_review. If a teacher sample is available, create a profile with layout_cli.py profile and pass it to build.

## Last-resort layout regression gate

The preventive style contract is the default. Keep the scanner as a final safety net on both MinerU and RapidOCR routes.

~~~powershell
& $py scripts\question_layout.py check --docx "job1\working.docx" --profile $profile --word-pages
& $py scripts\docx_layout_check.py check --docx "job1\working.docx" --output "job1\docx_layout_findings.json"
& $py scripts\docx_layout_check.py template --docx "job1\working.docx" --output "job1\layout_review.json"
~~~

It flags effective direct or inherited-style left/right/first-line/hanging indents above 1000 twips;
- direct or inherited-style paragraph before/after spacing above 720 twips;
- three or more consecutive spaces inside answer parentheses.

Compare findings with the source PDF. Do not automatically normalize them. Gray dots, return arrows and tab arrows in Word are nonprinting formatting marks, not document characters; hide formatting marks or inspect the exported PDF before deciding whether a real character exists.

## Render and source-bound acceptance

Render the result with genuine Microsoft Word (Windows COM):

~~~powershell
& $py scripts\layout_cli.py render --docx "job1\working.docx" --output "job1\word_preview"
~~~

Compare all source PDF pages with all actual output Word pages. Record the review accurately in job1/mineru_review.json, including the last candidate SHA-256 and all page checks. Never tick review fields automatically.

~~~powershell
& $py scripts\mineru_handoff.py audit --workspace "job1" --require-accepted --word-pages
~~~

Only deliver when audit status is accepted and the actual page/figure/math review has been performed. After editing the Word file again, the old review hash is invalid.

## Software

- scripts/mineru_handoff.py: preserve the raw MinerU DOCX and preformat the working copy before editing.
- scripts/mineru_entry.py: common GPT/Codex API CLI entry, delegating to the existing MinerU implementation. The Token remains in the ignored local secrets file.
- scripts/docx_styling.py: create/apply the preventive SIJI paragraph-style contract and extract safe styles from a teacher sample.
- scripts/question_layout.py: allocate Enter-created answer space, keep every complete question on one page, and verify real Word pagination.
- scripts/layout_cli.py: profile extraction, RapidOCR fallback, build, Word render and audit.
- scripts/review_coverage.py: require exactly one completed review record for each source page.
- scripts/docx_layout_check.py: last-resort DOCX paragraph-layout regression scanner.
- tests: deterministic local checks without Word or OCR.

Run from the Codex local terminal. All programs should operate on copies in a separate workspace and preserve existing original files. No remote computer connector is required.
