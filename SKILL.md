---
name: pdf-math-simple-layout
description: Convert mathematics PDF handouts into editable Word. Preformat DOCX with a controlled paragraph-style contract before Codex edits content; prefer MinerU and use RapidOCR only as fallback.
---

# Math PDF to Word: simple layout skill

## First: preserve the source and account for every question

Read [docs/EXERCISE_WORKFLOW.md](docs/EXERCISE_WORKFLOW.md) before preparing or editing a practice handout. Unless the user explicitly requests a subset, preserve the source's complete content and order: explanations, examples, every numbered question and full stem, conditions, choices, subquestions, formulas, tables, answer blanks, diagrams, illustrations and captions. A contents page, example heading or generated knowledge summary does not replace the actual question. Keep summaries already present in the source; do not add a new assistant-written summary or answers unless requested.

Before editing, make a source-page inventory that records each page's sections, examples, question numbers and stem anchors, choices/subquestions, tables and figures, including cross-page items and genuinely blank pages. Map each source page to its corresponding final content. Use MinerU as a conversion aid, then check the output against the original PDF; OCR output alone does not establish completeness. Correct conversion omissions and recognition errors against the source. Preserve unclear or inconsistent source content and flag it in the report rather than guessing.

For this user's local workflow, when no destination is specified and it exists, use `C:\Users\Lenovo\WPSDrive\1754907989\WPS企业云盘\深圳大学\我的企业文档\University\兼职2`. Preserve the original file and name. Use the same base name for the DOCX and report, and `_Word版.pdf` for the Word-exported PDF so it cannot overwrite the source. Current user instructions take priority.

## Practice-page layout: include real working space

- Put enough blank writing space after every question and independently solvable subquestion. Give comparable questions comparable space; allow more for long calculations, proofs and figure-based work. Do not compress questions to hit a target question count per page.
- Start the first question on each page at the top of the text area. When a section or chapter heading belongs with that first question, place the heading at the top and follow it with the question. Keep the question hierarchy and necessary subquestion indentation.
- If the final question on a page would not have enough room for its working space, move the question, its heading and any attached figure that fits with it to the next page. Keep a long question's parts and figures in source order.
- Use Enter-created empty paragraphs for working space. Do not add manual page breaks or new keep-with-next, keep-together, or page-break-before settings by default. Preserve the source document's effective direct and inherited pagination settings. Only change these settings when the current user explicitly asks.
- Put a question's figure after its complete stem and subquestions, adjacent to that question, with its original caption centered. Keep lecture illustrations with their explanatory text. Preserve all tables and images; do not replace a figure with a description.
- OfficeMath is the default editable formula format. Do not use an AxMath workflow unless the current user explicitly requests AxMath and identifies the applicable workflow.

## Core rule: prevent layout defects before editing

The working DOCX must be preformatted before Codex starts changing content. Do not create a bad layout and rely on the final audit to repair it.

The pipeline has two separate responsibilities:
1. **Preventive formatting:** create safe paragraph styles, clear unsafe source paragraph overrides, and make the intended hierarchy explicit before content editing.
2. **Last-resort regression gate:** detect accidental changes that occur later. This gate is not the primary formatting method.

## Default: MinerU-first

Read docs/MINERU_FIRST.md. Export the complete source PDF to editable DOCX using MinerU. Keep the original PDF, raw MinerU DOCX, and teacher style sample immutable. Set the Python executable once before any optional sample-profile step: `$py = 'D:\我的python\python3.12\python.exe'`.

If a teacher/reference Word sample exists, extract a controlled style profile first:

~~~powershell
$py = 'D:\我的python\python3.12\python.exe'
& $py scripts\layout_cli.py profile --sample "teacher-sample.docx" --output "style_profile.json"
~~~

Then prepare the handout using that profile:

~~~powershell
& $py scripts\mineru_handoff.py prepare --pdf "input.pdf" --mineru-docx "mineru.docx" --workspace "job1" --profile "style_profile.json"
~~~

When no teacher sample exists, omit --profile; the built-in references/style_profile.json supplies safe defaults.

The prepare command:
- preserves the raw MinerU export as mineru_original.docx;
- creates working.docx as a separate copy;
- applies the preventive SIJI paragraph-style contract to working.docx before Codex opens it;
- records the applied rules and hashes in style_policy.json.

Do not bypass this step or start editing the raw MinerU export.

## Preventive paragraph-style contract

The preformatter creates these named Word paragraph styles:

| Style | Use |
|---|---|
| SIJI Body | Ordinary explanatory text |
| SIJI Title | Lecture/title lines |
| SIJI Heading | Section headings |
| SIJI Question | Numbered question stems |
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

- Numbered questions use SIJI Question.
- Subquestions use SIJI Subquestion.
- The default subquestion indent is 600 twips. A provided teacher sample can override this through median per-kind style extraction.
- A sample paragraph with any indentation over 1000 twips is excluded from style learning. This prevents an anomalous source formula from contaminating the whole style profile.
- Do not flatten all paragraph indents to zero. Use the prepared hierarchy styles.

### Answer blanks and source text

- Answer parentheses and their spaces are source content, not decorative formatting.
- Preserve the exact parentheses and number of spaces shown in the source PDF.
- Never run global whitespace normalization, repeated-space collapsing, global strip(), or find/replace on answer blanks.
- The preformatter changes paragraph/run formatting only. It must not rewrite paragraph text or OfficeMath XML.
- When editing a paragraph containing OfficeMath, do not replace the entire paragraph text with paragraph.text = ...; that can discard embedded equation nodes. Edit only the relevant text runs and preserve the math XML.

### Codex editing contract

After prepare, Codex must reuse the named styles for every inserted paragraph. When using python-docx, assign the matching style and leave paragraph_format indentation/spacing properties alone. When using Word, select the matching named style from the style list.

Do not:
- set paragraph_format.left_indent, right_indent, first_line_indent, space_before, space_after, or line_spacing ad hoc;
- paste paragraph formatting from mineru_original.docx;
- use Format Painter from a suspicious formula paragraph;
- globally normalize spaces or paragraph text;
- create new formula paragraphs with Normal style and manual indentation.

Correct OCR/formula content by comparing the PDF original. Retain numbering, subquestions, data, tables, captions, and geometry diagrams in the original order. When geometry crops are needed, take them from the source PDF at high resolution and keep them adjacent to their corresponding question. Never infer unclear mathematics from context.

A formatting reference DOCX is for styles only, not content. Do not copy its question text into the handout.

## RapidOCR fallback

Only if MinerU is unavailable or its DOCX export fails. Do not repeat OCR after a valid MinerU output.

~~~powershell
& $py scripts\layout_cli.py prepare --pdf "input.pdf" --output "fallback_job"
& $py scripts\layout_cli.py recognize --workspace "fallback_job"
# Correct blocks and high-resolution figure bbox regions in fallback_job/layout.json.
& $py scripts\layout_cli.py build --workspace "fallback_job" --profile "references/style_profile.json"
~~~

The build command applies the same SIJI paragraph-style contract to the generated candidate before review. If a teacher sample is available, create a profile with layout_cli.py profile and pass it to build.

## Last-resort layout regression gate

The preventive style contract is the default. Keep the scanner as a final safety net on both MinerU and RapidOCR routes.

~~~powershell
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
& $py scripts\mineru_handoff.py audit --workspace "job1" --require-accepted
~~~

Only deliver when audit status is accepted and the actual page/figure/math review has been performed. After editing the Word file again, the old review hash is invalid.

## Software

- scripts/mineru_handoff.py: preserve the raw MinerU DOCX and preformat the working copy before editing.
- scripts/docx_styling.py: create/apply the preventive SIJI paragraph-style contract and extract safe styles from a teacher sample.
- scripts/layout_cli.py: profile extraction, RapidOCR fallback, build, Word render and audit.
- scripts/review_coverage.py: require exactly one completed review record for each source page.
- scripts/docx_layout_check.py: last-resort DOCX paragraph-layout regression scanner.
- tests: deterministic local checks without Word or OCR.

Run from the Codex local terminal. All programs should operate on copies in a separate workspace and preserve existing original files. No remote computer connector is required.
