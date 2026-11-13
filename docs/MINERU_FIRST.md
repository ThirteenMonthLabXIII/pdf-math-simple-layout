# MinerU-first PDF-to-Word handoff

Set the Python executable once:

~~~powershell
$py = 'D:\我的python\python3.12\python.exe'
~~~

1. Read [EXERCISE_WORKFLOW.md](EXERCISE_WORKFLOW.md). Before drafting, map every source page, example, full question, subquestion and its figures. Export the entire source PDF using MinerU; use the conversation's in-app browser when available, unless the user specifies another browser. Treat the export as an input to be checked, not as proof that the questions or figures are complete.

2. Select the paragraph-style profile BEFORE preparing the handout. Keep the bundled defaults unless a teacher sample is provided:

~~~powershell
$profile = "references\style_profile.json"
if (Test-Path "teacher-sample.docx") {
    $profile = Join-Path $env:TEMP "pdf-math-style-profile.json"
    & $py scripts\layout_cli.py profile --sample "teacher-sample.docx" --output $profile
}
~~~

The extractor learns margins, fonts, line spacing, question hierarchy, and paragraph alignment. Sample indents above 1000 twips are rejected instead of copied into the new document. The teacher sample cannot override the flush-left question-start style or the Enter-only answer-space rule. If no sample is provided, use references/style_profile.json.

3. Prepare the MinerU output. This automatically preserves a raw copy, applies the image-size policy to the working copy, applies the SIJI styles, and plans answer space and question pagination BEFORE Codex edits the document. Image sizing runs before the question planner and Word pagination preflight; its result is included in question_layout_plan.json and the prepare output.

With a teacher sample:

~~~powershell
& $py scripts\mineru_handoff.py prepare --pdf "input.pdf" --mineru-docx "mineru.docx" --workspace "job1" --profile $profile
~~~

Without a teacher sample:

~~~powershell
& $py scripts\mineru_handoff.py prepare --pdf "input.pdf" --mineru-docx "mineru.docx" --workspace "job1" --profile $profile
~~~

The untouched export remains mineru_original.docx. Codex edits working.docx. The prepare step:
- applies the shared image-size policy: per-lecture median of existing images up to 6.0 cm; images wider than 6.0 cm shrink to that median; document-wide median fallback; 4.05 cm final fallback; aspect ratio and image bytes preserved;
- detects each full question group, including its heading, number, stem, subquestions, formulas, options, figures, tables and captions;
- adds answer space using genuinely empty Word paragraphs (one Enter paragraph per writing line);
- gives each independent subquestion its own answer budget; big-question space grows with the number and complexity of its subquestions;
- uses the SIJI Question Start style for a flush-left question label;
- applies keep-together/keep-with-next only to bind the whole question and its answer paragraphs to one page;
- removes source manual page breaks/pageBreakBefore so the new Word pagination can flow naturally;
- runs a structural check and real Microsoft Word pagination preflight before it allows content editing;
- writes question_layout_plan.json, including the preflight result, and the applied rules/hashes to style_policy.json.

Only continue when prepare reports ready_for_codex_formatting. If it reports needs_question_mapping, needs_question_layout, needs_question_page_fit or needs_word_pagination_check, stop before content editing. Manually classify missed question stems/subquestions, tighten only redundant text spacing inside an oversized question, or restore Word COM, then rerun prepare and require a clean preflight. Its image dimensions have already been decided by the first-build policy; do not start a post-pagination image-resize pass.

4. Only after preformatting and question planning, edit content in working.docx. Reuse the prepared styles:
   - SIJI Body
   - SIJI Question Start
   - SIJI Question
   - SIJI Subquestion
   - SIJI Formula
   - SIJI AnswerBlank
   - SIJI Heading
   - SIJI Title
   - SIJI Caption
   - SIJI Figure

Do not manually assign paragraph indents or spacing. Do not copy paragraph formatting from the raw MinerU DOCX. Do not globally trim, collapse, or replace spaces. Preserve answer parentheses and their exact source spacing; correct text only against the PDF. Inline formulas stay in their question/body paragraph; standalone OfficeMath paragraphs use SIJI Formula.

When content editing adds, removes or restructures questions/subquestions, rerun the question planner before rendering. This is a normal layout-generation step for the changed content, not a repair-after-failure step. Text-only edits do not rerun image sizing; if a figure is added or replaced, apply scripts/image_sizing.py once to the updated working DOCX before rerunning the question planner:

~~~powershell
& $py scripts\question_layout.py apply --docx "job1\working.docx" --profile $profile
& $py scripts\question_layout.py check --docx "job1\working.docx" --profile $profile
~~~

Do not create answer space manually with Shift+Enter, underscores, borders, tabs, tables, shapes or paragraph spacing. Do not insert manual page breaks. If the last question cannot fit with all its answer space, the whole question group moves to the next page. If a question is taller than a page, compact inside that group; if it still cannot fit, stop and report it as blocked rather than splitting it.

5. Render and inspect the output using Microsoft Word:

~~~powershell
& $py scripts\layout_cli.py render --docx "job1\working.docx" --output "job1\word_preview"
~~~

6. Check actual Word pagination and question answer space. Every group must fit on one page; page-opening question labels must be flush left:

~~~powershell
& $py scripts\question_layout.py check --docx "job1\working.docx" --profile $profile --word-pages
~~~

A split group, insufficient answer space, non-Enter blank, missing pagination chain or non-flush question start blocks delivery.

7. Run the paragraph-layout regression gate:

~~~powershell
& $py scripts\docx_layout_check.py check --docx "job1\working.docx" --output "job1\docx_layout_findings.json"
& $py scripts\docx_layout_check.py template --docx "job1\working.docx" --output "job1\layout_review.json"
~~~

The scanner checks direct and inherited style indents, paragraph spacing, and answer-parenthesis spacing. Inspect every finding against the source PDF. A 600-twip subquestion indent may be intentional. Do not zero all indents or bulk-delete answer paragraphs. Gray dots and return arrows in Word are usually nonprinting formatting marks; verify with marks hidden or the exported PDF before editing them. If a finding is intentional, record the reason and source_pdf_checked=true in layout_review.json. The review file is bound to the final DOCX hash.

8. Compare every original PDF page with the corresponding content in the rendered Word pages. Update mineru_review.json only after checking content, formulas, figures and page starts. Set word_preview_checked=true only after reviewing all output pages, and set candidate_sha256 to the actual final DOCX hash.

9. Audit:

~~~powershell
& $py scripts\mineru_handoff.py audit --workspace "job1" --docx "job1\working.docx" --require-accepted --word-pages
~~~

Never deliver an unreviewed DOCX. OCR confidence is not proof of formula correctness; review mathematical symbols and diagrams against the source images. The real Word pagination check must report zero split question groups.

MinerU is developed independently and is not included in this repository. If MinerU is unavailable, see the RapidOCR fallback instructions in SKILL.md. If a task requires genuine AxMath objects, use a separately maintained AxMath conversion workflow.
