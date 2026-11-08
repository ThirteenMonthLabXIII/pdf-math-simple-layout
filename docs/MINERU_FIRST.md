# MinerU-first PDF-to-Word handoff

Set the Python executable once:

~~~powershell
$py = 'D:\\我的python\\python3.12\\python.exe'
~~~

1. Export the entire source PDF using MinerU. Verify that every page and key problem appears in the editable DOCX.

2. If a teacher/reference Word sample is provided, extract its paragraph-style contract BEFORE preparing the handout:

~~~powershell
& $py scripts\\layout_cli.py profile --sample "teacher-sample.docx" --output "style_profile.json"
~~~

The extractor learns margins, fonts, line spacing, question hierarchy, and paragraph alignment. Sample indents above 1000 twips are rejected instead of copied into the new document. If no sample is provided, use the safe defaults in references/style_profile.json.

3. Prepare the MinerU output. This automatically preserves a raw copy and preformats the working copy before Codex opens it.

With a teacher sample:

~~~powershell
& $py scripts\\mineru_handoff.py prepare --pdf "input.pdf" --mineru-docx "mineru.docx" --workspace "job1" --profile "style_profile.json"
~~~

Without a teacher sample:

~~~powershell
& $py scripts\\mineru_handoff.py prepare --pdf "input.pdf" --mineru-docx "mineru.docx" --workspace "job1"
~~~

The untouched export remains mineru_original.docx. Codex edits working.docx, which already has the SIJI paragraph-style contract applied. The rules and hashes are recorded in style_policy.json.

4. Only after preformatting, edit content in working.docx. Reuse the prepared styles:
   - SIJI Body
   - SIJI Question
   - SIJI Subquestion
   - SIJI Formula
   - SIJI AnswerBlank
   - SIJI Heading
   - SIJI Title
   - SIJI Caption
   - SIJI Figure

Do not manually assign paragraph left/right/first-line indents or before/after spacing. Do not copy paragraph formatting from the raw MinerU DOCX. Do not globally trim, collapse, or replace spaces. Preserve answer parentheses and their exact source spacing; correct text only against the PDF. Inline formulas stay in their question/body paragraph; standalone OfficeMath paragraphs use SIJI Formula.

5. Render and inspect the output using Microsoft Word:

~~~powershell
& $py scripts\\layout_cli.py render --docx "job1\\working.docx" --output "job1\\word_preview"
~~~

6. Run the last-resort paragraph-layout regression gate. This is a safety net after the preventive style contract, not the main formatting method:

~~~powershell
& $py scripts\\docx_layout_check.py check --docx "job1\\working.docx" --output "job1\\docx_layout_findings.json"
& $py scripts\\docx_layout_check.py template --docx "job1\\working.docx" --output "job1\\layout_review.json"
~~~

The scanner checks direct and inherited style indents, paragraph spacing, and answer-parenthesis spacing. Inspect every finding against the source PDF. A 600-twip subquestion indent may be intentional. Do not zero all indents or bulk-delete blank paragraphs. Gray dots and return arrows in Word are usually nonprinting formatting marks; verify with marks hidden or the exported PDF before editing them. If a finding is intentional, record the reason and source_pdf_checked=true in layout_review.json. The review file is bound to the final DOCX hash.

7. Compare every original PDF page with the corresponding content in the rendered Word pages. Update mineru_review.json only after checking content, formulas, and figures. Set word_preview_checked=true only after reviewing all output pages, and set candidate_sha256 to the actual final DOCX hash.

8. Audit:

~~~powershell
& $py scripts\\mineru_handoff.py audit --workspace "job1" --require-accepted
~~~

Never deliver an unreviewed DOCX. OCR confidence is not proof of formula correctness; review mathematical symbols and diagrams against the source images.

MinerU is developed independently and is not included in this repository. If MinerU is unavailable, see the RapidOCR fallback instructions in SKILL.md. If a task requires genuine AxMath objects, use a separately maintained AxMath conversion workflow.
