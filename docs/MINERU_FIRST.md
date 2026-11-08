# MinerU-first PDF-to-Word handoff

1. Export the entire source PDF using the supported MinerU web UI or supported local DOCX renderer. Verify that all pages and key problems appear in the downloaded editable DOCX.
2. Import source and MinerU output:

~~~powershell
python scripts\mineru_handoff.py prepare --pdf "input.pdf" --mineru-docx "mineru.docx" --workspace "job1"
~~~

This stores original hashes, PDF page reference images, a read-only raw export copy, and a working.docx.
3. Edit working.docx directly. Correct formula OCR, question numbers, superscripts, tables, and captions. Preserve clear geometric drawings as embedded images cropped from original PDF; keep body text editable. Format using your supplied Word style reference.
4. Render and actually check the output using Microsoft Word:

~~~powershell
python scripts\layout_cli.py render --docx "job1\working.docx" --output "job1\word_preview"
~~~

5. Run the mandatory paragraph-layout regression check on the final DOCX:

~~~powershell
python scripts/docx_layout_check.py check --docx "job1/working.docx" --output "job1/docx_layout_findings.json"
python scripts/docx_layout_check.py template --docx "job1/working.docx" --output "job1/layout_review.json"
~~~

The scanner flags indentation above 1000 twips, before/after spacing above 720 twips, and 3+ spaces inside answer parentheses. Inspect every finding against the source PDF. A 600-twip nested-question indent may be intentional; do not zero all indents or bulk-delete blank paragraphs. Gray dots and return arrows in Word are usually nonprinting formatting marks; verify with marks hidden or the exported PDF before editing them. Fix findings, save, and rerun both commands. Any finding intentionally retained must have a reason and source_pdf_checked=true; set checked=true and a reviewer in layout_review.json. The review file is bound to the final DOCX hash.

6. Compare each original PDF page against the corresponding content in reflowed Word pages. Update mineru_review.json only after checking content, formulas and figures. Set word_preview_checked=true only after reviewing output images, and candidate_sha256 to the actual final DOCX hash.
7. Audit:

~~~powershell
python scripts\mineru_handoff.py audit --workspace "job1" --require-accepted
~~~

Never deliver an unreviewed DOCX. Do not trust an OCR confidence number as proof of formula correctness; review math and diagrams with source images.

MinerU is developed independently and is not included in this repository. If MinerU is unavailable, see the RapidOCR fallback instructions in SKILL.md. If a task requires genuine AxMath objects, use a separately maintained AxMath conversion workflow.
