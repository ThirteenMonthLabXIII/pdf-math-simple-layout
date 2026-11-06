---
name: pdf-math-simple-layout
description: Convert mathematics PDF handouts into editable Word, preferring MinerU DOCX exports, retaining clear geometry diagrams and performing source-bound visual review. Fall back to RapidOCR when MinerU is unavailable.
---

# Math PDF → Word: simple layout skill

## When to use

Use for mathematics worksheets, scanned handouts, and exercise compilations when editable Word body text is required. Figures may be preserved as clear source images. Use a separate AxMath workflow only when the client explicitly requests genuine editable AxMath equations.

## Default: MinerU-first

Read docs/MINERU_FIRST.md. Export a complete editable DOCX using MinerU (web or compatible local renderer). Check that the source PDF was fully converted. The Skill does not install MinerU or ship its models.

Create a fresh workspace (do not overwrite any original):

~~~powershell
$py = ".\.venv\Scripts\python.exe"
& $py scripts\mineru_handoff.py prepare --pdf "input.pdf" --mineru-docx "mineru.docx" --workspace "job1"
~~~

This keeps an untouched mineru_original.docx and a working.docx. Keep editing the existing DOCX with Word/Word-compatible tooling rather than rebuilding its content from OCR.

Correct all OCR/formula problems by comparing the PDF original; retain numbering, subquestions, data, tables, captions, and geometry diagrams in the original order. When geometry crops are needed, take them from the source PDF at high resolution and keep them adjacent to their corresponding question. Never infer unclear mathematics from context.

A formatting reference DOCX is for styles only, not content. A generic style profile is included in references/style_profile.json. The layout_cli.py profile command can extract basic margins and fonts from another sample.

Render the result with genuine Microsoft Word (Windows COM):

~~~powershell
& $py scripts\layout_cli.py render --docx "job1\working.docx" --output "job1\word_preview"
~~~

Compare all source PDF pages with all actual output Word pages. Record the review accurately in job1/mineru_review.json, including the last candidate SHA-256 and all page checks. Never tick review fields automatically.

~~~powershell
& $py scripts\mineru_handoff.py audit --workspace "job1" --require-accepted
~~~

Only deliver when audit status is accepted **and** the actual page/figure/math review has been performed. After editing the Word file again, the old review hash is invalid.

## Fallback: local RapidOCR

Only if MinerU is unavailable or DOCX export fails. Do not repeat OCR after a valid MinerU output.

~~~powershell
& $py scripts\layout_cli.py prepare --pdf "input.pdf" --output "fallback_job"
& $py scripts\layout_cli.py recognize --workspace "fallback_job"
# Correct blocks and high-resolution figure bbox regions in fallback_job/layout.json.
& $py scripts\layout_cli.py build --workspace "fallback_job" --profile "references/style_profile.json"
& $py scripts\layout_cli.py render --docx "fallback_job/candidate.docx" --output "fallback_job/word_preview"
& $py scripts\layout_cli.py audit --workspace "fallback_job"
~~~

For block types, cropping, review and source integrity, see docs/WORKFLOW.md. All programs should operate on copies in a separate workspace and preserve existing original files.

## Software

- scripts/mineru_handoff.py: import existing MinerU DOCX and verify source hashes and review status.
- scripts/layout_cli.py: local OCR fallback, structure, formatting, Word render and audit.
- scripts/start_one_lecture.ps1 and scripts/finish_one_lecture.ps1: optional fallback shortcuts.
- tests: deterministic local checks without Word or OCR.

Run from the Codex local terminal. No remote computer connector is required.
