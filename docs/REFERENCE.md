# Formatting reference sample

The example Word document is a **style reference**, not a content source. It may be a completely different chapter from the input PDF.

To extract basic formatting into a profile:

~~~powershell
python scripts\layout_cli.py profile --sample "reference.docx" --output "style_profile.json"
~~~

For the RapidOCR fallback use the profile with the build command. For MinerU DOCX input, apply these styles directly to working.docx using Word or python-docx, and then inspect the rendered result. Verify that no original mathematics, figures or question order changed.

No teacher handouts or copyrighted source PDFs are distributed with this repository.
