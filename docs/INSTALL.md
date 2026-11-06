# Installation

Python 3.11+ recommended. The real Word preview feature needs Windows, Microsoft Word and pywin32; source parsing and tests do not require Word.

~~~powershell
py -3 -m venv .venv
.\.venv\Scripts\python.exe -m pip install -r requirements.txt
.\.venv\Scripts\python.exe -m unittest discover -s tests -v
~~~

Use MinerU's separate web UI or a supported local MinerU renderer to export the full PDF as editable DOCX. The public Skill does not include or install MinerU. rapidocr-onnxruntime is used for the local OCR fallback only.

To let Codex discover it automatically, link or copy this repository into your Codex skill discovery directory (often under ~/.agents/skills/). The agent may need a new conversation to discover a newly added skill.
