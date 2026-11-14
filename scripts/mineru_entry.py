#!/usr/bin/env python3
"""Single local entry to the existing MinerU API, without copying its implementation.

Both GPT and Codex may invoke this; only their preferred initial route differs.
The actual API script reads its own local Token file. This bridge never reads it.
"""
from __future__ import annotations

import os
import subprocess
import sys
from pathlib import Path

DEFAULT_API_SCRIPT = Path(r"D:\CHATGPT\skills\word-axmath-gpt-skill\scripts\mineru_api.py")


def resolve_api_script() -> Path:
    """Explicit location overrides the default, so relocation needs no Skill rewrite."""
    override = os.environ.get("MINERU_API_SCRIPT", "").strip()
    if override:
        candidate = Path(os.path.expandvars(override)).expanduser()
        if not candidate.is_file():
            raise FileNotFoundError(
                "MINERU_API_SCRIPT points to a missing MinerU API script: "
                + str(candidate)
            )
        return candidate.resolve()
    if DEFAULT_API_SCRIPT.is_file():
        return DEFAULT_API_SCRIPT.resolve()
    raise FileNotFoundError(
        "MinerU API script not found. Set MINERU_API_SCRIPT to the existing "
        "mineru_api.py path; the API implementation/Token need not be copied."
    )


def main(argv: list[str] | None = None) -> int:
    try:
        script = resolve_api_script()
    except FileNotFoundError as exc:
        print(str(exc), file=sys.stderr)
        return 2
    args = list(sys.argv[1:] if argv is None else argv)
    # No shell, no Token argument, and no second API implementation.
    return subprocess.call([sys.executable, str(script), *args])


if __name__ == "__main__":
    raise SystemExit(main())
