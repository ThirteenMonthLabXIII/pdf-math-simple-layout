"""The MinerU entry delegates to exactly one existing API implementation."""
import os
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "scripts"))
import mineru_entry


class MinerUEntryTest(unittest.TestCase):
    def test_explicit_override_preserves_args_and_exit_status(self):
        with tempfile.TemporaryDirectory() as d:
            folder = Path(d)
            script = folder / "api_stub.py"
            report = folder / "args.txt"
            script.write_text(
                "import pathlib,sys\n"
                f"pathlib.Path({str(report)!r}).write_text('|'.join(sys.argv[1:]),encoding='utf-8')\n"
                "raise SystemExit(7)\n",
                encoding="utf-8"
            )
            with patch.dict(os.environ, {"MINERU_API_SCRIPT": str(script)}):
                self.assertEqual(mineru_entry.resolve_api_script(), script.resolve())
                self.assertEqual(mineru_entry.main(["status", "--workspace", "w1"]), 7)
            self.assertEqual(report.read_text(encoding="utf-8"), "status|--workspace|w1")

    def test_invalid_override_does_not_silently_use_another_script(self):
        with tempfile.TemporaryDirectory() as d:
            missing = Path(d) / "nonexistent_api.py"
            with patch.dict(os.environ, {"MINERU_API_SCRIPT": str(missing)}):
                self.assertEqual(mineru_entry.main(["status"]), 2)


if __name__ == "__main__":
    unittest.main()
