"""Prevent the MinerU priority guideline from drifting into a GPT/Codex permission gate."""
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
DOCS = (
    "README.md",
    "START_HERE.md",
    "SKILL.md",
    "docs/CODEX_DIRECT.md",
    "docs/MINERU_API.md",
    "docs/MINERU_FIRST.md",
    "docs/EXERCISE_WORKFLOW.md",
)


class MinerURoutingContractTest(unittest.TestCase):
    def test_readme_leads_with_equal_access_and_default_priorities(self):
        head = (ROOT / "README.md").read_text(encoding="utf-8").splitlines()[:30]
        intro = "\n".join(head)
        for required in ("GPT 与 Codex 同权限", "默认先用", "本机 MinerU API", "当前 Codex 对话侧边栏", "共同出口"):
            self.assertIn(required, intro)

    def test_docs_do_not_reintroduce_codex_token_restrictions(self):
        forbidden = (
            "Codex 仅在本 Skill",
            "授权仅限",
            "Codex 默认不得",
            "Codex 默认禁止",
            "只有触发本机 API 备用通道时",
            "only when the local API fallback is triggered",
        )
        for name in DOCS:
            with self.subTest(doc=name):
                content = (ROOT / name).read_text(encoding="utf-8")
                for phrase in forbidden:
                    self.assertNotIn(phrase, content)

    def test_routing_is_central_and_both_actors_can_choose_api(self):
        readme = (ROOT / "README.md").read_text(encoding="utf-8")
        starter = (ROOT / "START_HERE.md").read_text(encoding="utf-8")
        skill = (ROOT / "SKILL.md").read_text(encoding="utf-8")
        self.assertIn("不需要重新取得用户授权", readme)
        self.assertIn("README.md 顶部", starter)
        self.assertIn("No additional Codex authorization is required", skill)
        self.assertIn("mineru_handoff.py prepare", readme)
        self.assertIn("docs/MINERU_API.md", skill)


if __name__ == "__main__":
    unittest.main()
