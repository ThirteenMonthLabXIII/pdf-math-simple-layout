# PDF Math Simple Layout

## GPT 与 Codex 同权限：MinerU 路线

| 执行者 | 默认先用 | 另一选择 |
|---|---|---|
| GPT | 本机 MinerU API | 已登录网页 |
| Codex | 当前 Codex 对话侧边栏的已登录网页 | 本机 API |

双方可按当前任务直接选择网页或 API；Codex 使用 API 不需要重新取得用户授权。网页/API 的完整转换结果进入同一条题目规划和验收工作流，共同出口是 `scripts\mineru_handoff.py prepare`；用户禁止分页属性时加 `--pagination-mode enter`。共享 API 入口是 `scripts\mineru_entry.py`。默认路线及切换条件见 [MinerU 流程](docs/MINERU_FIRST.md)，API 操作见 [API 说明](docs/MINERU_API.md)。

本 Skill 用于数学 PDF 转可编辑 Word 和现有数学讲义的局部修复。本机唯一实际目录：`D:\CHATGPT\skills\pdf-math-simple-layout`。从 [SKILL.md](SKILL.md) 开始；练习题规则见 [练习工作流](docs/EXERCISE_WORKFLOW.md)，常见问题可直接从 Skill 入口的索引定位。

Windows 和 Microsoft Word 用于真实分页及预览；MinerU 为独立服务。安装见 [docs/INSTALL.md](docs/INSTALL.md)。其他入口：[RapidOCR](docs/WORKFLOW.md)、[教师样例](docs/REFERENCE.md)。

MIT License.
