# 排版 Skill 统一入口（新对话先看这里）

适用仓库：D:\work\pdf-math-simple-layout

## 最高优先级：不要把默认路线误判为权限限制

**GPT 与 Codex 都有完整 MinerU 网页/API/本机 Token 使用权限，无需为 Codex 另行申请授权。**

- GPT **默认本机 API**，也可使用网页登录路线。
- Codex **默认当前对话侧边栏已登录网页**；网页排队预计超过 120 秒、60 秒无进展、前端卡住或网页登录态不可用时，马上改用本机 API。实际需要时也可以一开始就使用 API。
- **这是优先级，不是必须先失败才准切换。** 统一 API 实现位置与操作命令见 README.md 顶部和 docs\MINERU_API.md。

## 新任务固定读取顺序

1. 完整阅读根目录 SKILL.md。
2. 完整阅读 docs\EXERCISE_WORKFLOW.md。题目完整性、答题空段落、图片尺寸、分页的统一硬规则都在这里。
3. 根据实际转换路线读取对应文件：
   - MinerU 转换：docs\MINERU_FIRST.md
   - Codex 执行 MinerU：再读 docs\CODEX_DIRECT.md
   - 准备使用本机 API 时读 docs\MINERU_API.md（GPT/Codex 均可）
   - RapidOCR / layout.json：docs\WORKFLOW.md
4. 有老师样例时再读 docs\REFERENCE.md；只有安装或环境故障时才读 docs\INSTALL.md。
5. 按现有脚本执行。不要从旧任务目录抄旧流程，不要搜索无关文件夹，不要自行新建第二套处理链。

## 生产顺序（不得调换）

保留原 PDF 和原始 MinerU DOCX
→ 完整转换一次
→ 首次插入/载入图片时自动确定最终显示尺寸
→ 应用段落样式
→ 识别完整题组并插入真正的 Enter 答题空段落
→ Word 真实分页预检
→ 内容编辑
→ 只有题目/小题结构发生变化时才重跑题目规划
→ 最终按源稿进行一次必要验收。

## MinerU 转换入口：看 README 顶部权限与优先级表

GPT 默认本机 API，Codex 默认本对话已登录侧边栏网页；**双方都能随时使用任一路线**。网页慢就切 API，不要无意义等待。唯一 API 位于 D:\CHATGPT\skills\word-axmath-gpt-skill\scripts\mineru_api.py，Token 由现有脚本读取，绝不输出其内容。

从网页或 API 获取 DOCX 后，统一调用 scripts\mineru_handoff.py prepare。操作文件：docs\MINERU_FIRST.md、docs\CODEX_DIRECT.md、docs\MINERU_API.md。

## 图片尺寸规则（首次制作时执行）

生产代码统一使用 scripts\image_sizing.py，MinerU prepare 和 RapidOCR build 自动调用。

- 同一讲中宽度大于 0 且不超过 6.0 cm 的现有小图作为参照，取中位数。
- 本讲宽度超过 6.0 cm 的图片，首次制作时直接缩到本讲中位数；小图原尺寸保留，绝不放大。
- 本讲无参照图时，使用整份文档的小图中位数；整份文档也无参照图时使用 4.05 cm。
- 只改变 Word 显示尺寸，严格保持宽高比；不重采样、不替换图片素材。
- 规则默认值在 references\style_profile.json，可通过 image_size_policy 覆盖。

图片尺寸必须在题目布局和分页前确定。不要在终稿后另开一轮图片缩放，也不要反复缩小、检查、再缩小。只有后来新增或替换图片时，才在下一次题目规划前对工作稿应用同一规则一次；纯文字修改不触发图片处理。

## 发生问题时先看哪里

| 问题 | 先看 |
|---|---|
| 不知道完整生产流程 | SKILL.md → docs\EXERCISE_WORKFLOW.md |
| MinerU 优先级、网页/API 选择 | README.md 顶部 → docs\MINERU_FIRST.md → docs\CODEX_DIRECT.md / docs\MINERU_API.md |
| MinerU 导出后的 DOCX 如何准备 | docs\MINERU_FIRST.md → scripts\mineru_handoff.py |
| RapidOCR 图片、题目块、DOCX 如何生成 | docs\WORKFLOW.md → scripts\layout_cli.py |
| 图片尺寸如何计算/应用 | scripts\image_sizing.py → references\style_profile.json |
| 题目识别、答题空行、整题同页 | scripts\question_layout.py → tests\test_image_sizing.py |
| 段落样式和公式段落 | scripts\docx_styling.py → docs\REFERENCE.md |
| 安装依赖或环境问题 | docs\INSTALL.md |

## 简洁执行原则

- 先导航到上述入口，再按任务路线执行；不要无目的全仓搜索。
- 不重复审计已经有明确结果的步骤；不为“看起来稳妥”新增多轮相同检查。
- 规则在首次制作时落实，终稿验收只验证内容和分页，不承担事后补救。
- 原始文件不覆盖；只在工作副本上处理。
- 只交付用户要求的产物，不生成无关 PDF、报告或中间副本。
