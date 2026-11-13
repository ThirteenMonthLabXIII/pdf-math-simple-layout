# 排版 Skill 统一入口（新对话先看这里）

适用仓库：D:\work\pdf-math-simple-layout

## 新任务固定读取顺序

1. 完整阅读根目录 SKILL.md。
2. 完整阅读 docs\EXERCISE_WORKFLOW.md。题目完整性、答题空段落、图片尺寸、分页的统一硬规则都在这里。
3. 根据实际转换路线，只读对应工作流：
   - MinerU 已导出 DOCX：docs\MINERU_FIRST.md
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
