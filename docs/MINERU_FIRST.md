# MinerU 转换与 Word 交接

先读 [练习工作流](EXERCISE_WORKFLOW.md)。GPT 默认使用本机 MinerU API；Codex 默认使用当前对话侧边栏已登录网页。两者都可根据任务直接选择网页或 API。网页预计排队超过 120 秒、连续 60 秒无进展、卡住或登录态不可用时切换 API；已有完整导出时不重复转换。网页操作见 [CODEX_DIRECT.md](CODEX_DIRECT.md)，API 和任务恢复见 [MINERU_API.md](MINERU_API.md)。

在 Skill 根目录选择预装默认样式；有教师样例时按 [样例说明](REFERENCE.md) 生成临时 profile。以下路径按实际文件替换：

~~~powershell
$py = 'D:\我的python\python3.12\python.exe'
$profile = 'references\style_profile.json'
& $py scripts\mineru_handoff.py prepare --pdf 'input.pdf' --mineru-docx 'mineru.docx' --workspace 'job1' --profile $profile
~~~

标准 prepare 保留原始导出、建立工作稿，依次应用图片尺寸、受控段落样式、题组和 Enter 空段，建立整题同页链并运行真实 Word 分页预检。`question_layout_plan.json` 记录规划和预检结果。

仅在输出 `ready_for_codex_formatting` 后编辑题目内容。`needs_question_mapping`、`needs_question_layout`、`needs_question_page_fit` 或 `needs_word_pagination_check` 分别表示需要补题组映射、修复结构、调整题组高度或恢复 Word 分页检查；修好后再预检。用户禁止段落分页属性时，**不要运行上面的 prepare**，改走 [Enter 分页分支](EXERCISE_WORKFLOW.md#用户禁止段落分页或换行设置时)。

复用已有 SIJI 样式；保留公式节点和原有空格。仅当编辑新增、删除或调整题目/小问结构时重新运行规划器；新增或替换图片时先按练习工作流处理尺寸。标准路线：

~~~powershell
& $py scripts\question_layout.py apply --docx 'job1\working.docx' --profile $profile
& $py scripts\question_layout.py check --docx 'job1\working.docx' --profile $profile --word-pages
~~~

标准路线在完成编辑后渲染、按源页清单逐页核对，再对实际候选稿审计：

~~~powershell
& $py scripts\layout_cli.py render --docx 'job1\working.docx' --output 'job1\word_preview'
& $py scripts\docx_layout_check.py check --docx 'job1\working.docx' --output 'job1\docx_layout_findings.json'
& $py scripts\docx_layout_check.py template --docx 'job1\working.docx' --output 'job1\layout_review.json'
& $py scripts\mineru_handoff.py audit --workspace 'job1' --docx 'job1\working.docx' --require-accepted --word-pages
~~~

仅在实际核对后填写 `mineru_review.json` 的当前 DOCX 哈希、`word_preview_checked` 与每个源页三项检查；`layout_review.json` 填写当前哈希、`checked`、审阅者，并为保留的发现记录 `intentional`、原因和已核对源稿。审计警告必须保留并对照源稿处理，不能通过虚填消除。用户禁止分页属性时，默认结构审计仍要求分页链；保留它的原始结果，并按 Enter 分支记录实际 Word 分页和属性检查，不伪造 `accepted`。