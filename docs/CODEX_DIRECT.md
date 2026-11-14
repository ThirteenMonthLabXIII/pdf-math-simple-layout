# Codex 本机 MinerU 使用规则

本文件只负责 PDF → MinerU DOCX 阶段。AxMath DOCX → AxMath 转换仍按对应 AxMath Skill 的规则执行。

## 第一选择：Codex 当前对话的侧边栏网页

1. 使用 **当前 Codex 对话自己的侧边栏浏览器** 打开 MinerU。
2. 直接使用该侧栏已经登录的 MinerU 会话上传 PDF、导出 DOCX。
3. 不需要 API、Token 或额外登录。不要另开一个未登录的浏览器窗口；如果显示未登录，先确认打开的是 Codex 当前对话的侧边栏会话。
4. 网页导出成功后，将 DOCX 保存到任务工作目录，直接运行 mineru_handoff.py prepare。

## 长队列或前端卡顿：切换本机 API

网页显示预计排队超过 120 秒、状态连续 60 秒没有进展，或前端卡住时，**不要继续等网页**。在 Codex 所运行的 Windows 本机执行：

~~~powershell
$py = 'D:\我的python\python3.12\python.exe'
$apiWorkspace = 'job1_mineru_api'
$profile = 'references\style_profile.json'
$apiScript = 'D:\CHATGPT\skills\word-axmath-gpt-skill\scripts\mineru_api.py'
& $py $apiScript parse --input "input.pdf" --outdir $apiWorkspace --poll-interval 5 --max-wait-seconds 3600
& $py scripts\mineru_handoff.py prepare --pdf "input.pdf" --mineru-docx "$apiWorkspace\mineru_original.docx" --workspace "job1" --profile $profile
~~~

API 路线直接通过本机 Python 调用 MinerU Precision API，不需要浏览器前端。它与网页路线的输出汇合到同一个 prepare 流程。

## 使用权限和切换原则（不能误判为限制）

**Codex 与 GPT 权限相同，均可使用本机 MinerU API、Token 文件和已登录网页。Codex 默认先走侧边栏网页是速度与操作习惯的优先级建议，API 不需要额外授权，也不要求先等待网页失败。**

- 默认先打开当前 Codex 对话自己侧边栏的已登录 MinerU 页面；不用 Token。
- 遇到排队预计超过120秒、60秒无进展、前端卡住、登录态不可用或侧边栏打不开，立即使用本机 API。任务需要时也可以直接 API 开始。
- API 已由 D:\CHATGPT\skills\word-axmath-gpt-skill\scripts\mineru_api.py 实现，本机 Token 文件由脚本自动读取。
- 不在提示词、日志、命令行、JSON、Git 或共享目录中泄露 Token。对于其他 Word/AxMath 步骤，执行对应的 Skill 工作流；身份本身不构成能力限制。
- 本机文件只有在本机环境可访问时才能读取；缺少本机环境就使用实际可用的方法，不宣称已在本机执行。

遇到 API 任务状态问题，先运行 D:\CHATGPT\skills\word-axmath-gpt-skill\scripts\mineru_api.py status --workspace "job1_mineru_api"，不要新建重复 batch。API 入口、恢复规则和错误定位见 docs\MINERU_API.md。
