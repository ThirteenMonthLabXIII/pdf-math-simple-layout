# MinerU API 通用入口（GPT、Codex 同权）

## 目的

网页转换和本机 API 是两种获取同一份 MinerU DOCX 的方式。**GPT 默认本机 API；Codex 默认当前对话的已登录侧栏网页，但双方都可随时调用 API。网页慢或卡住就直接 API，不要继续排队。两者权限相同。** 两种方式的 DOCX 最后都进入同一个 mineru_handoff.py prepare 流程。

本仓库的 `scripts\mineru_entry.py` 是 GPT 与 Codex 共用的 CLI 入口，默认转发给 `D:\CHATGPT\skills\word-axmath-gpt-skill\scripts\mineru_api.py`。如实际安装路径改变，只需设置 `MINERU_API_SCRIPT` 指向现有 API 脚本；无需复制 API 代码或 Token。

API 使用 MinerU Precision API，直接从本机提交 PDF、上传文件、查询任务并下载 DOCX，不需要浏览器前端。

## Codex 默认顺序（不是权限限制）

1. 在 Codex 当前对话自己的侧栏浏览器打开 MinerU。使用该侧栏已经登录的会话，不要另开一个未登录的浏览器窗口。
2. 上传完整 PDF，启动 DOCX 导出。
3. 出现以下任一情况，停止等待网页，转本机 API；也可以在任务需要时直接选择 API：
   - 页面显示预计排队超过 **120 秒**；
   - 状态连续 **60 秒没有进展**；
   - 前端卡住、上传或导出控件无法继续操作。
4. 如果网页已经成功导出完整 DOCX，直接使用该文件，不再重复调用 API。
5. API 完成后得到 mineru_original.docx，再进入统一的 prepare 阶段。

本机 API 仍受 MinerU 服务端处理状态影响；它绕过的是网页前端和网页队列等待，通过 API 直接提交并轮询服务端任务。若服务端本身返回 pending / running，脚本会轮询同一个任务，不会重复提交。

## 唯一 API 命令

在本仓库根目录运行：

~~~powershell
$py = 'D:\我的python\python3.12\python.exe'
$apiWorkspace = 'job1_mineru_api'
$profile = 'references\style_profile.json'
$apiScript = 'scripts\mineru_entry.py'
& $py $apiScript parse --input "input.pdf" --outdir $apiWorkspace --poll-interval 5 --max-wait-seconds 3600
& $py scripts\mineru_handoff.py prepare --pdf "input.pdf" --mineru-docx "$apiWorkspace\mineru_original.docx" --workspace "job1" --profile $profile
~~~

mineru_api.py 会产生：

- mineru_original.docx：完整的 MinerU 原始 DOCX。
- mineru_result.zip：原始完整结果包。
- mineru_output/：解压结果。
- mineru_job.json：源文件哈希、batch_id、任务状态和恢复信息。
- mineru_ready.json：本机 API 转换完成记录。

随后 mineru_handoff.py prepare 会复制原始 DOCX，自动执行既有图片尺寸规则、题目布局、答题空间和真实 Word 分页预检。不要为 API 路线另造排版流程。

## Token 位置与权限

- Token 只保存在本机专用文件：D:\CHATGPT\skills\word-axmath-gpt-skill\secrets\mineru_token.txt。
- 现有 API 脚本默认从该文件读取；不要在本仓库复制第二份 Token 或第二份 API 实现。
- 命令行只出现 Token 文件路径，不出现 Token 值。不得把 Token 写入提示词、源码、日志、任务 JSON、Git 或共享文件。
- GPT 与 Codex 对本机 API 和 Token 有相同的使用权限，无需让 Codex 额外申请授权。GPT 默认直接 API，Codex 默认已登录网页；排队慢或界面不稳定立即 API，也可从一开始直接 API。实际任务和用户当次指令优先。
- API 运行必须在能访问本机 Token 文件的 Windows 本机环境。Codex 云端 Linux 环境无法读取该本机文件，不得假装已调用。

## 断点恢复

~~~powershell
$apiScript = 'scripts\mineru_entry.py'
& $py $apiScript status --workspace "job1_mineru_api"
~~~

网络中断或本机进程退出后，重新运行同一条 parse 命令。脚本读取 mineru_job.json 并恢复同一个 batch。

若任务状态为 submission_unknown 或 submitting，先到 MinerU 任务列表核实，不得盲目重新提交。HTTP 401/403、A0202、A0211 时停止重试，核对本机 Token 文件。
