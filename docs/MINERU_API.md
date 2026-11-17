# MinerU 本机 API

网页和 API 都生成完整 MinerU DOCX，执行者的默认选择及切换条件见 [MinerU 流程](MINERU_FIRST.md)。API 通过 `scripts\mineru_entry.py` 调用现有共享实现；必要时用 `MINERU_API_SCRIPT` 指向该实现。Token 由原实现从本机读取，不能写入命令、日志、JSON、Git 或共享文件。

在 Skill 根目录运行：

~~~powershell
$py = 'D:\我的python\python3.12\python.exe'
$apiWorkspace = 'job1_mineru_api'
& $py scripts\mineru_entry.py parse --input 'input.pdf' --outdir $apiWorkspace --poll-interval 5 --max-wait-seconds 3600
~~~

完成后工作区保留 `mineru_original.docx`、结果 ZIP、解压内容、`mineru_job.json` 和 `mineru_ready.json`。继续按 [MinerU 流程](MINERU_FIRST.md) 选择标准 prepare 或用户禁止分页属性时的 Enter 分支。

网络中断后重跑同一条 parse 命令以恢复原 batch；查看状态：

~~~powershell
& $py scripts\mineru_entry.py status --workspace 'job1_mineru_api'
~~~

`submission_unknown` 或 `submitting` 时先核对 MinerU 任务列表，避免重复提交。HTTP 401/403、A0202、A0211 时停止重试并检查本机 Token。缺少本机环境或凭据时如实报告，不宣称已执行。