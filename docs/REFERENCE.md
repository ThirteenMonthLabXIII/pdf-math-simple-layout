# 教师格式样例

样例只提供格式，不提供题目内容。生成临时 profile 后，在 [MinerU 流程](MINERU_FIRST.md) 的 prepare 或 [RapidOCR 流程](WORKFLOW.md) 的 build 中传入：

~~~powershell
$py = 'D:\我的python\python3.12\python.exe'
$profile = Join-Path $env:TEMP 'pdf-math-style-profile.json'
& $py scripts\layout_cli.py profile --sample 'reference.docx' --output $profile
~~~

提取器排除异常缩进；题目首段顶格、真实 Enter 答题空段和当前用户要求优先。用户禁止段落分页属性时按练习工作流的 Enter 分支处理，不调用会写入分页链的 prepare/build。