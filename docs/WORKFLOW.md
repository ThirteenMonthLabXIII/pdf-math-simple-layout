# RapidOCR 与 layout.json

仅在完整 MinerU DOCX 不可用时使用。先读 [练习工作流](EXERCISE_WORKFLOW.md)；用户禁止段落分页或换行设置时，在 `build` 命令末尾加 `--pagination-mode enter`。

~~~powershell
$py = 'D:\我的python\python3.12\python.exe'
& $py scripts\layout_cli.py prepare --pdf 'input.pdf' --output 'fallback_job'
& $py scripts\layout_cli.py recognize --workspace 'fallback_job'
# 对照原 PDF 修订 fallback_job\layout.json 后：
& $py scripts\layout_cli.py build --workspace 'fallback_job' --profile 'references\style_profile.json'
~~~

`layout.json` 是源页与内容块清单，`source_sha256` 绑定源 PDF。块按源稿顺序记录：

~~~json
{
  "page": 2,
  "verified": false,
  "blocks": [
    {"type": "heading", "text": "【例题 1】"},
    {"type": "question", "text": "1．如图，……"},
    {"type": "image", "bbox": [60, 125, 290, 230], "caption": ""},
    {"type": "subquestion", "text": "（1）求证：……"},
    {"type": "table", "rows": [["项目", "数据"], ["甲", "3"]]},
    {"type": "blank", "count": 3}
  ]
}
~~~

- 文本类型包括 `title`、`heading`、`question`、`subquestion`、`paragraph`、`math`；图的 `bbox` 使用 PDF 点坐标，默认 260 dpi，可按需提高。表格的行列非空。
- `blank` 的 count 为 1–100，表示相同数量的真实空段。只有目视确认无正文的源页才能标记 `source_blank: true`。原 PDF 页边界不是 Word 分页符；用户明确要求时才使用 `page_break`。
- OCR 可能混淆 `∠` / `<`、`△` / `A`、`∥` / `11`，以及撇号和负号。逐项对照原图，不凭置信度或同页码判断正确；图形裁切不能混入正文。

标准 build 必须报告 `ready_for_content_review` 才能进入内容审核。数学表达式须变成可编辑公式并对照源稿。按练习工作流渲染并逐页核对；`visual_review.json` 对每个真实源页恰有一条记录，四项检查均须实际完成后标真。最终哈希变化时重新核对受影响部分，再运行：

~~~powershell
& $py scripts\layout_cli.py render --docx 'fallback_job\candidate.docx' --output 'fallback_job\word_preview'
& $py scripts\layout_cli.py audit --workspace 'fallback_job' --docx 'fallback_job\candidate.docx' --word-pages
~~~

将示例 DOCX 路径替换为 build 实际输出。不要虚填审核字段，或把 OCR 审计结果当成逐页人工核对。
