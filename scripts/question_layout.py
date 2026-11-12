#!/usr/bin/env python3
"""Preventive question grouping, Enter-created answer space, and page-safety rules."""
from __future__ import annotations

import argparse
import hashlib
import json
import re
from pathlib import Path

from docx import Document
from docx.enum.style import WD_STYLE_TYPE
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.table import Table
from docx.text.paragraph import Paragraph

from docx_styling import DRAWING_TAGS, MATH_TAGS, STYLE_NAMES, classify_paragraph, iter_paragraphs

SCHEMA = "question-layout-plan/1"
QUESTION_HEADING_RE = re.compile(
    r"(例题|练习|拓展|课后练习|课后作业|巩固练习|综合练习|思考题|试一试|作业)"
)
INLINE_SUBQUESTION_RE = re.compile(r"[（(]\s*[1-9][0-9]*\s*[）)]")
PROOF_RE = re.compile(r"(证明|求证|论证|作图|画图|几何证明)")
COMPLEX_RE = re.compile(
    r"(应用题|工程问题|行程问题|顺水|逆水|年龄问题|数字问题|方案|联立|"
    r"综合|分类讨论|最值|函数图像|统计分析|概率分析)"
)
CALC_RE = re.compile(r"(计算|求解|解方程|列方程|方程组|求出|求值|化简)")


def default_answer_space_policy() -> dict:
    """Counts are empty Word paragraphs (Enter presses), never visual line tricks."""
    return {
        "simple_paragraphs": 2,
        "standard_paragraphs": 3,
        "calculation_paragraphs": 4,
        "complex_paragraphs": 5,
        "proof_paragraphs": 6,
        "max_per_subquestion": 8,
    }


def sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with Path(path).open("rb") as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def _numbering_format(paragraph):
    """Return an effective Word numbering format, if the paragraph is auto-numbered."""
    num_pr = None
    p_pr = paragraph._p.pPr
    if p_pr is not None:
        num_pr = p_pr.numPr
    style = paragraph.style
    visited = set()
    while num_pr is None and style is not None and style.style_id not in visited:
        visited.add(style.style_id)
        style_ppr = style._element.pPr
        if style_ppr is not None:
            num_pr = style_ppr.numPr
        style = style.base_style
    if num_pr is None or num_pr.numId is None:
        return None

    num_id = str(num_pr.numId.val)
    level = str(num_pr.ilvl.val if num_pr.ilvl is not None else 0)
    numbering = paragraph.part.numbering_part.element
    nums = numbering.xpath(f'./w:num[@w:numId="{num_id}"]')
    if not nums:
        return None
    abstract_id = nums[0].abstractNumId.val
    abstracts = numbering.xpath(f'./w:abstractNum[@w:abstractNumId="{abstract_id}"]')
    if not abstracts:
        return None
    levels = abstracts[0].xpath(f'./w:lvl[@w:ilvl="{level}"]')
    if not levels:
        return None
    num_fmt = levels[0].numFmt
    return num_fmt.val if num_fmt is not None else None


def _is_decimal_numbered(paragraph) -> bool:
    return _numbering_format(paragraph) in {
        "decimal", "decimalZero", "upperRoman", "lowerRoman",
        "upperLetter", "lowerLetter",
    }


def _is_question_heading(text: str) -> bool:
    return bool(QUESTION_HEADING_RE.search(text or ""))


def _is_empty_answer_paragraph(paragraph) -> bool:
    """Only a truly empty paragraph counts; spaces, tabs, breaks and drawings do not."""
    if paragraph.text != "":
        return False
    for node in paragraph._p.iter():
        if node.tag in DRAWING_TAGS or node.tag in MATH_TAGS:
            return False
        if node.tag in {qn("w:tab"), qn("w:br")}:
            return False
    return True


def _body_blocks(document):
    """Keep the original top-level paragraph/table order."""
    blocks = []
    for element in document.element.body.iterchildren():
        if element.tag == qn("w:p"):
            paragraph = Paragraph(element, document)
            kind = classify_paragraph(paragraph)
            text = paragraph.text or ""
            if not text and _is_decimal_numbered(paragraph):
                kind = "question_marker"
            blocks.append({
                "element": element, "paragraph": paragraph, "table": None,
                "kind": kind, "text": text, "index": len(blocks),
            })
        elif element.tag == qn("w:tbl"):
            table = Table(element, document)
            text = "\n".join(cell.text for row in table.rows for cell in row.cells)
            blocks.append({
                "element": element, "paragraph": None, "table": table,
                "kind": "table", "text": text, "index": len(blocks),
            })
    return blocks


def _new_group(start: int, question_index: int) -> dict:
    return {
        "start": start,
        "end": question_index,
        "question_index": question_index,
        "subquestion_starts": [],
    }


def build_question_groups(document) -> tuple[list[dict], list[dict]]:
    """Find complete logical questions, including example/exercise headings."""
    blocks = _body_blocks(document)
    groups = []
    current = None
    pending_question_heading = None

    def close(end_index):
        nonlocal current
        if current is not None:
            current["end"] = max(current["start"], end_index)
            groups.append(current)
            current = None

    for block in blocks:
        index = block["index"]
        kind = block["kind"]
        text = block["text"]

        if kind in {"title", "heading"}:
            if current is not None:
                close(index - 1)
            pending_question_heading = (
                index if kind == "heading" and _is_question_heading(text) else None
            )
            continue

        if kind in {"question", "question_start", "question_marker"}:
            if current is not None:
                close(index - 1)
            start = pending_question_heading if pending_question_heading is not None else index
            current = _new_group(start, index)
            if INLINE_SUBQUESTION_RE.search(text):
                current["subquestion_starts"].append(index)
            pending_question_heading = None
            continue

        if kind == "subquestion":
            if current is None:
                start = pending_question_heading if pending_question_heading is not None else index
                current = _new_group(start, index)
            current["subquestion_starts"].append(index)
            current["end"] = index
            pending_question_heading = None
            continue

        if current is None and pending_question_heading is not None:
            if kind in {"body", "figure", "caption", "formula", "table", "answer_blank"}:
                current = _new_group(pending_question_heading, index)
                pending_question_heading = None
                continue

        if current is not None:
            current["end"] = index

    if current is not None:
        close(len(blocks) - 1)
    return blocks, groups


def _segment_text(blocks, start, end):
    texts = []
    has_figure = False
    has_table = False
    for block in blocks[start:end + 1]:
        if block["text"]:
            texts.append(block["text"])
        has_figure = has_figure or block["kind"] == "figure"
        has_table = has_table or block["kind"] == "table"
    return "\n".join(texts), has_figure, has_table


def estimate_answer_paragraphs(text: str, has_figure=False, has_table=False,
                               policy: dict | None = None) -> int:
    """Estimate handwritten working space from the actual task complexity."""
    cfg = {**default_answer_space_policy(), **(policy or {})}
    text = text or ""
    if re.search(r"(选择题|填空题|判断题)", text):
        count = cfg["simple_paragraphs"]
    elif PROOF_RE.search(text):
        count = cfg["proof_paragraphs"]
    elif has_figure or has_table or COMPLEX_RE.search(text) or len(text) >= 140:
        count = cfg["complex_paragraphs"]
    elif CALC_RE.search(text) or len(text) >= 55:
        count = cfg["calculation_paragraphs"]
    elif re.search(r"(选择题|填空题|判断题)", text) or (len(text) < 60 and re.search(r"(选择|填空|判断|写出)", text)):
        count = cfg["simple_paragraphs"]
    elif len(text) < 32:
        count = cfg["simple_paragraphs"]
    else:
        count = cfg["standard_paragraphs"]
    return max(1, min(int(count), int(cfg["max_per_subquestion"])))


def _trailing_empty_count(blocks, start, end):
    count = 0
    for block in reversed(blocks[start:end + 1]):
        paragraph = block["paragraph"]
        if paragraph is not None and _is_empty_answer_paragraph(paragraph):
            count += 1
        else:
            break
    return count


def _answer_segments(blocks, group, policy):
    starts = sorted(set(group["subquestion_starts"]))
    if not starts:
        start = group["question_index"]
        end = group["end"]
        text, has_figure, has_table = _segment_text(blocks, start, end)
        required = estimate_answer_paragraphs(text, has_figure, has_table, policy)
        return [{
            "start": start, "end": end, "required": required,
            "existing": _trailing_empty_count(blocks, start, end),
            "kind": "whole_question",
        }]

    segments = []
    question_stem = blocks[group["question_index"]]["text"]
    group_has_figure = any(block["kind"] == "figure" for block in blocks[group["start"]:group["end"] + 1])
    group_has_table = any(block["kind"] == "table" for block in blocks[group["start"]:group["end"] + 1])
    for offset, start in enumerate(starts):
        end = starts[offset + 1] - 1 if offset + 1 < len(starts) else group["end"]
        text, has_figure, has_table = _segment_text(blocks, start, end)
        context_text = "\n".join(part for part in (question_stem, text) if part)
        inline_count = max(1, len(INLINE_SUBQUESTION_RE.findall(blocks[start]["text"])))
        per_subquestion = estimate_answer_paragraphs(
            context_text,
            has_figure or group_has_figure,
            has_table or group_has_table,
            policy,
        )
        required = per_subquestion * inline_count
        segments.append({
            "start": start, "end": end, "required": required,
            "existing": _trailing_empty_count(blocks, start, end),
            "kind": "subquestion",
        })
    return segments


def _style_for(document, kind):
    from docx_styling import STYLE_NAMES
    name = STYLE_NAMES[kind]
    try:
        return document.styles[name]
    except KeyError:
        style = document.styles.add_style(name, WD_STYLE_TYPE.PARAGRAPH)
        style.base_style = document.styles["Normal"]
        return style


def _insert_empty_paragraph_after(element, style):
    paragraph_xml = OxmlElement("w:p")
    p_pr = OxmlElement("w:pPr")
    p_style = OxmlElement("w:pStyle")
    p_style.set(qn("w:val"), style.style_id)
    p_pr.append(p_style)
    paragraph_xml.append(p_pr)
    element.addnext(paragraph_xml)
    return paragraph_xml


def _remove_manual_page_controls(document):
    removed = {
        "manual_page_breaks": 0,
        "page_break_before": 0,
        "empty_page_break_paragraphs": 0,
    }
    body = document.element.body
    # Remove manual page breaks from every body descendant and remember their
    # containing paragraphs so page-break-only wrappers can be deleted too.
    page_break_paragraphs = set()
    for node in list(body.iter(qn("w:br"))):
        if node.get(qn("w:type")) == "page":
            paragraph_node = node
            while paragraph_node is not None and paragraph_node.tag != qn("w:p"):
                paragraph_node = paragraph_node.getparent()
            if paragraph_node is not None:
                page_break_paragraphs.add(paragraph_node)
            node.getparent().remove(node)
            removed["manual_page_breaks"] += 1
    for node in list(body.iter(qn("w:lastRenderedPageBreak"))):
        node.getparent().remove(node)

    # A paragraph created only by Document.add_page_break() becomes empty
    # after its break is removed. Delete that wrapper to avoid a blank page top.
    for paragraph_node in page_break_paragraphs:
        if paragraph_node.getparent() is None:
            continue
        paragraph = Paragraph(paragraph_node, document)
        parent = paragraph_node.getparent()
        # Word table cells must retain at least one paragraph; only remove
        # page-break-only wrappers that are direct children of the document body.
        if parent is not None and parent.tag == qn("w:body") and _is_empty_answer_paragraph(paragraph):
            parent.remove(paragraph_node)
            removed["empty_page_break_paragraphs"] += 1

    for paragraph in iter_paragraphs(document):
        p_pr = paragraph._p.pPr
        if p_pr is not None:
            page_before = p_pr.find(qn("w:pageBreakBefore"))
            if page_before is not None:
                p_pr.remove(page_before)
                removed["page_break_before"] += 1
        paragraph.paragraph_format.page_break_before = False
    return removed


def _set_row_cannot_split(row):
    tr_pr = row._tr.get_or_add_trPr()
    if tr_pr.find(qn("w:cantSplit")) is None:
        tr_pr.append(OxmlElement("w:cantSplit"))


def _apply_group_pagination(document, blocks, groups):
    for group in groups:
        start, end = group["start"], group["end"]
        question_index = group["question_index"]
        question_paragraph = blocks[question_index]["paragraph"]
        if question_paragraph is not None:
            question_paragraph.style = _style_for(document, "question_start")

        for subquestion_index in group["subquestion_starts"]:
            if subquestion_index == question_index:
                continue
            paragraph = blocks[subquestion_index]["paragraph"]
            if paragraph is not None:
                paragraph.style = _style_for(document, "subquestion")

        for index in range(start, end + 1):
            block = blocks[index]
            is_not_last = index < end
            paragraph = block["paragraph"]
            if paragraph is not None:
                paragraph.paragraph_format.keep_together = True
                paragraph.paragraph_format.keep_with_next = is_not_last
                paragraph.paragraph_format.page_break_before = False
            table = block["table"]
            if table is not None:
                for row in table.rows:
                    _set_row_cannot_split(row)
                    for cell in row.cells:
                        for cell_paragraph in cell.paragraphs:
                            cell_paragraph.paragraph_format.keep_together = True
                            cell_paragraph.paragraph_format.keep_with_next = is_not_last
                            cell_paragraph.paragraph_format.page_break_before = False

        # Keep an immediately preceding chapter/section title with this question.
        previous = start - 1
        while previous >= 0:
            block = blocks[previous]
            paragraph = block["paragraph"]
            if paragraph is None:
                break
            if block["kind"] in {"title", "heading"}:
                paragraph.paragraph_format.keep_together = True
                paragraph.paragraph_format.keep_with_next = True
                paragraph.paragraph_format.page_break_before = False
                previous -= 1
                continue
            break


def _normalize_answer_blank(paragraph, answer_style):
    """Make an existing empty paragraph a plain Enter line, not a visual-format trick."""
    paragraph.style = answer_style
    for child in list(paragraph._p):
        if child.tag != qn("w:pPr"):
            paragraph._p.remove(child)
    p_pr = paragraph._p.get_or_add_pPr()
    for tag in ("pBdr", "tabs", "ind", "spacing", "shd", "framePr", "jc"):
        node = p_pr.find(qn("w:" + tag))
        if node is not None:
            p_pr.remove(node)


def _set_existing_answer_blanks(blocks, segments, document):
    answer_style = _style_for(document, "answer_blank")
    for segment in segments:
        start, end = segment["start"], segment["end"]
        index = end
        while index >= start:
            paragraph = blocks[index]["paragraph"]
            if paragraph is None or not _is_empty_answer_paragraph(paragraph):
                break
            _normalize_answer_blank(paragraph, answer_style)
            index -= 1


def apply_question_layout(docx_path: Path, profile: dict | None = None) -> dict:
    """Insert Enter-created answer paragraphs and bind each question into one page block."""
    docx_path = Path(docx_path).resolve(strict=True)
    profile = profile or {}
    policy = {
        **default_answer_space_policy(),
        **(profile.get("answer_space_policy") or {}),
    }
    document = Document(str(docx_path))
    removed_breaks = _remove_manual_page_controls(document)
    _style_for(document, "question_start")
    _style_for(document, "answer_blank")

    blocks, groups = build_question_groups(document)
    insertion_jobs = []
    segment_reports = []
    for group_number, group in enumerate(groups, start=1):
        segments = _answer_segments(blocks, group, policy)
        for segment in segments:
            missing = max(0, segment["required"] - segment["existing"])
            segment_reports.append({
                "group": group_number,
                "kind": segment["kind"],
                "required_empty_paragraphs": segment["required"],
                "existing_empty_paragraphs": segment["existing"],
                "inserted_empty_paragraphs": missing,
            })
            if missing:
                insertion_jobs.append((segment["end"], missing))
        _set_existing_answer_blanks(blocks, segments, document)

    # Insert from the bottom upwards so the original block indices remain valid.
    for end_index, missing in sorted(insertion_jobs, key=lambda item: item[0], reverse=True):
        anchor = blocks[end_index]["element"]
        answer_style = _style_for(document, "answer_blank")
        for _ in range(missing):
            _insert_empty_paragraph_after(anchor, answer_style)

    # Rebuild after insertion and apply one keep-with-next chain to each full question.
    blocks, groups = build_question_groups(document)
    _apply_group_pagination(document, blocks, groups)

    temp = docx_path.with_name(docx_path.stem + ".question-layout-tmp.docx")
    document.save(str(temp))
    temp.replace(docx_path)

    return {
        "schema": SCHEMA,
        "candidate": str(docx_path),
        "candidate_sha256": sha256(docx_path),
        "question_group_count": len(groups),
        "answer_space_segments": segment_reports,
        "inserted_empty_paragraphs": sum(
            item["inserted_empty_paragraphs"] for item in segment_reports
        ),
        "pagination": {
            "keep_together": True,
            "keep_with_next_chain": True,
            "manual_page_breaks_added": 0,
            "page_break_before_added": 0,
        },
        "removed_source_page_controls": removed_breaks,
        "answer_space_format": "genuinely empty Word paragraphs (Enter); no borders, tabs, underscores, manual line breaks, tables, or spacing tricks",
    }


def _effective_indent_twips(paragraph, name):
    value = getattr(paragraph.paragraph_format, name)
    if value is None and paragraph.style is not None:
        value = getattr(paragraph.style.paragraph_format, name)
    return int(round(value.twips)) if value is not None else 0


def check_question_layout(docx_path: Path, profile: dict | None = None) -> dict:
    """Check answer-space counts, real empty paragraphs, flush-left labels and page chains."""
    docx_path = Path(docx_path).resolve(strict=True)
    profile = profile or {}
    policy = {
        **default_answer_space_policy(),
        **(profile.get("answer_space_policy") or {}),
    }
    document = Document(str(docx_path))
    blocks, groups = build_question_groups(document)
    findings = []
    total_required = 0
    total_existing = 0

    if not groups:
        findings.append({
            "code": "no_question_groups_detected",
            "detail": "No numbered questions or example/exercise stems were detected; manually classify the source before delivery.",
        })

    for group_number, group in enumerate(groups, start=1):
        question_paragraph = blocks[group["question_index"]]["paragraph"]
        if question_paragraph is None:
            findings.append({"code": "question_start_not_paragraph", "group": group_number})
        else:
            if question_paragraph.style.name != STYLE_NAMES["question_start"]:
                findings.append({
                    "code": "question_start_style",
                    "group": group_number,
                    "actual": question_paragraph.style.name,
                    "expected": STYLE_NAMES["question_start"],
                })
            if (_effective_indent_twips(question_paragraph, "left_indent") != 0
                    or _effective_indent_twips(question_paragraph, "first_line_indent") != 0):
                findings.append({"code": "question_start_not_flush_left", "group": group_number})

        segments = _answer_segments(blocks, group, policy)
        for segment in segments:
            total_required += segment["required"]
            total_existing += segment["existing"]
            if segment["existing"] < segment["required"]:
                findings.append({
                    "code": "insufficient_answer_space",
                    "group": group_number,
                    "segment_kind": segment["kind"],
                    "required": segment["required"],
                    "existing": segment["existing"],
                })
            start, end = segment["start"], segment["end"]
            trailing = end
            while trailing >= start:
                paragraph = blocks[trailing]["paragraph"]
                if paragraph is None or not _is_empty_answer_paragraph(paragraph):
                    break
                if paragraph.style.name != STYLE_NAMES["answer_blank"]:
                    findings.append({
                        "code": "answer_space_not_enter_paragraph",
                        "group": group_number,
                        "block": trailing,
                        "style": paragraph.style.name,
                    })
                p_pr = paragraph._p.pPr
                if p_pr is not None:
                    forbidden = [tag for tag in ("pBdr", "tabs", "ind", "spacing", "shd", "framePr", "jc")
                                 if p_pr.find(qn("w:" + tag)) is not None]
                    if forbidden:
                        findings.append({
                            "code": "answer_space_visual_formatting",
                            "group": group_number,
                            "block": trailing,
                            "properties": forbidden,
                        })
                if any(child.tag != qn("w:pPr") for child in paragraph._p):
                    findings.append({
                        "code": "answer_space_not_empty_enter_paragraph",
                        "group": group_number,
                        "block": trailing,
                    })
                trailing -= 1

        for index in range(group["start"], group["end"] + 1):
            paragraph = blocks[index]["paragraph"]
            if paragraph is None:
                continue
            expected_keep_next = index < group["end"]
            if paragraph.paragraph_format.keep_together is not True:
                findings.append({"code": "question_paragraph_not_keep_together", "group": group_number, "block": index})
            if paragraph.paragraph_format.keep_with_next is not expected_keep_next:
                findings.append({
                    "code": "question_pagination_chain",
                    "group": group_number,
                    "block": index,
                    "expected_keep_with_next": expected_keep_next,
                    "actual_keep_with_next": paragraph.paragraph_format.keep_with_next,
                })
            if paragraph.paragraph_format.page_break_before is True:
                findings.append({"code": "page_break_before_present", "group": group_number, "block": index})

    manual_breaks = [
        node for node in document.element.body.iter(qn("w:br"))
        if node.get(qn("w:type")) == "page"
    ]
    if manual_breaks:
        findings.append({"code": "manual_page_break_present", "count": len(manual_breaks)})

    return {
        "schema": SCHEMA,
        "candidate": str(docx_path),
        "candidate_sha256": sha256(docx_path),
        "question_group_count": len(groups),
        "required_empty_paragraphs": total_required,
        "existing_empty_paragraphs": total_existing,
        "manual_page_breaks": len(manual_breaks),
        "findings": findings,
        "status": "ready_for_word_pagination_check" if not findings else "needs_layout_work",
    }


def check_word_pagination(docx_path: Path) -> dict:
    """Use real Word pagination to detect any question block that spans pages."""
    try:
        import win32com.client
    except ImportError as exc:
        raise RuntimeError("Word page check requires Windows Microsoft Word COM") from exc

    word = win32com.client.DispatchEx("Word.Application")
    word.Visible = False
    word.DisplayAlerts = 0
    document = None
    try:
        document = word.Documents.Open(
            str(Path(docx_path).resolve()),
            ReadOnly=True,
            AddToRecentFiles=False,
            Visible=False,
        )
        paragraphs = []
        for index in range(1, document.Paragraphs.Count + 1):
            paragraph = document.Paragraphs(index)
            try:
                style_name = paragraph.Range.Style.NameLocal
            except Exception:
                style_name = str(paragraph.Range.Style)
            text = paragraph.Range.Text.replace("\r", "").replace("\x07", "").strip()
            page = int(paragraph.Range.Information(3))
            paragraphs.append({
                "index": index, "style": style_name, "text": text, "page": page,
            })

        starts = [
            index for index, item in enumerate(paragraphs)
            if item["style"] == STYLE_NAMES["question_start"]
        ]
        splits = []
        page_first = {}
        for index in range(1, document.Paragraphs.Count + 1):
            paragraph = document.Paragraphs(index)
            page = int(paragraph.Range.Information(3))
            text = paragraph.Range.Text.replace("\r", "").replace("\x07", "").strip()
            has_visual = paragraph.Range.InlineShapes.Count > 0 or paragraph.Range.Tables.Count > 0
            if page not in page_first:
                page_first[page] = {
                    "paragraph": index,
                    "text": text[:100],
                    "is_empty": not text and not has_visual,
                    "style": str(paragraph.Range.Style.NameLocal) if hasattr(paragraph.Range.Style, "NameLocal") else str(paragraph.Range.Style),
                }
        for page, first in page_first.items():
            if first["is_empty"]:
                splits.append({
                    "reason": "page_starts_with_empty_paragraph",
                    "page": page,
                    "paragraph": first["paragraph"],
                })
        for group_number, start in enumerate(starts, start=1):
            end = starts[group_number] if group_number < len(starts) else len(paragraphs)
            for index in range(start + 1, end):
                item = paragraphs[index]
                if item["style"] == STYLE_NAMES["heading"]:
                    end = index
                    break
            group_rows = paragraphs[start:end]
            pages = sorted({row["page"] for row in group_rows})
            start_paragraph = document.Paragraphs(paragraphs[start]["index"])
            if abs(float(start_paragraph.Format.LeftIndent)) > 0.1 or abs(float(start_paragraph.Format.FirstLineIndent)) > 0.1:
                splits.append({
                    "group": group_number,
                    "reason": "question_start_not_flush_left",
                    "paragraph": paragraphs[start]["index"],
                    "left_indent": float(start_paragraph.Format.LeftIndent),
                    "first_line_indent": float(start_paragraph.Format.FirstLineIndent),
                })
            if len(pages) > 1:
                splits.append({
                    "group": group_number,
                    "start_paragraph": paragraphs[start]["index"],
                    "start_text": paragraphs[start]["text"][:100],
                    "pages": pages,
                })
            if start > 0:
                previous = paragraphs[start - 1]
                if previous["style"] == STYLE_NAMES["heading"] and previous["page"] != paragraphs[start]["page"]:
                    splits.append({
                        "group": group_number,
                        "start_paragraph": previous["index"],
                        "start_text": previous["text"][:100],
                        "pages": sorted({previous["page"], paragraphs[start]["page"]}),
                        "reason": "question_heading_separated_from_question",
                    })
        return {
            "schema": SCHEMA,
            "candidate": str(Path(docx_path).resolve()),
            "word_page_count": int(document.ComputeStatistics(2)),
            "question_group_count": len(starts),
            "split_question_groups": splits,
            "status": "accepted" if starts and not splits else "needs_layout_work",
        }
    finally:
        if document is not None:
            document.Close(False)
        word.Quit()


def main(argv=None):
    parser = argparse.ArgumentParser(description="Plan Enter-created answer space and keep complete questions together.")
    sub = parser.add_subparsers(dest="action", required=True)
    apply_parser = sub.add_parser("apply")
    apply_parser.add_argument("--docx", type=Path, required=True)
    apply_parser.add_argument("--profile", type=Path)
    check_parser = sub.add_parser("check")
    check_parser.add_argument("--docx", type=Path, required=True)
    check_parser.add_argument("--profile", type=Path)
    check_parser.add_argument("--word-pages", action="store_true")
    args = parser.parse_args(argv)
    profile = json.loads(args.profile.read_text("utf-8")) if args.profile else {}
    if args.action == "apply":
        report = apply_question_layout(args.docx, profile)
        print(json.dumps(report, ensure_ascii=False, indent=2))
        return 0
    report = check_question_layout(args.docx, profile)
    if args.word_pages and not report["findings"]:
        report["word_pagination"] = check_word_pagination(args.docx)
        if report["word_pagination"]["status"] != "accepted":
            report["status"] = "needs_layout_work"
    print(json.dumps(report, ensure_ascii=False, indent=2))
    return 0 if report["status"] in {"ready_for_word_pagination_check", "accepted"} else 2


if __name__ == "__main__":
    raise SystemExit(main())
