#!/usr/bin/env python3
"""Move complete questions using only real empty paragraphs and Word pagination."""
from __future__ import annotations

import os
import shutil
import tempfile
import uuid
from pathlib import Path

from docx import Document
from docx.enum.style import WD_STYLE_TYPE
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Pt
from docx.text.paragraph import Paragraph
from lxml import etree

SPACER_STYLE = "SIJI Enter Page Spacer"
PARAGRAPH_CONTROLS = (
    "keepNext", "keepLines", "pageBreakBefore", "widowControl",
    "suppressLineNumbers", "suppressAutoHyphens", "wordWrap", "kinsoku",
    "overflowPunct", "topLinePunct",
)
MAX_SPACERS_PER_GROUP = 300


def _roots(document):
    for part in document.part.package.parts:
        root = getattr(part, "_element", None)
        if root is not None:
            yield str(part.partname), root, part, False
        elif str(part.partname).startswith("/word/") and str(part.partname).endswith(".xml"):
            root = etree.fromstring(part.blob, parser=etree.XMLParser(resolve_entities=False, no_network=True))
            yield str(part.partname), root, part, True


def _inherited_pprs(document, style):
    visited = set()
    while style is not None and style.style_id not in visited:
        visited.add(style.style_id)
        if style._element.pPr is not None:
            yield style._element.pPr
        style = style.base_style
    defaults = document.styles.element.find(qn("w:docDefaults"))
    if defaults is not None:
        default_ppr = defaults.find("./" + qn("w:pPrDefault") + "/" + qn("w:pPr"))
        if default_ppr is not None:
            yield default_ppr


def _effective_space(document, style, side):
    for ppr in _inherited_pprs(document, style):
        spacing = ppr.find(qn("w:spacing"))
        if spacing is None:
            continue
        auto = spacing.get(qn("w:" + side + "Autospacing"))
        if auto is not None and auto.lower() not in {"0", "false", "off"}:
            return 1
        for name in (side + "Lines", side):
            value = spacing.get(qn("w:" + name))
            if value is not None:
                return int(value)
    return 0


def _normal_line_spacing(document, style):
    for ppr in _inherited_pprs(document, style):
        spacing = ppr.find(qn("w:spacing"))
        if spacing is not None and spacing.get(qn("w:line")) is not None:
            rule = spacing.get(qn("w:lineRule"), "auto")
            return rule == "auto" and 180 <= int(spacing.get(qn("w:line"))) <= 360
    return True


def _unsupported_content(document):
    body = document.element.body
    allowed = {qn("w:" + name) for name in (
        "p", "tbl", "sectPr", "bookmarkStart", "bookmarkEnd", "proofErr",
        "permStart", "permEnd", "commentRangeStart", "commentRangeEnd",
    )}
    unknown = sorted({etree.QName(child).localname for child in body if child.tag not in allowed})
    findings = []
    if unknown:
        findings.append({"code": "unsupported_top_level_content", "elements": unknown})
    if sum(1 for _ in body.iter(qn("w:sectPr"))) > 1:
        findings.append({"code": "unsupported_section_break"})
    if any(True for _ in body.iter(qn("wp:anchor"))):
        findings.append({"code": "floating_drawing_requires_manual_layout"})
    return findings


def find_forbidden_controls(document):
    """Inspect direct and style-inherited paragraph controls and manual breaks."""
    findings = _unsupported_content(document)
    for partname, root, _, _ in _roots(document):
        counts = {}
        for ppr in root.iter(qn("w:pPr")):
            for name in PARAGRAPH_CONTROLS:
                count = len(ppr.findall(qn("w:" + name)))
                if count:
                    counts[name] = counts.get(name, 0) + count
        for name in ("cantSplit", "br", "cr"):
            count = sum(1 for _ in root.iter(qn("w:" + name)))
            if count:
                counts[name] = count
        if counts:
            findings.append({"code": "forbidden_pagination_or_line_control", "part": partname, "controls": counts})
    from docx_styling import iter_paragraphs
    unsafe_styles = set()
    invalid_spacers = 0
    for paragraph in iter_paragraphs(document):
        if paragraph.style.name not in {SPACER_STYLE, "SIJI AnswerBlank"}:
            continue
        if paragraph.style.name == SPACER_STYLE and any(child.tag != qn("w:pPr") for child in paragraph._p):
            invalid_spacers += 1
        paragraph_style = paragraph.style
        style = paragraph_style
        visited = set()
        while style is not None and style.style_id not in visited:
            visited.add(style.style_id)
            ppr = style._element.pPr
            bad = [name for name in ("pBdr", "tabs", "shd", "framePr")
                   if ppr is not None and ppr.find(qn("w:" + name)) is not None]
            if bad:
                unsafe_styles.add(style.name + ":" + ",".join(bad))
            style = style.base_style
        try:
            answer_style = document.styles["SIJI AnswerBlank"]
        except KeyError:
            unsafe_styles.add("SIJI AnswerBlank missing")
            continue
        if (_effective_space(document, paragraph_style, "before")
                or _effective_space(document, paragraph_style, "after")
                or not _normal_line_spacing(document, paragraph_style)):
            unsafe_styles.add(paragraph_style.name + ":spacing")
    if unsafe_styles:
        findings.append({"code": "answer_or_spacer_style_visual_formatting", "styles": sorted(unsafe_styles)})
    if invalid_spacers:
        findings.append({"code": "page_spacer_not_empty_paragraph", "count": invalid_spacers})
    return findings


def _sanitize(document):
    """Remove source paging properties; never discard meaningful line breaks."""
    from question_layout import _is_empty_answer_paragraph

    unsupported = _unsupported_content(document)
    if unsupported:
        raise RuntimeError(f"unsupported_document_structure: {unsupported}")
    removed = {}
    for partname, root, part, raw in _roots(document):
        for node in list(root.iter(qn("w:br"))) + list(root.iter(qn("w:cr"))):
            if node.tag != qn("w:br") or node.get(qn("w:type")) != "page":
                raise RuntimeError(f"unsupported_line_or_column_break: {partname}; convert it to real paragraphs first")
            paragraph = node
            while paragraph is not None and paragraph.tag != qn("w:p"):
                paragraph = paragraph.getparent()
            node.getparent().remove(node)
            removed["manual_page_breaks"] = removed.get("manual_page_breaks", 0) + 1
            disposable = paragraph is not None and all(
                child.tag == qn("w:pPr") or (child.tag == qn("w:r") and all(
                    sub.tag == qn("w:rPr") for sub in child)) for child in paragraph)
            if (disposable and paragraph.getparent() is not None
                    and paragraph.getparent().tag == qn("w:body")
                    and _is_empty_answer_paragraph(Paragraph(paragraph, document))):
                paragraph.getparent().remove(paragraph)
                removed["empty_break_paragraphs"] = removed.get("empty_break_paragraphs", 0) + 1
        for ppr in root.iter(qn("w:pPr")):
            for name in PARAGRAPH_CONTROLS:
                for node in list(ppr.findall(qn("w:" + name))):
                    ppr.remove(node)
                    removed[name] = removed.get(name, 0) + 1
        for name in ("cantSplit", "lastRenderedPageBreak"):
            for node in list(root.iter(qn("w:" + name))):
                node.getparent().remove(node)
                removed[name] = removed.get(name, 0) + 1
        if raw:
            part._blob = etree.tostring(root, encoding="UTF-8", xml_declaration=True, standalone=True)
    return removed


def _remove_generated_spacers(document):
    from question_layout import _is_empty_answer_paragraph
    removed = 0
    for paragraph in list(document.paragraphs):
        if paragraph.style.name != SPACER_STYLE:
            continue
        if not _is_empty_answer_paragraph(paragraph) or any(
                node.tag != qn("w:pPr") for node in paragraph._p):
            raise RuntimeError("page_spacer_contains_content: refusing to delete it")
        paragraph._p.getparent().remove(paragraph._p)
        removed += 1
    return removed


def _spacer_style(document):
    from question_layout import _style_for
    answer = _style_for(document, "answer_blank")
    ppr = answer._element.get_or_add_pPr()
    for name in ("pBdr", "tabs", "ind", "spacing", "shd", "framePr", "jc"):
        for node in list(ppr.findall(qn("w:" + name))):
            ppr.remove(node)
    answer.paragraph_format.space_before = Pt(0)
    answer.paragraph_format.space_after = Pt(0)
    answer.paragraph_format.line_spacing = 1.15
    try:
        style = document.styles[SPACER_STYLE]
    except KeyError:
        style = document.styles.add_style(SPACER_STYLE, WD_STYLE_TYPE.PARAGRAPH)
    style.base_style = answer
    for name in ("pPr", "rPr"):
        node = style._element.find(qn("w:" + name))
        if node is not None:
            style._element.remove(node)
    return style


def sanitize_working_copy(path):
    """Sanitize a disposable working copy before any style preparation."""
    path = Path(path).resolve(strict=True)
    document = Document(str(path))
    removed_spacers = _remove_generated_spacers(document)
    removed = _sanitize(document)
    document.save(str(path))
    return {"removed_source_controls": removed, "removed_generated_spacers": removed_spacers}


def _bookmarked_copy(path, destination):
    """Add temporary position markers for each group's first and last paragraph."""
    from question_layout import build_question_groups
    doc = Document(str(path))
    blocks, groups = build_question_groups(doc)
    if not groups:
        raise RuntimeError("no_question_groups_detected: classify question starts before pagination")
    ids = [int(n.get(qn("w:id"))) for n in doc.element.iter(qn("w:bookmarkStart"))
           if (n.get(qn("w:id")) or "").isdigit()]
    next_id = max(ids, default=0) + 1
    prefix = "cx" + uuid.uuid4().hex[:8]
    anchors = []
    previous_end = -1

    def add_mark(paragraph, name, end=False):
        nonlocal next_id
        start = OxmlElement("w:bookmarkStart")
        start.set(qn("w:id"), str(next_id))
        start.set(qn("w:name"), name)
        finish = OxmlElement("w:bookmarkEnd")
        finish.set(qn("w:id"), str(next_id))
        next_id += 1
        if end:
            paragraph.append(start)
            paragraph.append(finish)
        else:
            offset = 1 if paragraph.pPr is not None else 0
            paragraph.insert(offset, start)
            paragraph.insert(offset + 1, finish)

    for number, group in enumerate(groups, 1):
        first = group["start"]
        while (first - 1 > previous_end and blocks[first - 1]["kind"] in {"title", "heading"}):
            first -= 1
        first_el = blocks[first]["element"]
        last_el = blocks[group["end"]]["element"]
        first_p = first_el if first_el.tag == qn("w:p") else next(first_el.iter(qn("w:p")), None)
        last_p = last_el if last_el.tag == qn("w:p") else next(reversed(list(last_el.iter(qn("w:p")))), None)
        if first_p is None or last_p is None:
            raise RuntimeError(f"group_has_no_paragraph: {number}")
        names = (f"{prefix}s{number}", f"{prefix}e{number}")
        add_mark(first_p, names[0])
        add_mark(last_p, names[1], end=True)
        anchors.append({"group": number, "start": names[0], "end": names[1],
                        "text": blocks[group["question_index"]]["text"][:100]})
        previous_end = group["end"]
    doc.save(str(destination))
    return prefix, anchors


def _remove_marks(document, prefix):
    ids = set()
    for node in list(document.element.iter(qn("w:bookmarkStart"))):
        if (node.get(qn("w:name")) or "").startswith(prefix):
            ids.add(node.get(qn("w:id")))
            node.getparent().remove(node)
    for node in list(document.element.iter(qn("w:bookmarkEnd"))):
        if node.get(qn("w:id")) in ids:
            node.getparent().remove(node)


def _relocate_bookmarks_from_special_paragraphs(document):
    """Keep original bookmark markers out of paragraphs reserved for empty space."""
    from question_layout import STYLE_NAMES

    special_styles = {SPACER_STYLE, STYLE_NAMES["answer_blank"]}
    marker_tags = {qn("w:bookmarkStart"), qn("w:bookmarkEnd")}
    paragraphs = list(document.paragraphs)
    special_indexes = [i for i, paragraph in enumerate(paragraphs)
                       if paragraph.style.name in special_styles]
    moved = 0

    for index in reversed(special_indexes):
        source = paragraphs[index]._p
        markers = [child for child in list(source) if child.tag in marker_tags]
        if not markers:
            continue
        target = next((p for p in paragraphs[index + 1:]
                       if p.style.name not in special_styles), None)
        if target is None:
            target = next((p for p in reversed(paragraphs[:index])
                           if p.style.name not in special_styles), None)
        if target is None:
            raise RuntimeError("bookmark_in_empty_paragraph_without_neighbor")

        for marker in markers:
            source.remove(marker)
        target_p = target._p
        offset = 1 if target_p.pPr is not None else 0
        for marker in markers:
            target_p.insert(offset, marker)
            offset += 1
        moved += len(markers)
    return moved


def _word():
    try:
        import win32com.client
        word = win32com.client.DispatchEx("Word.Application")
    except Exception as exc:
        raise RuntimeError("Microsoft Word COM is required for pure-Enter pagination") from exc
    word.Visible = False
    word.DisplayAlerts = 0
    return word


def _group_pages(doc, anchor, extra=0, start_pos=None):
    start = (int(doc.Bookmarks(anchor["start"]).Range.Start) if start_pos is None else start_pos) + extra
    end = int(doc.Bookmarks(anchor["end"]).Range.End)
    return int(doc.Range(start, start).Information(3)), int(doc.Range(end, end).Information(3))


def _first_paragraph_findings(doc):
    first = {}
    for index in range(1, doc.Paragraphs.Count + 1):
        para = doc.Paragraphs(index)
        rng = para.Range
        page = int(rng.Information(3))
        if page in first:
            continue
        text = rng.Text.replace("\r", "").replace("\x07", "").strip()
        visual = rng.InlineShapes.Count > 0 or rng.Tables.Count > 0
        try:
            visual = visual or rng.ShapeRange.Count > 0
        except Exception:
            pass
        first[page] = {"text": text, "visual": visual, "paragraph": index}
    return [{"reason": "page_starts_with_empty_paragraph", "page": page,
             "paragraph": row["paragraph"]}
            for page, row in first.items() if not row["text"] and not row["visual"]]


def _check_open_word(doc, anchors):
    doc.Repaginate()
    findings = _first_paragraph_findings(doc)
    for anchor in anchors:
        start, end = _group_pages(doc, anchor)
        if start != end:
            findings.append({"group": anchor["group"], "reason": "question_spans_pages",
                             "pages": [start, end], "start_text": anchor["text"]})
        pos = int(doc.Bookmarks(anchor["start"]).Range.Start)
        para = doc.Range(pos, pos).Paragraphs(1)
        if abs(float(para.Format.LeftIndent)) > 0.1 or abs(float(para.Format.FirstLineIndent)) > 0.1:
            findings.append({"group": anchor["group"], "reason": "group_start_not_flush_left"})
    return {"word_page_count": int(doc.ComputeStatistics(2)),
            "question_group_count": len(anchors), "split_question_groups": findings,
            "status": "accepted" if anchors and not findings else "needs_layout_work"}


def check_enter_word_pagination(path):
    path = Path(path).resolve(strict=True)
    with tempfile.TemporaryDirectory(prefix="codex-enter-check-") as folder:
        marked = Path(folder) / "marked.docx"
        _, anchors = _bookmarked_copy(path, marked)
        word = _word()
        doc = None
        try:
            doc = word.Documents.Open(str(marked), ReadOnly=True, AddToRecentFiles=False, Visible=False)
            result = _check_open_word(doc, anchors)
            result.update({"schema": "question-layout-plan/1", "candidate": str(path), "pagination_mode": "enter"})
            return result
        finally:
            if doc is not None:
                doc.Close(False)
            word.Quit()


def apply_enter_layout(path, profile=None):
    """Work on a copy; replace the input only after actual Word and XML checks pass."""
    from question_layout import apply_question_layout, check_question_layout, sha256
    path = Path(path).resolve(strict=True)
    profile = profile or {}
    with tempfile.TemporaryDirectory(prefix="codex-enter-layout-") as folder:
        working = Path(folder) / "working.docx"
        marked = Path(folder) / "marked.docx"
        shutil.copy2(path, working)
        model = Document(str(working))
        old_spacers = _remove_generated_spacers(model)
        removed = _sanitize(model)
        _spacer_style(model)
        model.save(str(working))
        report = apply_question_layout(working, profile, "_enter_prepare")
        prefix, anchors = _bookmarked_copy(working, marked)
        word = _word()
        doc = None
        inserted = 0
        try:
            doc = word.Documents.Open(str(marked), ReadOnly=False, AddToRecentFiles=False, Visible=False)
            doc.Repaginate()
            for anchor in anchors:
                start_page, end_page = _group_pages(doc, anchor)
                if start_page == end_page:
                    continue
                pos = int(doc.Bookmarks(anchor["start"]).Range.Start)
                if pos == 0:
                    raise RuntimeError(f"question_too_tall: group {anchor['group']} ({anchor['text']})")
                current = 0

                def set_spacers(count):
                    nonlocal current
                    if current:
                        doc.Range(pos, pos + current).Delete()
                    if count:
                        doc.Range(pos, pos).InsertBefore("\r" * count)
                        doc.Range(pos, pos + count).Style = doc.Styles(SPACER_STYLE)
                    current = count
                    doc.Repaginate()
                    return _group_pages(doc, anchor, current, pos)

                low, high = 0, 1
                while True:
                    moved_start, _ = set_spacers(high)
                    if moved_start > start_page:
                        break
                    low = high
                    if high >= MAX_SPACERS_PER_GROUP:
                        raise RuntimeError(f"cannot_move_question_with_enter: group {anchor['group']}")
                    high = min(high * 2, MAX_SPACERS_PER_GROUP)
                while high - low > 1:
                    mid = (low + high) // 2
                    moved_start, _ = set_spacers(mid)
                    if moved_start > start_page:
                        high = mid
                    else:
                        low = mid
                moved_start, moved_end = set_spacers(high)
                if moved_start != start_page + 1 or moved_start != moved_end:
                    raise RuntimeError(f"question_too_tall: group {anchor['group']} ({anchor['text']})")
                inserted += high
            doc.Save()
        finally:
            if doc is not None:
                doc.Close(False)
            word.Quit()

        model = Document(str(marked))
        _remove_marks(model, prefix)
        relocated_bookmark_markers = _relocate_bookmarks_from_special_paragraphs(model)
        model.save(str(working))
        structure = check_question_layout(working, profile, "enter")
        if structure["status"] != "ready_for_word_pagination_check":
            raise RuntimeError(f"enter_structure_check_failed: {structure['findings']}")
        actual = check_enter_word_pagination(working)
        if actual["status"] != "accepted":
            raise RuntimeError(f"enter_word_pagination_failed: {actual['split_question_groups']}")

        temp_name = None
        try:
            with tempfile.NamedTemporaryFile(prefix=path.stem + ".enter-", suffix=".docx",
                                             dir=str(path.parent), delete=False) as stream:
                temp_name = stream.name
            shutil.copyfile(working, temp_name)
            os.replace(temp_name, path)
            temp_name = None
        finally:
            if temp_name and os.path.exists(temp_name):
                os.unlink(temp_name)
        report.update({"candidate": str(path), "candidate_sha256": sha256(path),
                       "pagination": {"mode": "enter", "inserted_page_spacers": inserted,
                                      "replaced_prior_page_spacers": old_spacers,
                                      "relocated_bookmark_markers": relocated_bookmark_markers},
                       "removed_source_page_controls": removed,
                       "structural_check": structure, "word_pagination": actual})
        report["structural_check"]["candidate"] = str(path)
        report["word_pagination"]["candidate"] = str(path)
        return report
