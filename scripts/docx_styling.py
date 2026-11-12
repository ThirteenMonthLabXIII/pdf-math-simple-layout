#!/usr/bin/env python3
"""Preventive, deterministic Word paragraph styling for math handouts.

The source DOCX is retained separately. This module applies a controlled style
contract before Codex edits the working copy; it does not rewrite paragraph text,
spaces, equation XML, numbering, or figures.
"""
from __future__ import annotations

import argparse
import copy
import hashlib
import json
import re
import statistics
from collections import Counter, defaultdict
from pathlib import Path

from docx import Document
from docx.enum.style import WD_STYLE_TYPE
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Pt, Twips

MATH_TAGS = {
    "{http://schemas.openxmlformats.org/officeDocument/2006/math}oMath",
    "{http://schemas.openxmlformats.org/officeDocument/2006/math}oMathPara",
}
DRAWING_TAGS = {
    "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}drawing",
    "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}pict",
}
STYLE_NAMES = {
    "body": "SIJI Body",
    "title": "SIJI Title",
    "heading": "SIJI Heading",
    "question": "SIJI Question",
    "question_start": "SIJI Question Start",
    "subquestion": "SIJI Subquestion",
    "formula": "SIJI Formula",
    "answer_blank": "SIJI AnswerBlank",
    "caption": "SIJI Caption",
    "figure": "SIJI Figure",
}
ALIGNMENTS = {
    "left": WD_ALIGN_PARAGRAPH.LEFT,
    "center": WD_ALIGN_PARAGRAPH.CENTER,
    "right": WD_ALIGN_PARAGRAPH.RIGHT,
    "justify": WD_ALIGN_PARAGRAPH.JUSTIFY,
}
ALIGNMENT_NAMES = {
    WD_ALIGN_PARAGRAPH.LEFT: "left",
    WD_ALIGN_PARAGRAPH.CENTER: "center",
    WD_ALIGN_PARAGRAPH.RIGHT: "right",
    WD_ALIGN_PARAGRAPH.JUSTIFY: "justify",
}
CAPTION_RE = re.compile(r"^图\s*\d+")
TITLE_RE = re.compile(r"^第\s*[一二三四五六七八九十百千万\d]+\s*讲")
HEADING_RE = re.compile(r"^(?:【?(?:例题|练习|拓展|知识点|课后练习|课后作业|巩固练习|综合练习)|知识点[一二三四五六七八九十0-9])")
SUBQUESTION_RE = re.compile(r"^\s*[（(]\s*[1-9][0-9]*\s*[）)]")
QUESTION_RE = re.compile(r"^\s*\d+\s*[．.、)）]\s*(?=\S)")
ANSWER_ONLY_RE = re.compile(r"^(?:答案|答)\s*[:：]?\s*[（(][ \t\u3000]*[）)]$")
MAX_SAMPLE_INDENT_TWIPS = 1000


def default_paragraph_policy() -> dict:
    """Safe defaults: zero body/formula indent; one controlled subquestion level."""
    common = {
        "left_indent_twips": 0,
        "right_indent_twips": 0,
        "first_line_indent_twips": 0,
        "space_before_pt": 0,
        "line_spacing": 1.15,
        "alignment": "left",
    }
    def style(after, **extra):
        value = {**common, "space_after_pt": after}
        value.update(extra)
        return value
    return {
        "body": style(7),
        "title": style(10, alignment="center", font_size_key="title_pt", bold=True),
        "heading": style(4, space_before_pt=6, font_size_key="heading_pt", bold=True),
        "question": style(7),
        "question_start": style(7),
        "subquestion": style(10, left_indent_twips=600),
        "formula": style(7, alignment="left"),
        "answer_blank": style(0),
        "caption": style(4, alignment="center"),
        "figure": style(4, alignment="center"),
    }


def _sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with Path(path).open("rb") as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def _has_math(paragraph) -> bool:
    return any(node.tag in MATH_TAGS for node in paragraph._p.iter())


def classify_paragraph(paragraph) -> str:
    """Classify without mutating or normalizing any source text."""
    text = (paragraph.text or "").strip()
    style_name = paragraph.style.name if paragraph.style is not None else ""
    if style_name == STYLE_NAMES["question_start"]:
        return "question_start"
    if style_name == STYLE_NAMES["answer_blank"] and not text:
        return "answer_blank"
    if any(node.tag in DRAWING_TAGS for node in paragraph._p.iter()):
        return "figure"
    if paragraph.style is not None and "caption" in paragraph.style.name.lower():
        return "caption"
    if CAPTION_RE.match(text):
        return "caption"
    if _has_math(paragraph) and not text:
        return "formula"
    if TITLE_RE.match(text):
        return "title"
    if HEADING_RE.match(text):
        return "heading"
    if SUBQUESTION_RE.match(text):
        return "subquestion"
    if QUESTION_RE.match(text):
        return "question"
    if ANSWER_ONLY_RE.match(text):
        return "answer_blank"
    return "body"


def iter_paragraphs(document):
    """Yield paragraphs in body, tables (including nested), headers and footers once."""
    seen = set()

    def emit(paragraph):
        marker = id(paragraph._p)
        if marker not in seen:
            seen.add(marker)
            return paragraph
        return None

    def walk_table(table):
        for row in table.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    value = emit(paragraph)
                    if value is not None:
                        yield value
                for nested in cell.tables:
                    yield from walk_table(nested)

    for paragraph in document.paragraphs:
        value = emit(paragraph)
        if value is not None:
            yield value
    for table in document.tables:
        yield from walk_table(table)
    for section in document.sections:
        for container in (section.header, section.footer, section.first_page_header,
                          section.first_page_footer, section.even_page_header, section.even_page_footer):
            for paragraph in container.paragraphs:
                value = emit(paragraph)
                if value is not None:
                    yield value
            for table in container.tables:
                yield from walk_table(table)


def _effective_paragraph_value(paragraph, name):
    value = getattr(paragraph.paragraph_format, name)
    if value is None and paragraph.style is not None:
        value = getattr(paragraph.style.paragraph_format, name)
    return value


def extract_paragraph_policy(sample_docx: Path, base_policy: dict | None = None) -> tuple[dict, list[str]]:
    """Learn per-kind paragraph spacing/alignment from a sample, rejecting extreme indents."""
    policy = copy.deepcopy(base_policy or default_paragraph_policy())
    groups = defaultdict(list)
    warnings = []
    sample = Document(str(sample_docx))

    for index, paragraph in enumerate(iter_paragraphs(sample), start=1):
        kind = classify_paragraph(paragraph)
        if kind not in policy:
            continue
        if not (paragraph.text or "").strip() and kind != "formula":
            continue

        values = {}
        unsafe = False
        for key, attr in (
            ("left_indent_twips", "left_indent"),
            ("right_indent_twips", "right_indent"),
            ("first_line_indent_twips", "first_line_indent"),
        ):
            length = _effective_paragraph_value(paragraph, attr)
            twips = int(round(length.twips)) if length is not None else 0
            values[key] = twips
            if abs(twips) > MAX_SAMPLE_INDENT_TWIPS:
                unsafe = True
        if unsafe:
            warnings.append(
                f"sample paragraph {index} ({kind}) has an indent over "
                f"{MAX_SAMPLE_INDENT_TWIPS} twips; its layout was not learned"
            )
            continue

        for key, attr in (("space_before_pt", "space_before"), ("space_after_pt", "space_after")):
            length = _effective_paragraph_value(paragraph, attr)
            values[key] = round(length.pt, 2) if length is not None else 0.0

        line_spacing = _effective_paragraph_value(paragraph, "line_spacing")
        if isinstance(line_spacing, (int, float)):
            values["line_spacing"] = float(line_spacing)

        alignment = _effective_paragraph_value(paragraph, "alignment")
        if alignment in ALIGNMENT_NAMES:
            values["alignment"] = ALIGNMENT_NAMES[alignment]
        groups[kind].append(values)

    for kind, rows in groups.items():
        for key in ("left_indent_twips", "right_indent_twips", "first_line_indent_twips",
                    "space_before_pt", "space_after_pt", "line_spacing"):
            nums = [row[key] for row in rows if key in row]
            if nums:
                policy[kind][key] = round(float(statistics.median(nums)), 2) if key.endswith("_pt") or key == "line_spacing" else int(round(statistics.median(nums)))
        alignments = [row["alignment"] for row in rows if "alignment" in row]
        if alignments:
            policy[kind]["alignment"] = Counter(alignments).most_common(1)[0][0]

    # Formula indentation is never learned from an extreme source paragraph.
    formula = policy.setdefault("formula", {})
    if abs(int(formula.get("left_indent_twips", 0))) > MAX_SAMPLE_INDENT_TWIPS:
        formula["left_indent_twips"] = 0
        warnings.append("unsafe formula indent in sample policy reset to zero")
    return policy, warnings


def _set_east_asia_font(style_or_run, font_name: str, east_asia_font: str):
    element = style_or_run._element
    if hasattr(element, "get_or_add_rPr"):
        rpr = element.get_or_add_rPr()
    else:
        rpr = element.get_or_add_rPr()
    rfonts = rpr.rFonts
    if rfonts is None:
        rfonts = OxmlElement("w:rFonts")
        rpr.insert(0, rfonts)
    rfonts.set(qn("w:ascii"), font_name)
    rfonts.set(qn("w:hAnsi"), font_name)
    rfonts.set(qn("w:eastAsia"), east_asia_font)
    rfonts.set(qn("w:cs"), font_name)


def _get_or_create_style(document, name: str):
    try:
        style = document.styles[name]
    except KeyError:
        style = document.styles.add_style(name, WD_STYLE_TYPE.PARAGRAPH)
        style.base_style = document.styles["Normal"]
    return style


PAGINATION_PROPERTIES = ("keep_with_next", "keep_together", "page_break_before", "widow_control")


def _inherited_pagination(style):
    """Read style-chain pagination values without changing them."""
    values = {}
    visited = set()
    while style is not None and style.style_id not in visited:
        visited.add(style.style_id)
        for name in PAGINATION_PROPERTIES:
            if name not in values:
                value = getattr(style.paragraph_format, name)
                if value is not None:
                    values[name] = value
        style = style.base_style
    return {name: values.get(name) for name in PAGINATION_PROPERTIES}


def _configure_style(style, spec, profile, allow_pagination):
    font_name = profile.get("font_name", "宋体")
    east_asia_font = profile.get("east_asia_font", "宋体")
    body_pt = float(profile.get("body_pt", 10.5))
    font = style.font
    font.name = font_name
    font.size = Pt(float(profile.get(spec.get("font_size_key", ""), body_pt)))
    if "bold" in spec:
        font.bold = bool(spec["bold"])
    _set_east_asia_font(style, font_name, east_asia_font)

    fmt = style.paragraph_format
    fmt.left_indent = Twips(int(spec.get("left_indent_twips", 0)))
    fmt.right_indent = Twips(int(spec.get("right_indent_twips", 0)))
    fmt.first_line_indent = Twips(int(spec.get("first_line_indent_twips", 0)))
    fmt.space_before = Pt(float(spec.get("space_before_pt", 0)))
    fmt.space_after = Pt(float(spec.get("space_after_pt", 7)))
    fmt.line_spacing = float(spec.get("line_spacing", profile.get("line_spacing", 1.15)))
    fmt.alignment = ALIGNMENTS.get(spec.get("alignment", "left"), WD_ALIGN_PARAGRAPH.LEFT)
    if allow_pagination and "keep_with_next" in spec:
        fmt.keep_with_next = bool(spec["keep_with_next"])
    if allow_pagination and "keep_together" in spec:
        fmt.keep_together = bool(spec["keep_together"])


def _apply_style_contract(document, profile: dict) -> dict:
    policy = copy.deepcopy(default_paragraph_policy())
    supplied = profile.get("paragraph_policy") or {}
    allow_pagination = profile.get("allow_pagination_overrides") is True
    for kind, values in supplied.items():
        if kind in policy and isinstance(values, dict):
            policy[kind].update({key: value for key, value in values.items()
                if allow_pagination or key not in {"keep_with_next", "keep_together"}})

    # Question starts are always flush-left, regardless of a teacher sample.
    policy["question_start"]["left_indent_twips"] = 0
    policy["question_start"]["first_line_indent_twips"] = 0
    policy["question_start"]["alignment"] = "left"

    # Answer space must be real empty paragraphs, never spacing inherited from a sample.
    policy["answer_blank"] = default_paragraph_policy()["answer_blank"]

    styles = {}

    for kind, style_name in STYLE_NAMES.items():
        spec = policy[kind]
        style = _get_or_create_style(document, style_name)
        _configure_style(style, spec, profile, allow_pagination)
        styles[kind] = style
    return {"policy": policy, "styles": styles}


def _clear_direct_paragraph_layout(paragraph):
    """Remove source DOCX direct layout overrides so the controlled style is authoritative."""
    fmt = paragraph.paragraph_format
    fmt.left_indent = None
    fmt.right_indent = None
    fmt.first_line_indent = None
    fmt.space_before = None
    fmt.space_after = None
    fmt.line_spacing = None
    paragraph.alignment = None


def _set_run_font(run, profile: dict, kind: str):
    policy = profile.get("paragraph_policy") or default_paragraph_policy()
    spec = policy.get(kind, {})
    size = float(profile.get(spec.get("font_size_key", ""), profile.get("body_pt", 10.5)))
    run.font.name = profile.get("font_name", "宋体")
    run.font.size = Pt(size)
    _set_east_asia_font(run, profile.get("font_name", "宋体"), profile.get("east_asia_font", "宋体"))


def apply_docx_styles(docx_path: Path, profile: dict | None = None, output_path: Path | None = None) -> dict:
    """Apply controlled paragraph styles before editing; text and OfficeMath XML remain untouched."""
    docx_path = Path(docx_path).resolve(strict=True)
    output_path = Path(output_path).resolve() if output_path else docx_path
    profile = copy.deepcopy(profile or {})
    profile.setdefault("paragraph_policy", default_paragraph_policy())
    source_hash = _sha256(docx_path)

    document = Document(str(docx_path))
    contract = _apply_style_contract(document, profile)
    counts = Counter()
    variants = {}
    for paragraph in iter_paragraphs(document):
        kind = classify_paragraph(paragraph)
        if kind not in contract["styles"]:
            kind = "body"
        original_style = paragraph.style
        style = contract["styles"][kind]
        spec = contract["policy"][kind]
        allow_pagination = profile.get("allow_pagination_overrides") is True
        overridden = {name for name in ("keep_with_next", "keep_together")
                      if allow_pagination and name in spec}
        original_settings = _inherited_pagination(original_style)
        styled_settings = _inherited_pagination(style)
        desired = {name: styled_settings[name] if name in overridden else original_settings[name]
                   for name in PAGINATION_PROPERTIES}
        if desired != styled_settings:
            # Keep the source style's effective pagination without paragraph-level overrides.
            signature = json.dumps(desired, sort_keys=True)
            key = (kind, signature)
            if key not in variants:
                suffix = hashlib.sha256(signature.encode("utf-8")).hexdigest()[:12]
                name = STYLE_NAMES[kind] + " inherited " + suffix
                try:
                    variant = document.styles[name]
                except KeyError:
                    variant = document.styles.add_style(name, WD_STYLE_TYPE.PARAGRAPH)
                    for prop, value in desired.items():
                        if value is not None:
                            setattr(variant.paragraph_format, prop, value)
                if _inherited_pagination(variant) != desired:
                    raise ValueError("Existing SIJI inherited style has unexpected pagination")
                _configure_style(variant, spec, profile, allow_pagination)
                variants[key] = variant
            style = variants[key]
        _clear_direct_paragraph_layout(paragraph)
        paragraph.style = style
        # Run-level font normalization makes the paragraph style effective while
        # preserving bold/italic/underline/superscript/subscript and all text.
        for run in paragraph.runs:
            _set_run_font(run, profile, kind)
        counts[kind] += 1

    output_path.parent.mkdir(parents=True, exist_ok=True)
    temp_path = output_path.with_name(output_path.stem + ".style-tmp.docx")
    document.save(str(temp_path))
    temp_path.replace(output_path)
    return {
        "schema": "docx-style-policy/1",
        "source_docx_sha256": source_hash,
        "styled_docx_sha256": _sha256(output_path),
        "paragraph_count": sum(counts.values()),
        "paragraphs_by_style": dict(sorted(counts.items())),
        "policy": contract["policy"],
        "style_names": STYLE_NAMES,
        "text_mutated": False,
        "office_math_xml_preserved_by_design": True,
    }


def main(argv=None):
    parser = argparse.ArgumentParser(description="Apply the preventive SIJI Word style contract.")
    sub = parser.add_subparsers(dest="action", required=True)
    apply_parser = sub.add_parser("apply", help="apply safe paragraph styles to a DOCX copy")
    apply_parser.add_argument("--docx", type=Path, required=True)
    apply_parser.add_argument("--profile", type=Path)
    apply_parser.add_argument("--output", type=Path)
    args = parser.parse_args(argv)
    profile = json.loads(args.profile.read_text(encoding="utf-8")) if args.profile else {}
    report = apply_docx_styles(args.docx, profile, args.output)
    print(json.dumps(report, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
