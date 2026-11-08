#!/usr/bin/env python3
"""Detect suspicious Word paragraph layout before a math handout is accepted.

This is a review gate, not an auto-formatter: findings must be fixed or explicitly
justified against the source PDF. The script never edits a DOCX.
"""
from __future__ import annotations

import argparse
import hashlib
import json
import re
import sys
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
M = "http://schemas.openxmlformats.org/officeDocument/2006/math"
NS = {"w": W, "m": M}
DEFAULT_INDENT_TWIPS = 1000
DEFAULT_SPACING_TWIPS = 720
ANSWER_SPACE_RE = re.compile(r"[（(]([ \t\u3000]{3,})[）)]")


def sha256_file(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def _int_attr(element, name: str):
    if element is None:
        return None
    value = element.get(f"{{{W}}}{name}")
    if value is None:
        return None
    try:
        return int(value)
    except (TypeError, ValueError):
        return None


def _text(paragraph) -> str:
    values = []
    for node in paragraph.iter():
        if node.tag in (f"{{{W}}}t", f"{{{M}}}t") and node.text:
            values.append(node.text)
        elif node.tag == f"{{{W}}}tab":
            values.append("\t")
    return "".join(values)


def _read_style_map(archive):
    if "word/styles.xml" not in archive.namelist():
        return {}, None
    root = ET.fromstring(archive.read("word/styles.xml"))
    styles = {}
    default_paragraph_style = None
    for style in root.findall(f"{{{W}}}style"):
        if style.get(f"{{{W}}}type") != "paragraph":
            continue
        style_id = style.get(f"{{{W}}}styleId")
        if not style_id:
            continue
        ppr = style.find(f"{{{W}}}pPr")
        based_on = style.find(f"{{{W}}}basedOn")
        name = style.find(f"{{{W}}}name")
        styles[style_id] = {
            "ppr": ppr,
            "based_on": based_on.get(f"{{{W}}}val") if based_on is not None else None,
            "name": name.get(f"{{{W}}}val") if name is not None else style_id,
            "default": style.get(f"{{{W}}}default") == "1",
        }
        if styles[style_id]["default"]:
            default_paragraph_style = style_id
    return styles, default_paragraph_style


def _inherited_style_value(style_map, style_id, property_name, attribute):
    visited = set()
    while style_id and style_id in style_map and style_id not in visited:
        visited.add(style_id)
        info = style_map[style_id]
        ppr = info["ppr"]
        element = ppr.find(f"{{{W}}}{property_name}") if ppr is not None else None
        value = _int_attr(element, attribute)
        if value is not None:
            return value, info["name"]
        style_id = info["based_on"]
    return None, None


def inspect_docx(
    docx_path: Path,
    indent_threshold_twips: int = DEFAULT_INDENT_TWIPS,
    spacing_threshold_twips: int = DEFAULT_SPACING_TWIPS,
) -> dict:
    """Return deterministic findings for unusually large paragraph properties."""
    docx_path = Path(docx_path).resolve(strict=True)
    with zipfile.ZipFile(docx_path) as archive:
        if "word/document.xml" not in archive.namelist():
            raise ValueError("DOCX is missing word/document.xml")
        root = ET.fromstring(archive.read("word/document.xml"))
        style_map, default_paragraph_style = _read_style_map(archive)

    paragraphs = list(root.iter(f"{{{W}}}p"))
    findings = []

    def add(index, code, severity, detail, excerpt, **extra):
        finding_id = f"p{index:04d}:{code}:{extra.get('attribute', '')}:{extra.get('value_twips', extra.get('space_count', ''))}"
        findings.append({
            "id": finding_id,
            "paragraph": index,
            "code": code,
            "severity": severity,
            "detail": detail,
            "excerpt": excerpt[:120],
            **extra,
        })

    for index, paragraph in enumerate(paragraphs, start=1):
        text = _text(paragraph)
        excerpt = re.sub(r"\s+", " ", text).strip()
        has_math = any(node.tag in (f"{{{M}}}oMath", f"{{{M}}}oMathPara") for node in paragraph.iter())
        ppr = paragraph.find(f"{{{W}}}pPr")

        ind = ppr.find(f"{{{W}}}ind") if ppr is not None else None
        pstyle = ppr.find(f"{{{W}}}pStyle") if ppr is not None else None
        style_id = pstyle.get(f"{{{W}}}val") if pstyle is not None else default_paragraph_style
        for attribute in ("left", "right", "firstLine", "hanging"):
            value = _int_attr(ind, attribute)
            source = "paragraph"
            style_name = None
            if value is None:
                value, style_name = _inherited_style_value(style_map, style_id, "ind", attribute)
                source = "style"
            if value is not None and abs(value) > indent_threshold_twips:
                code = "formula_paragraph_indent" if has_math and attribute in ("left", "firstLine") else "paragraph_indent"
                origin = f"style {style_name}" if source == "style" else "direct paragraph formatting"
                add(
                    index,
                    code,
                    "high" if has_math else "medium",
                    f"Paragraph {index} has effective {attribute}={value} twips from {origin} (threshold {indent_threshold_twips}); compare with the source PDF and neighbouring paragraphs.",
                    excerpt,
                    attribute=attribute,
                    value_twips=value,
                    property_source=source,
                    style_name=style_name,
                    contains_math=has_math,
                    source_pdf_required=True,
                )

        spacing = ppr.find(f"{{{W}}}spacing") if ppr is not None else None
        for attribute in ("before", "after"):
            value = _int_attr(spacing, attribute)
            source = "paragraph"
            style_name = None
            if value is None:
                value, style_name = _inherited_style_value(style_map, style_id, "spacing", attribute)
                source = "style"
            if value is not None and value > spacing_threshold_twips:
                origin = f"style {style_name}" if source == "style" else "direct paragraph formatting"
                add(
                    index,
                    "large_paragraph_spacing",
                    "medium",
                    f"Paragraph {index} has effective {attribute} spacing={value} twips from {origin} (threshold {spacing_threshold_twips}); verify that the resulting whitespace is intentional.",
                    excerpt,
                    attribute=attribute,
                    value_twips=value,
                    property_source=source,
                    style_name=style_name,
                    contains_math=has_math,
                    source_pdf_required=True,
                )

        for occurrence, match in enumerate(ANSWER_SPACE_RE.finditer(text), start=1):
            count = len(match.group(1))
            add(
                index,
                "answer_blank_spacing",
                "medium",
                f"Answer parentheses contain {count} spaces; compare the answer blank with the source PDF before changing it.",
                excerpt,
                occurrence=occurrence,
                space_count=count,
                contains_math=has_math,
                source_pdf_required=True,
            )

    return {
        "schema": "docx-layout-check/1",
        "docx": str(docx_path),
        "candidate_sha256": sha256_file(docx_path),
        "thresholds": {
            "indent_twips": indent_threshold_twips,
            "paragraph_spacing_twips": spacing_threshold_twips,
            "answer_blank_spaces": 3,
        },
        "paragraph_count": len(paragraphs),
        "finding_count": len(findings),
        "findings": findings,
    }


def make_review_template(docx_path: Path, output_path: Path) -> dict:
    report = inspect_docx(docx_path)
    template = {
        "schema": "docx-layout-review/1",
        "candidate_sha256": report["candidate_sha256"],
        "checked": False,
        "reviewer": "",
        "findings": [
            {
                "id": finding["id"],
                "disposition": "",
                "reason": "",
                "source_pdf_checked": False,
            }
            for finding in report["findings"]
        ],
    }
    output_path = Path(output_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    output_path.write_text(json.dumps(template, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    return template


def evaluate_layout_review(docx_path: Path, review_path: Path, report: dict | None = None) -> dict:
    docx_path = Path(docx_path).resolve(strict=True)
    candidate_hash = sha256_file(docx_path)
    report = report or inspect_docx(docx_path)
    try:
        review = json.loads(Path(review_path).read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError):
        review = {}

    entries = review.get("findings", []) if isinstance(review, dict) else []
    by_id = {entry.get("id"): entry for entry in entries if isinstance(entry, dict)}
    current_ids = {finding["id"] for finding in report["findings"]}
    reviewed_ids = set(by_id)
    unresolved = []
    for finding in report["findings"]:
        entry = by_id.get(finding["id"], {})
        if entry.get("disposition") != "intentional" or not str(entry.get("reason", "")).strip():
            unresolved.append(finding["id"])
        if finding.get("source_pdf_required") and entry.get("source_pdf_checked") is not True:
            unresolved.append(finding["id"] + ":source_pdf_not_checked")

    hash_matches = review.get("candidate_sha256") == candidate_hash
    exact_findings = reviewed_ids == current_ids
    complete = (
        review.get("schema") == "docx-layout-review/1"
        and review.get("checked") is True
        and bool(str(review.get("reviewer", "")).strip())
        and hash_matches
        and exact_findings
        and not unresolved
    )
    return {
        "complete": complete,
        "candidate_sha256": candidate_hash,
        "hash_matches": hash_matches,
        "all_findings_accounted_for": exact_findings,
        "finding_count": len(report["findings"]),
        "unresolved": sorted(set(unresolved)),
    }


def main():
    parser = argparse.ArgumentParser(description="Check DOCX paragraph layout anomalies.")
    sub = parser.add_subparsers(dest="action", required=True)
    check = sub.add_parser("check", help="scan a DOCX and print findings")
    check.add_argument("--docx", type=Path, required=True)
    check.add_argument("--output", type=Path)
    template = sub.add_parser("template", help="create a source-bound layout review template")
    template.add_argument("--docx", type=Path, required=True)
    template.add_argument("--output", type=Path, required=True)
    args = parser.parse_args()

    if args.action == "check":
        report = inspect_docx(args.docx)
        if args.output:
            args.output.parent.mkdir(parents=True, exist_ok=True)
            args.output.write_text(json.dumps(report, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
        print(json.dumps(report, ensure_ascii=False))
    else:
        template = make_review_template(args.docx, args.output)
        print(json.dumps({"output": str(args.output), "candidate_sha256": template["candidate_sha256"],
                          "findings": len(template["findings"]), "checked": False}, ensure_ascii=False))


if __name__ == "__main__":
    main()
