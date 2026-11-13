#!/usr/bin/env python3
"""Deterministic image sizing applied during the first DOCX build, before pagination."""
from __future__ import annotations

import math
import os
import re
from collections import defaultdict
from pathlib import Path
from statistics import median

from docx import Document
from docx.oxml.ns import qn
from docx.shared import Emu

EMU_PER_CM = 360000
LECTURE_RE = re.compile(r"^\s*第\s*([0-9一二三四五六七八九十百千万]+)\s*讲")


def default_image_size_policy() -> dict:
    """Use the proven small-image reference rule; never enlarge an image."""
    return {
        "enabled": True,
        "reference_min_width_cm": 0.0,
        "reference_max_width_cm": 6.0,
        "resize_threshold_cm": 6.0,
        "fallback_width_cm": 4.05,
        "group_by_lecture": True,
        "preserve_aspect_ratio": True,
        "only_shrink": True,
    }


def lecture_group_from_text(text: str, current_group: str = "__document__") -> str:
    """Return a stable lecture group when a paragraph/block begins a lecture title."""
    match = LECTURE_RE.match(text or "")
    return f"lecture:{match.group(1)}" if match else current_group


def plan_image_widths(image_items, policy: dict | None = None) -> dict:
    """Plan final widths from small images' per-lecture median, with a document fallback.

    Each item is a mapping with key, group, and width_cm. Images at or below the
    reference ceiling are left untouched. Images above the ceiling are set to the
    median of eligible small images in the same lecture. If that lecture has no
    reference image, use the document-wide median; if the document has none, use
    the configured 4.05 cm fallback.
    """
    config = {**default_image_size_policy(), **(policy or {})}
    items = []
    for item in image_items:
        key = str(item["key"])
        group = (
            str(item.get("group") or "__document__")
            if config["group_by_lecture"]
            else "__document__"
        )
        width = float(item["width_cm"])
        if not math.isfinite(width) or width <= 0:
            raise ValueError(f"Invalid image width for {key}: {width}")
        items.append({"key": key, "group": group, "width_cm": width})

    if not config["enabled"] or not items:
        return {
            "schema": "image-size-policy/1",
            "enabled": bool(config["enabled"]),
            "items": {
                item["key"]: {
                    **item,
                    "final_width_cm": item["width_cm"],
                    "resized": False,
                }
                for item in items
            },
            "groups": {},
            "total_images": len(items),
            "resized_count": 0,
        }

    min_ref = float(config["reference_min_width_cm"])
    max_ref = float(config["reference_max_width_cm"])
    threshold = float(config["resize_threshold_cm"])
    fallback = min(float(config["fallback_width_cm"]), max_ref)
    if min_ref < 0 or max_ref <= min_ref or threshold <= 0 or fallback <= 0:
        raise ValueError("Invalid image_size_policy limits")

    by_group = defaultdict(list)
    all_references = []
    for item in items:
        width = item["width_cm"]
        if min_ref <= width <= max_ref:
            by_group[item["group"]].append(width)
            all_references.append(width)

    document_reference = median(all_references) if all_references else fallback
    group_plans = {}
    for group in sorted({item["group"] for item in items}):
        local = by_group.get(group, [])
        if local:
            reference = median(local)
            source = "lecture_median" if group.startswith("lecture:") else "document_group_median"
        elif all_references:
            reference = document_reference
            source = "document_median_fallback"
        else:
            reference = fallback
            source = "fixed_fallback"
        group_plans[group] = {
            "reference_width_cm": float(reference),
            "reference_count": len(local),
            "reference_source": source,
        }

    planned = {}
    for item in items:
        reference = group_plans[item["group"]]["reference_width_cm"]
        final_width = item["width_cm"]
        resized = False
        if config["only_shrink"] and item["width_cm"] > threshold:
            final_width = min(item["width_cm"], reference)
            resized = final_width < item["width_cm"]
        planned[item["key"]] = {
            **item,
            "reference_width_cm": reference,
            "final_width_cm": float(final_width),
            "resized": resized,
        }

    return {
        "schema": "image-size-policy/1",
        "enabled": True,
        "reference_min_width_cm": min_ref,
        "reference_max_width_cm": max_ref,
        "resize_threshold_cm": threshold,
        "fallback_width_cm": fallback,
        "preserve_aspect_ratio": True,
        "only_shrink": bool(config["only_shrink"]),
        "items": planned,
        "groups": group_plans,
        "total_images": len(items),
        "resized_count": sum(1 for item in planned.values() if item["resized"]),
    }


def _paragraph_groups(document) -> dict:
    """Map body paragraph XML nodes to the most recent lecture heading."""
    groups = {}
    current = "__document__"
    for paragraph in document.element.body.iter(qn("w:p")):
        text = "".join(node.text or "" for node in paragraph.iter(qn("w:t")))
        if LECTURE_RE.match(text or ""):
            current = lecture_group_from_text(text, current)
        groups[paragraph] = current
    return groups


def _shape_paragraph(shape):
    node = shape._inline
    paragraph_tag = qn("w:p")
    while node is not None and node.tag != paragraph_tag:
        node = node.getparent()
    return node


def apply_docx_image_size_policy(docx_path: Path, policy: dict | None = None) -> dict:
    """Apply image sizing to a working DOCX before question layout/pagination."""
    path = Path(docx_path).resolve(strict=True)
    document = Document(str(path))
    shapes = list(document.inline_shapes)
    paragraph_groups = _paragraph_groups(document)
    items = []
    for index, shape in enumerate(shapes, start=1):
        paragraph = _shape_paragraph(shape)
        group = paragraph_groups.get(paragraph, "__document__")
        items.append({
            "key": str(index),
            "group": group,
            "width_cm": float(shape.width) / EMU_PER_CM,
        })

    plan = plan_image_widths(items, policy)
    changes = []
    for index, shape in enumerate(shapes, start=1):
        item = plan["items"][str(index)]
        if not item["resized"]:
            continue
        old_width = int(shape.width)
        old_height = int(shape.height)
        new_width = int(round(item["final_width_cm"] * EMU_PER_CM))
        new_height = int(round(old_height * new_width / old_width))
        shape.width = Emu(new_width)
        shape.height = Emu(new_height)
        changes.append({
            "image_index": index,
            "group": item["group"],
            "old_width_cm": round(old_width / EMU_PER_CM, 2),
            "new_width_cm": round(new_width / EMU_PER_CM, 2),
            "old_height_cm": round(old_height / EMU_PER_CM, 2),
            "new_height_cm": round(new_height / EMU_PER_CM, 2),
        })

    if changes:
        temp_path = path.with_name(path.stem + ".image-size-tmp.docx")
        document.save(str(temp_path))
        os.replace(temp_path, path)

    return {
        "schema": plan["schema"],
        "enabled": plan["enabled"],
        "image_count": len(shapes),
        "resized_count": len(changes),
        "groups": plan["groups"],
        "changes": changes,
        "applied_before_question_layout": True,
        "aspect_ratio_preserved": True,
    }
