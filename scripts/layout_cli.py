#!/usr/bin/env python3
"""Make reviewable, editable Word handouts from PDF page evidence.

This CLI deliberately keeps OCR recognition separate from human-reviewed content.
No model output is silently marked as verified.
"""
from __future__ import annotations

import argparse
import hashlib
import json
import re
import shutil
import sys

# Windows service shells may expose cp1252 even with Chinese file paths.
for _stream in (sys.stdout, sys.stderr):
    if hasattr(_stream, "reconfigure"):
        try:
            _stream.reconfigure(encoding="utf-8", errors="replace")
        except (OSError, ValueError):
            pass
import tempfile
from collections import Counter
from datetime import datetime, timezone
from pathlib import Path

try:
    import pymupdf as fitz
except ImportError:
    import fitz
from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt
from docx_layout_check import evaluate_layout_review, inspect_docx
from docx_styling import apply_docx_styles, default_paragraph_policy, extract_paragraph_policy
from image_sizing import apply_docx_image_size_policy, default_image_size_policy
from question_layout import apply_question_layout, check_question_layout, check_word_pagination
from review_coverage import complete_page_review

SCHEMA = "math-simple-layout/1"
TEXT_TYPES = {"title", "heading", "paragraph", "question", "subquestion", "math"}
DEFAULT_PROFILE = {
    "page_width_mm": 210.0,
    "page_height_mm": 297.0,
    "margin_top_pt": 62.35,
    "margin_bottom_pt": 62.35,
    "margin_left_pt": 68.05,
    "margin_right_pt": 68.05,
    "font_name": "宋体",
    "east_asia_font": "宋体",
    "body_pt": 10.5,
    "title_pt": 16.0,
    "heading_pt": 10.5,
    "line_spacing": 1.15,
    "body_after_pt": 7,
    "question_after_pt": 7,
    "subquestion_after_pt": 10,
    "heading_before_pt": 6,
    "heading_after_pt": 4,
    "title_after_pt": 10,
    "image_max_width_cm": 14.7,
    "image_size_policy": default_image_size_policy(),
    "paragraph_policy": default_paragraph_policy(),
}


def sha256(path: Path) -> str:
    hasher = hashlib.sha256()
    with path.open("rb") as stream:
        for b in iter(lambda: stream.read(1024 * 1024), b""):
            hasher.update(b)
    return hasher.hexdigest()


def load(path: Path) -> dict:
    with path.open("r", encoding="utf-8") as stream:
        return json.load(stream)


def save(path: Path, value: dict) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", encoding="utf-8") as stream:
        json.dump(value, stream, ensure_ascii=False, indent=2)
        stream.write("\n")


def rel_or_absolute(root: Path, text: str) -> Path:
    path = Path(text)
    return path if path.is_absolute() else root / path


def classify(text: str) -> str:
    s = text.strip().replace(" ", "")
    if re.match(r"^第[一二三四五六七八九十百千万\d]+讲", s):
        return "title"
    if re.match(r"^(【?(例题|练习|知识点|课后练习)|课后练习|知识点[一二三四五六七八九十0-9])", s):
        return "heading"
    if re.match(r"^[（(][1-9][0-9]*[）)]", s):
        return "subquestion"
    if re.match(r"^\d+[．.、](?=\S)", s):
        return "question"
    return "paragraph"


def effective_font(font, default) -> dict:
    name = font.name or default["font_name"]
    ea = font._element.rPr.rFonts.get(qn("w:eastAsia")) if font._element.rPr is not None and font._element.rPr.rFonts is not None else None
    return {"font_name": name, "east_asia_font": ea or default["east_asia_font"], "body_pt": round(font.size.pt, 2) if font.size else default["body_pt"]}


def cmd_profile(args):
    doc = Document(str(args.sample))
    profile = dict(DEFAULT_PROFILE)
    section = doc.sections[0]
    profile.update({
        "page_width_mm": round(section.page_width.cm * 10, 2),
        "page_height_mm": round(section.page_height.cm * 10, 2),
        "margin_top_pt": round(section.top_margin.pt, 2),
        "margin_bottom_pt": round(section.bottom_margin.pt, 2),
        "margin_left_pt": round(section.left_margin.pt, 2),
        "margin_right_pt": round(section.right_margin.pt, 2),
    })
    profile.update(effective_font(doc.styles["Normal"].font, profile))
    for paragraph in doc.paragraphs:
        if not paragraph.text.strip():
            continue
        if paragraph.alignment == WD_ALIGN_PARAGRAPH.CENTER and paragraph.runs:
            sizes = [r.font.size.pt for r in paragraph.runs if r.font.size]
            if sizes:
                profile["title_pt"] = round(max(sizes), 2)
            if paragraph.paragraph_format.space_after:
                profile["title_after_pt"] = round(paragraph.paragraph_format.space_after.pt, 2)
            break
    profile["paragraph_policy"], policy_warnings = extract_paragraph_policy(
        args.sample, profile.get("paragraph_policy")
    )
    profile["paragraph_policy_warnings"] = policy_warnings
    profile["sample_sha256"] = sha256(args.sample)
    profile["source_reference_only"] = True
    save(args.output, profile)
    print(json.dumps({"profile": str(args.output), "body_pt": profile["body_pt"], "title_pt": profile["title_pt"]}, ensure_ascii=False))


def pdf_info(pdf: Path) -> dict:
    doc = fitz.open(str(pdf))
    result = {"source_pdf": str(pdf.resolve()), "source_sha256": sha256(pdf), "page_count": len(doc), "pages": []}
    for ix, page in enumerate(doc, 1):
        char_count = len(page.get_text().strip())
        result["pages"].append({"page": ix, "width_pt": round(page.rect.width, 3), "height_pt": round(page.rect.height, 3), "text_chars": char_count, "scanned": char_count < 30})
    doc.close()
    return result


def cmd_inspect(args):
    info = pdf_info(args.pdf)
    print(json.dumps(info, ensure_ascii=False, indent=2))
    if args.output:
        save(args.output, info)


def native_blocks(page, pnum: int) -> list[dict]:
    words = page.get_text("blocks")
    items = []
    for block in words:
        x0, y0, x1, y1, content, *_ = block
        for text_line in content.splitlines():
            s = text_line.strip()
            if s:
                items.append({"id": f"p{pnum:03d}-b{len(items)+1:03d}", "type": classify(s), "text": s,
                              "bbox": [round(v, 2) for v in (x0,y0,x1,y1)], "source": "pdf-text-layer", "confidence": 1.0})
    return items


def cmd_prepare(args):
    pdf = args.pdf.resolve()
    out = args.output.resolve()
    manifest_path = out / "layout.json"
    if manifest_path.exists() and not args.overwrite:
        raise RuntimeError(f"Workspace already has layout.json; refusing overwrite: {out}")
    info = pdf_info(pdf)
    out.mkdir(parents=True, exist_ok=True)
    (out / "pages").mkdir(exist_ok=True)
    info["schema"] = SCHEMA
    info["created_at_utc"] = datetime.now(timezone.utc).isoformat()
    info["dpi"] = args.dpi
    info["working_status"] = "draft"
    doc = fitz.open(str(pdf))
    for entry, page in zip(info["pages"], doc):
        dest = out / "pages" / f"page-{entry['page']:03d}.png"
        pix = page.get_pixmap(matrix=fitz.Matrix(args.dpi / 72, args.dpi / 72), alpha=False)
        pix.save(str(dest))
        entry["render"] = dest.relative_to(out).as_posix()
        entry["blocks"] = native_blocks(page, entry["page"]) if not entry["scanned"] else []
        entry["verified"] = False
        entry["source_blank"] = False
        entry["review_notes"] = "" if not entry["scanned"] else "扫描件：需要运行OCR并按页面图片核对公式、题号、图形区域"
    doc.close()
    save(manifest_path, info)
    review = {"schema": SCHEMA, "source_sha256": info["source_sha256"], "candidate_sha256": None,
              "reviewer": "", "pages": [{"page": v["page"], "source_checked": False, "output_checked": False, "text_math_checked": False, "figures_checked": False, "notes": ""} for v in info["pages"]]}
    save(out / "visual_review.json", review)
    save(out / "layout_review.json", {"schema": "docx-layout-review/1", "candidate_sha256": "",
         "checked": False, "reviewer": "", "findings": []})
    print(json.dumps({"workspace": str(out), "layout": str(manifest_path), "pages": len(info["pages"]),
                      "scan_pages": sum(v["scanned"] for v in info["pages"]), "next": "recognize --workspace; revise layout.json; build --workspace; render; audit"}, ensure_ascii=False))


def ocr_predict(image: Path):
    try:
        from rapidocr_onnxruntime import RapidOCR
        if not hasattr(ocr_predict, "engine"):
            ocr_predict.engine = RapidOCR()
        return ocr_predict.engine(str(image))
    except ImportError:
        try:
            from rapidocr import RapidOCR
            if not hasattr(ocr_predict, "engine"):
                ocr_predict.engine = RapidOCR()
            return ocr_predict.engine(str(image))
        except ImportError as exc:
            raise RuntimeError("Scanned PDF requires local RapidOCR; see docs/INSTALL.md. Or transcribe page images into layout.json with visual verification.") from exc


def cmd_recognize(args):
    ws = args.workspace.resolve()
    manifest_path = ws / "layout.json"
    data = load(manifest_path)
    check_source(data)
    pages = [p for p in data["pages"] if p["scanned"] and not p["blocks"]]
    if not pages:
        print(json.dumps({"recognized_pages": 0, "message": "No unrecognized scanned pages"}))
        return
    # Write a crash-safe checkpoint after every completed page. Refuse to erase corrected pages.
    for entry in pages:
        image = rel_or_absolute(ws, entry["render"])
        result = ocr_predict(image)
        raw = result[0] if isinstance(result, (tuple, list)) and len(result) == 2 and (result[0] is None or isinstance(result[0], list)) else result
        raw = raw or []
        blocks = []
        from PIL import Image
        iw, ih = Image.open(image).size
        for item in raw:
            if not isinstance(item, (tuple, list)) or len(item) < 2:
                continue
            points, text = item[:2]
            conf = item[2] if len(item) > 2 else None
            if not isinstance(text, str) or not text.strip():
                continue
            coords = [(float(x), float(y)) for x, y in points]
            box = [min(x for x, y in coords) / iw * entry["width_pt"], min(y for x, y in coords) / ih * entry["height_pt"],
                   max(x for x, y in coords) / iw * entry["width_pt"], max(y for x, y in coords) / ih * entry["height_pt"]]
            blocks.append({"id": f"p{entry['page']:03d}-b{len(blocks)+1:03d}", "type": classify(text), "text": text.strip(),
                           "bbox": [round(v,2) for v in box], "source": "rapidocr", "confidence": round(float(conf), 3) if conf is not None else None})
        blocks.sort(key=lambda x: (x["bbox"][1] // 10, x["bbox"][0]))
        entry["blocks"] = blocks
        entry["verified"] = False
        entry["review_notes"] = f"RapidOCR初稿({len(blocks)}行)。需要对照原页校正文句、数学符号并补齐图形裁剪框；不可直接交付。"
        save(manifest_path, data)
        print(f"page {entry['page']:03d}: {len(blocks)} OCR lines",flush=True)
    print(json.dumps({"recognized_pages":len(pages),"layout":str(manifest_path)},ensure_ascii=False))


def check_source(data):
    if data.get("schema") != SCHEMA:
        raise ValueError("Invalid layout schema")
    source = Path(data["source_pdf"])
    if not source.is_file() or sha256(source) != data["source_sha256"]:
        raise ValueError("Source PDF missing or changed; do not silently proceed")
    page_count = data.get("page_count")
    pages = data.get("pages")
    with fitz.open(str(source)) as original:
        actual_page_count = len(original)
    if (type(page_count) is not int or page_count < 1 or page_count != actual_page_count
            or not isinstance(pages, list) or len(pages) != page_count
            or any(not isinstance(p, dict) or type(p.get("page")) is not int for p in pages)
            or [p["page"] for p in pages] != list(range(1, page_count + 1))):
        raise ValueError("Page mapping incomplete")
    return source


def check_bbox(page: dict, bbox):
    if not isinstance(bbox, list) or len(bbox) != 4:
        raise ValueError(f"Invalid bbox: {bbox}")
    x0, y0, x1, y1 = [float(v) for v in bbox]
    if not (0 <= x0 < x1 <= page["width_pt"] and 0 <= y0 < y1 <= page["height_pt"]):
        raise ValueError(f"bbox outside original PDF page {page['page']}: {bbox}")
    return fitz.Rect(x0,y0,x1,y1)


def validate_layout(data, allow_draft=False):
    check_source(data)
    bad = []
    for p in data["pages"]:
        if (not any(b.get("type") not in {"blank", "page_break"} for b in p["blocks"])
                and p.get("source_blank") is not True):
            bad.append(f"page {p['page']}: no content blocks")
        for i, b in enumerate(p["blocks"]):
            typ = b.get("type")
            if typ not in TEXT_TYPES | {"image", "table", "blank", "page_break"}:
                bad.append(f"page {p['page']} block {i}: unsupported type {typ!r}")
            if typ in TEXT_TYPES and not str(b.get("text", "")).strip():
                bad.append(f"page {p['page']} block {i}: empty text")
            if typ == "image":
                try:check_bbox(p,b.get("bbox"))
                except Exception as e: bad.append(str(e))
            if typ == "table" and not (isinstance(b.get("rows"), list) and b["rows"] and all(isinstance(r,list) for r in b["rows"])):
                bad.append(f"page {p['page']} block {i}: invalid table rows")
            if typ == "blank" and (type(b.get("count", 1)) is not int or not 1 <= b.get("count", 1) <= 100):
                bad.append(f"page {p['page']} block {i}: blank count must be an integer from 1 to 100")
    if bad and not allow_draft:
        raise ValueError("Layout incomplete:\n" + "\n".join(bad[:20]))
    return bad


def set_font(run, profile, size=None, bold=None):
    run.font.name = profile["font_name"]
    run.font.size = Pt(size if size is not None else profile["body_pt"])
    if bold is not None:
        run.font.bold = bold
    rPr = run._element.get_or_add_rPr()
    rFonts = rPr.rFonts
    if rFonts is None:
        rFonts = OxmlElement("w:rFonts")
        rPr.insert(0,rFonts)
    rFonts.set(qn("w:eastAsia"), profile["east_asia_font"])


def set_paragraph(paragraph, typ, profile):
    f = paragraph.paragraph_format
    f.line_spacing = profile["line_spacing"]
    if typ == "title":
        paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
        f.space_after = Pt(profile["title_after_pt"])
    elif typ == "heading":
        f.space_before = Pt(profile["heading_before_pt"])
        f.space_after = Pt(profile["heading_after_pt"])
    elif typ == "subquestion":
        f.space_after = Pt(profile["subquestion_after_pt"])
    elif typ == "question":
        f.space_after = Pt(profile["question_after_pt"])
    else:
        f.space_after = Pt(profile["body_after_pt"])


def make_document(data, profile: dict, result: Path, workdir: Path):
    pdf = fitz.open(data["source_pdf"])
    doc = Document()
    section = doc.sections[0]
    section.page_width = Cm(profile["page_width_mm"] / 10)
    section.page_height = Cm(profile["page_height_mm"] / 10)
    section.top_margin = Pt(profile["margin_top_pt"])
    section.bottom_margin = Pt(profile["margin_bottom_pt"])
    section.left_margin = Pt(profile["margin_left_pt"])
    section.right_margin = Pt(profile["margin_right_pt"])
    norm = doc.styles["Normal"]
    norm.font.name = profile["font_name"]
    norm.font.size = Pt(profile["body_pt"])
    norm.paragraph_format.line_spacing = profile["line_spacing"]
    norm.paragraph_format.space_after = Pt(profile["body_after_pt"])
    figures_dir = workdir / "figures"
    figures_dir.mkdir(parents=True, exist_ok=True)
    for page in data["pages"]:
        for i, block in enumerate(page["blocks"]):
            typ = block.get("type")
            if typ in TEXT_TYPES:
                paragraph = doc.add_paragraph()
                set_paragraph(paragraph,typ,profile)
                run = paragraph.add_run(block["text"])
                set_font(run,profile,profile["title_pt"] if typ == "title" else profile["heading_pt"] if typ == "heading" else None,typ in {"title","heading"})
            elif typ == "image":
                region = check_bbox(page,block["bbox"])
                zoom = max(float(block.get("dpi", 260)) / 72,1.0)
                pix = pdf[page["page"]-1].get_pixmap(matrix=fitz.Matrix(zoom,zoom),clip=region,alpha=False)
                target = figures_dir / f"page-{page['page']:03d}-figure-{i+1:03d}.png"
                pix.save(str(target))
                width = min(profile["image_max_width_cm"],region.width/72*2.54)
                paragraph = doc.add_paragraph()
                paragraph.alignment=WD_ALIGN_PARAGRAPH.CENTER
                paragraph.paragraph_format.space_after=Pt(4)
                paragraph.add_run().add_picture(str(target),width=Cm(width))
                if block.get("caption"):
                    cp=doc.add_paragraph()
                    cp.style=doc.styles["Caption"]
                    cp.alignment=WD_ALIGN_PARAGRAPH.CENTER
                    set_font(cp.add_run(block["caption"]),profile)
            elif typ == "table":
                rows=block["rows"]
                cols=max(len(r) for r in rows)
                table=doc.add_table(rows=len(rows),cols=cols)
                table.style="Table Grid"
                table.alignment=WD_TABLE_ALIGNMENT.CENTER
                for row_i, row in enumerate(rows):
                    for col_i, val in enumerate(row):
                        c=table.cell(row_i,col_i)
                        c.text=""
                        set_font(c.paragraphs[0].add_run(str(val)),profile)
            elif typ == "blank":
                for _ in range(block.get("count", 1)):
                    doc.add_paragraph()
            elif typ == "page_break":
                # Source PDF page boundaries are inventory markers, not Word page breaks.
                continue
    pdf.close()
    result.parent.mkdir(parents=True, exist_ok=True)
    doc.save(str(result))
    # Fix final figure dimensions once, before question grouping and Word pagination.
    image_size_report = apply_docx_image_size_policy(
        result, profile.get("image_size_policy")
    )
    style_report = apply_docx_styles(result, profile)
    question_layout_report = apply_question_layout(result, profile)
    question_layout_report["image_sizing"] = image_size_report
    structural_check = check_question_layout(result, profile)
    word_pagination_report = None
    word_pagination_error = None
    if structural_check["status"] == "ready_for_word_pagination_check":
        try:
            word_pagination_report = check_word_pagination(result)
        except Exception as exc:
            word_pagination_error = f"{type(exc).__name__}: {exc}"
    if question_layout_report["question_group_count"] == 0:
        preflight_status = "needs_question_mapping"
    elif structural_check["status"] != "ready_for_word_pagination_check":
        preflight_status = "needs_question_layout"
    elif word_pagination_report is None:
        preflight_status = "needs_word_pagination_check"
    elif word_pagination_report["status"] != "accepted":
        preflight_status = "needs_question_page_fit"
    else:
        preflight_status = "ready_for_content_review"
    question_layout_report["structural_check"] = structural_check
    question_layout_report["word_pagination"] = word_pagination_report
    question_layout_report["word_pagination_error"] = word_pagination_error
    question_layout_report["preflight_status"] = preflight_status
    return {
        "style_policy": style_report,
        "question_layout": question_layout_report,
        "image_sizing": image_size_report,
    }


def cmd_build(args):
    ws = args.workspace.resolve()
    data = load(ws / "layout.json")
    invalid = validate_layout(data,allow_draft=args.allow_incomplete)
    profile = load(args.profile) if args.profile else dict(DEFAULT_PROFILE)
    output = (args.output or ws / "candidate.docx").resolve()
    if output.exists() and not args.overwrite:
        raise FileExistsError(f"Refusing to overwrite existing candidate: {output}; pass --output new_name.docx")
    build_styles = make_document(data,profile,output,ws)
    style_report = build_styles["style_policy"]
    question_layout_report = build_styles["question_layout"]
    ignored_source_page_breaks=sum(1 for page in data["pages"] for block in page["blocks"] if block.get("type")=="page_break")
    report={"schema":SCHEMA,"source_sha256":data["source_sha256"],"candidate":str(output),"candidate_sha256":sha256(output),
            "style_policy":style_report,"question_layout":question_layout_report,
            "image_sizing":build_styles["image_sizing"],
            "question_layout_preflight_status":question_layout_report["preflight_status"],
            "answer_space_policy":profile.get("answer_space_policy", {}),
            "ignored_source_page_breaks":ignored_source_page_breaks,
            "page_count":data["page_count"],"blocks":sum(len(p["blocks"]) for p in data["pages"]),"layout_errors":invalid,
            "unverified_pages":[p["page"] for p in data["pages"] if p.get("verified") is not True],"status":"candidate_not_approved"}
    save(ws / "build_report.json",report)
    print(json.dumps({"output":str(output),"status":report["status"],
        "question_layout_preflight_status":question_layout_report["preflight_status"],
        "question_groups":question_layout_report["question_group_count"],
        "image_sizing":build_styles["image_sizing"],
        "word_page_count":question_layout_report["word_pagination"]["word_page_count"] if question_layout_report["word_pagination"] else None,
        "split_question_groups":question_layout_report["word_pagination"]["split_question_groups"] if question_layout_report["word_pagination"] else None,
        "unverified_pages":report["unverified_pages"],"layout_errors":invalid},ensure_ascii=False))


def cmd_crop(args):
    data=load(args.workspace / "layout.json")
    src=check_source(data)
    page=data["pages"][args.page-1]
    box=check_bbox(page,[float(x) for x in args.bbox.split(",")])
    with fitz.open(str(src)) as f:
        pix=f[args.page-1].get_pixmap(matrix=fitz.Matrix(args.dpi/72,args.dpi/72),clip=box,alpha=False)
        args.output.parent.mkdir(parents=True,exist_ok=True)
        pix.save(str(args.output))
    print(json.dumps({"figure":str(args.output),"bbox":[box.x0,box.y0,box.x1,box.y1]},ensure_ascii=False))


def extract_docx_text(doc):
    paragraphs=[p.text for p in doc.paragraphs if p.text.strip()]
    table_text=["\t".join(c.text for c in r.cells) for t in doc.tables for r in t.rows]
    return paragraphs,table_text


def cmd_audit(args):
    ws=args.workspace.resolve()
    data=load(ws / "layout.json")
    check_source(data)
    output=args.docx.resolve() if args.docx else (ws / "candidate.docx").resolve()
    if not output.is_file():raise FileNotFoundError(output)
    doc=Document(str(output))
    paras,tables=extract_docx_text(doc)
    body_text="\n".join(paras+tables)
    expected_texts=[b["text"] for p in data["pages"] for b in p["blocks"] if b["type"] in TEXT_TYPES]
    expected_tables=[str(v) for p in data["pages"] for b in p["blocks"] if b["type"]=="table" for row in b["rows"] for v in row]
    missing=[s for s in expected_texts if s not in body_text]
    missing_tables=[s for s in expected_tables if s not in body_text]
    exp_images=sum(b["type"]=="image" for p in data["pages"] for b in p["blocks"])
    actual_images=len(doc.inline_shapes)
    bad_layout=validate_layout(data,allow_draft=True)
    unverified=[p["page"] for p in data["pages"] if p.get("verified") is not True]
    review_path=ws/"visual_review.json"
    review=load(review_path) if review_path.exists() else {}
    expected_hash=sha256(output)
    review_pages=review.get("pages",[])
    review_fields=("source_checked","output_checked","text_math_checked","figures_checked")
    checked=[r.get("page") for r in review_pages if isinstance(r,dict) and all(r.get(k) is True for k in review_fields)] if isinstance(review_pages,list) else []
    if review.get("candidate_sha256") != expected_hash:
        checked=[]
    reviewer=review.get("reviewer")
    review_complete=(review.get("candidate_sha256")==expected_hash
        and complete_page_review(review_pages,data["page_count"],"page",review_fields)
        and isinstance(reviewer,str) and bool(reviewer.strip()))
    docx_layout = inspect_docx(output)
    save(ws / "docx_layout_findings.json", docx_layout)
    layout_review = evaluate_layout_review(output, ws / "layout_review.json", docx_layout)
    build_report = load(ws / "build_report.json") if (ws / "build_report.json").exists() else {}
    question_layout_report = check_question_layout(output, {"answer_space_policy": build_report.get("answer_space_policy", {})})
    word_pagination_report = check_word_pagination(output) if getattr(args, "word_pages", False) else None
    source_unchanged=sha256(Path(data["source_pdf"]))==data["source_sha256"]
    checks={"source_unchanged":source_unchanged,"all_pages_have_content":not bad_layout,"all_source_pages_verified":not unverified,
            "all_expected_text_present":not missing,"all_table_values_present":not missing_tables,"figure_count_matches":exp_images==actual_images,
            "visual_review_complete":review_complete,"docx_layout_review_complete":layout_review["complete"],
            "question_layout_complete":question_layout_report["status"]=="ready_for_word_pagination_check",
            "question_groups_same_page":word_pagination_report is not None and word_pagination_report["status"]=="accepted"}
    ready=all(checks.values())
    report={"schema":SCHEMA,"status":"accepted" if ready else "needs_review","checks":checks,"page_count":data["page_count"],
            "unverified_pages":unverified,"missing_text_blocks":missing[:50],"missing_table_values":missing_tables[:50],
            "expected_figures":exp_images,"actual_figures":actual_images,"layout_errors":bad_layout,"visually_checked_pages":checked,
            "docx_layout_findings":docx_layout["findings"],"docx_layout_review":layout_review,
            "question_layout":question_layout_report,"word_pagination":word_pagination_report,
            "source_sha256":data["source_sha256"],"candidate_sha256":expected_hash,"candidate":str(output)}
    save(ws/"audit_report.json",report)
    print(json.dumps({"report":str(ws/"audit_report.json"),"status":report["status"],"checks":checks,"layout_errors":bad_layout[:4]},ensure_ascii=False))


def cmd_render(args):
    try:
        import win32com.client
    except ImportError as e:
        raise RuntimeError("render requires Windows Word COM; use Microsoft Word on 13-PC") from e
    docx=args.docx.resolve()
    out=args.output.resolve()
    out.mkdir(parents=True,exist_ok=True)
    pdf=out/"word-preview.pdf"
    word=win32com.client.DispatchEx("Word.Application")
    word.Visible=False
    word.DisplayAlerts=0
    document=None
    try:
        document=word.Documents.Open(str(docx),ConfirmConversions=False,ReadOnly=True,AddToRecentFiles=False)
        document.ExportAsFixedFormat(str(pdf),17)
    finally:
        if document is not None:document.Close(False)
        word.Quit()
    f=fitz.open(str(pdf))
    for i,p in enumerate(f,1):
        p.get_pixmap(matrix=fitz.Matrix(args.dpi/72,args.dpi/72),alpha=False).save(str(out/f"word-page-{i:03d}.png"))
    count=len(f)
    f.close()
    print(json.dumps({"preview":str(pdf),"page_images":count,"dir":str(out)},ensure_ascii=False))


def main(argv=None):
    parser=argparse.ArgumentParser(description="Scanned-math PDF → editable Word, source-bounded & review-gated")
    sub=parser.add_subparsers(dest="command",required=True)
    p=sub.add_parser("profile",help="Extract formatting ONLY from reference Word sample")
    p.add_argument("--sample",required=True,type=Path);p.add_argument("--output",required=True,type=Path)
    p.set_defaults(func=cmd_profile)
    p=sub.add_parser("inspect",help="Read-only PDF text-layer and page preflight")
    p.add_argument("--pdf",required=True,type=Path);p.add_argument("--output",type=Path)
    p.set_defaults(func=cmd_inspect)
    p=sub.add_parser("prepare",help="Freeze source hash, render all pages, create review manifest")
    p.add_argument("--pdf",required=True,type=Path);p.add_argument("--output",required=True,type=Path)
    p.add_argument("--dpi",type=int,default=180);p.add_argument("--overwrite",action="store_true")
    p.set_defaults(func=cmd_prepare)
    p=sub.add_parser("recognize",help="Run RapidOCR for only unrecognized scanned pages")
    p.add_argument("--workspace",required=True,type=Path);p.set_defaults(func=cmd_recognize)
    p=sub.add_parser("crop",help="Crop a figure from the original PDF using PDF-point bbox")
    p.add_argument("--workspace",required=True,type=Path);p.add_argument("--page",type=int,required=True)
    p.add_argument("--bbox",required=True,help="PDF points x0,y0,x1,y1");p.add_argument("--output",required=True,type=Path)
    p.add_argument("--dpi",type=int,default=300);p.set_defaults(func=cmd_crop)
    p=sub.add_parser("build",help="Build a Word candidate from verified/reviewable page blocks")
    p.add_argument("--workspace",required=True,type=Path);p.add_argument("--profile",type=Path);p.add_argument("--output",type=Path)
    p.add_argument("--allow-incomplete",action="store_true");p.add_argument("--overwrite",action="store_true")
    p.set_defaults(func=cmd_build)
    p=sub.add_parser("audit",help="Verify source, text, table, figures, and manually approved pages")
    p.add_argument("--workspace",required=True,type=Path);p.add_argument("--docx",type=Path);p.add_argument("--word-pages",action="store_true");p.set_defaults(func=cmd_audit)
    p=sub.add_parser("render",help="Word COM preview, isolated Word instance; no user documents touched")
    p.add_argument("--docx",required=True,type=Path);p.add_argument("--output",required=True,type=Path)
    p.add_argument("--dpi",type=int,default=140);p.set_defaults(func=cmd_render)
    args=parser.parse_args(argv)
    try:
        args.func(args)
        return 0
    except Exception as e:
        print(f"ERROR: {e}",file=sys.stderr)
        return 2


if __name__=="__main__":
    raise SystemExit(main())
