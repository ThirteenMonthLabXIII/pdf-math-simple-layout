#!/usr/bin/env python3
"""Import MinerU's editable Word output into a source-bound review workspace."""
import argparse, hashlib, json, shutil, sys
from pathlib import Path
import pymupdf
from docx import Document
from docx_layout_check import evaluate_layout_review, inspect_docx
from docx_styling import apply_docx_styles
from question_layout import apply_question_layout, check_question_layout, check_word_pagination
from review_coverage import complete_page_review

for stream in (sys.stdout, sys.stderr):
    if hasattr(stream, "reconfigure"): stream.reconfigure(encoding="utf-8", errors="replace")

def sha(path):
    h=hashlib.sha256()
    with open(path,"rb") as f:
        for chunk in iter(lambda:f.read(1048576),b""): h.update(chunk)
    return h.hexdigest()

def stats(path):
    doc=Document(str(path))
    return {"paragraphs":len(doc.paragraphs),
            "chars":sum(len(p.text.strip()) for p in doc.paragraphs),
            "tables":len(doc.tables),"images":sum(x.reltype.endswith("/image") for x in doc.part.rels.values())}

def save(path,obj): path.write_text(json.dumps(obj,ensure_ascii=False,indent=2)+"\n",encoding="utf-8")

def prepare(args):
    pdf=args.pdf.resolve(strict=True); src=args.mineru_docx.resolve(strict=True)
    if pdf.suffix.lower()!=".pdf" or src.suffix.lower()!=".docx": raise ValueError("Expected PDF and DOCX")
    s=stats(src)
    if s["chars"]<10 and s["tables"]==0: raise ValueError("MinerU DOCX seems empty")
    ws=args.workspace.resolve()
    if ws.exists() and any(ws.iterdir()): raise FileExistsError("Workspace not empty: "+str(ws))
    ws.mkdir(parents=True,exist_ok=True)
    doc=pymupdf.open(str(pdf))
    total=len(doc); images=ws/"source_pages";images.mkdir()
    for i,page in enumerate(doc,1):
        page.get_pixmap(dpi=120,alpha=False).save(str(images/f"page-{i:03}.png"))
    doc.close()
    shutil.copy2(src,ws/"mineru_original.docx")
    shutil.copy2(src,ws/"working.docx")
    default_profile = Path(__file__).resolve().parents[1] / "references" / "style_profile.json"
    profile_path = args.profile.resolve(strict=True) if args.profile else default_profile
    profile = json.loads(profile_path.read_text(encoding="utf-8")) if profile_path.is_file() else {}
    style_report = apply_docx_styles(ws/"working.docx", profile)
    question_layout_report = apply_question_layout(ws/"working.docx", profile)
    question_structure_check = check_question_layout(ws/"working.docx", profile)
    word_pagination_report = None
    word_pagination_error = None
    if question_structure_check["status"] == "ready_for_word_pagination_check":
        try:
            word_pagination_report = check_word_pagination(ws/"working.docx")
        except Exception as exc:
            word_pagination_error = f"{type(exc).__name__}: {exc}"
    if question_layout_report["question_group_count"] == 0:
        preflight_status = "needs_question_mapping"
    elif question_structure_check["status"] != "ready_for_word_pagination_check":
        preflight_status = "needs_question_layout"
    elif word_pagination_report is None:
        preflight_status = "needs_word_pagination_check"
    elif word_pagination_report["status"] != "accepted":
        preflight_status = "needs_question_page_fit"
    else:
        preflight_status = "ready_for_codex_formatting"
    question_layout_report["structural_check"] = question_structure_check
    question_layout_report["word_pagination"] = word_pagination_report
    question_layout_report["word_pagination_error"] = word_pagination_error
    question_layout_report["preflight_status"] = preflight_status
    save(ws/"question_layout_plan.json", question_layout_report)
    save(ws/"style_policy.json", {
        "profile_path": str(profile_path) if profile_path.is_file() else None,
        "profile_sha256": sha(profile_path) if profile_path.is_file() else None,
        **style_report,
        "question_layout": question_layout_report,
    })
    save(ws/"mineru_manifest.json",{"source_pdf":str(pdf),"source_sha256":sha(pdf),
        "source_pages":total,"raw_sha256":sha(ws/"mineru_original.docx"),
        "working_docx_sha256":sha(ws/"working.docx"),
        "style_policy_schema":style_report["schema"],
        "question_layout_schema":question_layout_report["schema"],
        "question_layout_preflight_status":preflight_status,
        "raw_stats":s})
    save(ws/"mineru_review.json",{"candidate_sha256":"","word_preview_checked":False,
        "pages":[{"pdf_page":i,"content_checked":False,"formulas_checked":False,"figures_checked":False}
                 for i in range(1,total+1)]})
    save(ws/"layout_review.json",{"schema":"docx-layout-review/1","candidate_sha256":"",
        "checked":False,"reviewer":"","findings":[]})
    print(json.dumps({"status":preflight_status,"workspace":str(ws),
        "working_docx":str(ws/"working.docx"),"source_pages":total,"mineru_stats":s,
        "preformatted":True,"style_policy":style_report["paragraphs_by_style"],
        "question_layout_groups":question_layout_report["question_group_count"],
        "inserted_answer_paragraphs":question_layout_report["inserted_empty_paragraphs"],
        "word_page_count":word_pagination_report["word_page_count"] if word_pagination_report else None,
        "split_question_groups":word_pagination_report["split_question_groups"] if word_pagination_report else None,
        "word_pagination_error":word_pagination_error},ensure_ascii=False))

def audit(args):
    ws=args.workspace.resolve(strict=True)
    m=json.loads((ws/"mineru_manifest.json").read_text(encoding="utf-8"))
    r=json.loads((ws/"mineru_review.json").read_text(encoding="utf-8"))
    doc=args.docx.resolve(strict=True) if args.docx else ws/"working.docx"
    s=stats(doc)
    checked=complete_page_review(r.get("pages"),m["source_pages"],"pdf_page",
        ("content_checked","formulas_checked","figures_checked"))
    with pymupdf.open(str(m["source_pdf"])) as source_pdf:
        source_page_count_matches=len(source_pdf)==m["source_pages"]
    layout_report=inspect_docx(doc)
    save(ws/"docx_layout_findings.json",layout_report)
    layout_review=evaluate_layout_review(doc,ws/"layout_review.json",layout_report)
    policy_data=json.loads((ws/"style_policy.json").read_text(encoding="utf-8"))
    profile_path=Path(policy_data["profile_path"]) if policy_data.get("profile_path") else None
    profile=json.loads(profile_path.read_text(encoding="utf-8")) if profile_path and profile_path.is_file() else {}
    question_layout_report=check_question_layout(doc,profile)
    word_pagination_report=check_word_pagination(doc) if args.word_pages else None
    checks={"source_unchanged":sha(Path(m["source_pdf"]))==m["source_sha256"],
        "source_page_count_matches":source_page_count_matches,
        "mineru_raw_unchanged":sha(ws/"mineru_original.docx")==m["raw_sha256"],
        "editable_text_present":s["chars"]>10,
        "current_document_reviewed":r.get("candidate_sha256")==sha(doc),
        "word_preview_checked":r.get("word_preview_checked") is True,
        "all_pdf_pages_reviewed":checked,
        "docx_layout_review_complete":layout_review["complete"],
        "question_layout_complete":question_layout_report["status"]=="ready_for_word_pagination_check",
        "question_groups_same_page":word_pagination_report is not None and word_pagination_report["status"]=="accepted"}
    warnings=[]
    if s["images"]==0: warnings.append("No images in Word: check geometry drawings")
    if s["chars"]<m["raw_stats"]["chars"]//2: warnings.append("Over 50% of MinerU text missing")
    status="accepted" if all(checks.values()) and not warnings else "needs_review"
    report={"status":status,"checks":checks,"warnings":warnings,"docx_stats":s,
        "word_pagination":word_pagination_report,
        "question_layout":question_layout_report,
        "docx_layout_findings":layout_report["findings"],"docx_layout_review":layout_review}
    save(ws/"mineru_audit.json",report)
    print(json.dumps(report,ensure_ascii=False))
    if args.require_accepted and status!="accepted": sys.exit(2)

if __name__=="__main__":
    ap=argparse.ArgumentParser()
    cmd=ap.add_subparsers(dest="action",required=True)
    p=cmd.add_parser("prepare")
    p.add_argument("--pdf",type=Path,required=True);p.add_argument("--mineru-docx",type=Path,required=True)
    p.add_argument("--workspace",type=Path,required=True);p.add_argument("--profile",type=Path)
    p.set_defaults(f=prepare)
    a=cmd.add_parser("audit");a.add_argument("--workspace",type=Path,required=True)
    a.add_argument("--docx",type=Path);a.add_argument("--require-accepted",action="store_true");a.add_argument("--word-pages",action="store_true");a.set_defaults(f=audit)
    o=ap.parse_args();o.f(o)
