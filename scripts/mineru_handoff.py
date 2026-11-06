#!/usr/bin/env python3
"""Import MinerU's editable Word output into a source-bound review workspace."""
import argparse, hashlib, json, shutil, sys
from pathlib import Path
import pymupdf
from docx import Document

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
    save(ws/"mineru_manifest.json",{"source_pdf":str(pdf),"source_sha256":sha(pdf),
        "source_pages":total,"raw_sha256":sha(ws/"mineru_original.docx"),
        "raw_stats":s})
    save(ws/"mineru_review.json",{"candidate_sha256":"","word_preview_checked":False,
        "pages":[{"pdf_page":i,"content_checked":False,"formulas_checked":False,"figures_checked":False}
                 for i in range(1,total+1)]})
    print(json.dumps({"status":"ready_for_codex_formatting","workspace":str(ws),
        "working_docx":str(ws/"working.docx"),"source_pages":total,"mineru_stats":s},ensure_ascii=False))

def audit(args):
    ws=args.workspace.resolve(strict=True)
    m=json.loads((ws/"mineru_manifest.json").read_text(encoding="utf-8"))
    r=json.loads((ws/"mineru_review.json").read_text(encoding="utf-8"))
    doc=args.docx.resolve(strict=True) if args.docx else ws/"working.docx"
    s=stats(doc)
    checked=(len(r.get("pages",[]))==m["source_pages"] and
             all(p.get("content_checked") is True and p.get("formulas_checked") is True
                 and p.get("figures_checked") is True for p in r.get("pages",[])))
    checks={"source_unchanged":sha(Path(m["source_pdf"]))==m["source_sha256"],
        "mineru_raw_unchanged":sha(ws/"mineru_original.docx")==m["raw_sha256"],
        "editable_text_present":s["chars"]>10,
        "current_document_reviewed":r.get("candidate_sha256")==sha(doc),
        "word_preview_checked":r.get("word_preview_checked") is True,
        "all_pdf_pages_reviewed":checked}
    warnings=[]
    if s["images"]==0: warnings.append("No images in Word: check geometry drawings")
    if s["chars"]<m["raw_stats"]["chars"]//2: warnings.append("Over 50% of MinerU text missing")
    status="accepted" if all(checks.values()) and not warnings else "needs_review"
    report={"status":status,"checks":checks,"warnings":warnings,"docx_stats":s}
    save(ws/"mineru_audit.json",report)
    print(json.dumps(report,ensure_ascii=False))
    if args.require_accepted and status!="accepted": sys.exit(2)

if __name__=="__main__":
    ap=argparse.ArgumentParser()
    cmd=ap.add_subparsers(dest="action",required=True)
    p=cmd.add_parser("prepare")
    p.add_argument("--pdf",type=Path,required=True);p.add_argument("--mineru-docx",type=Path,required=True)
    p.add_argument("--workspace",type=Path,required=True);p.set_defaults(f=prepare)
    a=cmd.add_parser("audit");a.add_argument("--workspace",type=Path,required=True)
    a.add_argument("--docx",type=Path);a.add_argument("--require-accepted",action="store_true");a.set_defaults(f=audit)
    o=ap.parse_args();o.f(o)
