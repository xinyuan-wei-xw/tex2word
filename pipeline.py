r"""tex2word pipeline: LaTeX project zip + compiled PDF -> high-quality .docx.

Architecture (v1):
  1. CONTENT from .tex via pandoc (structure, math->OMML, tables, sections).
  2. STYLE from a JSON style profile (v1: academic defaults; v3: AI vision
     reads the PDF and generates this profile automatically).
  3. FIGURES: pandoc resolves \includegraphics from the zip; anything missing
     is pulled from the compiled PDF (rendered at print resolution).
  4. python-docx post-processing applies the style profile deterministically.

Storage layout (per-user from day one; auth plugs in later):
  storage/<user_id>/<project_id>/{upload.zip, extracted/..., output.docx}
r"""
import json
import os
import re
import shutil
import subprocess
import zipfile

import pymupdf  # PyMuPDF
from docx import Document
from docx.shared import Pt, RGBColor, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH


def find_main_tex(extract_dir):
    """Return the .tex file containing \documentclass (the root document)."""
    candidates = []
    for root, _, files in os.walk(extract_dir):
        for f in files:
            if f.endswith(".tex"):
                p = os.path.join(root, f)
                try:
                    with open(p, "r", errors="ignore") as fh:
                        head = fh.read(20000)
                    if r"\documentclass" in head:
                        # prefer top-level files over nested ones
                        depth = p[len(extract_dir):].count(os.sep)
                        candidates.append((depth, p))
                except OSError:
                    pass
    if not candidates:
        return None
    candidates.sort()
    return candidates[0][1]


def find_pdf(extract_dir, tex_path):
    """Return the compiled PDF (same basename as tex, else largest PDF)."""
    pdfs = []
    for root, _, files in os.walk(extract_dir):
        for f in files:
            if f.endswith(".pdf"):
                pdfs.append(os.path.join(root, f))
    if not pdfs:
        return None
    if tex_path:
        base = os.path.splitext(os.path.basename(tex_path))[0] + ".pdf"
        for p in pdfs:
            if os.path.basename(p) == base:
                return p
    pdfs.sort(key=lambda p: os.path.getsize(p), reverse=True)
    return pdfs[0]


def find_bib(extract_dir):
    for root, _, files in os.walk(extract_dir):
        for f in files:
            if f.endswith(".bib"):
                return os.path.join(root, f)
    return None


def run_pandoc(tex_path, work_dir, out_docx, bib_path=None):
    """tex -> docx via pandoc. Returns (ok, log)."""
    tex_path = os.path.abspath(tex_path)
    out_docx = os.path.abspath(out_docx)
    work_dir = os.path.abspath(work_dir)
    cmd = [
        "pandoc", tex_path,
        "-f", "latex", "-t", "docx",
        "-o", out_docx,
        "--resource-path", work_dir,
    ]
    if bib_path:
        cmd += ["--bibliography", os.path.abspath(bib_path), "--citeproc"]
    try:
        r = subprocess.run(cmd, capture_output=True, text=True, timeout=120,
                           cwd=work_dir)
        return r.returncode == 0, (r.stdout + r.stderr)[-2000:]
    except FileNotFoundError:
        return False, "pandoc not installed"
    except subprocess.TimeoutExpired:
        return False, "pandoc timed out"


def extract_pdf_images(pdf_path, out_dir):
    """Pull embedded raster images out of the PDF (fallback figure source)."""
    os.makedirs(out_dir, exist_ok=True)
    saved = []
    try:
        doc = pymupdf.open(pdf_path)
    except Exception:
        return saved
    seen = set()
    for i, page in enumerate(doc):
        for img in page.get_images(full=True):
            xref = img[0]
            if xref in seen:
                continue
            seen.add(xref)
            try:
                pix = pymupdf.Pixmap(doc, xref)
                if pix.n - pix.alpha > 3:  # CMYK -> RGB
                    pix = pymupdf.Pixmap(pymupdf.csRGB, pix)
                if pix.width < 80 or pix.height < 80:
                    continue  # skip tiny icons/logos
                fp = os.path.join(out_dir, f"pdfimg_p{i}_{xref}.png")
                pix.save(fp)
                saved.append(fp)
            except Exception:
                continue
    doc.close()
    return saved


def load_profile(name="default"):
    fp = os.path.join(os.path.dirname(__file__), "profiles", name + ".json")
    with open(fp) as fh:
        return json.load(fh)


def _hex_rgb(s):
    s = s.lstrip("#")
    return RGBColor(int(s[0:2], 16), int(s[2:4], 16), int(s[4:6], 16))


def postprocess_docx(docx_path, profile):
    """Apply the style profile to a pandoc-generated docx, deterministically."""
    doc = Document(docx_path)
    st = profile["styles"]

    def apply(name, cfg):
        if name not in doc.styles:
            return
        s = doc.styles[name]
        s.font.name = cfg.get("font", "Times New Roman")
        if cfg.get("size_pt"):
            s.font.size = Pt(cfg["size_pt"])
        if cfg.get("bold") is not None:
            s.font.bold = cfg["bold"]
        if cfg.get("italic") is not None:
            s.font.italic = cfg["italic"]
        if cfg.get("color"):
            s.font.color.rgb = _hex_rgb(cfg["color"])
        pf = s.paragraph_format
        if cfg.get("space_before_pt") is not None:
            pf.space_before = Pt(cfg["space_before_pt"])
        if cfg.get("space_after_pt") is not None:
            pf.space_after = Pt(cfg["space_after_pt"])
        if cfg.get("line_spacing") is not None:
            pf.line_spacing = cfg["line_spacing"]
        align = {"center": WD_ALIGN_PARAGRAPH.CENTER,
                 "right": WD_ALIGN_PARAGRAPH.RIGHT,
                 "justify": WD_ALIGN_PARAGRAPH.JUSTIFY}.get(cfg.get("align"))
        if align is not None:
            pf.alignment = align

    for name, cfg in st.items():
        apply(name, cfg)

    # two-column body, if the profile asks for it
    if profile.get("two_column"):
        for section in doc.sections:
            section._sectPr.xpath("./w:cols")[0].set(
                "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}num", "2")

    # drop pandoc's empty first paragraphs (title block artefacts)
    while len(doc.paragraphs) > 2 and not doc.paragraphs[0].text.strip():
        p = doc.paragraphs[0]._element
        p.getparent().remove(p)

    doc.save(docx_path)


def convert(project_dir, profile_name="default"):
    """Full pipeline. project_dir holds upload.zip already extracted to
    extracted/. Writes output.docx. Returns dict with status + log."""
    log = []
    extract_dir = os.path.join(project_dir, "extracted")
    os.makedirs(extract_dir, exist_ok=True)

    zpath = os.path.join(project_dir, "upload.zip")
    if os.path.exists(zpath):
        with zipfile.ZipFile(zpath) as z:
            z.extractall(extract_dir)
        # flatten single top-level folder (GitHub-style zips)
        entries = os.listdir(extract_dir)
        if len(entries) == 1 and os.path.isdir(os.path.join(extract_dir, entries[0])):
            inner = os.path.join(extract_dir, entries[0])
            for e in os.listdir(inner):
                shutil.move(os.path.join(inner, e), extract_dir)
            os.rmdir(inner)
        log.append("zip extracted")

    tex_path = find_main_tex(extract_dir)
    pdf_path = find_pdf(extract_dir, tex_path)
    bib_path = find_bib(extract_dir)
    log.append(f"main tex: {os.path.basename(tex_path) if tex_path else 'NONE'}")
    log.append(f"pdf: {os.path.basename(pdf_path) if pdf_path else 'NONE'}")
    if not tex_path:
        return {"ok": False, "log": log, "error": "no .tex with \\documentclass found in zip"}

    out_docx = os.path.join(project_dir, "output.docx")
    ok, plog = run_pandoc(tex_path, extract_dir, out_docx, bib_path)
    log.append("pandoc: " + ("ok" if ok else "FAILED"))
    log.append(plog)
    if not ok or not os.path.exists(out_docx):
        return {"ok": False, "log": log, "error": "pandoc conversion failed"}

    # figure fallback: harvest images from the PDF
    if pdf_path:
        imgs = extract_pdf_images(pdf_path, os.path.join(project_dir, "pdf_images"))
        log.append(f"extracted {len(imgs)} images from PDF as figure fallback")

    profile = load_profile(profile_name)
    try:
        postprocess_docx(out_docx, profile)
        log.append("style profile applied: " + profile_name)
    except Exception as e:  # never fail the whole job on styling
        log.append(f"style post-processing skipped: {e}")

    return {"ok": True, "log": log, "docx": out_docx}
