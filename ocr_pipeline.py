"""Mode B: scanned / poor-quality PDF -> OCR -> clean structured .docx.

Does NOT replicate the original layout. Produces a standard, readable,
editable document: Title + Heading 1/2 + body paragraphs.

Pipeline:
  1. Render each PDF page to image (PyMuPDF, 300 DPI).
  2. Tesseract OCR with TSV output (word boxes + bboxes).
     Languages: chi_sim+eng (mixed Chinese/English academic docs).
  3. Rebuild lines/blocks; estimate body font size per page from median
     line height; lines clearly larger than body -> headings
     (largest -> Heading 1, next tier -> Heading 2, first-page largest -> Title).
  4. Emit docx with the standard style profile.
"""
import csv
import os
import statistics
import subprocess

import pymupdf
from docx import Document
from docx.shared import Pt

DPI = 300
TESS_LANG = "chi_sim+eng"


def render_pages(pdf_path, out_dir, dpi=DPI):
    os.makedirs(out_dir, exist_ok=True)
    doc = pymupdf.open(pdf_path)
    paths = []
    for i, page in enumerate(doc):
        pix = page.get_pixmap(dpi=dpi)
        fp = os.path.join(out_dir, f"page_{i:03d}.png")
        pix.save(fp)
        paths.append(fp)
    doc.close()
    return paths


def ocr_tsv(image_path):
    """Run tesseract, return list of word dicts from TSV output."""
    base = image_path + ".tsv"
    r = subprocess.run(
        ["tesseract", image_path, image_path, "-l", TESS_LANG, "tsv"],
        capture_output=True, text=True, timeout=180)
    if r.returncode != 0:
        raise RuntimeError("tesseract failed: " + (r.stderr or r.stdout)[-500:])
    words = []
    with open(base, newline="", encoding="utf-8") as fh:
        reader = csv.DictReader(fh, delimiter="\t")
        for row in reader:
            if row.get("level") == "5" and row.get("text", "").strip():
                try:
                    words.append({
                        "page": int(row["page_num"]),
                        "block": int(row["block_num"]),
                        "par": int(row["par_num"]),
                        "line": int(row["line_num"]),
                        "left": int(row["left"]), "top": int(row["top"]),
                        "width": int(row["width"]),
                        "height": int(row["height"]),
                        "conf": float(row["conf"]),
                        "text": row["text"].strip(),
                    })
                except (ValueError, KeyError):
                    continue
    try:
        os.remove(base)
    except OSError:
        pass
    return words


def build_lines(words):
    """Group words -> lines -> list of (page, block, par, line_idx, text, height, width, left)."""
    lines = {}
    for w in words:
        key = (w["page"], w["block"], w["par"], w["line"])
        lines.setdefault(key, []).append(w)
    out = []
    for key, ws in sorted(lines.items()):
        ws.sort(key=lambda w: w["left"])
        text = " ".join(w["text"] for w in ws)
        # CJK: tesseract often emits chars with spaces; collapse for CJK runs
        text = collapse_cjk_spaces(text)
        h = statistics.median(w["height"] for w in ws)
        width = max(w["left"] + w["width"] for w in ws) - min(w["left"] for w in ws)
        out.append({"key": key, "text": text,
                    "height": h, "width": width,
                    "left": min(w["left"] for w in ws),
                    "conf": statistics.median(w["conf"] for w in ws)})
    return out


def collapse_cjk_spaces(text):
    import re
    # remove spaces between two CJK chars (tesseract word-split artefact)
    return re.sub(r"(?<=[\u4e00-\u9fff]) (?=[\u4e00-\u9fff])", "", text)


def body_height(lines):
    """Most common line height across lines; robust body-size estimate."""
    from collections import Counter
    cnt = Counter(round(l["height"]) for l in lines)
    # break ties toward the smaller height (body, not headings)
    best = sorted(cnt.items(), key=lambda kv: (-kv[1], kv[0]))[0][0]
    return float(best) or 1.0


def classify(lines, page_width_px, body_h=None, allow_title=False):
    """Attach a style name to each line: Title/Heading 1/Heading 2/Normal."""
    if not lines:
        return lines
    if body_h is None:
        body_h = body_height(lines)
    # title: first line of the document, clearly largest and short
    title_h = None
    l0 = lines[0]
    if allow_title and l0["height"] > body_h * 1.4 \
            and len(l0["text"].split()) <= 12 and len(l0["text"]) < 80:
        title_h = round(l0["height"])
    # candidate heading sizes: distinct heights clearly above body,
    # excluding the title's own size so H1 isn't swallowed by the title
    sizes = sorted({round(l["height"]) for l in lines
                    if l["height"] > body_h * 1.25
                    and round(l["height"]) != title_h},
                   reverse=True)
    for i, l in enumerate(lines):
        h = l["height"]
        words_n = len(l["text"].split())
        is_short = words_n <= 12 and len(l["text"]) < 80
        centered = abs((l["left"] + l["width"] / 2) - page_width_px / 2) < page_width_px * 0.2
        style = "Normal"
        if title_h is not None and i == 0:
            style = "Title"
        elif sizes and round(h) == sizes[0] and is_short:
            style = "Heading 1"
        elif len(sizes) > 1 and round(h) == sizes[1] and is_short:
            style = "Heading 2"
        elif h > body_h * 1.25 and is_short and centered and i < 6:
            style = "Heading 1"  # centered section titles on early pages
        l["style"] = style
    return lines


def merge_paragraphs(lines):
    """Join consecutive Normal lines in the same block/par into paragraphs."""
    out = []
    buf = None
    for l in lines:
        if l["style"] != "Normal":
            if buf:
                out.append(buf)
                buf = None
            out.append(l)
            continue
        key = l["key"][:3]  # page, block, par
        if buf and buf["key"][:3] == key:
            sep = "" if needs_no_space(buf["text"], l["text"]) else " "
            buf["text"] = buf["text"] + sep + l["text"]
        else:
            if buf:
                out.append(buf)
            buf = dict(l)
    if buf:
        out.append(buf)
    return out


def needs_no_space(a, b):
    import re
    # CJK text flows without spaces
    return bool(re.search(r"[\u4e00-\u9fff]$", a or "")) or \
           bool(re.search(r"^[\u4e00-\u9fff]", b or ""))


def ocr_convert(pdf_path, project_dir, profile_name="default"):
    """Full Mode-B pipeline. Writes output.docx in project_dir."""
    log = []
    import sys
    sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
    from pipeline import load_profile, _hex_rgb  # reuse style profiles
    from docx.enum.text import WD_ALIGN_PARAGRAPH

    pages_dir = os.path.join(project_dir, "ocr_pages")
    imgs = render_pages(pdf_path, pages_dir)
    log.append(f"rendered {len(imgs)} pages at {DPI} DPI")

    try:
        subprocess.run(["tesseract", "--version"], capture_output=True,
                       check=True, timeout=10)
    except Exception:
        return {"ok": False, "log": log,
                "error": "tesseract not installed on server"}

    all_lines = []
    page_w = 2480  # A4 @300dpi fallback
    page_lines = []
    for img in imgs:
        try:
            pix_w = pymupdf.Pixmap(img).width
            page_w = pix_w
        except Exception:
            pass
        words = [w for w in ocr_tsv(img) if w["conf"] >= 30]
        page_lines.append(build_lines(words))
    # global body height across all pages (per-page mode breaks on
    # heading-heavy pages)
    flat = [l for pls in page_lines for l in pls]
    g_body = body_height(flat) if flat else 1.0
    for i, lines in enumerate(page_lines):
        all_lines.extend(classify(lines, page_w, body_h=g_body,
                                  allow_title=(i == 0)))
    log.append(f"OCR lines: {len(all_lines)}")
    if not all_lines:
        return {"ok": False, "log": log, "error": "no text recognized"}

    paras = merge_paragraphs(all_lines)
    n_head = sum(1 for p in paras if p["style"] != "Normal")
    log.append(f"paragraphs: {len(paras)}, headings detected: {n_head}")

    profile = load_profile(profile_name)
    doc = Document()
    styles = doc.styles
    # ensure our profile styles exist with the right look
    cfg_map = profile["styles"]
    for sname, cfg in cfg_map.items():
        if sname not in styles:
            continue
        s = styles[sname]
        s.font.name = cfg.get("font", "Calibri")
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
        align = {"center": WD_ALIGN_PARAGRAPH.CENTER,
                 "right": WD_ALIGN_PARAGRAPH.RIGHT,
                 "justify": WD_ALIGN_PARAGRAPH.JUSTIFY}.get(cfg.get("align"))
        if align is not None:
            pf.alignment = align

    for p in paras:
        if not p["text"].strip():
            continue
        doc.add_paragraph(p["text"], style=p["style"])

    out = os.path.join(project_dir, "output.docx")
    doc.save(out)
    log.append("wrote " + out)
    return {"ok": True, "log": log, "docx": out}
