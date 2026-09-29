"""AI layout pass for OCR mode (tex2word).

When the GEMINI_API_KEY environment variable is set, each page's layout is
reconstructed by a vision LLM (Google Gemini) from:
  * the rendered page image,
  * the OCR text lines with bounding boxes (from tesseract),
  * the embedded raster images with bounding boxes (from PyMuPDF),

so the .docx mirrors the PDF's original layout: reading order,
multi-column flow, real Word tables, and figures placed where they appear.

Any failure (no key, API error, malformed JSON) raises, and the caller
(ocr_pipeline.ocr_convert) falls back to the classic heuristic pipeline.
The key is read from the environment on every conversion (never stored).
"""

import base64
import json
import os
import statistics
import urllib.error
import urllib.request

import pymupdf
from docx import Document
from docx.enum.section import WD_SECTION
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.shared import Inches, Pt

LLM_DPI = 120  # render size for the vision call: small = fast & cheap
GEMINI_MODELS = ["gemini-2.5-flash", "gemini-2.0-flash"]


# --------------------------------------------------------------------------
# inputs: OCR lines -> normalized boxes, embedded images -> files + boxes
# --------------------------------------------------------------------------

def page_line_boxes(page_lines, page_w_px, page_h_px):
    """build_lines() output -> [{id, text, box:[x0,y0,x1,y1] in 0-1000}]."""
    out = []
    for i, l in enumerate(page_lines):
        x0 = l["left"] / page_w_px * 1000
        y0 = l["top"] / page_h_px * 1000
        x1 = (l["left"] + l["width"]) / page_w_px * 1000
        y1 = (l["top"] + l["height"]) / page_h_px * 1000
        out.append({"id": i, "text": l["text"],
                    "box": [round(x0), round(y0), round(x1), round(y1)]})
    return out


def extract_images(pdf_path, out_dir):
    """Extract embedded raster images, one entry per page.

    Returns a list (one per PDF page) of lists of
    {id, file, box:[x0,y0,x1,y1] 0-1000, w_pt, h_pt}.
    """
    os.makedirs(out_dir, exist_ok=True)
    doc = pymupdf.open(pdf_path)
    all_pages = []
    for pno, page in enumerate(doc):
        pw, ph = page.rect.width, page.rect.height
        imgs = []
        seen = set()
        for item in page.get_images(full=True):
            xref = item[0]
            if xref in seen:
                continue
            seen.add(xref)
            try:
                # NOTE: pass the full item tuple, not the bare xref —
                # get_image_bbox(xref) raises "bad image name" here
                bbox = page.get_image_bbox(item)
            except Exception:
                continue
            try:
                pix = pymupdf.Pixmap(doc, xref)
            except Exception:
                continue
            if pix.width < 8 or pix.height < 8:
                continue  # tracking pixels / noise
            fp = os.path.join(out_dir, f"img_p{pno:03d}_{xref}.png")
            try:
                with open(fp, "wb") as fh:
                    fh.write(pix.tobytes("png"))
            except Exception:
                continue
            imgs.append({
                "id": len(imgs), "file": fp,
                "box": [round(bbox.x0 / pw * 1000),
                        round(bbox.y0 / ph * 1000),
                        round(bbox.x1 / pw * 1000),
                        round(bbox.y1 / ph * 1000)],
                "w_pt": bbox.width, "h_pt": bbox.height,
            })
        all_pages.append(imgs)
    doc.close()
    return all_pages


def render_small(pdf_path, index, out_dir, dpi=LLM_DPI):
    """Render one page small for the vision call; returns the PNG path."""
    os.makedirs(out_dir, exist_ok=True)
    doc = pymupdf.open(pdf_path)
    pix = doc[index].get_pixmap(dpi=dpi)
    fp = os.path.join(out_dir, f"llm_{index:03d}.png")
    tmp = fp + ".tmp"
    with open(tmp, "wb") as fh:
        fh.write(pix.tobytes("png"))
    os.replace(tmp, fp)
    doc.close()
    return fp


# --------------------------------------------------------------------------
# Gemini vision call
# --------------------------------------------------------------------------

PROMPT = """You are a document layout analyst. The image shows one page of a scanned document (page {pno} of {total}). You are also given OCR text lines with bounding boxes and embedded figure images with bounding boxes. Coordinates are 0-1000, origin top-left.

Goal: describe the page's layout as JSON so the page can be rebuilt in Microsoft Word with the same reading order, columns, tables and figure placement.

Return ONLY this JSON (no markdown, no commentary):
{{
  "columns": 1,
  "blocks": [ ... ],
  "skip_line_ids": [ ... ]
}}

Block kinds, listed in READING ORDER (top to bottom; on multi-column pages finish one column top-to-bottom before starting the next):
- {{"kind": "heading", "level": 1, "line_ids": [3]}} — level 1 = main section titles, 2 = subsections, 3 = smaller heads
- {{"kind": "paragraph", "line_ids": [4, 5, 6]}}
- {{"kind": "bullets", "items": [[7], [8, 9]]}} — each item is the line ids forming one bullet or numbered item
- {{"kind": "table", "line_ids": [10, 11, 12], "rows": [["a", "b"], ["c", "d"]]}} — reconstruct the table grid from the image
- {{"kind": "image", "image_id": 0, "caption_line_ids": [13]}} — image_id refers to the embedded-image list

Rules:
- "columns" is 1 or 2: the number of text columns on this page.
- Use each line id at most once across all blocks. Page numbers, running headers/footers and OCR noise go in "skip_line_ids".
- Table cell text MUST be copied VERBATIM from the given OCR line texts — never invent, rephrase, translate or "fix" text. Split one line into cells only at clear column gaps visible in the image. Every row must have the same number of cells (pad with "").
- A figure caption belongs to its figure's image block via caption_line_ids.
- If the page is a full-page figure or has no text lines, "blocks" may contain only image blocks.

OCR lines:
{lines}

Embedded images:
{images}
"""


def _gemini_post(model, api_key, payload):
    url = ("https://generativelanguage.googleapis.com/v1beta/models/"
           f"{model}:generateContent?key={api_key}")
    req = urllib.request.Request(
        url, data=json.dumps(payload).encode("utf-8"),
        headers={"Content-Type": "application/json"})
    try:
        with urllib.request.urlopen(req, timeout=180) as r:
            return json.loads(r.read().decode("utf-8")), None
    except urllib.error.HTTPError as e:
        return None, e.code
    except Exception as e:
        raise RuntimeError(f"Gemini request failed: {e}")


def gemini_page_layout(api_key, png_path, line_boxes, image_boxes,
                       pno, total):
    """Ask Gemini for one page's layout JSON. Returns the parsed dict."""
    with open(png_path, "rb") as fh:
        img_b64 = base64.b64encode(fh.read()).decode("ascii")
    prompt = PROMPT.format(
        pno=pno + 1, total=total,
        lines=json.dumps(line_boxes, ensure_ascii=False),
        images=json.dumps(image_boxes, ensure_ascii=False))
    payload = {
        "contents": [{"parts": [
            {"text": prompt},
            {"inline_data": {"mime_type": "image/png",
                             "data": img_b64}}]}],
        "generationConfig": {"responseMimeType": "application/json",
                             "temperature": 0.1,
                             "maxOutputTokens": 16384},
    }
    last_code = None
    for model in GEMINI_MODELS:
        data, code = _gemini_post(model, api_key, payload)
        if data is not None:
            try:
                text = data["candidates"][0]["content"]["parts"][0]["text"]
                return json.loads(text)
            except (KeyError, IndexError, ValueError) as e:
                raise RuntimeError(f"Gemini returned unusable JSON: {e}")
        last_code = code
        if code != 404:
            raise RuntimeError(f"Gemini API error (HTTP {code})")
    raise RuntimeError(f"Gemini models unavailable (HTTP {last_code})")


# --------------------------------------------------------------------------
# validation (strict: any doubt -> caller falls back to heuristics)
# --------------------------------------------------------------------------

def _check_ids(ids, n_lines, seen, what="line_ids"):
    if not isinstance(ids, list):
        raise ValueError(f"bad {what}")
    out = []
    for v in ids:
        try:
            v = int(v)
        except (TypeError, ValueError):
            raise ValueError(f"bad {what} entry {v!r}")
        if not 0 <= v < n_lines:
            raise ValueError(f"{what} id {v} out of range")
        if v in seen:
            raise ValueError(f"line id {v} used twice")
        seen.add(v)
        out.append(v)
    return out


def validate_layout(layout, n_lines, n_images):
    """Validate + normalize one page's layout dict. Raises ValueError."""
    if not isinstance(layout, dict):
        raise ValueError("layout is not an object")
    cols = layout.get("columns", 1)
    if cols not in (1, 2):
        raise ValueError(f"bad columns value {cols!r}")
    blocks = layout.get("blocks")
    if not isinstance(blocks, list) or not blocks:
        raise ValueError("blocks must be a non-empty list")
    seen_lines, seen_imgs = set(), set()
    clean = []
    for b in blocks:
        if not isinstance(b, dict):
            raise ValueError("block is not an object")
        kind = b.get("kind")
        if kind == "heading":
            level = b.get("level", 1)
            if level not in (1, 2, 3):
                raise ValueError(f"bad heading level {level!r}")
            clean.append({"kind": "heading", "level": level,
                          "line_ids": _check_ids(b.get("line_ids"),
                                                 n_lines, seen_lines)})
        elif kind == "paragraph":
            clean.append({"kind": "paragraph",
                          "line_ids": _check_ids(b.get("line_ids"),
                                                 n_lines, seen_lines)})
        elif kind == "bullets":
            items = b.get("items")
            if not isinstance(items, list) or not items:
                raise ValueError("bullets need non-empty items")
            clean.append({"kind": "bullets",
                          "items": [_check_ids(it, n_lines, seen_lines,
                                              "bullet item")
                                    for it in items]})
        elif kind == "table":
            rows = b.get("rows")
            if not isinstance(rows, list) or not rows:
                raise ValueError("table needs non-empty rows")
            width = None
            for r in rows:
                if not isinstance(r, list) or not r:
                    raise ValueError("table row must be a non-empty list")
                if width is None:
                    width = len(r)
                elif len(r) != width:
                    raise ValueError("ragged table rows")
                for c in r:
                    if not isinstance(c, str):
                        raise ValueError("table cell must be a string")
            clean.append({"kind": "table",
                          "line_ids": _check_ids(b.get("line_ids"),
                                                 n_lines, seen_lines),
                          "rows": rows})
        elif kind == "image":
            iid = b.get("image_id")
            if not isinstance(iid, int) or not 0 <= iid < n_images:
                raise ValueError(f"bad image_id {iid!r}")
            if iid in seen_imgs:
                raise ValueError(f"image {iid} used twice")
            seen_imgs.add(iid)
            clean.append({
                "kind": "image", "image_id": iid,
                "caption_line_ids": _check_ids(
                    b.get("caption_line_ids", []), n_lines, seen_lines,
                    "caption_line_ids")})
        else:
            raise ValueError(f"unknown block kind {kind!r}")
    skip = layout.get("skip_line_ids", []) or []
    skip_ids = set()
    for v in skip:
        try:
            v = int(v)
        except (TypeError, ValueError):
            raise ValueError(f"bad skip_line_id {v!r}")
        if 0 <= v < n_lines and v not in seen_lines:
            skip_ids.add(v)
    return {"columns": cols, "blocks": clean, "skip": skip_ids,
            "seen_images": seen_imgs}


# --------------------------------------------------------------------------
# .docx builder: native Word constructs mirroring the PDF layout
# --------------------------------------------------------------------------

def _set_columns(section, n):
    cols = section._sectPr.xpath("./w:cols")
    el = cols[0] if cols else section._sectPr.makeelement(qn("w:cols"), {})
    if not cols:
        section._sectPr.append(el)
    el.set(qn("w:num"), str(n))
    el.set(qn("w:space"), "720")  # 0.5" gutter


def _cjk(run, name="Calibri"):
    """Set both ascii and East-Asian fonts so CJK text renders properly."""
    run.font.name = name
    try:
        run._element.rPr.rFonts.set(qn("w:eastAsia"), name)
    except Exception:
        pass


def _px_to_pt(px, dpi):
    return max(6.0, min(px * 72.0 / dpi, 72.0))


def _median_height(lines, ids, dpi):
    hs = [lines[i]["height"] for i in ids if i < len(lines)]
    if not hs:
        return 11.0
    return _px_to_pt(statistics.median(hs), dpi)


def _join_text(lines, ids):
    parts = [lines[i]["text"] for i in ids if i < len(lines)]
    # CJK-aware join: no space between CJK chars
    out = ""
    for p in parts:
        if not p.strip():
            continue
        if out and not _cjk_boundary(out, p):
            out += " "
        out += p
    return out.strip()


def _cjk_boundary(a, b):
    import re
    return bool(re.search(r"[\u4e00-\u9fff]$", a)) or \
        bool(re.search(r"^[\u4e00-\u9fff]", b))


def build_ai_docx(layouts, page_images, page_lines, page_dpis, out_path,
                   log, base_font="Calibri"):
    """Build the Word document from validated per-page layouts."""
    doc = Document()
    cur_cols = None
    for pno, layout in enumerate(layouts):
        cols = layout["columns"]
        if pno == 0:
            _set_columns(doc.sections[0], cols)
        elif cols != cur_cols:
            doc.add_section(WD_SECTION.CONTINUOUS)
            _set_columns(doc.sections[-1], cols)
        cur_cols = cols

        lines = page_lines[pno]
        dpi = page_dpis[pno]
        # body size from paragraph/bullet/table lines on this page
        body_ids = [i for b in layout["blocks"]
                    for i in (b.get("line_ids", []) if b["kind"] != "bullets"
                              else [x for it in b["items"] for x in it])]
        body_pt = _median_height(lines, body_ids, dpi) if body_ids else 11.0
        body_pt = max(8.0, min(body_pt, 13.0))

        imgs = {im["id"]: im for im in page_images[pno]}
        sec = doc.sections[-1]
        col_w_in = ((sec.page_width - sec.left_margin - sec.right_margin)
                    / cols / 914400.0)

        for b in layout["blocks"]:
            kind = b["kind"]
            if kind == "heading":
                text = _join_text(lines, b["line_ids"])
                if not text:
                    continue
                h = doc.add_heading(level=min(b["level"], 3))
                run = h.add_run(text)
                run.bold = True
                run.font.size = Pt(
                    min(max(_median_height(lines, b["line_ids"], dpi),
                                body_pt * 1.15), 28.0))
                _cjk(run, base_font)
            elif kind == "paragraph":
                text = _join_text(lines, b["line_ids"])
                if not text:
                    continue
                p = doc.add_paragraph()
                run = p.add_run(text)
                run.font.size = Pt(body_pt)
                _cjk(run, base_font)
            elif kind == "bullets":
                for item_ids in b["items"]:
                    text = _join_text(lines, item_ids)
                    if not text:
                        continue
                    p = doc.add_paragraph(style="List Bullet")
                    run = p.add_run(text)
                    run.font.size = Pt(body_pt)
                    _cjk(run, base_font)
            elif kind == "table":
                rows = b["rows"]
                ncols = max(len(r) for r in rows)
                t = doc.add_table(rows=len(rows), cols=ncols)
                t.style = "Table Grid"
                for ri, row in enumerate(rows):
                    for ci in range(ncols):
                        cell = t.cell(ri, ci)
                        cell.text = ""
                        txt = row[ci] if ci < len(row) else ""
                        if not txt.strip():
                            continue
                        par = cell.paragraphs[0]
                        run = par.add_run(txt)
                        run.font.size = Pt(max(body_pt - 1.0, 7.0))
                        _cjk(run, base_font)
                        if ri == 0:
                            # header shading
                            tcPr = cell._tc.get_or_add_tcPr()
                            shd = tcPr.makeelement(
                                qn("w:shd"),
                                {qn("w:val"): "clear",
                                 qn("w:fill"): "D9E2F3"})
                            tcPr.append(shd)
                doc.add_paragraph().add_run().add_break()  # breathing room
            elif kind == "image":
                im = imgs.get(b["image_id"])
                if im and os.path.exists(im["file"]):
                    p = doc.add_paragraph()
                    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    w_in = min(im["w_pt"] / 72.0, col_w_in)
                    try:
                        p.add_run().add_picture(
                            im["file"], width=Inches(max(w_in, 0.5)))
                    except Exception as e:
                        log.append(f"page {pno + 1}: image skipped ({e})")
                for cid in b["caption_line_ids"]:
                    if cid < len(lines):
                        cap = doc.add_paragraph()
                        cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
                        run = cap.add_run(lines[cid]["text"])
                        run.italic = True
                        run.font.size = Pt(max(body_pt - 1.5, 7.0))
                        _cjk(run, base_font)

        # safety net: lines the LLM didn't place (and didn't skip) are
        # appended as plain paragraphs so no content is silently lost
        used = {i for b in layout["blocks"]
                for i in (b.get("line_ids", []) if b["kind"] != "bullets"
                          else [x for it in b["items"] for x in it])
                } | {c for b in layout["blocks"]
                      for c in b.get("caption_line_ids", [])}
        for i, l in enumerate(lines):
            if i not in used and i not in layout["skip"] and l["text"].strip():
                p = doc.add_paragraph()
                run = p.add_run(l["text"])
                run.font.size = Pt(body_pt)
                _cjk(run, base_font)

        # figures the LLM didn't place go at the end of the page
        for im in page_images[pno]:
            if im["id"] not in layout["seen_images"] \
                    and os.path.exists(im["file"]):
                p = doc.add_paragraph()
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                w_in = min(im["w_pt"] / 72.0, col_w_in)
                try:
                    p.add_run().add_picture(
                        im["file"], width=Inches(max(w_in, 0.5)))
                except Exception:
                    pass

        if pno < len(layouts) - 1:
            doc.add_page_break()

    doc.save(out_path)
    log.append(f"AI layout docx written: {out_path}")


# --------------------------------------------------------------------------
# entry point used by ocr_pipeline
# --------------------------------------------------------------------------

def ai_layout_convert(pdf_path, project_dir, pages_dir, page_lines,
                      page_sizes, page_dpis, imgs, progress_cb, log):
    """Full AI-layout conversion. Raises on any failure (caller falls back).

    Returns the output .docx path.
    """
    api_key = os.environ.get("GEMINI_API_KEY", "")
    if not api_key:
        raise RuntimeError("GEMINI_API_KEY is not set")
    n = len(imgs)

    img_dir = os.path.join(pages_dir, "figures")
    page_images = extract_images(pdf_path, img_dir)
    n_img = sum(len(p) for p in page_images)
    log.append(f"AI layout: extracted {n_img} embedded images")

    llm_dir = os.path.join(pages_dir, "llm")
    layouts = []
    for i in range(n):
        small = render_small(pdf_path, i, llm_dir)
        boxes = page_line_boxes(page_lines[i], *page_sizes[i])
        imboxes = [{"id": im["id"], "box": im["box"]}
                   for im in page_images[i]]
        raw = gemini_page_layout(api_key, small, boxes, imboxes, i, n)
        layout = validate_layout(raw, len(boxes), len(imboxes))
        layouts.append(layout)
        log.append(f"AI layout: page {i + 1}/{n} — {layout['columns']} "
                   f"column(s), {len(layout['blocks'])} blocks")
        if progress_cb:
            try:
                progress_cb(n + i + 1, 2 * n,
                            f"AI layout: page {i + 1} of {n}")
            except Exception:
                pass

    out = os.path.join(project_dir, "output.docx")
    build_ai_docx(layouts, page_images, page_lines, page_dpis, out, log,
                  base_font="Calibri")
    return out
