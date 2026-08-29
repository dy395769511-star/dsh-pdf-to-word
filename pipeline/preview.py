#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""PIL-based DOCX page preview renderer (best-effort, for LLM verification).

Renders a .docx to per-page JPEGs using declared styles (font/size/alignment/
indent/line-spacing/tables/images). It does NOT simulate Word's reflow or
page-break decisions exactly; page breaks happen when content overflows the
page or on explicit breaks.
"""
import os
import re
import sys

sys.stdout.reconfigure(encoding="utf-8")

from PIL import Image, ImageDraw, ImageFont

FONT_FILES = {
    "宋体": ["simsun.ttc", "simkai.ttf"],
    "黑体": ["simhei.ttf", "simkai.ttf", "msyh.ttc"],
    "楷体": ["simkai.ttf", "simhei.ttf"],
    "仿宋": ["simfang.ttf", "simsun.ttc"],
    "微软雅黑": ["msyh.ttc", "msyhbd.ttc"],
    "等线": ["msyh.ttc", "simsun.ttc"],
    "Times New Roman": ["times.ttf", "timesbd.ttf"],
    "Arial": ["arial.ttf"],
    "Calibri": ["calibri.ttf", "arial.ttf"],
    "Courier New": ["cour.ttf", "consola.ttf"],
    "Lucida Console": ["lucon.ttf", "cour.ttf", "consola.ttf"],
    "Segoe UI": ["seguisb.ttf", "arial.ttf"],
}
FONT_DIR = r"C:\Windows\Fonts"
_font_cache = {}


def load_font(name, size_px, bold=False):
    """Returns (font, real_bold) where real_bold means the file itself is a bold face."""
    key = (name, int(size_px), bold)
    if key in _font_cache:
        return _font_cache[key]
    name = name or "Times New Roman"
    cands = FONT_FILES.get(name, [])
    if bold and name == "Times New Roman":
        cands = ["timesbd.ttf"] + cands
    for c in cands:
        p = os.path.join(FONT_DIR, c)
        if os.path.exists(p):
            try:
                f = ImageFont.truetype(p, max(6, int(size_px)))
                real_bold = ("bd" in c.lower()) or ("bold" in c.lower())
                _font_cache[key] = (f, real_bold)
                return (f, real_bold)
            except Exception:
                pass
    try:
        f = ImageFont.truetype(os.path.join(FONT_DIR, "arial.ttf"), max(6, int(size_px)))
    except Exception:
        f = ImageFont.load_default()
    _font_cache[key] = (f, False)
    return (f, False)


def is_cjk(ch):
    o = ord(ch)
    return (0x4E00 <= o <= 0x9FFF) or (0x3400 <= o <= 0x4DBF) or (0x3000 <= o <= 0x303F) \
        or (0xFF00 <= o <= 0xFFEF) or (0x2000 <= o <= 0x206F)


def wrap_line(text, font, max_w, draw):
    """Greedy wrap; CJK may break anywhere, latin breaks at spaces.

    Returns a list of (line_text, start_index_in_text) so callers can map
    drawn characters back to per-run styles without index drift.
    """
    lines = []
    cur = ""
    cur_start = 0
    for i, ch in enumerate(text):
        trial = cur + ch
        w = draw.textlength(trial, font=font)
        if w <= max_w or not cur:
            cur = trial
        else:
            if ch == " ":
                lines.append((cur, cur_start))
                cur = ""
                cur_start = i + 1
                continue
            lines.append((cur, cur_start))
            cur = ch
            cur_start = i
    if cur:
        lines.append((cur, cur_start))
    return lines or [("", 0)]


def _emu_to_pt(v):
    if v is None:
        return None
    try:
        return float(v) / 12700.0
    except Exception:
        return None


def dump_docx(docx_path):
    from docx import Document
    from docx.table import Table
    from docx.text.paragraph import Paragraph
    from docx.oxml.ns import qn

    doc = Document(docx_path)
    sec = doc.sections[0]
    info = {
        "page_w_pt": _emu_to_pt(sec.page_width),
        "page_h_pt": _emu_to_pt(sec.page_height),
        "ml_pt": _emu_to_pt(sec.left_margin),
        "mr_pt": _emu_to_pt(sec.right_margin),
        "mt_pt": _emu_to_pt(sec.top_margin),
        "mb_pt": _emu_to_pt(sec.bottom_margin),
        "blocks": [],
    }
    for child in doc.element.body.iterchildren():
        if child.tag == qn("w:p"):
            p = Paragraph(child, doc)
            runs = []
            for r in p.runs:
                rd = {"text": r.text or ""}
                rd["size_pt"] = _emu_to_pt(r.font.size) if r.font.size else None
                rd["bold"] = bool(r.font.bold)
                rd["italic"] = bool(r.font.italic)
                rd["underline"] = bool(r.font.underline)
                rd["font"] = r.font.name
                ea = None
                rPr = r._element.rPr
                if rPr is not None:
                    rFonts = rPr.find(qn("w:rFonts"))
                    if rFonts is not None:
                        ea = rFonts.get(qn("w:eastAsia"))
                rd["ea"] = ea
                if r.font.color and r.font.color.type is not None:
                    try:
                        rd["color"] = str(r.font.color.rgb)
                    except Exception:
                        rd["color"] = None
                else:
                    rd["color"] = None
                # image?
                blip = r._element.find('.//' + qn('a:blip'))
                if blip is not None:
                    rid = blip.get(qn('r:embed'))
                    try:
                        part = doc.part.related_parts[rid]
                        ext = os.path.splitext(part.partname)[1] or ".png"
                        rd["image"] = part.blob + b""  # keep bytes in memory
                        ext_el = r._element.find('.//' + qn('wp:extent'))
                        if ext_el is not None:
                            rd["image_w_pt"] = _emu_to_pt(ext_el.get("cx"))
                            rd["image_h_pt"] = _emu_to_pt(ext_el.get("cy"))
                    except Exception:
                        rd["image"] = None
                runs.append(rd)
            # explicit page break?
            page_break = False
            for br in p._element.iter():
                if br.tag == qn("w:br"):
                    if br.get(qn("w:type")) == "page":
                        page_break = True
            pf = p.paragraph_format
            style = p.style.name if p.style else ""
            bg_fill = None
            pPr = p._p.pPr
            if pPr is not None:
                shd_el = pPr.find(qn("w:shd"))
                if shd_el is not None and shd_el.get(qn("w:fill")) not in (None, "auto"):
                    bg_fill = shd_el.get(qn("w:fill"))
            borders = {}
            if pPr is not None:
                bdr_el = pPr.find(qn("w:pBdr"))
                if bdr_el is not None:
                    for side_el in bdr_el:
                        side = side_el.tag.split("}")[1]
                        if side_el.get(qn("w:val")) in (None, "nil"):
                            continue
                        col = side_el.get(qn("w:color"))
                        if not col or col == "auto":
                            continue
                        try:
                            w_pt = int(side_el.get(qn("w:sz") or 0)) / 8.0
                        except Exception:
                            w_pt = 0.75
                        borders[side] = {"color": col, "w_pt": w_pt}
            info["blocks"].append({
                "type": "para",
                "runs": runs,
                "style": style,
                "align": str(pf.alignment) if pf.alignment is not None else None,
                "first_indent_pt": _emu_to_pt(pf.first_line_indent),
                "line_spacing": pf.line_spacing if isinstance(pf.line_spacing, float) else 1.0,
                "exact_line_h_pt": (_emu_to_pt(pf.line_spacing)
                                    if pf.line_spacing is not None
                                    and not isinstance(pf.line_spacing, float) else None),
                "space_before_pt": _emu_to_pt(pf.space_before),
                "space_after_pt": _emu_to_pt(pf.space_after),
                "page_break": page_break,
                "bg_fill": bg_fill,
                "borders": borders,
            })
        elif child.tag == qn("w:tbl"):
            t = Table(child, doc)
            style = t.style.name if t.style else ""
            # grid column widths from w:tblGrid (twips -> pt)
            grid_w = []
            try:
                tblGrid = t._tbl.find(qn("w:tblGrid"))
                if tblGrid is not None:
                    for gc in list(tblGrid):
                        wv = gc.get(qn("w:w"))
                        grid_w.append(int(wv) / 20.0 if wv else None)
            except Exception:
                grid_w = []
            rows = []
            for row in t.rows:
                cells = []
                prev_tc = None
                for cell in row.cells:
                    tc = cell._tc
                    ctexts = []
                    csize = None
                    cbold = False
                    for cp in cell.paragraphs:
                        ct = cp.text
                        if ct:
                            ctexts.append(ct)
                        for r in cp.runs:
                            if r.font.size:
                                csize = _emu_to_pt(r.font.size)
                                break
                            if r.font.bold:
                                cbold = True
                        if csize:
                            break
                    if cells and tc is prev_tc:
                        # gridSpan: same cell object repeats per covered grid col
                        cells[-1]["span"] += 1
                        continue
                    cells.append({"text": "\n".join(ctexts), "size_pt": csize,
                                  "bold": cbold, "span": 1})
                    prev_tc = tc
                rows.append(cells)
            info["blocks"].append({"type": "table", "rows": rows, "style": style,
                                   "grid_w_pt": grid_w})
    return info


ALIGN_LEFT = 0
ALIGN_CENTER = 1
ALIGN_RIGHT = 2
ALIGN_JUSTIFY = 3


def _align_code(s):
    if not s:
        return ALIGN_LEFT
    s = s.upper()
    if "CENTER" in s:
        return ALIGN_CENTER
    if "RIGHT" in s:
        return ALIGN_RIGHT
    if "JUSTIFY" in s or "DISTRIBUTE" in s:
        return ALIGN_JUSTIFY
    return ALIGN_LEFT


def render_docx_pages(docx_path, assets_dir, dpi=140):
    info = dump_docx(docx_path)
    scale = dpi / 72.0
    W = int((info["page_w_pt"] or 595) * scale)
    H = int((info["page_h_pt"] or 842) * scale)
    ml = int((info["ml_pt"] or 56) * scale)
    mt = int((info["mt_pt"] or 56) * scale)
    usable_w = W - ml - int((info["mr_pt"] or 56) * scale)
    usable_h = H - mt - int((info["mb_pt"] or 56) * scale)

    pages = []
    cur = None
    draw = None
    y = 0

    def new_page():
        nonlocal cur, draw, y
        cur = Image.new("RGB", (W, H), "white")
        draw = ImageDraw.Draw(cur)
        y = mt
        pages.append((cur, []))
        return cur

    new_page()

    def ensure(h):
        nonlocal y
        if y + h > mt + usable_h:
            new_page()
            return True
        return False

    def page_text(s):
        pages[-1][1].append(s)

    for block in info["blocks"]:
        if block["type"] == "para":
            if block.get("page_break"):
                new_page()
            sb = int((block.get("space_before_pt") or 0) * scale)
            sa = int((block.get("space_after_pt") or 0) * scale)
            runs = [r for r in block["runs"] if (r["text"] or r.get("image"))]
            if not runs:
                # empty paragraph: advance by its exact line height when set
                # (shaded code-box gap filler), else skip as before
                h_pt = block.get("exact_line_h_pt") or 0.0
                if h_pt > 0:
                    h = int(h_pt * scale)
                    ensure(h)
                    bg = block.get("bg_fill")
                    if bg:
                        draw.rectangle([int(ml), int(y), int(ml + usable_w), int(y + h)],
                                       fill="#" + bg)
                    bdr = block.get("borders")
                    if bdr:
                        for side in ("left", "right"):
                            b = bdr.get(side)
                            if not b:
                                continue
                            bw = max(1, int(round((b.get("w_pt") or 0.75) * scale)))
                            xx = int(ml) if side == "left" else int(ml + usable_w)
                            draw.line((xx, int(y), xx, int(y + h)),
                                      fill="#" + b["color"], width=bw)
                    y += h
                continue
            # pick dominant run for metrics
            dom = runs[0]
            size_pt = dom.get("size_pt") or 10.5
            if block.get("style", "").startswith("Heading"):
                size_pt = dom.get("size_pt") or max(size_pt, 16)
            font_size_px = size_pt * scale
            ls = block.get("line_spacing")
            try:
                ls = float(ls) if ls is not None else 1.0
            except Exception:
                ls = 1.0
            line_h = font_size_px * max(1.0, ls) + 2
            indent = int((block.get("first_indent_pt") or 0) * scale)
            align = _align_code(block.get("align"))

            # build full text with per-char style map
            chars = []  # (ch, fontkey)
            for r in runs:
                if r.get("image"):
                    chars.append(("__IMG__", r))
                    continue
                for ch in (r["text"] or ""):
                    chars.append((ch, r))
            # merge text runs: wrap by segments
            segments = []
            for ch, r in chars:
                if ch == "__IMG__":
                    segments.append({"image": r.get("image"),
                                     "w_pt": r.get("image_w_pt"),
                                     "h_pt": r.get("image_h_pt")})
                elif segments and segments[-1].get("text") is not None and segments[-1]["r"] is r:
                    segments[-1]["text"] += ch
                else:
                    segments.append({"text": ch, "r": r})

            # pre-render lines: wrap each text segment independently is wrong for
            # mixed runs; instead wrap the whole paragraph with the dominant font,
            # drawing each segment with its own font.
            full = "".join(s["text"] if s.get("text") is not None else "\uFFFC" for s in segments)
            img_only = all(s.get("text") is None for s in segments)
            if not full.strip() and not img_only:
                continue
            # build per-char font list for the whole paragraph
            char_fonts = []
            for s in segments:
                if s.get("text") is None:
                    char_fonts.append(s)
                else:
                    for ch in s["text"]:
                        char_fonts.append(s["r"])
            # wrap using dominant font width
            dom_r = runs[0]
            dom_font, _dom_rb = load_font(dom_r.get("ea") or dom_r.get("font"), font_size_px, bool(dom_r.get("bold")))
            raw_lines = wrap_line(full, dom_font, usable_w - indent, draw)
            text_lines = [t for t, _ in raw_lines]
            line_starts = [st for _, st in raw_lines]
            # line height from the real font metrics (Word multipliers scale
            # the single-line height, not the em size)
            try:
                _asc, _desc = dom_font.getmetrics()
            except Exception:
                _asc, _desc = int(font_size_px * 1.1), int(font_size_px * 0.15)
            line_h = max(int(font_size_px * max(1.0, ls) + 2),
                         int((_asc + _desc) * max(1.0, ls)) + 2)
            y += sb
            para_top = y
            _pg0 = len(pages)
            _draw0 = draw
            line_img_max = max((int((s.get("h_pt") or (s.get("w_pt") or 100) * 0.6) * scale)
                                for s in segments if s.get("text") is None), default=0)
            para_text = full
            # paragraph background shading (code boxes) — draw before the text
            bg = block.get("bg_fill")
            if bg:
                _adv_sum = 0
                for line in text_lines:
                    if not line.strip():
                        _adv_sum += line_h
                    else:
                        _adv = line_h
                        if line_img_max and "\uFFFC" in line:
                            _adv = max(line_h, line_img_max + 4)
                        _adv_sum += _adv
                draw.rectangle([int(ml), int(y), int(ml + usable_w), int(y + _adv_sum)],
                               fill="#" + bg)
            for li, line in enumerate(text_lines):
                if not line.strip():
                    ensure(line_h)
                    y += line_h
                    continue
                if li == 0:
                    page_text(para_text)
                ensure(max(line_h, line_img_max + 4))
                # per-line width
                lw = draw.textlength(line, font=dom_font)
                if align == ALIGN_CENTER:
                    x0 = ml + (usable_w - lw) / 2
                elif align == ALIGN_RIGHT:
                    x0 = ml + usable_w - lw
                elif align == ALIGN_JUSTIFY and li < len(text_lines) - 1:
                    x0 = ml
                else:
                    x0 = ml
                if li == 0 and indent:
                    x0 += indent
                # draw chars with individual fonts
                ci = 0
                x = x0
                start = line_starts[li]
                line_img_h = 0
                for ch in line:
                    r = char_fonts[start + ci] if start + ci < len(char_fonts) else None
                    ci += 1
                    if r is None or r.get("image"):
                        # image placeholder: draw the image
                        img_r = r if (r is not None and r.get("image")) else None
                        if img_r is None:
                            x += 40
                            continue
                        try:
                            from PIL import Image as _I
                            import io as _io
                            im = _I.open(_io.BytesIO(img_r["image"]))
                            w_pt = img_r.get("w_pt") or 100
                            h_pt = img_r.get("h_pt") or w_pt * 0.6
                            iw = int(w_pt * scale)
                            ih = int(h_pt * scale)
                            if img_only and align == ALIGN_CENTER:
                                x = int(ml + (usable_w - max(10, iw)) / 2)
                            elif img_only and align == ALIGN_RIGHT:
                                x = int(ml + usable_w - max(10, iw))
                            im = im.convert("RGB").resize((max(10, iw), max(10, ih)))
                            cur.paste(im, (int(x), int(y)))
                            x += max(10, iw)
                            line_img_h = max(line_img_h, max(10, ih))
                        except Exception:
                            x += 40
                        continue
                    fnt, fnt_rb = load_font(r.get("ea") or r.get("font"), font_size_px, bool(r.get("bold")))
                    color = "#000000"
                    if r.get("color"):
                        color = "#" + r["color"]
                    sw = 1 if (bool(r.get("bold")) and not fnt_rb) else 0
                    draw.text((x, y), ch, font=fnt, fill=color, stroke_width=sw, stroke_fill=color)
                    if r.get("underline"):
                        draw.line((x, y + line_h - 3, x + fnt.getlength(ch), y + line_h - 3), fill=color)
                    x += fnt.getlength(ch)
                y += max(line_h, line_img_h + 4)
            # paragraph borders (code-box frames)
            _borders = block.get("borders")
            if _borders:
                _bw = 1
                for _b in _borders.values():
                    _bw = max(_bw, max(1, int(round((_b.get("w_pt") or 0.75) * scale))))
                _para_bottom = y
                _pg_last = len(pages)
                if _pg_last == _pg0:
                    for _side, _b in _borders.items():
                        _col = "#" + _b["color"]
                        if _side == "top":
                            draw.line([int(ml), int(para_top), int(ml + usable_w), int(para_top)],
                                       fill=_col, width=_bw)
                        elif _side == "bottom":
                            draw.line([int(ml), int(_para_bottom), int(ml + usable_w), int(_para_bottom)],
                                       fill=_col, width=_bw)
                        elif _side == "left":
                            draw.line([int(ml), int(para_top), int(ml), int(_para_bottom)],
                                       fill=_col, width=_bw)
                        elif _side == "right":
                            draw.line([int(ml + usable_w), int(para_top), int(ml + usable_w), int(_para_bottom)],
                                       fill=_col, width=_bw)
                else:
                    _page_h = mt + usable_h
                    for _side, _b in _borders.items():
                        _col = "#" + _b["color"]
                        if _side == "top":
                            _draw0.line([int(ml), int(para_top), int(ml + usable_w), int(para_top)],
                                        fill=_col, width=_bw)
                        elif _side == "bottom":
                            draw.line([int(ml), int(_para_bottom), int(ml + usable_w), int(_para_bottom)],
                                       fill=_col, width=_bw)
                        elif _side in ("left", "right"):
                            _xx = int(ml) if _side == "left" else int(ml + usable_w)
                            _draw0.line([_xx, int(para_top), _xx, int(_page_h)], fill=_col, width=_bw)
                            draw.line([_xx, int(mt), _xx, int(_para_bottom)], fill=_col, width=_bw)
            y += sa
        elif block["type"] == "table":
            rows = block["rows"]
            if not rows:
                continue
            s = scale
            grid_w = [w for w in (block.get("grid_w_pt") or []) if w]
            if not grid_w:
                n_c = max(len(r) for r in rows)
                grid_w = [usable_w / scale / n_c] * n_c
            n_c = len(grid_w)
            col_ws = [max(10, int(w * s)) for w in grid_w]
            # wrap cell texts (span-aware)
            cell_lines = []
            for row in rows:
                crow = []
                ci = 0
                for c in row:
                    if ci >= n_c:
                        break
                    span = int(c.get("span", 1))
                    wsum = sum(col_ws[ci:ci + span]) or usable_w
                    size_pt = c.get("size_pt") or 10.5
                    fnt, fnt_rb = load_font("宋体", size_pt * s, bool(c.get("bold")))
                    cw = max(20, wsum - 8)
                    lines = []
                    for piece in (c.get("text") or "").split("\n"):
                        lines.extend(t for t, _ in wrap_line(piece, fnt, cw, draw))
                    crow.append({"lines": lines or [""], "size_pt": size_pt,
                                 "bold": bool(c.get("bold")), "real_bold": fnt_rb,
                                 "x": ci, "span": span})
                    ci += span
                cell_lines.append(crow)
            # row heights
            row_hs = []
            for crow in cell_lines:
                maxn = max((len(c["lines"]) for c in crow), default=1)
                maxsz = max((c["size_pt"] for c in crow), default=10.5)
                row_hs.append(max(18, int(maxn * (maxsz * s * 1.3 + 6))))
            total_h = sum(row_hs) + 4
            ensure(min(total_h, 60))
            for ri in range(len(rows)):
                h = row_hs[ri]
                if y + h > mt + usable_h:
                    new_page()
                page_text(" | ".join(c["lines"][0] for c in cell_lines[ri] if c["lines"]))
                for c in cell_lines[ri]:
                    w = sum(col_ws[c["x"]:c["x"] + c["span"]])
                    x = ml + sum(col_ws[:c["x"]])
                    draw.rectangle([x, y, x + w, y + h], outline="#000000", width=1)
                    fnt, _rb = load_font("宋体", c["size_pt"] * s, c["bold"])
                    sw = 1 if (c["bold"] and not c.get("real_bold")) else 0
                    ty = y + 3
                    for tl in c["lines"]:
                        draw.text((x + 4, ty), tl, font=fnt, fill="#000000",
                                  stroke_width=sw, stroke_fill="#000000")
                        ty += c["size_pt"] * s * 1.3
                y += h
            y += 6
    out = []
    page_texts = []
    for i, (pg, texts) in enumerate(pages):
        p = os.path.join(assets_dir, "word_p%d.jpg" % i)
        pg.save(p, quality=85)
        out.append(p)
        page_texts.append("".join(texts))
    return len(pages), page_texts


if __name__ == "__main__":
    import argparse
    ap = argparse.ArgumentParser()
    ap.add_argument("docx")
    ap.add_argument("out")
    ap.add_argument("--dpi", type=int, default=140)
    a = ap.parse_args()
    os.makedirs(a.out, exist_ok=True)
    n = render_docx_pages(a.docx, a.out, a.dpi)
    print("rendered %d pages" % n)
