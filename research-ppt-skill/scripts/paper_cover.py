"""Literature-report cover: left navy panel, centered title block, paper's first page below.

Usage:
    python paper_cover.py paper.pdf deck.pptx \
      --title "分布式置换流水车间调度问题" \
      --subtitle "DPFSP：模型、算法与基准" \
      --meta "Naderi & Ruiz · Computers & Operations Research, 2010" \
      --tagline $'分布式车间\n多工厂调度\nDPFSP 的起点' \
      --date "2026 年 10 月"

Layout (the user-approved version, 2026-10): 3.83 in navy left panel; on the right a
centered two-line title, a rule, one metadata line and the date; below them the paper's
first page as a clean, fully visible 3:2 crop -- not a faded background, not a full-page
screenshot. All shapes on slide 1 are replaced; speaker notes stay.
"""
import argparse
import tempfile
from pathlib import Path

import pymupdf
from PIL import Image, ImageOps
from lxml import etree
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Inches, Pt

from theme import C as _THEME_C
NAVY = _THEME_C['navy']
INK = "1B2533"
WHITE = "FFFFFF"
PANEL_W = 3.83
SLIDE_W = 13.333333
SLIDE_H = 7.5


def _font(run, size, color, bold=False):
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.color.rgb = RGBColor.from_string(color)
    run.font.name = "微软雅黑"
    rpr = run._r.get_or_add_rPr()
    rpr.set("lang", "zh-CN")
    ea = rpr.find(qn("a:ea"))
    if ea is None:
        ea = etree.SubElement(rpr, qn("a:ea"))
    ea.set("typeface", "微软雅黑")


def _tx(slide, x, y, w, h, text, size, color, bold=False, align=PP_ALIGN.LEFT,
        spacing=1.0, wrap=False):
    box = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = box.text_frame
    tf.word_wrap = wrap
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    tf.vertical_anchor = 1
    p = tf.paragraphs[0]
    p.alignment = align
    p.line_spacing = spacing
    r = p.add_run()
    r.text = text
    _font(r, size, color, bold)


def _paper_crop(pdf, image, out_png, dpi=200):
    """Return a 3:2 crop from the top of the paper page (title / authors / abstract)."""
    if image:
        im = Image.open(image).convert("RGB")
    else:
        doc = pymupdf.open(pdf)
        tmp = Path(out_png).with_name("paper-page.png")
        doc[0].get_pixmap(dpi=dpi).save(tmp)
        im = Image.open(tmp).convert("RGB")
    im = ImageOps.fit(im, (1800, 1200), method=Image.Resampling.LANCZOS,
                      centering=(0.5, 0.0))
    im.save(out_png, optimize=True)


def cover(pdf, deck, title, subtitle, meta, date, tagline, image=None,
          panel_title="文献精读", dpi=200):
    with tempfile.TemporaryDirectory() as td:
        paper = Path(td) / "paper-crop.png"
        _paper_crop(pdf, image, paper, dpi)

        prs = Presentation(deck)
        slide = prs.slides[0]
        for shape in list(slide.shapes):
            shape._element.getparent().remove(shape._element)
        # Drop any background picture/fill left by an earlier cover.
        bg = slide._element.find(qn("p:bg"))
        if bg is not None:
            slide._element.remove(bg)

        # Left navy panel
        panel = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0,
                                       Inches(PANEL_W), Inches(SLIDE_H))
        panel.fill.solid()
        panel.fill.fore_color.rgb = RGBColor.from_string(NAVY)
        panel.line.fill.background()
        _tx(slide, 0.48, 0.62, 2.9, 0.62, panel_title, 30, WHITE, True)
        line = slide.shapes.add_connector(1, Inches(0.53), Inches(1.42),
                                          Inches(3.27), Inches(1.42))
        line.line.color.rgb = RGBColor.from_string(WHITE)
        line.line.width = Pt(0.8)
        _tx(slide, 0.49, 1.72, 3.05, 1.7, tagline, 21.6, WHITE, True,
            spacing=1.3, wrap=False)

        # Centered title block over the right area
        x, w = PANEL_W, SLIDE_W - PANEL_W
        _tx(slide, x, 0.55, w, 0.65, title, 32, NAVY, True, PP_ALIGN.CENTER)
        _tx(slide, x, 1.22, w, 0.6, subtitle, 28, NAVY, True, PP_ALIGN.CENTER)
        rule = slide.shapes.add_connector(1, Inches(4.20), Inches(1.98),
                                          Inches(12.95), Inches(1.98))
        rule.line.color.rgb = RGBColor.from_string(NAVY)
        rule.line.width = Pt(1.2)
        _tx(slide, x, 2.12, w, 0.35, meta, 16, INK, False, PP_ALIGN.CENTER)
        _tx(slide, x, 2.55, w, 0.35, date, 17, INK, False, PP_ALIGN.CENTER)

        # Paper's first page: clean, 3:2, centered under the metadata
        pw, ph = 6.4, 4.2667
        slide.shapes.add_picture(str(paper), Inches(PANEL_W + (SLIDE_W - PANEL_W - pw) / 2),
                                 Inches(3.02), width=Inches(pw), height=Inches(ph))

        # Keep the house cover flat: default connectors inherit theme shadows.
        for shape in slide.shapes:
            style = shape._element.find(qn("p:style"))
            if style is not None:
                shape._element.remove(style)
            sppr = shape._element.find(qn("p:spPr"))
            if sppr is not None and sppr.find(qn("a:effectLst")) is None:
                etree.SubElement(sppr, qn("a:effectLst"))
        prs.save(deck)


if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("pdf")
    ap.add_argument("deck")
    ap.add_argument("--image", default=None, help="optional first-page image; default renders PDF page 1")
    ap.add_argument("--title", default="论文精读")
    ap.add_argument("--subtitle", default="")
    ap.add_argument("--meta", default="")
    ap.add_argument("--tagline", default="研究问题\n方法路线\n结果与边界")
    ap.add_argument("--date", default="")
    ap.add_argument("--panel-title", default="文献精读")
    ap.add_argument("--dpi", type=int, default=200)
    args = ap.parse_args()
    cover(args.pdf, args.deck, args.title, args.subtitle, args.meta, args.date,
          args.tagline, args.image, args.panel_title, args.dpi)
    print(f"paper cover -> {args.deck}")
