# -*- coding: utf-8 -*-
"""封面与结束页的蓝色波浪底（矢量化：矩形渐变 + 自由曲线）。"""
import math
from pptx.enum.shapes import MSO_SHAPE
from pptx.util import Inches
from .lib import shape, gradient, textbox, para, BLUE, BLUE_L, WHITE, sline
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR


def _curve(slide, y0, amp, phase, w=13.333, fill='FFFFFF', steps=120, h_extra=4.0):
    """在 y0 上下做一条正弦底边，向下填色到画布底部。"""
    pts = []
    for i in range(steps + 1):
        x = w * i / steps
        y = y0 + amp * math.sin(math.pi * x / w + phase)
        pts.append((x, y))
    ff = slide.shapes.build_freeform(Inches(pts[0][0]), Inches(pts[0][1]))
    ff.add_line_segments([(Inches(x), Inches(y)) for x, y in pts[1:]], close=False)
    ff.add_line_segments([(Inches(w), Inches(y0 + h_extra)), (Inches(0), Inches(y0 + h_extra))],
                         close=True)
    s = ff.convert_to_shape()
    s.fill.solid()
    s.fill.fore_color.rgb = __import__('pptx').dml.color.RGBColor.from_string(fill)
    s.line.fill.background()
    s.shadow.inherit = False
    return s


def cover_background(slide, navy='001A4D', deep=BLUE, light='D8E3F5'):
    top = shape(slide, MSO_SHAPE.RECTANGLE, 0, 0, 13.333, 5.6, fill=deep, line=None)
    gradient(top, navy, deep, angle=90)
    _curve(slide, 4.35, 0.55, 0.35, fill=light)
    _curve(slide, 4.72, 0.42, 0.9, fill=WHITE)


def end_background(slide, navy='001A4D', deep=BLUE, light='D8E3F5'):
    top = shape(slide, MSO_SHAPE.RECTANGLE, 0, 0, 13.333, 6.1, fill=deep, line=None)
    gradient(top, navy, deep, angle=90)
    _curve(slide, 4.55, 0.5, 0.3, fill=light)
    _curve(slide, 4.95, 0.38, 0.95, fill=WHITE)


def top_bar(slide, left_text, right_text=None):
    sline(slide, 0, 0.88, 13.333, 0.88, 'FFFFFF', 0.9)
    box = textbox(slide, 0.55, 0.42, 8.6, 0.42, anchor=MSO_ANCHOR.MIDDLE)
    para(box.text_frame, left_text, size=17, color=WHITE, bold=True, first=True,
         space_after=0, line=1.0)
    if right_text:
        box = textbox(slide, 9.4, 0.42, 3.4, 0.42, align=PP_ALIGN.RIGHT,
                      anchor=MSO_ANCHOR.MIDDLE)
        para(box.text_frame, right_text, size=12, color='CBD8EE', first=True,
             align=PP_ALIGN.RIGHT, space_after=0, line=1.0)
