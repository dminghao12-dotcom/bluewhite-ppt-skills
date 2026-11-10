# -*- coding: utf-8 -*-
"""正文重点标红：用方括号标记手工指定的红词，渲染时把段内 run 切开。

用法（在项目的 deck/pages.py 里）：

    from rich_text import segments, marked_runs, write_para
    write_para(box.text_frame, '酯化交联把“氢键搭接”的纸变成【“共价交联”】的复合纸', size=16)

规则：方括号内 = 红色加粗 #C00000，其余保持给定颜色/粗细；不再跑自动挑词。
尾条（深蓝底）上的强调走 TAIL_RED，不要用 #C00000（会糊在深蓝里）。
"""
from pptx.util import Pt
from pptx.enum.text import PP_ALIGN

try:
    from deck.lib import set_font, RED, TXT
    from deck.lib import WHITE as W  # noqa: F401  (字符串色值)
except Exception:  # 独立使用时的兜底
    RED, TXT = 'C00000', '1A1A1A'

    def set_font(run, size=16, bold=False, color=TXT, font='微软雅黑', italic=False):
        run.font.size = Pt(size)
        run.font.bold = bold
        run.font.name = font
        run.font.color.rgb = __import__('pptx').dml.color.RGBColor.from_string(color)

OPEN, CLOSE = '【', '】'
TAIL_RED = 'FF6060'      # 深蓝尾条上的强调红


def segments(text):
    """'前【红】后' -> [('前', False), ('红', True), ('后', False)]；缺右括号时剩下的全当红。"""
    out, rest = [], text
    while rest:
        i = rest.find(OPEN)
        if i < 0:
            out.append((rest, False))
            break
        if i:
            out.append((rest[:i], False))
        j = rest.find(CLOSE, i)
        if j < 0:
            out.append((rest[i + 1:], True))
            break
        out.append((rest[i + 1:j], True))
        rest = rest[j + 1:]
    return out


def marked_runs(text, color=TXT, bold=False, red=RED):
    """转成 lib.para() 需要的 [(文字, 色值, 粗细)] 列表。"""
    return [(seg, red if hit else color, True if hit else bold) for seg, hit in segments(text)]


def write_para(tf, text, size=16, color=TXT, red=RED, bold=False, first=True,
               align=None, space_after=0, line=1.15):
    """在 text_frame 里写一段带红词的正文（不处理标题、图注、指标小字）。"""
    p = tf.paragraphs[0] if first else tf.add_paragraph()
    if align is not None:
        p.alignment = align
    p.line_spacing = line
    p.space_after = Pt(space_after)
    for seg, col, bd in marked_runs(text, color=color, bold=bold, red=red):
        r = p.add_run()
        r.text = seg
        set_font(r, size=size, bold=bd, color=col)
    return p


def check_leftover(pdf_path):
    """交付前自检：PDF 全文不能残留方括号（说明某处没走本模块）。"""
    import pymupdf
    d = pymupdf.open(pdf_path)
    bad = [i + 1 for i in range(len(d)) if OPEN in d[i].get_text() or CLOSE in d[i].get_text()]
    return bad
