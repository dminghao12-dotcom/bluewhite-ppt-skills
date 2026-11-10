# -*- coding: utf-8 -*-
"""仿「国家重点研发计划（蓝色）」参考稿的组件库：页眉、判断句条、卡片、色带、结论条、表格。

尺寸单位一律英寸，画布 13.333 x 7.5。
"""
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.oxml.ns import qn
from lxml import etree

FONT = '微软雅黑'

BLUE = '003195'
BLUE_D = '00205F'
BLUE_L = '2A62C4'
BLUE_PALE = 'DCE6F7'
RED = 'C00000'
RED_L = 'DE3B3B'
BORDER = 'C1CCE4'
BG_SOFT = 'F2F5FB'
TXT = '1A1A1A'
GRAY = '5A6A85'
GRAY_L = '8A98B0'
WHITE = 'FFFFFF'

# 字号阶梯（pt）：图注 12 / 标签 14 / 正文 18 / 分点小标题 19 / 首结论 20 / 尾结论 21 / 页眉 26 / 封面计划名 72。
# 正文 18 是用户定稿值（再大嫌大）；放不下先压条目或拆页，实在放不下才降到 16，并接受该页小一号。
FS = {'cap': 12, 'label': 14, 'body': 18, 'sub': 19, 'judg': 20, 'concl': 21, 'head': 26, 'cover': 72}
BODY_LINE = 1.44   # 中文正文行距
CARD_LINE = 1.40   # 卡片内行距
HANG = 0.30        # 分点悬挂缩进（in），卡片内可传 0.20


def rgb(h):
    return RGBColor.from_string(h)


# ---------------------------------------------------------------- 基础文本框
def set_font(run, size=14, bold=False, color=TXT, font=FONT, italic=False):
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.italic = italic
    run.font.name = font
    run.font.color.rgb = rgb(color)
    rPr = run._r.get_or_add_rPr()
    for tag in ('a:ea', 'a:cs'):
        el = rPr.find(qn(tag))
        if el is None:
            el = etree.SubElement(rPr, qn(tag))
        el.set('typeface', font)
    rPr.set('lang', 'zh-CN')


def textbox(slide, x, y, w, h, align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP,
            margins=(0.02, 0.02, 0.02, 0.02), wrap=True):
    box = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = box.text_frame
    tf.word_wrap = wrap
    tf.auto_size = None
    tf.vertical_anchor = anchor
    l, t, r, b = margins
    tf.margin_left, tf.margin_top = Inches(l), Inches(t)
    tf.margin_right, tf.margin_bottom = Inches(r), Inches(b)
    tf.paragraphs[0].alignment = align
    return box


def para(tf, runs, size=FS['body'], color=TXT, bold=False, align=PP_ALIGN.LEFT,
         space_before=0, space_after=3, line=1.12, first=False, level=0):
    """runs: str 或 [(text, color, bold), ...]"""
    p = tf.paragraphs[0] if first else tf.add_paragraph()
    p.alignment = align
    p.level = level
    p.line_spacing = line
    p.space_before = Pt(space_before)
    p.space_after = Pt(space_after)
    if isinstance(runs, str):
        runs = [(runs, color, bold)]
    for text, c, b in runs:
        r = p.add_run()
        r.text = text
        set_font(r, size=size, bold=b, color=c)
    return p


def shape(slide, kind, x, y, w, h, fill=None, line=None, line_w=1.0, radius=None,
          shadow=False):
    s = slide.shapes.add_shape(kind, Inches(x), Inches(y), Inches(w), Inches(h))
    if radius is not None:
        try:
            s.adjustments[0] = radius
        except Exception:
            pass
    if fill is None:
        s.fill.background()
    else:
        s.fill.solid()
        s.fill.fore_color.rgb = rgb(fill)
    if line is None:
        s.line.fill.background()
    else:
        s.line.color.rgb = rgb(line)
        s.line.width = Pt(line_w)
    # 去掉 p:style（effectRef 会让 LibreOffice 及部分 PowerPoint 版本加阴影），
    # 只留一个空 effectLst，确保形状不带任何主题效果
    el = s._element
    style = el.find(qn('p:style'))
    if style is not None:
        el.remove(style)
    s.shadow.inherit = False
    s.text_frame.word_wrap = True
    return s


def gradient(s, c1, c2, angle=0.0, c3=None):
    s.fill.gradient()
    stops = s.fill.gradient_stops
    stops[0].color.rgb = rgb(c1)
    stops[0].position = 0.0
    if c3:
        if len(stops) < 3:
            s.fill.gradient()  # 保证两档
            stops = s.fill.gradient_stops
        stops[1].color.rgb = rgb(c3)
        stops[1].position = 0.55
        stops[2].color.rgb = rgb(c2)
        stops[2].position = 1.0
    else:
        stops[1].color.rgb = rgb(c2)
        stops[1].position = 1.0
    s.fill.gradient_angle = angle


def sline(slide, x1, y1, x2, y2, color=BLUE, width=1.0):
    from pptx.enum.shapes import MSO_CONNECTOR
    ln = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Inches(x1), Inches(y1),
                                    Inches(x2), Inches(y2))
    ln.line.color.rgb = rgb(color)
    ln.line.width = Pt(width)
    return ln


# ---------------------------------------------------------------- 页面骨架
def header(slide, num, title, right=None, logo='brand/logo_blue.png'):
    """页眉：左侧编号 + 页标题 + 通栏细线 + 右上校徽（每页都有）。"""
    box = textbox(slide, 0.41, 0.15, 0.92, 0.45, anchor=MSO_ANCHOR.MIDDLE)
    para(box.text_frame, num, size=26, color=BLUE, bold=True, first=True,
         space_after=0, line=1.0)
    box = textbox(slide, 1.34, 0.15, 9.2, 0.45, anchor=MSO_ANCHOR.MIDDLE)
    para(box.text_frame, title, size=26, color=BLUE, bold=True, first=True,
         space_after=0, line=1.0)
    import os
    from pptx.util import Inches as _In
    root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    lp = os.path.join(root, logo)
    if os.path.exists(lp):
        slide.shapes.add_picture(lp, _In(11.02), _In(0.02), height=_In(0.72))
    if right:
        pass  # 章节名不再放页眉（右上让给校徽），由目录与页标题承担
    sline(slide, 0, 0.80, 13.333, 0.80, BLUE, 1.25)


TAIL_RED = 'FF6060'


def segments(text):
    """把【…】包住的核心词切成 (文字, 是否强调)。首尾结论与正文都用这套标记。"""
    out, rest = [], text
    while rest:
        i = rest.find('【')
        if i < 0:
            out.append((rest, False))
            break
        if i:
            out.append((rest[:i], False))
        j = rest.find('】', i)
        if j < 0:
            out.append((rest[i + 1:], True))
            break
        out.append((rest[i + 1:j], True))
        rest = rest[j + 1:]
    return out


def rich(text, color=TXT, bold=False, red=RED):
    """正文富文本：把【…】渲染成红色加粗，其余保持给定颜色与粗细。"""
    return [(seg, red if hit else color, True if hit else bold) for seg, hit in segments(text)]


def judgment(slide, text, size=20, top=0.93, h=0.58):
    """首结论：白底通栏 + 浅蓝细边 + 左侧蓝色三角引导；20pt 黑体加粗左对齐。

    text 里用【…】标核心词，只有这些词标红（每页 2–3 处），不要整句发红。
    框名以 la 开头，tidy 的“短句居中”才会跳过它——否则 28 字以内的判断句会被居中。
    """
    shape(slide, MSO_SHAPE.ROUNDED_RECTANGLE, 0.42, top, 12.5, h,
          fill=WHITE, line=BORDER, line_w=0.75, radius=0.06)
    tri = shape(slide, MSO_SHAPE.ISOSCELES_TRIANGLE, 0.60, top + h / 2 - 0.085, 0.17, 0.17,
                fill=BLUE, line=None)
    tri.rotation = 90
    box = textbox(slide, 0.90, top, 11.90, h, align=PP_ALIGN.LEFT,
                  anchor=MSO_ANCHOR.MIDDLE, margins=(0.02, 0, 0.04, 0))
    para(box.text_frame, rich(text, TXT, True), size=size, first=True,
         space_after=0, line=1.0)
    box.name = 'la-jd'
    return box


def conclude(slide, text, top=6.58, h=0.64, size=21):
    """尾结论：整条深蓝渐变 + 白字加粗居中；强调短语用亮红。

    深蓝底上用正文的 #C00000 会发黑，强调色改为 TAIL_RED；叫法保留 conclusion 兼容旧稿。
    """
    bar = shape(slide, MSO_SHAPE.ROUNDED_RECTANGLE, 0.42, top, 12.5, h, fill=BLUE,
                line=None, radius=0.14)
    gradient(bar, '002A80', '1F55B8', angle=0)
    box = textbox(slide, 0.60, top, 12.14, h, align=PP_ALIGN.CENTER,
                  anchor=MSO_ANCHOR.MIDDLE, margins=(0.04, 0, 0.04, 0))
    para(box.text_frame, rich(text, WHITE, True, red=TAIL_RED), size=size, first=True,
         align=PP_ALIGN.CENTER, space_after=0, line=1.0)
    box.name = 'tail'
    return bar


def card(slide, x, y, w, h, ribbon=None, ribbon_color='blue', ribbon_size=18,
         fill=WHITE, line=BORDER):
    body = shape(slide, MSO_SHAPE.ROUNDED_RECTANGLE, x, y, w, h, fill=fill,
                 line=line, line_w=1.0, radius=0.04)
    if ribbon:
        rh = 0.44
        cap = shape(slide, MSO_SHAPE.ROUND_2_SAME_RECTANGLE, x, y, w, rh,
                    fill=BLUE, line=None, radius=0.12)
        if ribbon_color == 'blue':
            gradient(cap, BLUE, BLUE_L, angle=0)
        else:
            gradient(cap, RED, RED_L, angle=0)
        box = textbox(slide, x + 0.1, y, w - 0.2, rh, align=PP_ALIGN.CENTER,
                      anchor=MSO_ANCHOR.MIDDLE, margins=(0.04, 0, 0.04, 0))
        para(box.text_frame, ribbon, size=ribbon_size, color=WHITE, bold=True,
             first=True, align=PP_ALIGN.CENTER, space_after=0, line=1.0)
    return body


def bullets(tf, items, size=FS['body'], color=TXT, marker='➤', mcolor=BLUE,
            space_after=8, line=CARD_LINE, first=True, indent=0.0, hang=HANG):
    """以 marker（默认蓝色 ➤）起行的分点，带悬挂缩进，换行文字与正文对齐。

    items 每项可以是：
      - str：普通正文；
      - {'label': '体系对照', 'text': '……'}：蓝粗标签（比正文大一号）+ 黑色正文；
      - (text, color, bold)：单段富文本；
      - [(text, color, bold), ...]：多段富文本（如 rich() 的返回值）。
    marker='' 时不加箭头、也不加悬挂缩进。
    """
    from pptx.util import Inches
    for i, item in enumerate(items):
        runs, label_idx = [], None
        if marker:
            runs.append((marker + ' ', mcolor, True))
        if isinstance(item, str):
            runs.append((item, color, False))
        elif isinstance(item, dict):
            label_idx = len(runs)
            runs.append((item['label'].rstrip('：:') + '：', BLUE, True))
            runs.append((item['text'], color, False))
        elif isinstance(item, tuple) and len(item) == 3 and isinstance(item[0], str):
            runs.append(item)                       # 单段富文本 (text, color, bold)
        else:
            for seg in item:                        # 多段富文本元组
                runs.append(seg)
        p = para(tf, runs, size=size, first=(first and i == 0), space_after=space_after,
                 line=line)
        if label_idx is not None:
            p.runs[label_idx].font.size = Pt(size + 1)
        if marker and hang:
            pPr = p._p.get_or_add_pPr()
            pPr.set('marL', str(int(Inches(hang + indent))))
            pPr.set('indent', str(-int(Inches(hang))))


def conclusion(slide, *args, top=6.58, h=0.64, size=21):
    """尾结论（兼容旧调用：conclusion(slide, 标签, 正文) 的标签已废弃，只取最后一个参数）。"""
    return conclude(slide, args[-1], top=top, h=h, size=size)


def caption(slide, x, y, w, text, size=10.5, color=GRAY, align=PP_ALIGN.CENTER, h=0.3):
    """图名：默认居中放在图片正下方。"""
    box = textbox(slide, x, y, w, h, align=align, anchor=MSO_ANCHOR.TOP,
                  margins=(0.02, 0, 0.02, 0))
    para(box.text_frame, text, size=size, color=color, first=True, align=align,
         space_after=0, line=1.05)


def picture(slide, path, x, y, w, h, mode='fit', border=None):
    """mode='fit' 等比缩放放入框内居中；'cover' 裁切填满。"""
    from PIL import Image
    iw, ih = Image.open(path).size
    ar = iw / ih
    if mode == 'fit':
        if w / h > ar:
            hh, ww = h, h * ar
        else:
            ww, hh = w, w / ar
        px, py = x + (w - ww) / 2, y + (h - hh) / 2
        pic = slide.shapes.add_picture(path, Inches(px), Inches(py), Inches(ww), Inches(hh))
    else:
        # cover：先按比例放大再裁切
        if w / h > ar:
            ww, hh = w, w / ar
        else:
            ww, hh = h * ar, h
        px, py = x - (ww - w) / 2, y - (hh - h) / 2
        pic = slide.shapes.add_picture(path, Inches(px), Inches(py), Inches(ww), Inches(hh))
        # 用 picture.crop_* 精确裁到框内
        cl = (px - x) / ww if ww > w else 0
        cr = cl
        ct = (py - y) / hh if hh > h else 0
        cb = ct
        try:
            pic.crop_left = max(0, cl)
            pic.crop_right = max(0, cr)
            pic.crop_top = max(0, ct)
            pic.crop_bottom = max(0, cb)
            pic.left, pic.top = Inches(x), Inches(y)
            pic.width, pic.height = Inches(w), Inches(h)
        except Exception:
            pass
    if border:
        pic.line.color.rgb = rgb(border)
        pic.line.width = Pt(0.75)
    return pic


def table(slide, x, y, w, rows, col_w, row_h=0.34, head_h=0.4, size=12.5,
          head_size=13, head_fill=BLUE, first_col_fill='E9EFF9', red_cells=(),
          align=None, head_align=PP_ALIGN.CENTER, border=BORDER):
    """rows[0] 为表头。col_w 为各列宽度（英寸），合计应≈w。"""
    nrow, ncol = len(rows), len(rows[0])
    heights = [head_h] + [row_h] * (nrow - 1)
    gt = slide.shapes.add_table(nrow, ncol, Inches(x), Inches(y), Inches(w),
                                Inches(sum(heights)))
    tbl = gt.table
    # 关闭默认样式条纹
    tblPr = tbl._tbl.tblPr
    tblPr.set('firstRow', '1')
    tblPr.set('bandRow', '0')
    for i, cw in enumerate(col_w):
        tbl.columns[i].width = Inches(cw)
    for r, hh in enumerate(heights):
        tbl.rows[r].height = Inches(hh)
    for r in range(nrow):
        for c in range(ncol):
            cell = tbl.cell(r, c)
            cell.text = ''
            cell.margin_left = Inches(0.06)
            cell.margin_right = Inches(0.06)
            cell.margin_top = Inches(0.02)
            cell.margin_bottom = Inches(0.02)
            cell.vertical_anchor = MSO_ANCHOR.MIDDLE
            cell.fill.solid()
            if r == 0:
                cell.fill.fore_color.rgb = rgb(head_fill)
            elif c == 0 and first_col_fill:
                cell.fill.fore_color.rgb = rgb(first_col_fill)
            else:
                cell.fill.fore_color.rgb = rgb(WHITE)
            tf = cell.text_frame
            tf.word_wrap = True
            val = rows[r][c]
            al = head_align if r == 0 else (align[c] if align else PP_ALIGN.LEFT)
            first = True
            for lineText in str(val).split('\n'):
                p = tf.paragraphs[0] if first else tf.add_paragraph()
                first = False
                p.alignment = al
                p.line_spacing = 1.03
                p.space_after = Pt(1)
                run = p.add_run()
                run.text = lineText
                set_font(run, size=(head_size if r == 0 else size),
                         bold=(r == 0 or c == 0),
                         color=(WHITE if r == 0 else (RED if (r, c) in red_cells else TXT)))
    # 单元格边框
    for r in range(nrow):
        for c in range(ncol):
            tcPr = tbl.cell(r, c)._tc.get_or_add_tcPr()
            for tag in ('a:lnL', 'a:lnR', 'a:lnT', 'a:lnB'):
                ln = tcPr.find(qn(tag))
                if ln is not None:
                    tcPr.remove(ln)
                ln = etree.SubElement(tcPr, qn(tag))
                ln.set('w', '9525')
                ln.set('cap', 'flat')
                ln.set('cmpd', 'sng')
                ln.set('algn', 'ctr')
                sf = etree.SubElement(ln, qn('a:solidFill'))
                clr = etree.SubElement(sf, qn('a:srgbClr'))
                clr.set('val', border)
                etree.SubElement(ln, qn('a:prstDash')).set('val', 'solid')
    return gt


def chip(slide, x, y, w, h, text, fill=BLUE, color=WHITE, size=13, bold=True,
         radius=0.2, line=None):
    s = shape(slide, MSO_SHAPE.ROUNDED_RECTANGLE, x, y, w, h, fill=fill, line=line,
              radius=radius)
    box = textbox(slide, x, y, w, h, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE,
                  margins=(0.03, 0, 0.03, 0))
    para(box.text_frame, text, size=size, color=color, bold=bold, first=True,
         align=PP_ALIGN.CENTER, space_after=0, line=1.0)
    return s


def arrow_right(slide, x, y, w, h, fill=BLUE):
    s = shape(slide, MSO_SHAPE.RIGHT_ARROW, x, y, w, h, fill=fill, line=None)
    return s


def badge(slide, x, y, d, text, fill=BLUE, size=15):
    s = shape(slide, MSO_SHAPE.OVAL, x, y, d, d, fill=fill, line=None)
    box = textbox(slide, x, y, d, d, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE,
                  margins=(0, 0, 0, 0))
    para(box.text_frame, text, size=size, color=WHITE, bold=True, first=True,
         align=PP_ALIGN.CENTER, space_after=0, line=1.0)
    return s
