# -*- coding: utf-8 -*-
"""P17 年代流向图：横向时间轴 + 年份节点 + 上下交替的论文卡片。

在研究背景页用（“领域十年发展脉络”）：8 张卡片里放的是**原稿/论文的真实插图**，
不是重画的示意图——这一页的价值就在“哪一年、哪篇、什么期刊”的可核对性。

在 dense_deck.py 之后跑（它会把目标页清空重画），并写进 build.sh：

    $PY $S/scripts/dense_deck.py --content content.ready.json --output out/slides.pptx
    $PY $S/scripts/timeline_cards.py out/slides.pptx cards.json --page 1

cards.json（页眉、引导句仍走 info_style，所以带 section 就自动是章节带页眉）：

    {"section": "研究背景", "title": "磁性软材料的发展",
     "lead": {"text": "从形状可编程、多尺度结构设计，走到本征磁性高分子",
              "highlights": ["本征磁性高分子"]},
     "cards": [{"year": "2016", "title": "形状可编程磁性软物质",
                "journal": "Adv. Mater. 2016", "text": "", "image": "images/y2016.png"},
               ...]}

卡片顺序 = 时间轴顺序；第 1、3、5、7 张放上半页，第 2、4、6、8 张放下半页。
`text` 留空时图占满“图 + 说明”两格（原稿没有配文的那几张就这样处理）。

尺寸现实：卡片宽 2.80 in，缩略图成稿约 1.7–1.9 in 宽，**过不了 spec_audit 的 2.6 in 图片下限**。
一行 4 张卡片、卡片 2.9 in，物理上放不下 2.6 in 的图，参考稿本身也只有 2.0 in；
按用户给的版式保留，在交付说明里写清这条为什么不由审计裁定。要过它只能减到 6 张或拆两页，先问用户。
"""
import argparse
import json
import sys
from pathlib import Path

from PIL import Image
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.util import Inches, Pt

_HERE = Path(__file__).resolve().parent
if str(_HERE) not in sys.path:
    sys.path.insert(0, str(_HERE))
import info_style as S  # noqa: E402

# ---- 几何（16:9，英寸）。卡片数 6–8 都能用，间距按数量自适应
AXIS_Y = 4.40          # 轴线中心
CARD_W = 2.80
HEAD_H = 0.50          # 卡片标题条
BODY_H = 0.95          # 图区
TEXT_H = 0.65          # 说明（12 pt 约 3 行）
JOUR_H = 0.24          # 期刊行
TOP_Y = 1.66
BOT_Y = 4.86
FS_HEAD, FS_TEXT, FS_JOUR, FS_YEAR = 13, 12, 12, 12.5


def clear(slide):
    """清空目标页（dense_deck 已经画过页眉与占位），再由本脚本重画。"""
    spTree = slide.shapes._spTree
    for el in list(spTree):
        if el.tag.split('}')[-1] in ('sp', 'pic', 'graphicFrame', 'cxnSp'):
            spTree.remove(el)


def italic_text(slide, value, x, y, w, h, size, color, bold=False, italic=True, name=None):
    box = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = box.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_right = 0
    tf.margin_top = tf.margin_bottom = 0
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    r = p.add_run()
    r.text = value
    r.font.size = Pt(size)
    r.font.bold = bold
    r.font.italic = italic
    r.font.name = '微软雅黑'
    r.font.color.rgb = RGBColor.from_string(color)
    r._r.get_or_add_rPr().set('lang', 'zh-CN')     # 语言是属性，不是子元素
    if name:
        box.name = name
    return box


def picture_contain(slide, path, x, y, w, h):
    """等比放进 (x, y, w, h) 居中，不裁切、不拉伸。"""
    iw, ih = Image.open(path).size
    k = min(w / iw, h / ih)
    pw, ph = iw * k, ih * k
    return slide.shapes.add_picture(str(path), Inches(x + (w - pw) / 2),
                                    Inches(y + (h - ph) / 2), Inches(pw), Inches(ph))


def draw_card(slide, cx, card, upward, base):
    year, title = card['year'], card['title']
    body, journal = card.get('text') or '', card.get('journal') or ''
    img = base / card['image']
    left = min(max(cx - CARD_W / 2, 0.30), 13.03 - CARD_W)      # 两端卡片贴边
    y = TOP_Y if upward else BOT_Y
    total_h = HEAD_H + BODY_H + TEXT_H + JOUR_H + 0.10
    S.rect(slide, left, y, CARD_W, total_h, 'FFFFFF', 'C9D6E6', 0.75, name=f'卡片-{year}')
    S.rect(slide, left, y, CARD_W, HEAD_H, S.LIGHT, None, name=f'卡片标题条-{year}')
    S.tx(slide, title, left + 0.08, y, CARD_W - 0.16, HEAD_H, FS_HEAD, S.NAVY, True, 'center',
         name=f'卡片标题-{year}')
    iy = y + HEAD_H + 0.05
    if body:
        picture_contain(slide, img, left + 0.10, iy, CARD_W - 0.20, BODY_H)
        S.tx(slide, body, left + 0.10, iy + BODY_H + 0.02, CARD_W - 0.20, TEXT_H - 0.04,
             FS_TEXT, S.INK, False, 'center', name=f'卡片说明-{year}')
    else:
        picture_contain(slide, img, left + 0.10, iy, CARD_W - 0.20, BODY_H + TEXT_H - 0.04)
    italic_text(slide, journal, left + 0.10, y + HEAD_H + BODY_H + TEXT_H + 0.05,
                CARD_W - 0.20, JOUR_H, FS_JOUR, S.NAVY, name=f'期刊-{year}')
    # 卡片与轴线之间的细线
    if upward:
        y0 = y + total_h
        S.rect(slide, cx - 0.007, y0, 0.014, max(AXIS_Y - 0.24 - y0, 0.02), S.LINE, None,
               name=f'连线-{year}')
    else:
        S.rect(slide, cx - 0.007, AXIS_Y + 0.24, 0.014, max(y - (AXIS_Y + 0.24), 0.02), S.LINE,
               None, name=f'连线-{year}')


def draw(pptx_path, spec, page=1, base=None):
    base = Path(base) if base else Path(pptx_path).resolve().parent
    prs = Presentation(pptx_path)
    slide = list(prs.slides)[page - 1]
    clear(slide)
    meta = {'section': spec['section']} if spec.get('section') else {}
    S.headline(slide, spec['title'], meta)
    if spec.get('lead'):
        S.lead(slide, spec['lead'], 0.94)
    # 轴线与趋势箭头
    S.rect(slide, 0.30, AXIS_Y - 0.055, 12.25, 0.11, S.ARROW, None, name='时间轴')
    arrow = slide.shapes.add_shape(MSO_SHAPE.UP_ARROW, Inches(12.45), Inches(2.55),
                                   Inches(0.55), Inches(3.5))
    arrow.fill.solid()
    arrow.fill.fore_color.rgb = RGBColor.from_string(S.ARROW)
    arrow.line.fill.background()
    arrow.name = '趋势箭头'
    # 年份节点 + 卡片
    cards = spec['cards']
    n = len(cards)
    step = 1.56 if n >= 8 else 12.0 / max(n - 1, 1)
    x0 = 1.06 if n >= 8 else 1.0
    for i, card in enumerate(cards):
        cx = x0 + i * step
        d = 0.46
        S.rect(slide, cx - d / 2, AXIS_Y - d / 2, d, d, 'FFFFFF', S.NAVY, 1.0,
               shape=MSO_SHAPE.OVAL, name=f'年份节点-{card["year"]}')
        S.tx(slide, card['year'], cx - d / 2, AXIS_Y - d / 2, d, d, FS_YEAR, S.NAVY, True,
             'center', name=f'年份-{card["year"]}')
        draw_card(slide, cx, card, upward=(i % 2 == 0), base=base)
    prs.save(pptx_path)
    print(f'年代流向图已重绘：第 {page} 页，{n} 张卡片 -> {pptx_path}')


def main():
    ap = argparse.ArgumentParser(description='P17 年代流向图（时间轴 + 论文卡片）')
    ap.add_argument('pptx')
    ap.add_argument('spec', help='cards.json：section / title / lead / cards')
    ap.add_argument('--page', type=int, default=1)
    ap.add_argument('--base', default=None, help='图片相对目录（默认取 PPTX 所在目录）')
    a = ap.parse_args()
    draw(a.pptx, json.loads(Path(a.spec).read_text(encoding='utf-8')), a.page, a.base)


if __name__ == '__main__':
    main()
