# -*- coding: utf-8 -*-
"""在内容页底部写一行页脚文字（如「姓名｜单位｜会议名」），跳过封面与结束页。

meta.info_footer: true 只画深色底带 + 页码（y=7.24、高 0.26），不画文字；
这个脚本在底带上补白字。它必须排在 dense_deck.py（以及 paper_cover.py）之后，
写进 build.sh——每次重建都会把它冲掉。重复运行是安全的：先删旧的同名形状再写。

用法：
    python scripts/add_footer.py out/slides.pptx --text "张三｜某某大学某某学院｜2026 年开题答辩"
    python scripts/add_footer.py out/slides.pptx --text "…" --color 111722 --y 7.1   # 没开 info_footer 时用深色字
可选：--size 10.5  --x 0.4  --w 9.0  --skip-first 1  --skip-last 1  --out 另存路径
"""
import argparse

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Inches, Pt

NAME = '页脚信息'


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument('pptx')
    ap.add_argument('--text', required=True)
    ap.add_argument('--y', type=float, default=7.255)
    ap.add_argument('--h', type=float, default=0.23)
    ap.add_argument('--x', type=float, default=0.4)
    ap.add_argument('--w', type=float, default=9.0)
    ap.add_argument('--size', type=float, default=10.5)
    ap.add_argument('--color', default='FFFFFF')
    ap.add_argument('--skip-first', type=int, default=1)
    ap.add_argument('--skip-last', type=int, default=1)
    ap.add_argument('--out')
    a = ap.parse_args()

    prs = Presentation(a.pptx)
    slides = list(prs.slides)
    targets = slides[a.skip_first:len(slides) - a.skip_last if a.skip_last else None]
    for slide in targets:
        for sh in [s for s in slide.shapes if s.name == NAME]:
            sh._element.getparent().remove(sh._element)
        tb = slide.shapes.add_textbox(Inches(a.x), Inches(a.y), Inches(a.w), Inches(a.h))
        tb.name = NAME
        tf = tb.text_frame
        tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
        tf.vertical_anchor = MSO_ANCHOR.MIDDLE
        tf.word_wrap = False
        p = tf.paragraphs[0]
        p.alignment = PP_ALIGN.LEFT
        r = p.add_run()
        r.text = a.text
        r.font.size = Pt(a.size)
        r.font.color.rgb = RGBColor.from_string(a.color.lstrip('#').upper())
        r.font.name = 'Arial'
        rpr = r._r.get_or_add_rPr()
        rpr.set('lang', 'zh-CN')
        ea = rpr.find(qn('a:ea'))
        if ea is None:
            ea = rpr.makeelement(qn('a:ea'), {})
            rpr.append(ea)
        ea.set('typeface', '微软雅黑')
    out = a.out or a.pptx
    prs.save(out)
    print(f'footer written on {len(targets)} slides -> {out}')


if __name__ == '__main__':
    main()
