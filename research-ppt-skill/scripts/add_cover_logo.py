# -*- coding: utf-8 -*-
"""在封面（paper_cover.py 重建后的第 1 页）右上角补机构 logo。

paper_cover.py 会整页重画 slide 1，页面 meta 里的 header_logo 不会跟过去；
所以在 build.sh 里要接在 paper_cover 之后、apply_notes 之前跑一次。

用法：
    python scripts/add_cover_logo.py out/slides.pptx images/logo.png [高度in] [top_in]

默认高 0.42 in、top 0.13 in：0.5/0.42 这个尺寸配上 y≈0.42 会压住居中标题的末字（实测），
往上提到 0.13 才不压字；右缘固定 13.02 in，与内页页眉对齐。
"""
import sys

from pptx import Presentation
from pptx.util import Inches

NAME = '机构标识'  # spec_audit.py 按这个名字跳过 logo 的 2.6 in 下限检查


def main():
    deck = sys.argv[1] if len(sys.argv) > 1 else 'out/slides.pptx'
    logo = sys.argv[2] if len(sys.argv) > 2 else 'images/logo.png'
    height = float(sys.argv[3]) if len(sys.argv) > 3 else 0.42
    top = float(sys.argv[4]) if len(sys.argv) > 4 else 0.13

    prs = Presentation(deck)
    slide = prs.slides[0]
    if any(sh.name == NAME for sh in slide.shapes):
        print('cover logo already present, skipped')
        return
    pic = slide.shapes.add_picture(logo, Inches(10.0), Inches(top), height=Inches(height))
    pic.left = Inches(13.02) - pic.width
    pic.name = NAME
    prs.save(deck)
    print(f'cover logo added: {height} in tall, top {top} in, right edge 13.02 in')


if __name__ == '__main__':
    main()
