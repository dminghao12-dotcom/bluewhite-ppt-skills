# -*- coding: utf-8 -*-
"""在用户提供的蓝色模板上生成答辩稿：复用模板页的版式/文字框/图片位，其余页沿用我们的自绘版式。

产物：out/申报答辩-中相微乳液-套模板.pptx
"""
import copy
import os
import sys

from pptx import Presentation
from pptx.oxml.ns import qn

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
sys.path.insert(0, ROOT)
os.chdir(ROOT)

from deck import pages1, pages2, tpluse  # noqa: E402

TPL = 'tpl2/t.pptx'

# 模板中要保留的原页（1-based）：封面、目录、背景、指标、资源、成果、结束
KEEP = [1, 3, 4, 11, 17, 18]


def drop_slides(prs, keep_idx):
    lst = prs.slides._sldIdLst
    entries = list(lst)
    for i, ent in enumerate(entries):
        if i in keep_idx:
            continue
        rId = ent.get(qn('r:id'))
        prs.part.drop_rel(rId)
        lst.remove(ent)


def reorder(prs, order):
    """order: 以 0 为起点的旧顺序 -> 新顺序列表"""
    lst = prs.slides._sldIdLst
    entries = list(lst)
    new = [entries[i] for i in order]
    for e in entries:
        lst.remove(e)
    for e in new:
        lst.append(e)


def build(out_path='out/申报答辩-中相微乳液-套模板.pptx'):
    prs = Presentation(TPL)
    keep0 = [i - 1 for i in KEEP]
    drop_slides(prs, keep0)

    cover, toc, bg, metric, result, end = list(prs.slides)
    tpluse.fill_cover(cover)
    tpluse.fill_toc(toc)
    tpluse.fill_bg(bg)
    tpluse.fill_metric(metric)
    tpluse.fill_result(result)
    tpluse.fill_end(end)

    # 其余 13 页沿用自绘版式
    for fn in [pages1.s03, pages1.s05, pages1.s06, pages1.s07, pages1.s08,
               pages1.s09, pages1.s10, pages2.s11, pages2.s12, pages2.s13,
               pages2.s15, pages2.s16, pages2.s17r, pages2.s19]:
        fn(prs)

    # 旧顺序：0 封面 1 目录 2 背景 3 指标 4 成果 5 结束 6.. 自绘
    order = [0, 1, 6, 2, 7, 8, 9, 10, 11, 12, 13, 14, 15, 3, 16, 17, 18, 4, 19, 5]
    assert len(order) == len(prs.slides.__iter__.__self__._sldIdLst), len(order)
    reorder(prs, order)

    from deck.tidy import apply as tidy_apply
    tidy_apply(prs)
    from deck.emphasis import apply as emphasize
    emphasize(prs)

    from deck.notes2 import SCRIPT
    assert len(SCRIPT) == 20, len(SCRIPT)
    import copy as _copy
    slides = list(prs.slides)
    # 找一个自带备注正文占位符的页做模板（自绘页都有）
    donor = None
    for s in slides:
        if s.notes_slide.notes_text_frame is not None:
            donor = s.notes_slide.notes_placeholder._element
            break
    total = 0
    for slide, (head, body) in zip(slides, SCRIPT):
        ns = slide.notes_slide
        if ns.notes_text_frame is None and donor is not None:
            ns.shapes._spTree.append(_copy.deepcopy(donor))
        ns.notes_text_frame.text = head + '\n' + body
        total += len(body)

    # 修掉 add_slide 产生的 partname 冲突（删页后包内仍留有旧名）
    from pptx.opc.packuri import PackURI
    parts = [s.part for s in prs.slides]
    for i, p in enumerate(parts):
        p.partname = PackURI(f'/ppt/slides/tmp{i}.xml')
    for i, p in enumerate(parts, 1):
        p.partname = PackURI(f'/ppt/slides/slide{i}.xml')

    os.makedirs(os.path.dirname(out_path), exist_ok=True)
    prs.save(out_path)
    print(f'saved {out_path}  slides={len(prs.slides)} 讲稿 {total} 字')
    return out_path


if __name__ == '__main__':
    build(*(sys.argv[1:] or []))
