# -*- coding: utf-8 -*-
"""套模板的通用工具：按「形状名」定位、按文本框原样式改写。

为什么有这一层：原作者项目里的 `tpluse.fill_cover/fill_toc/fill_result` 是
**那个项目的实现**——文字写死在函数体里（"中相微乳液"、"焦妍"、长江大学）。
换项目时不要照抄那些函数，用这里的 by_name + set_text，按模板里实际的形状名改。

用法见 deck/build2.py。
"""
import copy
import os

from pptx.oxml.ns import qn
from pptx.util import Pt
from pptx.dml.color import RGBColor

from .lib import TXT


# ---------------------------------------------------------------- 定位形状
def iter_shapes(shapes):
    """递归展开组合（GROUP）里的子形状。"""
    for s in shapes:
        yield s
        if s.shape_type == 6:  # MSO_SHAPE_TYPE.GROUP
            yield from iter_shapes(s.shapes)


def named(slide, name):
    return [s for s in iter_shapes(slide.shapes) if s.name == name]


def by_name(slide, name, required=False):
    got = named(slide, name)
    if got:
        return got[0]
    if required:
        raise KeyError(
            f'模板里没有名为 {name!r} 的形状。'
            f'先打印实际名字：python -c "import sys;from pptx import Presentation;'
            f'p=Presentation(\'tpl2/t.pptx\');[print(i,repr(s.name)) for i,s in '
            f'enumerate(p.slides[0].shapes)]"')
    print(f'  ! 跳过：模板里没有形状 {name!r}')
    return None


# ---------------------------------------------------------------- 改文字
def set_text(shape, text, size=None, color=None, bold=None, align=None):
    """改写文本框内容，默认沿用模板里第一个 run 的字体/字号/颜色。

    text 里的 \\n 会拆成多个段落（同一文本框内换行）。
    """
    if shape is None:
        return None
    tf = shape.text_frame
    tf.word_wrap = True

    proto = None
    for p in tf.paragraphs:
        if p.runs:
            proto = copy.deepcopy(p.runs[0]._r)
            break

    for p in list(tf.paragraphs)[1:]:
        p._p.getparent().remove(p._p)
    p0 = tf.paragraphs[0]
    for r in list(p0.runs):
        r._r.getparent().remove(r._r)
    if align is not None:
        p0.alignment = align

    for i, line in enumerate(text.split('\n')):
        p = p0 if i == 0 else tf.add_paragraph()
        run = p.add_run()
        run.text = line
        if proto is not None:
            r = run._r
            old = r.find(qn('a:rPr'))
            if old is not None:
                r.remove(old)
            pr = proto.find(qn('a:rPr'))
            if pr is not None:
                r.insert(0, copy.deepcopy(pr))
        if size is not None:
            run.font.size = Pt(size)
        if color is not None:
            run.font.color.rgb = RGBColor.from_string(color)
        if bold is not None:
            run.font.bold = bold
        else:
            run.font.bold = run.font.bold
    return shape


def replace_all(slide, mapping):
    """mapping: {形状名: 新文字}。缺的形状打印提示，不中断。"""
    for name, text in mapping.items():
        set_text(by_name(slide, name), text)


# ---------------------------------------------------------------- 目录页
def fill_toc(slide, items, num_style='一二三四五六七八九十'):
    """items: [(序号或None, 标题), ...]，按模板里 toc-n / toc-title 的先后顺序改。

    模板条目少于 items 时只改能改的；多出来的模板条目原样保留并提示。
    """
    nums = named(slide, 'toc-n')
    titles = named(slide, 'toc-title')
    if len(titles) < len(items):
        print(f'  ! 目录页只有 {len(titles)} 条，items 给了 {len(items)} 条：多余的没写进去')
    for i, item in enumerate(items):
        num, title = item if isinstance(item, (tuple, list)) else (None, item)
        if num is None:
            num = num_style[i] if i < len(num_style) else str(i + 1)
        if i < len(nums):
            set_text(nums[i], num)
        if i < len(titles):
            set_text(titles[i], title)
    if len(titles) > len(items):
        print(f'  ! 模板目录还有 {len(titles) - len(items)} 条没被覆盖，记得在 fill_toc 里补上或删页')


# ---------------------------------------------------------------- 页序
def keep_slides(prs, indices):
    """indices: 保留页的 0 基下标（其余删掉，连带关系一起清）。"""
    lst = prs.slides._sldIdLst
    for i, ent in enumerate(list(lst)):
        if i in indices:
            continue
        prs.part.drop_rel(ent.get(qn('r:id')))
        lst.remove(ent)


def reorder(prs, order):
    """order: 以 0 为起点的当前顺序 → 期望顺序。"""
    lst = prs.slides._sldIdLst
    entries = list(lst)
    assert len(order) == len(entries), f'order 给了 {len(order)} 项，幻灯片有 {len(entries)} 页'
    for e in entries:
        lst.remove(e)
    for i in order:
        lst.append(entries[i])


def renumber_partnames(prs):
    """删页 + 加页之后，包内 partname 可能撞名；先挪到 tmp 再顺序编号。"""
    from pptx.opc.packuri import PackURI
    parts = [s.part for s in prs.slides]
    for i, p in enumerate(parts):
        p.partname = PackURI(f'/ppt/slides/tmp{i}.xml')
    for i, p in enumerate(parts, 1):
        p.partname = PackURI(f'/ppt/slides/slide{i}.xml')


def report(prs, path):
    print(f'saved {path}  slides={len(prs.slides)}  size={os.path.getsize(path)/1024:.0f} KB')
