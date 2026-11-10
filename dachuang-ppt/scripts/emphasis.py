# -*- coding: utf-8 -*-
"""把正文里的核心词挑出来标红（每段最多两处），用于突出想传达的关键判断。"""
import copy

from pptx.dml.color import RGBColor
from pptx.oxml.ns import qn

RED = RGBColor(0xC0, 0x00, 0x00)
KEYS = ['中相微乳液', 'Winsor Ⅲ', '中相窗口', '超低界面张力', '原位成相', '双向增溶',
        '洗油效率', '波及体积', '注入性', '配伍性', '重复性', '判据', '低浓度', '无碱',
        '耐温耐盐', '采收率', '压差', '药剂用量', '段塞']


def _segments(text):
    """返回 [(片段, 是否标红)]，按最靠前的关键词切分。"""
    out, rest, hits = [], text, 0
    while rest and hits < 2:
        best, bi = None, -1
        for k in KEYS:
            i = rest.find(k)
            if i >= 0 and (bi < 0 or i < bi or (i == bi and len(k) > len(best))):
                best, bi = k, i
        if best is None:
            break
        if bi:
            out.append((rest[:bi], False))
        out.append((best, True))
        rest = rest[bi + len(best):]
        hits += 1
    if rest:
        out.append((rest, False))
    return out


def _restyle(run_el, color, bold=False):
    rPr = run_el.find(qn('a:rPr'))
    if rPr is None:
        rPr = run_el.makeelement(qn('a:rPr'), {})
        run_el.insert(0, rPr)
    fill = rPr.find(qn('a:solidFill'))
    if fill is not None:
        rPr.remove(fill)
    sf = rPr.makeelement(qn('a:solidFill'), {})
    clr = rPr.makeelement(qn('a:srgbClr'), {'val': '%02X%02X%02X' % (color[0], color[1], color[2])})
    sf.append(clr)
    rPr.append(sf)
    rPr.set('b', '1' if bold else '0')


def apply(prs, skip_first=2):
    for idx, slide in enumerate(prs.slides):
        if idx < skip_first:
            continue
        for sh in slide.shapes:
            if not sh.has_text_frame:
                continue
            if sh.top is not None and sh.top < 0.85 * 914400:
                continue                      # 页眉标题不动
            for p in sh.text_frame.paragraphs:
                for r in list(p.runs):
                    text = r.text
                    if len(text) < 12 or not any(k in text for k in KEYS):
                        continue
                    segs = _segments(text)
                    if not any(flag for _, flag in segs):
                        continue
                    base = copy.deepcopy(r._r)
                    r._r.getparent().remove(r._r)
                    for seg, hot in segs:
                        new = copy.deepcopy(base)
                        p._p.append(new)
                        run = p.runs[-1]
                        run.text = seg
                        if hot:
                            run.font.color.rgb = RED
                            run.font.bold = True
    return prs
