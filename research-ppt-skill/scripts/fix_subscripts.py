# -*- coding: utf-8 -*-
"""把 CODMn 的 n 改成真下标（baseline run），文本框与表格单元格都处理。

用法：
    python scripts/fix_subscripts.py out/汇报.pptx

放在 dense_deck.py / paper_cover.py 之后跑（会就地改写文件），写进 build.sh。
重复跑是安全的：第一次拆完，'CODMn' 已不存在于单个 run 里，第二次找不到就不动。
Unicode 下标字符（ₙ、₃）在这里一律不用——微软雅黑缺字形会渲成方框。
"""
import copy
import sys

from pptx import Presentation
from pptx.oxml.ns import qn

# (整串, 正体, 下标)；要加 NO3 之类，同样把下标部分写进来
PAIRS = [('CODMn', 'CODM', 'n')]
BASELINE = '-25000'


def fix_paragraph(para):
    n = 0
    for run in list(para.runs):
        txt = run.text
        for whole, base, sub in PAIRS:
            if whole not in txt:
                continue
            parts = txt.split(whole)
            r = run._r
            new_elems = []
            for i, seg in enumerate(parts):
                if seg:
                    e = copy.deepcopy(r)
                    e.find(qn('a:t')).text = seg
                    new_elems.append(e)
                if i < len(parts) - 1:
                    e1 = copy.deepcopy(r)
                    e1.find(qn('a:t')).text = base
                    e1.find(qn('a:rPr')).attrib.pop('baseline', None)
                    new_elems.append(e1)
                    e2 = copy.deepcopy(r)
                    e2.find(qn('a:t')).text = sub
                    e2.find(qn('a:rPr')).set('baseline', BASELINE)
                    new_elems.append(e2)
            parent = r.getparent()
            idx = parent.index(r)
            parent.remove(r)
            for k, e in enumerate(new_elems):
                parent.insert(idx + k, e)
            n += 1
            break
    return n


def main(path):
    prs = Presentation(path)
    total = 0
    for slide in prs.slides:
        for shape in slide.shapes:
            if shape.has_text_frame:
                for para in shape.text_frame.paragraphs:
                    total += fix_paragraph(para)
            if shape.has_table:
                for row in shape.table.rows:
                    for cell in row.cells:
                        for para in cell.text_frame.paragraphs:
                            total += fix_paragraph(para)
    prs.save(path)
    print('CODMn 下标修正：', total, '处 ->', path)


if __name__ == '__main__':
    if len(sys.argv) < 2:
        sys.exit('用法: python scripts/fix_subscripts.py <deck.pptx>')
    main(sys.argv[1])
