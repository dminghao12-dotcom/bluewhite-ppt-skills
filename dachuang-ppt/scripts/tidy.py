# -*- coding: utf-8 -*-
"""收尾统一：文本框垂直居中、表格整格居中。"""
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.oxml.ns import qn

HAS_MARKER = ('■', '·', '▸', '●', '◆', '–', '—')


def _body_pr(tf):
    bodyPr = tf._txBody.find(qn('a:bodyPr'))
    return bodyPr


def center_boxes(prs, long_chars=60):
    """把每个文本框内容在框内整体居中：垂直居中；短文本水平也居中。"""
    for slide in prs.slides:
        for sh in slide.shapes:
            if not sh.has_text_frame:
                continue
            tf = sh.text_frame
            if (sh.name or '').startswith(('toc', 'la')) or (sh.top is not None and sh.top < 0.85 * 914400):
                continue          # 页眉区（编号+页标题）保持左对齐
            text = tf.text.strip()
            if not text:
                continue
            bodyPr = _body_pr(tf)
            if bodyPr is not None:
                bodyPr.set('anchor', 'ctr')
                for attr in ('vert', 'vert270'):
                    if bodyPr.get(attr) is not None:
                        del bodyPr.attrib[attr]
            # 短标签/短句：水平居中；长段落与项目符号列表保持左对齐
            short = len(text.replace('\n', '')) <= 28
            bulleted = any(m in text for m in HAS_MARKER) or '\u2022' in text
            if short and not bulleted:
                for p in tf.paragraphs:
                    p.alignment = PP_ALIGN.CENTER


def center_tables(prs, long_chars=200):
    """表格：所有单元格水平、垂直都居中。"""
    for slide in prs.slides:
        for sh in slide.shapes:
            if not sh.has_table:
                continue
            for row in sh.table.rows:
                for cell in row.cells:
                    cell.vertical_anchor = MSO_ANCHOR.MIDDLE
                    txt = cell.text_frame.text.strip()
                    al = PP_ALIGN.CENTER if len(txt) <= long_chars else PP_ALIGN.LEFT
                    for p in cell.text_frame.paragraphs:
                        p.alignment = al


def apply(prs):
    center_boxes(prs)
    center_tables(prs)
    return prs
