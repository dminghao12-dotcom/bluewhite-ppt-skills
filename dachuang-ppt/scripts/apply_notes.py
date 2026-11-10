# -*- coding: utf-8 -*-
"""把讲稿 md 逐页写回 PPT 备注（讲稿单一来源，备注不另维护）。

用法：
    $PY scripts/apply_notes.py \\
        out/xxx-20页.pptx 讲稿-xxx-20页.md

讲稿 md 约定：每页一节，标题行以 `## ` 开头（`## 第 3 页 · 1.1 研究背景`），
节内的「建议用时：x.x 分钟」行会被丢弃，其余段落原样写进该页 notes，
首行是节标题。页数不一致直接抛 AssertionError——宁可构建失败，
也不要交出一份备注与讲稿错位的稿。
"""
import re
import sys

from pptx import Presentation


def parse(md_path):
    """返回 [(标题, 正文), ...]，只取 `## ` 开头的页段。"""
    txt = open(md_path, encoding='utf-8').read().replace('\r\n', '\n')
    pages = []
    for chunk in re.split(r'(?m)^## ', txt)[1:]:
        lines = chunk.split('\n')
        title = lines[0].strip()
        body = [l.strip() for l in lines[1:]
                if l.strip() and not l.startswith('建议用时')]
        pages.append((title, '\n'.join(body)))
    return pages


def main(pptx, md):
    pages = parse(md)
    prs = Presentation(pptx)
    assert len(pages) == len(prs.slides), \
        f'讲稿 {len(pages)} 段，幻灯片 {len(prs.slides)} 页：先对齐页序再写备注'
    for slide, (title, text) in zip(prs.slides, pages):
        slide.notes_slide.notes_text_frame.text = f'{title}\n{text}'
    prs.save(pptx)
    print(f'已写回备注：{len(pages)} 页 -> {pptx}')


if __name__ == '__main__':
    if len(sys.argv) < 3:
        raise SystemExit(__doc__)
    main(sys.argv[1], sys.argv[2])
