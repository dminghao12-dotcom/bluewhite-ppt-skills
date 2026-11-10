# -*- coding: utf-8 -*-
"""从一份成稿里抽出「默认模板壳」（封面 / 目录 / 成果 / 结束），压图并重编 partname。

用法：
    $PY scripts/make_default.py <成稿.pptx> <输出.pptx>

没有用户模板时，就把输出这份当底：自绘页走 deck.merge / deck.lib，模板页只换文字与图片。
"""
import io
import os
import sys

from PIL import Image
from pptx import Presentation
from pptx.opc.packuri import PackURI
from pptx.oxml.ns import qn

# 成稿里这四页的位置（0 基）：封面、目录、成果、结束
KEEP = (0, 1, -3, -1)


def build(src, dst, max_w=1800, max_bytes=600_000):
    prs = Presentation(src)
    n = len(prs.slides)
    keep = {i if i >= 0 else n + i for i in KEEP}

    lst = prs.slides._sldIdLst
    for i, e in enumerate(list(lst)):
        if i in keep:
            continue
        prs.part.drop_rel(e.get(qn('r:id')))
        lst.remove(e)

    # 压图：>600KB 或 >1800px 的图重采样成 jpg 再换 r:embed
    tmp = os.path.join(os.path.dirname(dst) or '.', '_tmp_pic.jpg')
    for s in prs.slides:
        for sh in s.shapes:
            if sh.shape_type != 13:
                continue
            blob = sh.image.blob
            im = Image.open(io.BytesIO(blob)).convert('RGB')
            if len(blob) < max_bytes and im.width <= max_w:
                continue
            w = min(max_w, im.width)
            im = im.resize((w, round(im.height * w / im.width)), Image.LANCZOS)
            im.save(tmp, quality=82, optimize=True)
            _, rId = s.part.get_or_add_image_part(tmp)
            sh._element.blipFill.blip.rEmbed = rId
    if os.path.exists(tmp):
        os.remove(tmp)

    parts = [s.part for s in prs.slides]
    for i, p in enumerate(parts):
        p.partname = PackURI(f'/ppt/slides/tmp{i}.xml')
    for i, p in enumerate(parts, 1):
        p.partname = PackURI(f'/ppt/slides/slide{i}.xml')

    prs.save(dst)
    print(f'saved {dst}  slides={len(prs.slides)}  {os.path.getsize(dst)/1e6:.1f} MB')
    return dst


if __name__ == '__main__':
    if len(sys.argv) < 3:
        print(__doc__)
        sys.exit(1)
    build(sys.argv[1], sys.argv[2])
