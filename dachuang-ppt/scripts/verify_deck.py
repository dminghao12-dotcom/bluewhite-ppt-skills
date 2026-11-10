# -*- coding: utf-8 -*-
"""交付前核验：页数一致、16:9、每页备注非空、方括号残留、每页正文红字上限。

用法：
    # 有 PDF（最稳）：页数/画幅/方括号残留都能交叉核对
    $PY scripts/verify_deck.py \
        --pptx out/x.pptx --pdf out/x.pdf --docx out/x-讲稿.docx

    # 没有 PDF（PowerPoint COM 导出失败、或本机只有 LibreOffice 懒得转）：只查 PPTX 自身
    $PY scripts/verify_deck.py --pptx out/x.pptx

`--pdf` 是可选的。缺 PDF 时跳过「PPTX 与 PDF 页数一致」和「PDF 里有没有残留【】」这两项，
画幅改用 PPTX 自身尺寸判，并在输出里明确提示哪几项没查——不要把跳过当成通过。

判据全部可机判，别用「我看过了」代替。退出码 0 才算过检。
"""
import argparse
import os
import re
import sys

EMU_PER_IN = 914400


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--pptx', required=True)
    ap.add_argument('--pdf', default=None,
                    help='可选。没有 PDF 时跳过页数交叉核对与方括号残留检查')
    ap.add_argument('--docx', default=None)
    ap.add_argument('--red-max', type=int, default=4,
                    help='每页正文红字（#C00000）上限，尾条的 #FF6060 不计入')
    a = ap.parse_args()

    from pptx import Presentation

    errs, warns, skipped = [], [], []
    prs = Presentation(a.pptx)
    n_ppt = len(prs.slides)

    # ---------------------------------------------------------------- PDF（可选）
    pdf = None
    if a.pdf:
        if not os.path.exists(a.pdf):
            print('ERROR 指定的 PDF 不存在：%s' % a.pdf)
            return 1
        import pymupdf
        pdf = pymupdf.open(a.pdf)
        n_pdf = len(pdf)
        if n_ppt != n_pdf:
            errs.append('页数不一致：PPTX %d / PDF %d' % (n_ppt, n_pdf))
        w = pdf[0].rect.width / 72.0
        h = pdf[0].rect.height / 72.0
        size_src = 'PDF'
    else:
        n_pdf = None
        w = prs.slide_width / EMU_PER_IN
        h = prs.slide_height / EMU_PER_IN
        size_src = 'PPTX'
        skipped.append('PPTX 与 PDF 页数一致（需要 --pdf）')
        skipped.append('PDF 里是否残留【】标记（需要 --pdf）')

    if abs(w / h - 16 / 9.0) > 0.02:
        errs.append('页面比例不是 16:9：%.3f × %.3f in（按 %s 判）' % (w, h, size_src))

    # ---------------------------------------------------------------- PPTX 自身
    for i, s in enumerate(prs.slides, 1):
        if not s.notes_slide.notes_text_frame.text.strip():
            errs.append('第 %d 页没有备注' % i)
        red = 0
        for sh in s.shapes:
            if not sh.has_text_frame:
                continue
            for p in sh.text_frame.paragraphs:
                for r in p.runs:
                    try:
                        c = str(r.font.color.rgb)
                    except Exception:
                        continue
                    if c == 'C00000' and r.text.strip():
                        red += 1
        if red > a.red_max:
            errs.append('第 %d 页正文红字 %d 处，超过 %d 处' % (i, red, a.red_max))

    if pdf is not None:
        txt = ''.join(p.get_text() for p in pdf)
        left = re.findall(r'[【】]', txt)
        if left:
            errs.append('PDF 里残留 %d 个方括号标记：%s' % (len(left), left[:6]))

    if a.docx:
        import docx
        d = docx.Document(a.docx)
        heads = [p.text for p in d.paragraphs if p.text.strip().startswith('第 ')]
        if len(heads) != n_ppt:
            warns.append('讲稿节数 %d ≠ PPT 页数 %d' % (len(heads), n_ppt))

    # ---------------------------------------------------------------- 报告
    pages = 'PPTX %d 页' % n_ppt + (' / PDF %d 页' % n_pdf if n_pdf is not None else '')
    print('%s / %.3f × %.3f in（按 %s 判）' % (pages, w, h, size_src))
    for e in errs:
        print('ERROR', e)
    for x in warns:
        print('WARN ', x)
    for s in skipped:
        print('SKIP ', s)

    if not a.pdf:
        print('\n拿 PDF 的三条路（任选一条，拿到后加 --pdf 重跑可全项核对）：')
        print('  macOS/LibreOffice: /Applications/LibreOffice.app/Contents/MacOS/soffice '
              '--headless --convert-to pdf --outdir out <pptx>')
        print('  Windows/PowerPoint: 手动打开 → 另存为 PDF（COM 报 0x80004005 时就用这条）')
        print('  已装 pymupdf 的场景：先跑 render.py 产出的 out/render/<同名>.pdf 可直接复用')

    print('全部通过' if not errs else '%d 条错误' % len(errs))
    return 1 if errs else 0


if __name__ == '__main__':
    sys.exit(main())
