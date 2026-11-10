#!/usr/bin/env python3
"""交付前三份文件的一致性验收：PPTX / PDF / Word 讲稿。

    python scripts/verify_delivery.py --pptx out/slides.pptx --pdf out/slides.pdf \
        --docx out/讲稿.docx [--rate 225]

检查：PPTX 页数 == PDF 页数；16:9；每页有备注；Word 里能按顺序找到每页备注的首句；
统计讲稿字符数与估算时长；打印三份文件的 sha256 前 12 位。
任一项失败打印 FAIL 并以 1 退出——不要只用“文件生成了”宣布交付。
"""
import argparse
import hashlib
import json
import re
import sys
from pathlib import Path


def sha(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()[:12]


def main() -> int:
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument('--pptx', type=Path, required=True)
    ap.add_argument('--pdf', type=Path, required=True)
    ap.add_argument('--docx', type=Path)
    ap.add_argument('--rate', type=int, default=225, help='讲稿语速，字/分钟')
    ap.add_argument('--min-note-chars', type=int, default=10,
                    help='低于这个长度的备注视为空（结束页 15 字左右是正常的）')
    args = ap.parse_args()

    from pptx import Presentation
    import pymupdf

    report = {'errors': [], 'warnings': []}
    prs = Presentation(args.pptx)
    pdf = pymupdf.open(args.pdf)
    slides = list(prs.slides)
    notes = [(s.notes_slide.notes_text_frame.text.strip() if s.has_notes_slide else '')
             for s in slides]

    report['pptx_pages'] = len(slides)
    report['pdf_pages'] = pdf.page_count
    ratio = prs.slide_width / prs.slide_height
    report['aspect'] = round(ratio, 3)
    if len(slides) != pdf.page_count:
        report['errors'].append(f'PPTX {len(slides)} 页与 PDF {pdf.page_count} 页不一致')
    if abs(ratio - 16 / 9) > 0.01:
        report['errors'].append(f'页面不是 16:9（{report["aspect"]}）')

    empty = [i + 1 for i, n in enumerate(notes) if len(n) < args.min_note_chars]
    if empty:
        report['errors'].append(f'以下页备注为空：{empty}')

    plain = [re.sub(r'\s', '', n) for n in notes]
    total = sum(len(t) for t in plain)
    report['note_chars'] = total
    report['est_minutes'] = round(total / args.rate, 1)

    if args.docx:
        from docx import Document
        body = '\n'.join(p.text for p in Document(args.docx).paragraphs)
        missing = [i + 1 for i, n in enumerate(notes)
                   if n and n.split('\n')[0][:20] not in body]
        if missing:
            report['errors'].append(
                f'Word 未包含以下页的备注首句（备注与讲稿不同源）：{missing}')
        if '讲稿' not in body[:200]:
            report['warnings'].append('Word 开头不像 notes_docx 的标题行，确认用的是导出的那一份')

    report['files'] = {p.name: {'mb': round(p.stat().st_size / 1048576, 2), 'sha256': sha(p)}
                       for p in (args.pptx, args.pdf, args.docx) if p}
    report['ok'] = not report['errors']
    print(json.dumps(report, ensure_ascii=False, indent=2))
    return 0 if report['ok'] else 1


if __name__ == '__main__':
    sys.exit(main())
