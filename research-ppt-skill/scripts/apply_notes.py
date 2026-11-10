# -*- coding: utf-8 -*-
"""把讲稿单一源写进 PPT 每页备注（Word 讲稿随后由 notes_docx.py 从备注导出）。

用法（在项目目录里，与 script_content.py 同级）：
    python <技能目录>/scripts/apply_notes.py out/slides.pptx
    python <技能目录>/scripts/notes_docx.py --input out/slides.pptx --output out/讲稿.docx

约定：项目目录下的 script_content.py 里写
    SCRIPT = [("第 1 页 · 封面", "正文…"), ...]
条数必须等于幻灯片页数，否则 assert 直接失败——这就是“备注与 Word 不会对不上”的保证。
"""
import sys
from pathlib import Path

from pptx import Presentation



def main() -> None:
    path = Path(sys.argv[1] if len(sys.argv) > 1 else 'out/slides.pptx')
    sys.path.insert(0, str(Path.cwd()))
    from script_content import SCRIPT  # noqa: E402  项目目录下的单一源

    prs = Presentation(str(path))
    assert len(prs.slides) == len(SCRIPT), (
        f'页数 {len(prs.slides)} 与 SCRIPT 条数 {len(SCRIPT)} 不一致：先同步 script_content.py'
    )
    for slide, (_title, body) in zip(prs.slides, SCRIPT):
        tf = slide.notes_slide.notes_text_frame
        tf.clear()
        tf.text = body
    total = sum(len(body) for _title, body in SCRIPT)
    prs.save(str(path))
    print(f'notes applied: {len(SCRIPT)} pages, {total} 字, 约 {total / 250:.1f} 分钟')


if __name__ == '__main__':
    main()
