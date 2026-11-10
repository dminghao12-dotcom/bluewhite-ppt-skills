# -*- coding: utf-8 -*-
"""把 PPTX 里所有 typeface 统一成同一款字体（默认场景：课程要求全稿黑体）。

用法：
    python set_font.py out/slides.pptx 黑体
    python set_font.py out/slides.pptx 宋体 --only 微软雅黑,等线

为什么用后处理而不是改 info_style.py：脚本里 latin 写死 Arial、东亚写死微软雅黑，
改源码会污染其他稿子；替换所有 XML 的 typeface 只影响这一份成品。
排在 dense_deck.py 之后跑，并写进 build.sh——重建会把它冲掉。
"""
import sys, zipfile, shutil

DEFAULT_OLD = ['微软雅黑', 'Arial', 'Calibri', '等线', 'Helvetica']


def main():
    if len(sys.argv) < 3:
        print(__doc__)
        raise SystemExit(2)
    src, target = sys.argv[1], sys.argv[2]
    olds = DEFAULT_OLD
    if '--only' in sys.argv:
        olds = [x for x in sys.argv[sys.argv.index('--only') + 1].split(',') if x]

    tmp = src + '.tmp'
    cnt = 0
    with zipfile.ZipFile(src) as zin, zipfile.ZipFile(tmp, 'w', zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename.endswith('.xml'):
                text = data.decode('utf-8')
                for old in olds:
                    if old in text:
                        cnt += text.count(old)
                        text = text.replace(old, target)
                data = text.encode('utf-8')
            zout.writestr(item, data)
    shutil.move(tmp, src)
    print('typeface -> %s：%d 处' % (target, cnt))


if __name__ == '__main__':
    main()
