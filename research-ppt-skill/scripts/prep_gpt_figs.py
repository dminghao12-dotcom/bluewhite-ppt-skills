"""把网页端 GPT 的插画整理成能直接进页面的小图。

用法：
  python prep_gpt_figs.py 源图.png --name g-detect --cut-title --max-width 1200 --colors 200
  python prep_gpt_figs.py 合并图.png --name g-point --split        # 一张左右对照拆成两半
  python prep_gpt_figs.py 目录/ --cut-title --max-width 1200        # 批量（文件名即输出名）

做四件事，按需开关：
  1) 去白边：按像素差找内容框，留 pad 白边。
  2) 裁标题条：按行统计蓝像素，找到顶部那条标题带，在之后的第一个空白行切开。
     插画模型很喜欢在顶上画一条“电量要够”“预测给区间”——不裁会和页里的标题带重复。
  3) 左右拆半：一张画里放了两栏对照时，从中间切开当两张用（看它画的中缝在哪，必要时用 --split-at 微调）。
  4) 压图：降到 max_width、量化到 N 色。10 MB 插画能降到 5 MB，投影看不出差别。

依赖：Pillow + numpy（见技能根目录 INSTALL.md）。
"""
import argparse
import sys
from pathlib import Path

import numpy as np
from PIL import Image, ImageChops


def trim(im, thresh=8, pad=14):
    """去白边（白底图）。返回内容框外扩 pad 的图。"""
    bg = Image.new('RGB', im.size, (255, 255, 255))
    diff = ImageChops.difference(im, bg).convert('L')
    box = diff.point(lambda p: 255 if p > thresh else 0).getbbox()
    if not box:
        return im
    x0, y0, x1, y1 = box
    return im.crop((max(0, x0 - pad), max(0, y0 - pad),
                    min(im.width, x1 + pad), min(im.height, y1 + pad)))


def cut_title(im, band=0.45, on=0.08, off=0.005, pad_below=3):
    """裁掉顶部自带的标题条。band：只在前多少比例里找；on/off：蓝像素行占比阈值。

    机理：标题条是一整块蓝，行占比高；它和下面的内容之间必有一行白。
    找不到“蓝带后紧跟空白”就原图返回，宁可留着也不要切掉主体。
    """
    a = np.asarray(im.convert('RGB')).astype(int)
    blue = (a[:, :, 2] > 110) & (a[:, :, 0] < 90) & (a[:, :, 1] < 110)
    dens = blue.mean(axis=1)
    limit = int(len(dens) * band)
    start = next((i for i in range(limit) if dens[i] > on), None)
    if start is None:
        return im
    end = next((i for i in range(start, limit) if dens[i] < off), None)
    if end is None or end - start < 8:
        return im
    return im.crop((0, min(end + pad_below, im.height - 1), im.width, im.height))


def split_halves(im, at=0.5):
    """从中间切开，返回左右两张。"""
    w = im.width
    cut = int(w * at)
    return im.crop((0, 0, cut, im.height)), im.crop((cut, 0, w, im.height))


def compress(im, max_width=1200, colors=200):
    """降采样 + 量化；colors=0 表示只降采样。"""
    if max_width and im.width > max_width:
        im = im.resize((max_width, round(im.height * max_width / im.width)), Image.LANCZOS)
    if colors:
        im = im.quantize(colors=colors, method=Image.MEDIANCUT, dither=Image.NONE)
    return im


def process(src, out_dir, name=None, do_cut=True, max_width=1200, colors=200, do_split=False, split_at=0.5):
    im = Image.open(src).convert('RGB')
    im = trim(im)
    if do_cut:
        im = trim(cut_title(im))
    stem = name or Path(src).stem
    paths = []
    if do_split:
        left, right = split_halves(im, split_at)
        for suffix, part in (('-a', left), ('-b', right)):
            p = out_dir / f'{stem}{suffix}.png'
            compress(trim(part), max_width, colors).save(p, optimize=True)
            paths.append(p)
    else:
        p = out_dir / f'{stem}.png'
        compress(im, max_width, colors).save(p, optimize=True)
        paths.append(p)
    for p in paths:
        with Image.open(p) as o:
            print(f'{p}  {o.size[0]}x{o.size[1]}  {o.width / o.height:.2f}  {p.stat().st_size // 1024} KB')
    return paths


def main():
    ap = argparse.ArgumentParser(description='整理 GPT 插画：去白边 / 裁标题条 / 拆半 / 压图')
    ap.add_argument('src', help='源图或目录')
    ap.add_argument('--out', default=None, help='输出目录（默认与源图同目录）')
    ap.add_argument('--name', default=None, help='输出文件名（批量时忽略）')
    ap.add_argument('--no-cut-title', action='store_true', help='不裁顶部标题条')
    ap.add_argument('--split', action='store_true', help='左右拆成两张')
    ap.add_argument('--split-at', type=float, default=0.5, help='拆分位置比例，默认 0.5')
    ap.add_argument('--max-width', type=int, default=1200)
    ap.add_argument('--colors', type=int, default=200, help='量化色数，0 = 不量化')
    args = ap.parse_args()

    src = Path(args.src).expanduser()
    if not src.exists():
        sys.exit(f'找不到 {src}')
    out_dir = Path(args.out).expanduser() if args.out else src.parent
    out_dir.mkdir(parents=True, exist_ok=True)
    files = sorted(src.glob('*.png')) if src.is_dir() else [src]
    for f in files:
        process(f, out_dir, name=None if src.is_dir() else args.name,
                do_cut=not args.no_cut_title, max_width=args.max_width,
                colors=args.colors, do_split=args.split, split_at=args.split_at)


if __name__ == '__main__':
    main()
