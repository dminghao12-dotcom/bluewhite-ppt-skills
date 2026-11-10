#!/usr/bin/env python3
"""只裁白边，不裁标题条 —— 生图插画贴进 PPT 前的默认整理器。

与 prep_gpt_figs.py 的分工：那个默认会把图上方独立标题条一起裁掉（对论文版式合适），
本脚本**只去掉四周纯白/近白边距**，保留完整画布、图内所有标签与图例。
生图端交付的图常带 10–30% 的白边，直接贴进图位会让主体显得只有“邮票大”。

用法：
    $PY scripts/trim_margins.py images/raw/ --out images            # 批量，文件名不变
    $PY scripts/trim_margins.py a.png b.png --out images --pad 0.015
    $PY scripts/trim_margins.py images/ --out images --dry-run      # 只报尺寸，不写文件

参数：
    --thresh  0–255 的灰度差阈值，默认 10（≈每通道 10 级以内算白）
              浅蓝衬底、浅灰网格要保留时降到 6–8；扫描件噪点多时提到 14–18
    --pad     内边距 = pad × min(宽,高)，默认 0.012；不要设成 0，贴边会切到投影描边
    --max-w   可选，裁完超宽时等比缩到该像素宽（默认不缩，普通插图另行压到 ~1200 px 再量化 200 色）
    --quality JPEG 质量，默认 93；PNG 输入仍存 PNG

输出一份 JSON 摘要（原尺寸 → 裁后尺寸、裁掉比例、输出路径），便于批量核对。
"""
import argparse, json, sys
from pathlib import Path
from PIL import Image, ImageChops, ImageOps

EXTS = {'.png', '.jpg', '.jpeg', '.webp'}


def trim(path: Path, out_dir: Path, thresh: int, pad_ratio: float,
         max_w: int | None, quality: int, dry: bool) -> dict:
    im = Image.open(path)
    im = ImageOps.exif_transpose(im)
    rgb = im.convert('RGB')
    w0, h0 = rgb.size

    bg = Image.new('RGB', rgb.size, (255, 255, 255))
    diff = ImageChops.difference(rgb, bg).convert('L')
    mask = diff.point(lambda v: 255 if v > thresh else 0)
    box = mask.getbbox()
    if box is None:                      # 全白图：原样输出，不裁成 1×1
        box = (0, 0, w0, h0)

    pad = int(round(pad_ratio * min(w0, h0)))
    l = max(0, box[0] - pad); t = max(0, box[1] - pad)
    r = min(w0, box[2] + pad); b = min(h0, box[3] + pad)
    out = rgb.crop((l, t, r, b))

    if max_w and out.width > max_w:
        ratio = max_w / out.width
        out = out.resize((max_w, int(round(out.height * ratio))), Image.Resampling.LANCZOS)

    dst = out_dir / path.name
    kept = (out.width * out.height) / float(w0 * h0)
    rec = {'src': str(path), 'dst': str(dst), 'from': [w0, h0], 'to': list(out.size),
           'kept_ratio': round(kept, 3), 'box': [l, t, r, b]}
    if dry:
        return rec

    out_dir.mkdir(parents=True, exist_ok=True)
    if dst.suffix.lower() in ('.jpg', '.jpeg'):
        out.save(dst, quality=quality, optimize=True, subsampling=1)
    else:
        out.save(dst, optimize=True)
    return rec


def collect(targets):
    files = []
    for t in targets:
        p = Path(t)
        if p.is_dir():
            files += sorted(f for f in p.iterdir() if f.suffix.lower() in EXTS)
        elif p.suffix.lower() in EXTS:
            files.append(p)
    return files


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('targets', nargs='+', help='图片文件或目录')
    ap.add_argument('--out', required=True, help='输出目录（写回源目录时显式写同一路径）')
    ap.add_argument('--thresh', type=int, default=10)
    ap.add_argument('--pad', type=float, default=0.012)
    ap.add_argument('--max-w', type=int, default=None)
    ap.add_argument('--quality', type=int, default=93)
    ap.add_argument('--dry-run', action='store_true')
    a = ap.parse_args()

    files = collect(a.targets)
    if not files:
        print('没有找到图片', file=sys.stderr)
        return 1

    recs = []
    for f in files:
        try:
            recs.append(trim(f, Path(a.out), a.thresh, a.pad, a.max_w, a.quality, a.dry_run))
        except Exception as e:                        # 单张失败不拖垮整批
            recs.append({'src': str(f), 'error': repr(e)})
    print(json.dumps({'count': len(recs), 'dry_run': a.dry_run, 'items': recs},
                     ensure_ascii=False, indent=1))
    bad = [r for r in recs if 'error' in r]
    return 1 if bad else 0


if __name__ == '__main__':
    sys.exit(main())
