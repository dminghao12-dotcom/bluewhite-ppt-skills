#!/usr/bin/env python
"""从资料 PPT 抽取的图里挑出本次申报稿要用的，裁白边、转 RGB、压到合适尺寸。"""
import os
from PIL import Image, ImageChops

SRC = os.path.join(os.path.dirname(__file__), '素材', 'media')
OUT = os.path.join(os.path.dirname(__file__), 'images')
os.makedirs(OUT, exist_ok=True)

# (源文件, 输出名, 是否裁白边, 最大宽度)
PICKS = [
    ('s100_b45d53bc.jpeg', 'wellsite_raw', False, 1600),  # 井场航拍（再裁横条）
    ('s083_d5ff3f12.png',  'apparatus', True, 1100),     # 岩心驱替装置
    ('s082_04c5801c.jpeg', 'recovery', True, 1100),      # 采收率曲线
    ('s061_1ffc8458.png',  'betaine', True, 1400),       # 甜菜碱结构
    ('s086_6c8a7408.png',  'prod-curve', True, 1400),    # 胜利井生产曲线
    ('s094_4063b52e.png',  'ift-salt', True, 1500),      # 增溶指数与IFT vs 盐度
    ('s069_5adf5f27.png',  'phase-vol', True, 1200),     # 相体积分数 vs NaCl
    ('s094_dc35501f.png',  'tubes', True, 1400),         # 相态试管照片
    ('s067_73029d1b.png',  'mixed-film', True, 1200),    # 复配界面膜示意
    ('s090_c850ece1.png',  'micro', True, 1200),         # 微乳液-原油显微照片
    ('s080_bbfe9e90.png',  'coreflood-rig', True, 1200), # 岩心驱替装置
    ('s094_1d227c6a.png',  'ift-conc', True, 1400),      # IFT vs 浓度 三条
    ('s083_75fa8b10.jpeg', 'field-tank', True, 1000),    # 现场配液罐
    ('s090_a23207d9.png',  'emulsion-micro', True, 1000),
]


def trim(im, thresh=248, pad=6):
    rgb = im.convert('RGB')
    bg = Image.new('RGB', rgb.size, (255, 255, 255))
    diff = ImageChops.difference(rgb, bg).convert('L')
    mask = diff.point(lambda v: 255 if v > 255 - thresh else 0)
    bbox = mask.getbbox()
    if not bbox:
        return rgb
    l, t, r, b = bbox
    l, t = max(0, l - pad), max(0, t - pad)
    r, b = min(rgb.width, r + pad), min(rgb.height, b + pad)
    if (r - l) < rgb.width * 0.15 or (b - t) < rgb.height * 0.15:
        return rgb
    return rgb.crop((l, t, r, b))


for src, name, do_trim, maxw in PICKS:
    p = os.path.join(SRC, src)
    im = Image.open(p).convert('RGB')
    if do_trim:
        im = trim(im)
    if im.width > maxw:
        h = round(im.height * maxw / im.width)
        im = im.resize((maxw, h), Image.LANCZOS)
    out = os.path.join(OUT, name + '.png')
    im.save(out, optimize=True)
    print(f'{name:16s} {im.width}x{im.height}  {os.path.getsize(out)//1024} KB')

# 井场图裁成 1.78:1 横条（纵向取中部偏上，保留井架）
src = os.path.join(OUT, 'wellsite_raw.png')
im = Image.open(src).convert('RGB')
w, h = im.size
target = 1.78
nh = int(w / target)
top = min(max(0, int(h * 0.30)), h - nh)
im = im.crop((0, top, w, top + nh))
if im.width > 1400:
    im = im.resize((1400, round(im.height * 1400 / im.width)), Image.LANCZOS)
im.save(os.path.join(OUT, 'wellsite.png'), optimize=True)
os.remove(src)
print(f"{'wellsite':16s} {im.width}x{im.height}")
