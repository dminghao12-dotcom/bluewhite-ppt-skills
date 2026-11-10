# -*- coding: utf-8 -*-
"""讲稿 md 按页计时：统计每页正文字数，重写「建议用时」，重建汇总表。

    $PY scripts/script_timing.py 讲稿.md [--rate 4.3] [--dry]

- 只统计「## 第 N 页 · 标题」段落里的正文：跳过 `建议用时：` 行与标题行。
- 重写每页的 `建议用时：x.x 分钟`，并把页首汇总表按「章」分组重建。
- 速率默认 4.3 字/秒（= 258 字/分钟）；口语改书面语后同一套内容约涨 15%，所以要重跑。
"""
import re
import sys

RATE = 4.3 * 60  # 字/分钟
DEFAULT_RATE = 4.3


def parse_sections(lines):
    """返回 [(起, 止, 页号, 标题)]，段落以 '## ' 起行。"""
    idx = [i for i, l in enumerate(lines) if l.startswith('## ')]
    secs = []
    for k, i in enumerate(idx):
        end = idx[k + 1] if k + 1 < len(idx) else len(lines)
        title = lines[i][3:].strip()
        m = re.search(r'第\s*(\d+)\s*页', title)
        secs.append((i, end, int(m.group(1)) if m else None, title))
    return secs


def body_chars(lines, start, end):
    n = 0
    for l in lines[start + 1:end]:
        s = l.strip()
        if not s or s.startswith('建议用时') or s == '---':
            continue
        n += len(re.sub(r'\s', '', s))
    return n


def chapter_of(title):
    m = re.search(r'·\s*(\d+)\.\d+', title)
    return m.group(1) if m else None


def main(argv):
    if not argv:
        raise SystemExit(__doc__)
    path = argv[0]
    rate = DEFAULT_RATE
    dry = '--dry' in argv
    if '--rate' in argv:
        rate = float(argv[argv.index('--rate') + 1])
    per_min = rate * 60

    text = open(path, encoding='utf-8').read()
    lines = text.replace('\r\n', '\n').split('\n')
    secs = parse_sections(lines)
    if not secs:
        raise SystemExit('没找到 "## 第 N 页" 段落，检查讲稿结构。')

    stats = []
    for start, end, page, title in secs:
        n = body_chars(lines, start, end)
        stats.append({'page': page, 'title': title, 'chars': n, 'min': n / per_min})
    total = sum(s['chars'] for s in stats)

    # 重写每页「建议用时」
    out = list(lines)
    for (start, end, page, title), st in zip(secs, stats):
        new_line = f"建议用时：{st['min']:.1f} 分钟"
        replaced = False
        for i in range(start + 1, end):
            if lines[i].startswith('建议用时'):
                out[i] = new_line
                replaced = True
                break
        if not replaced:
            out.insert(start + 1, new_line)
            out.insert(start + 2, '')

    # 重建汇总表：按章分组，没有章号的页单独一行
    groups = []
    for st in stats:
        ch = chapter_of(st['title'])
        if ch and groups and groups[-1][0] == ch:
            groups[-1][1].append(st)
        else:
            groups.append((ch, [st]))
    tbl = ['| 页 | 内容 | 建议用时 |', '|---|---|---|']
    for ch, items in groups:
        pages = [s['page'] for s in items if s['page'] is not None]
        rng = f'{min(pages)}' if len(pages) == 1 else f'{min(pages)}–{max(pages)}'
        name = re.sub(r'^第\s*\d+\s*页\s*·\s*', '', items[0]['title'])
        name = re.sub(r'^\d+\.\d+\s*', '', name)
        tbl.append(f"| {rng} | {name} | {sum(s['min'] for s in items):.1f} 分 |")
    tbl.append(f"| — | **合计** | **{total / per_min:.1f} 分** |")

    # 把首页到第一个 '---' 之间的表格换成新表
    first_sec = secs[0][0]
    head_end = next((i for i in range(first_sec) if lines[i].strip() == '---'), first_sec)
    keep_head = [l for l in lines[:head_end] if not l.startswith('|')]
    while keep_head and not keep_head[-1].strip():
        keep_head.pop()
    head = keep_head + tbl + ['', '---', '']
    new_text = '\n'.join(head + out[first_sec:])
    new_text = re.sub(r'\*\*总字数\*\*：约[^　]*', f'**总字数**：约 {round(total, -1):.0f} 字', new_text)
    new_text = re.sub(r'\*\*估算时长\*\*：[^（(]*',
                     f'**估算时长**：{total / per_min:.1f} 分钟', new_text)

    if not dry:
        open(path, 'w', encoding='utf-8').write(new_text)
    for s in stats:
        print(f"{str(s['page']):>3}  {s['title'][:26]:28s} {s['chars']:5d} 字  {s['min']:4.1f} 分")
    print(f"合计 {total} 字 ≈ {total / per_min:.1f} 分钟（{rate} 字/秒）")
    if dry:
        print('（--dry：只统计，未写回）')


if __name__ == '__main__':
    main(sys.argv[1:])
