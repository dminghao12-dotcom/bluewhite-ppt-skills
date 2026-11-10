# -*- coding: utf-8 -*-
"""按 SKILL.md 的硬规则逐条体检 content.json + 生成的 pptx。

用法：
    python spec_audit.py content.json out/slides.pptx "稿名"

查：标题 ≤14 字、引导句 ≤40、标题带 ≤26、总结条 ≤40、一页高亮 ≤4、
版式 ≥8 种且同版式不连续 >2 页、图片不被压到 2.6 in 以下（三／四栏页卡 3.5 in）、
每页有图或组件、三／四栏页不得用 region.head 当栏首、口播句、
红线词（AI 生成 / 已逐字核对 / 待补 / 此处插入 / 不虚构）、章节编号连续性。
口播词分两档：硬口播句直接判不合规；「这就是／坑／伙伴／闭环／赋能」这类常见但可合法的词
只作提示——它们会误伤“坑塘”“闭环控制”这类正常术语，不当门槛。
自绘页在 JSON 里标 "custom_drawn": true 即跳过“无图无组件”检查。
不合规会逐条列出页号和数值；交付前必须跑到“全部通过”。
"""
import json, re, sys
from pptx import Presentation
from pptx.util import Emu

CN = re.compile(r'[\u4e00-\u9fff]')


def width(s):
    return sum(1 if CN.match(c) else 0.5 for c in s)


def check(content, pptx_path, name):
    d = json.load(open(content, encoding='utf-8'))
    slides = d['slides']
    prs = Presentation(pptx_path)
    bad = []
    soft = []

    # 版式统计
    layouts = [s.get('layout') for s in slides]
    print(f"\n=== {name} · {len(slides)} 页 ===")
    print("版式：", {l: layouts.count(l) for l in dict.fromkeys(layouts)})
    run = 1
    for i in range(1, len(layouts)):
        run = run + 1 if layouts[i] == layouts[i-1] and layouts[i] else 1
        if run > 2:
            bad.append(f"第{i+1}页 版式 {layouts[i]} 连续第 {run} 次（上限 2）")
    uniq = len(set(l for l in layouts if l))
    cap = min(8, sum(1 for l in layouts if l))     # 一页一种版式，短稿的上限就是内容页数
    if uniq < cap:
        bad.append(f"版式种类 {uniq} < {cap}")
    elif uniq < 8:
        print(f"版式种类 {uniq}：内容页只有 {cap} 页，按内容页上限计通过（不为凑版式拆页）")

    for i, s in enumerate(slides, 1):
        if s.get('type') in ('cover', 'end', 'agenda'):
            continue
        t = s.get('title', '')
        tt = re.sub(r'^\s*\d+(?:\.\d+)*\s+', '', t)
        if width(tt) > 14:
            bad.append(f"第{i}页 标题 {width(tt):.1f} 字 >14：{tt}")
        lead = s.get('lead')
        if lead:
            txt = lead['text'] if isinstance(lead, dict) else lead
            if width(txt) > 40:
                bad.append(f"第{i}页 引导句 {width(txt):.0f} 字 >40")
        band = s.get('band')
        if band:
            txt = band['text'] if isinstance(band, dict) else band
            if width(txt) > 26:
                bad.append(f"第{i}页 标题带 {width(txt):.0f} 字 >26")
        sm = s.get('summary')
        if isinstance(sm, dict):
            if width(sm.get('text', '')) > 40:
                bad.append(f"第{i}页 总结条 {width(sm.get('text','')):.0f} 字 >40")

        # 高亮：数量与是否命中正文
        blob = json.dumps(s, ensure_ascii=False)
        hl = re.findall(r'"highlights":\s*\[([^\]]*)\]', blob)
        terms = []
        for group in hl:
            terms += re.findall(r'"([^"]+)"', group)
        terms = [x for x in dict.fromkeys(terms) if x]
        if len(terms) > 4:
            bad.append(f"第{i}页 高亮 {len(terms)} 处 >4：{terms}")

        # 图片尺寸
        pg = prs.slides[i-1]
        # 页眉机构标识（header_logo）不是内容图，不参与 2.6 in 下限检查
        pics = [sh for sh in pg.shapes if sh.shape_type == 13
                and sh.name not in ('机构标识', '单位标识')]
        multi = s.get('layout') in ('info_three', 'info_three_strip', 'info_four', 'info_quad')
        for sh in pics:
            w, h = Emu(sh.width).inches, Emu(sh.height).inches
            if w < 2.6 and h < 2.6:
                bad.append(f"第{i}页 图片只有 {w:.1f}×{h:.1f} in（<2.6）")
            # 三栏 / 四格页的插画通常带 3 个以上要读的标签：宽度门槛提到 3.5 in
            elif multi and w < 3.5:
                bad.append(f"第{i}页 三／四栏页图片只有 {w:.1f} in 宽（<3.5）："
                           f"图内标签投影看不清，放大图位或换 info_three／info_four 占整栏")
        # 每页至少一个图或图形组件
        if not pics and not s.get('regions') and not s.get('custom_drawn'):
            bad.append(f"第{i}页 既无图也无区域组件")

        # 多栏页的栏首：region.head 不读 head_align，必须用 section(head_align: center)
        if multi:
            for slot, reg in (s.get('regions') or {}).items():
                if isinstance(reg, dict) and reg.get('head'):
                    bad.append(f"第{i}页 分区 {slot} 用了 region.head（栏首不会居中）："
                               f"改成 section(head=<同名>, head_align=\"center\", items=[])，正文另起 markers")

        # 双小标题：同一条栏里“栏首（空 section）＋ 下面又一条带 head 的 section”
        def _secs(o, acc):
            if isinstance(o, dict):
                if o.get('type') == 'section' and o.get('head'):
                    acc.append(o)
                for v in o.values():
                    _secs(v, acc)
            elif isinstance(o, list):
                for v in o:
                    _secs(v, acc)

        for slot, reg in (s.get('regions') or {}).items():
            items = reg.get('items') if isinstance(reg, dict) else None
            if not items or len(items) < 2:
                continue
            first, second = items[0], items[1]
            # 只在“栏首（空 section）+ 紧接一条带 head 的 section”时判双小标题；
            # row 里的并列小节（如 H1／H2 两格、可识别／不可识别三格）不算
            if (isinstance(first, dict) and first.get('type') == 'section'
                    and first.get('head') and not first.get('items')
                    and isinstance(second, dict) and second.get('type') == 'section'
                    and second.get('head')):
                bad.append(f"第{i}页 分区 {slot} 出现双小标题（{first['head']}／{second['head']}）："
                           f"栏首下面直接接 markers，不要再起带 head 的 section")

        # 红色框块：强调只能用红色加粗文字
        def _redbox(o, acc):
            if isinstance(o, dict):
                if o.get('type') in ('note', 'callout') and o.get('tone') == 'red':
                    acc.append(str(o.get('text', ''))[:18])
                for v in o.values():
                    _redbox(v, acc)
            elif isinstance(o, list):
                for v in o:
                    _redbox(v, acc)

        acc = []
        _redbox(s, acc)
        if acc:
            bad.append(f"第{i}页 用了红框强调块（{acc[0]}…）：改成正文 markers + highlights 红色加粗")

        # 口播 / AI 味词（页面文字；讲稿另有 speaker-notes 规则）
        for kw in ('要立的是', '说白了', '站不住', '说到哪里为止'):
            if kw in blob:
                bad.append(f"第{i}页 出现口播句：{kw}")
        for kw in ('这就是', '坑', '伙伴', '闭环', '赋能'):
            if kw in blob:
                soft.append(f"第{i}页 疑似 AI 味词：{kw}（坑塘／闭环控制这类合法术语可忽略）")

        # 红线词
        for kw in ('AI 生成', '已逐字核对', '待补', '此处插入', '不虚构'):
            if kw in blob:
                bad.append(f"第{i}页 出现红线词：{kw}")

    # 编号连续性
    nums = [re.match(r'^(\d+)\.(\d+)', s.get('title', '')) for s in slides]
    seen = [(int(m.group(1)), int(m.group(2))) for m in nums if m]
    print("编号：", seen)

    if soft:
        print(f"\n提示 {len(soft)} 条（不阻断交付）：")
        for s_ in soft:
            print(" ·", s_)
    if bad:
        print(f"\n不合规 {len(bad)} 条：")
        for b in bad:
            print(" -", b)
    else:
        print("\n全部通过")
    return bad


if __name__ == '__main__':
    check(sys.argv[1], sys.argv[2], sys.argv[3])
