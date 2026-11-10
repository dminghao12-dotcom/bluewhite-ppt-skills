# -*- coding: utf-8 -*-
"""合并/精简后的四页：3.2（不足+关键问题）、4.1（五项创新）、5.1（路线+进度）、5.2（配方+界面）。

直接用 deck.lib 画在用户那份 pptx 上（新增页），不动模板页。
"""
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE

from deck.lib import (header, judgment, card, textbox, para, bullets, shape, picture,
                      caption, conclusion, chip, badge, arrow_right, sline, gradient,
                      BLUE, BLUE_L, BLUE_D, RED, TXT, GRAY, WHITE, BORDER, BG_SOFT)


# --------------------------------------------------------------- 3.2 不足 → 问题 → 验证
def p32(prs):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    header(s, '3.2', '现有研究不足与拟解决的关键问题')
    judgment(s, '三处不足与三个关键问题一一对应，每个问题都指定了回答它的实验')

    # 列头
    for x, w, t in ((1.02, 2.60, '现有研究的不足'), (4.18, 3.90, '本项目拟解决的关键问题'),
                    (8.58, 3.90, '对应的验证方式')):
        box = textbox(s, x, 1.54, w, 0.34, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
        para(box.text_frame, t, size=13.5, color=BLUE, bold=True, first=True,
             align=PP_ALIGN.CENTER, space_after=0)

    rows = [
        ('1', '窗口容错性不明确',
         '混入、稀释与吸附后相区偏移，最优配方可能失效',
         '复配比例、浓度与盐度如何共同决定中相形成，怎样获得有容错性的稳定相区？',
         '盐度、浓度与复配比例扫描；混入、稀释、吸附后复测相态与相体积'),
        ('2', '注入性缺少协同评价',
         '增油与驱替压差未同时记录，流度调整与孔喉堵塞难以区分',
         '较低界面张力与较强增溶能否在吸附、稀释与温度变化后保持，并转化为可测增油？',
         '动态与平衡界面张力、增溶参数；温度、钙镁、老化与吸附前后对比'),
        ('3', '配方筛选依赖试错',
         '缺少统一记录口径的相态与性能数据，成相边界难以预测',
         '如何兼顾洗油能力、药剂用量与注入性，识别有利流度调整与不利孔喉堵塞？',
         '岩心驱替同步记录采收率、含水、压差与注入体积；对照组与平行样核查'),
    ]
    y = 1.96
    for i, (num, title, gap, q, ver) in enumerate(rows):
        if i % 2 == 0:
            shape(s, MSO_SHAPE.ROUNDED_RECTANGLE, 0.42, y - 0.06, 12.5, 1.02,
                  fill=BG_SOFT, line=None, radius=0.06)
        badge(s, 0.52, y + 0.26, 0.44, num, fill=BLUE, size=14)
        box = textbox(s, 1.02, y, 2.52, 0.42)
        box.name = 'la'
        para(box.text_frame, title, size=16, color=BLUE, bold=True, first=True,
             align=PP_ALIGN.LEFT, space_after=0)
        box = textbox(s, 1.02, y + 0.40, 2.52, 0.56)
        box.name = 'la'
        para(box.text_frame, gap, size=13, color=TXT, first=True, align=PP_ALIGN.LEFT,
             space_after=0, line=1.1)
        arrow_right(s, 3.66, y + 0.34, 0.34, 0.22, fill=BLUE)
        box = textbox(s, 4.18, y + 0.02, 3.86, 0.90, anchor=MSO_ANCHOR.MIDDLE)
        box.name = 'la'
        para(box.text_frame, q, size=15, color=TXT, first=True, align=PP_ALIGN.LEFT,
             space_after=0, line=1.16)
        arrow_right(s, 8.16, y + 0.34, 0.30, 0.22, fill=BLUE)
        box = textbox(s, 8.58, y + 0.02, 3.90, 0.90, anchor=MSO_ANCHOR.MIDDLE)
        box.name = 'la'
        para(box.text_frame, ver, size=13.5, color=TXT, first=True, align=PP_ALIGN.LEFT,
             space_after=0, line=1.14)
        y += 1.10

    # 项目切入点
    shape(s, MSO_SHAPE.ROUNDED_RECTANGLE, 0.42, 5.36, 12.5, 0.96, fill=WHITE,
          line=BLUE, line_w=1.0, radius=0.05)
    box = textbox(s, 0.62, 5.40, 1.55, 0.88, anchor=MSO_ANCHOR.MIDDLE)
    para(box.text_frame, '项目切入点', size=15, color=BLUE, bold=True, first=True,
         space_after=0, line=1.05)
    box = textbox(s, 2.24, 5.40, 10.5, 0.88, anchor=MSO_ANCHOR.MIDDLE)
    box.name = 'la'
    para(box.text_frame, '以甜菜碱类两性表面活性剂为主要候选组分，按相态筛选、界面性能与环境适应性'
                         '评价、岩心驱替三段推进，建立配方组成—相态—性能—驱替效果之间的关联。',
         size=15, color=TXT, first=True, align=PP_ALIGN.LEFT, space_after=0, line=1.15)

    conclusion(s, '研究结论', '不足、问题与验证一一对应，没有实验支撑的提问不写进申报材料')
    return s


# --------------------------------------------------------------- 4.1 五项创新
def p41(prs):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    header(s, '4.1', '项目创新点')
    judgment(s, '五项创新对应申请书研究内容 2.1.1—2.1.5，第五项为探索性内容')

    items = [
        ('01', '主剂选择与复配设计',
         '以磺基、羧基甜菜碱为主，考察内盐型偶极头基与疏水结构对界面吸附、增溶和钙镁'
         '耐受的影响；与石油磺酸盐、内烯烃磺酸盐、阴非复配三类对照同温同盐比较。'),
        ('02', '配方评价方法',
         '相态、界面、增溶、吸附与岩心驱替联合约束：鱼状相图定 Winsor Ⅲ 中相区间，'
         '界面张力以 10⁻³ mN/m 量级为优选参考，增溶比与增幅口径统一。'),
        ('03', '应用路线（原位成相）',
         '考察低浓度注入液与地层油水接触后原位成相的必要条件：油水比、盐度、稀释程度'
         '与接触时间；不把试管成相直接等同于储层有效成相。'),
        ('04', '功能评价（降压增注协同）',
         '同一流程同步记录采收率、含水、压差与注入体积，区分界面阻力降低、乳化流度调整'
         '与不利孔喉堵塞，避免只凭压力变化下结论。'),
    ]
    x0, y0, cw, ch, gx, gy = 0.42, 1.52, 6.16, 1.62, 0.18, 0.16
    for i, (num, head, body) in enumerate(items):
        col, row = i % 2, i // 2
        x = x0 + col * (cw + gx)
        y = y0 + row * (ch + gy)
        shape(s, MSO_SHAPE.ROUNDED_RECTANGLE, x, y, cw, ch, fill=WHITE, line=BORDER, radius=0.04)
        shape(s, MSO_SHAPE.ROUNDED_RECTANGLE, x, y, 0.86, ch, fill=BLUE, line=None, radius=0.06)
        box = textbox(s, x, y, 0.86, ch, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE,
                      margins=(0.02, 0, 0.02, 0))
        para(box.text_frame, num, size=22, color=WHITE, bold=True, first=True,
             align=PP_ALIGN.CENTER, space_after=0)
        box = textbox(s, x + 1.0, y + 0.10, cw - 1.16, 0.36)
        box.name = 'la'
        para(box.text_frame, head, size=16.5, color=BLUE, bold=True, first=True,
             align=PP_ALIGN.LEFT, space_after=0)
        box = textbox(s, x + 1.0, y + 0.48, cw - 1.16, ch - 0.56, anchor=MSO_ANCHOR.MIDDLE)
        box.name = 'la'
        para(box.text_frame, body, size=14, color=TXT, first=True, align=PP_ALIGN.LEFT,
             space_after=0, line=1.18)

    # 第五项：整行一条（探索性）
    y = y0 + 2 * (ch + gy) + 0.06
    shape(s, MSO_SHAPE.ROUNDED_RECTANGLE, 0.42, y, 12.5, 1.16, fill=WHITE, line=BORDER,
          radius=0.04)
    shape(s, MSO_SHAPE.ROUNDED_RECTANGLE, 0.42, y, 0.86, 1.16, fill=BLUE_L, line=None,
          radius=0.06)
    box = textbox(s, 0.42, y, 0.86, 1.16, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE,
                  margins=(0.02, 0, 0.02, 0))
    para(box.text_frame, '05', size=22, color=WHITE, bold=True, first=True,
         align=PP_ALIGN.CENTER, space_after=0)
    box = textbox(s, 1.42, y + 0.10, 11.3, 0.34)
    box.name = 'la'
    para(box.text_frame, '配方设计手段（探索性）', size=16.5, color=BLUE, bold=True,
         first=True, align=PP_ALIGN.LEFT, space_after=0)
    box = textbox(s, 1.42, y + 0.46, 11.2, 0.64, anchor=MSO_ANCHOR.MIDDLE)
    box.name = 'la'
    para(box.text_frame, '建立含盐度、醇浓度、表活浓度、复配比例与温度的数据表，用因素分析、'
                         '回归或响应面探索组成与成相边界的关系；数据量与重复性满足时预测配方区间，'
                         '并以新增实验检验。该项为探索性研究，不承诺通用模型。',
         size=14, color=TXT, first=True, align=PP_ALIGN.LEFT, space_after=0, line=1.15)

    conclusion(s, '研究结论', '四项为实验可验证的创新，第五项为探索性方法积累，均对应具体研究内容')
    return s


# --------------------------------------------------------------- 5.1 路线 + 进度
def p51(prs):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    header(s, '5.1', '实施方案 · 技术路线与进度安排')
    judgment(s, '五段推进、四阶段 12 个月：配方筛选与性能评价占主体，驱替与机理在后段交叠')

    stages = [('文献与条件', '资料梳理 · 样品与仪器落实'),
              ('配方与相态', '浓度·配比·盐度扫描 · 相态图'),
              ('界面与适应性', '界面张力·增溶 · 温度钙镁稀释吸附'),
              ('岩心驱替', '水驱对照 · 化学段塞 · 后续水驱'),
              ('数据整合', '指标汇总 · 综合优选 · 应用建议')]
    x, w, gap = 0.42, 2.32, 0.20
    for i, (name, sub) in enumerate(stages):
        cx = x + i * (w + gap)
        head = shape(s, MSO_SHAPE.ROUND_2_SAME_RECTANGLE, cx, 1.58, w, 0.52, fill=BLUE,
                     line=None, radius=0.12)
        gradient(head, BLUE_D, BLUE_L, angle=0)
        box = textbox(s, cx, 1.58, w, 0.52, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE,
                      margins=(0.04, 0, 0.04, 0))
        para(box.text_frame, name, size=14.5, color=WHITE, bold=True, first=True,
             align=PP_ALIGN.CENTER, space_after=0)
        shape(s, MSO_SHAPE.ROUNDED_RECTANGLE, cx, 2.14, w, 0.66, fill=BG_SOFT, line=BORDER,
              radius=0.10)
        box = textbox(s, cx + 0.06, 2.14, w - 0.12, 0.66, align=PP_ALIGN.CENTER,
                      anchor=MSO_ANCHOR.MIDDLE, margins=(0.03, 0, 0.03, 0))
        para(box.text_frame, sub, size=11, color=TXT, first=True, align=PP_ALIGN.CENTER,
             space_after=0, line=1.05)
        if i < len(stages) - 1:
            arrow_right(s, cx + w + 0.02, 1.80, 0.16, 0.16, fill=BLUE_L)

    # 进度：12 个月
    axis_y = 3.34
    left_w = 2.78
    track_x = 0.42 + left_w + 0.12
    track_w = 12.9 - track_x
    mw = track_w / 12
    box = textbox(s, 0.42, axis_y - 0.30, left_w, 0.26)
    box.name = 'la'
    para(box.text_frame, '阶段与时间（月）', size=11.5, color=GRAY, first=True,
         align=PP_ALIGN.LEFT, space_after=0)
    for m in range(12):
        box = textbox(s, track_x + m * mw, axis_y - 0.30, mw, 0.26, align=PP_ALIGN.CENTER)
        para(box.text_frame, str(m + 1), size=10.5, color=GRAY, first=True,
             align=PP_ALIGN.CENTER, space_after=0)
        sline(s, track_x + m * mw, axis_y - 0.02, track_x + m * mw, axis_y + 2.72,
              BORDER, 0.5)

    rows = [('阶段一', '资料调研与方案设计', 1, 2, '实施方案与记录表'),
            ('阶段二', '配方筛选与性能评价', 3, 6, '候选配方 2～3 组、相态图与性能数据'),
            ('阶段三', '驱替验证与机理分析', 7, 10, '优选体系与驱替结果'),
            ('阶段四', '总结与成果整理', 11, 12, '报告、数据表与论文初稿')]
    y = axis_y + 0.12
    for i, (tag, name, m0, m1, out) in enumerate(rows):
        shape(s, MSO_SHAPE.ROUNDED_RECTANGLE, 0.42, y, left_w, 0.52,
              fill=BG_SOFT if i % 2 == 0 else WHITE, line=BORDER, radius=0.10)
        box = textbox(s, 0.52, y, left_w - 0.2, 0.52, anchor=MSO_ANCHOR.MIDDLE)
        box.name = 'la'
        para(box.text_frame, [(tag + '　', BLUE, True), (name, TXT, False)], size=12.5,
             first=True, align=PP_ALIGN.LEFT, space_after=0, line=1.0)
        bar = shape(s, MSO_SHAPE.ROUNDED_RECTANGLE, track_x + (m0 - 1) * mw + 0.03,
                    y + 0.10, (m1 - m0 + 1) * mw - 0.06, 0.32,
                    fill=BLUE if i % 2 == 0 else BLUE_L, line=None, radius=0.18)
        box = textbox(s, track_x, y + 0.56, track_w, 0.24)
        box.name = 'la'
        para(box.text_frame, '输出：' + out, size=10.5, color=GRAY, first=True,
             align=PP_ALIGN.LEFT, space_after=0)
        y += 0.72

    conclusion(s, '研究结论', '阶段间以输出与阶段评审衔接；起始月份按立项通知确定，各阶段预留 2 周缓冲')
    return s


# --------------------------------------------------------------- 5.2 配方 + 界面
def p52(prs):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    header(s, '5.2', '实施方案 · 配方筛选与界面性能评价')
    judgment(s, '先定中相窗口，再验界面与增溶能否在温度、钙镁、稀释与吸附后保持')

    steps = [('①', '配方与盐度扫描',
              ['磺基／羧基甜菜碱为主，三类体系对照',
               '浓度、复配比例 1:3／1:1／3:1、盐度梯度',
               '记录相数、相体积、成相时间，绘鱼状相图',
               '定中相区间与最低有效浓度']),
             ('②', '界面与增溶',
              ['旋转滴法测动态与平衡界面张力',
               '以 10⁻³ mN/m 量级或更低为优选参考',
               '由相体积与有效含量算油、水增溶参数',
               '比较最佳盐度附近双向增溶是否平衡']),
             ('③', '环境适应性',
              ['温度梯度与 Ca²⁺、Mg²⁺ 梯度复测',
               '地层水按 10%／30%／50% 混入、注入水稀释 2／5／10 倍',
               '老化与吸附前后复测相态与界面张力',
               '记录窗口偏移方向，供前置段塞调盐'])]
    for i, (num, head, items) in enumerate(steps):
        x = 0.42 + i * 4.20
        shape(s, MSO_SHAPE.ROUNDED_RECTANGLE, x, 1.56, 4.05, 1.52, fill=WHITE, line=BORDER,
              radius=0.05)
        box = textbox(s, x + 0.14, 1.60, 0.36, 0.34)
        para(box.text_frame, num, size=16, color=BLUE, bold=True, first=True, space_after=0)
        box = textbox(s, x + 0.52, 1.62, 3.3, 0.34)
        box.name = 'la'
        para(box.text_frame, head, size=15.5, color=BLUE, bold=True, first=True,
             align=PP_ALIGN.LEFT, space_after=0)
        box = textbox(s, x + 0.16, 1.98, 3.74, 1.04, anchor=MSO_ANCHOR.MIDDLE)
        bullets(box.text_frame, items, size=11.5, marker='·', space_after=1, line=1.06)

    imgs = [('phase-vol.png', '相体积分数随 NaCl 质量分数变化（资料示例）'),
            ('ift-salt.png', '界面张力随盐度变化，Winsor Ⅲ 区取最低值（资料图）'),
            ('tubes.png', '不同盐度下的相态观察（资料图）')]
    for i, (fn, cap) in enumerate(imgs):
        x = 0.42 + i * 4.20
        picture(s, 'images/' + fn, x, 3.22, 4.05, 2.22, mode='fit', border=BORDER)
        caption(s, x, 5.48, 4.05, cap, size=10)

    conclusion(s, '研究结论', '相态窗口是入口，界面与适应性是筛选条件；两项不成立即回到复配调整')
    return s


# --------------------------------------------------------------- 5.3 岩心驱替
def p53(prs):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    header(s, '5.3', '岩心驱替与增油评价')
    judgment(s, '静态筛出的配方必须回到岩心：增油与压差要在同一流程里同时记录')

    # 装置与流向（整幅）
    import os as _os
    rig = 'images/gpt-rigflow.jpg'
    if _os.path.exists(rig):
        picture(s, rig, 0.42, 1.50, 12.50, 2.56, mode='fit')
        caption(s, 0.42, 4.14, 12.50,
                '岩心驱替装置与流向（示意）：注入泵—中间容器—岩心夹持器（恒温）—压力与产出液计量',
                size=9.5)
    else:
        picture(s, 'images/coreflood-rig.png', 0.42, 1.50, 6.20, 2.56, mode='fit', border=BORDER)

    # 流程一行
    box = textbox(s, 0.42, 4.44, 12.50, 0.36, align=PP_ALIGN.CENTER)
    para(box.text_frame, [('驱替流程：', BLUE, True),
                          ('饱和水 → 饱和油 → 水驱至含水 98% → 化学段塞 0.3–0.5 PV '
                           '→ 后续水驱（同一岩心连续进行）', TXT, False)],
         size=13, first=True, align=PP_ALIGN.CENTER, space_after=0)

    # 四项同步记录
    rec = [('采收率', '转驱后增幅，用百分点'),
           ('含水率', '与采收率同步给出'),
           ('驱替压差', '区分流度调整与堵塞'),
           ('注入体积', '折算孔隙体积倍数')]
    for i, (k, v) in enumerate(rec):
        x = 0.42 + i * 3.16
        shape(s, MSO_SHAPE.ROUNDED_RECTANGLE, x, 4.90, 3.00, 0.74, fill=BG_SOFT,
              line=BORDER, radius=0.10)
        box = textbox(s, x + 0.12, 4.94, 2.76, 0.28)
        box.name = 'la'
        para(box.text_frame, k, size=13, color=BLUE, bold=True, first=True,
             align=PP_ALIGN.LEFT, space_after=0)
        box = textbox(s, x + 0.12, 5.20, 2.76, 0.40, anchor=MSO_ANCHOR.MIDDLE)
        box.name = 'la'
        para(box.text_frame, v, size=11.5, color=TXT, first=True, align=PP_ALIGN.LEFT,
             space_after=0, line=1.0)

    # 判读口径
    box = textbox(s, 0.42, 5.76, 12.50, 0.60, anchor=MSO_ANCHOR.MIDDLE)
    box.name = 'la'
    para(box.text_frame, [('判读口径：', RED, True),
                          ('同矿化度配液的水驱对照、同流量同温度；至少 2 块平行样，增幅口径一致；'
                           '压差升高不等于调剖见效，需结合产出液与孔喉堵塞判断。', TXT, False)],
         size=12, first=True, align=PP_ALIGN.LEFT, space_after=0, line=1.2)

    conclusion(s, '研究结论', '配方能否进入下一阶段，取决于增油幅度与注入压差能否同时接受')
    return s


# --------------------------------------------------------------- 2.1 研究背景（自绘）
def p21(prs):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    header(s, '2.1', '研究背景')
    judgment(s, '高含水后期与低渗储量已成为提高采收率的主战场，亟需兼顾洗油效率与注入性的体系')

    shape(s, MSO_SHAPE.ROUNDED_RECTANGLE, 0.42, 1.54, 4.62, 4.72, fill=WHITE, line=BORDER,
          radius=0.04)
    picture(s, 'images/wellsite.png', 0.58, 1.68, 4.30, 2.30, mode='cover', border=BORDER)
    caption(s, 0.58, 4.02, 4.30, '国内东部老油田井场（资料图）', size=10)
    box = textbox(s, 0.62, 4.36, 4.22, 0.30)
    box.name = 'la'
    para(box.text_frame, '开发形势', size=14, color=BLUE, bold=True, first=True,
         align=PP_ALIGN.LEFT, space_after=0)
    box = textbox(s, 0.62, 4.64, 4.22, 1.50, anchor=MSO_ANCHOR.MIDDLE)
    bullets(box.text_frame, ['主力油田进入高含水后期，水驱可采储量动用程度高',
                             '低渗—中低渗储量占比上升，孔喉细小、启动压力高',
                             '提采空间同时受洗油效率与注入性限制'],
            size=12, marker='·', space_after=5, line=1.1)

    picture(s, 'images/gpt-remaining-oil.jpg', 5.34, 1.54, 3.70, 2.16, mode='fit', border=BORDER)
    caption(s, 5.34, 3.76, 3.70, '剩余油赋存形态：油膜、孤立油滴与孔喉残余（示意）', size=10)
    picture(s, 'images/prod-curve.png', 9.26, 1.54, 3.62, 2.16, mode='fit', border=BORDER)
    caption(s, 9.26, 3.76, 3.62, '典型井日产油、含水率与日产液变化（资料图）', size=10)

    shape(s, MSO_SHAPE.ROUNDED_RECTANGLE, 5.34, 4.28, 7.54, 1.98, fill=BG_SOFT, line=BORDER,
          radius=0.04)
    box = textbox(s, 5.52, 4.36, 7.2, 0.32)
    box.name = 'la'
    para(box.text_frame, '两处约束指向同一个问题', size=14, color=BLUE, bold=True, first=True,
         align=PP_ALIGN.LEFT, space_after=0)
    box = textbox(s, 5.52, 4.72, 7.2, 1.44, anchor=MSO_ANCHOR.MIDDLE)
    box.name = 'la'
    para(box.text_frame, '注入水沿高渗透条带突进，剩余油以油膜、孤立油滴与孔喉残余形式滞留，'
                         '常规水驱难以启动；低渗储层孔喉细小、毛管阻力大，提高注入压力又接近'
                         '破裂压力。扩大波及体积与提高洗油效率必须同时考虑，常规单一体系难以兼顾。',
         size=13.5, color=TXT, first=True, align=PP_ALIGN.LEFT, space_after=0, line=1.2)

    conclusion(s, '研究结论', '背景问题的落点是“兼顾洗油与注入”，这也决定了后面配方与评价的设计')
    return s
