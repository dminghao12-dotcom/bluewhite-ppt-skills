# -*- coding: utf-8 -*-
"""自绘页（沿用 lib 的版式组件），不依赖模板里的形状。

这两个函数是**骨架示例**：把文字换成你项目的就行，版式不用重排。
"""
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR

from .lib import (BLUE, BLUE_D, BLUE_L, BLUE_PALE, RED, BORDER, BG_SOFT, TXT, GRAY,
                  WHITE, FS, textbox, para, shape, card, bullets, conclude, caption,
                  chip, arrow_right, header, judgment, rich)


# ------------------------------------------------------------------ 2.1 背景与问题
def s03(prs):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    header(s, '2.1', '研究背景与问题')
    judgment(s, '【工程需求明确】，但现有方法在【精度】与【成本】上难以兼顾')

    xs = [0.42, 4.62, 8.82]
    titles = ['行业背景', '现有方法', '关键缺口']
    items = [
        ['需求侧增速快，年增约 8%', '政策与考核压力同步加码', '标定成本占运维投入较高'],
        ['机理模型精度高但标定慢', '数据驱动模型依赖大样本', '两类方法难以兼顾成本'],
        ['小样本工况下误差偏高', '跨场景迁移能力不足', '缺少统一的误差评价口径'],
    ]
    for x, t, its in zip(xs, titles, items):
        card(s, x, 1.72, 4.10, 2.75, ribbon=t)
        box = textbox(s, x + 0.22, 2.40, 3.66, 1.95)
        bullets(box.text_frame, its, hang=0.24)

    # 下半区：一条“缺口 → 目标”的因果条
    shape(s, MSO_SHAPE.ROUNDED_RECTANGLE, 0.42, 4.78, 12.48, 1.20,
          fill=BG_SOFT, line=BORDER, line_w=0.75, radius=0.06)
    box = textbox(s, 0.62, 4.84, 12.08, 1.08, anchor=MSO_ANCHOR.MIDDLE)
    para(box.text_frame, [('缺口：', RED, True),
                          ('小样本 + 跨场景下模型不易迁移', TXT, False),
                          ('   →   目标：', BLUE, True),
                          ('在样本量减少 50% 的条件下，误差不高于基线', TXT, False)],
         first=True, space_after=0, line=1.2)

    conclude(s, '本项目从【迁移学习】切入，把标定成本压到原来的三成以内')


# ------------------------------------------------------------------ 3.1 技术路线
def s05(prs):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    header(s, '3.1', '技术路线')
    judgment(s, '四步闭环：数据治理 → 模型构建 → 迁移适配 → 现场验证')

    steps = [('① 数据治理', '清洗/对齐/缺失补全'),
             ('② 模型构建', '机理约束 + 数据驱动'),
             ('③ 迁移适配', '小样本微调策略'),
             ('④ 现场验证', '实测对比 + 误差分析')]
    x = 0.55
    for i, (head, sub) in enumerate(steps):
        card(s, x, 1.80, 2.72, 1.50)
        box = textbox(s, x + 0.14, 1.88, 2.44, 1.34, align=PP_ALIGN.CENTER,
                      anchor=MSO_ANCHOR.MIDDLE)
        para(box.text_frame, head, size=FS['sub'], color=BLUE, bold=True, first=True,
             align=PP_ALIGN.CENTER, space_after=3)
        para(box.text_frame, sub, size=FS['label'], color=GRAY, align=PP_ALIGN.CENTER,
             space_after=0, line=1.1)
        if i < 3:
            arrow_right(s, x + 2.80, 2.42, 0.42, 0.30, fill=BLUE_L)
        x += 3.22

    # 每步的关键产出
    rows = [('交付物', ['标准化数据集', '基线模型 + 训练脚本', '迁移权重与配置', '验证报告'])]
    box = textbox(s, 0.55, 3.50, 12.3, 1.40)
    bullets(box.text_frame, [rich('数据治理阶段交付【标准化数据集】，指标口径与甲方对齐'),
                             '模型构建阶段交付基线模型，误差较现有方法下降 15% 以上',
                             '迁移适配阶段交付权重与配置，支持新场景快速上线'],
            space_after=6)

    shape(s, MSO_SHAPE.ROUNDED_RECTANGLE, 0.42, 5.06, 12.48, 1.30,
          fill=WHITE, line=BORDER, line_w=0.75, radius=0.06)
    box = textbox(s, 0.62, 5.12, 12.08, 1.18)
    para(box.text_frame, '关键节点', size=FS['sub'], color=BLUE, bold=True, first=True, space_after=4)
    para(box.text_frame, [('第 1–4 月 ', GRAY, True), ('数据与方法确定　　', TXT, False),
                          ('第 5–8 月 ', GRAY, True), ('模型开发与内测　　', TXT, False),
                          ('第 9–12 月 ', GRAY, True), ('现场验证与结题', TXT, False)],
         size=16, space_after=0, line=1.2)   # 一行三段放不下 18，按规则留 16

    conclude(s, '路线按【月度节点】推进，任一步延误都能在双周例会暴露')
