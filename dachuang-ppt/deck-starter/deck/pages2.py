# -*- coding: utf-8 -*-
"""自绘页（后半段）：指标表与预期成果。同样是骨架示例，只换内容。"""
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR

from .lib import (FS, BLUE, BLUE_L, RED, BORDER, BG_SOFT, TXT, GRAY, WHITE,
                  textbox, para, shape, card, bullets, conclude, caption,
                  table, header, judgment, rich)


# ------------------------------------------------------------------ 4.1 关键指标
def s11(prs):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    header(s, '4.1', '关键指标与验收标准')
    judgment(s, '三项硬指标全部可量化：误差、耗时、成本')

    rows = [
        ['指标', '现有水平', '本项目目标', '验收方式'],
        ['预测误差 MAPE', '8.5%', '≤ 6.0%', '留出集盲测'],
        ['单次标定耗时', '6 h', '≤ 2 h', '现场计时记录'],
        ['样本需求量', '12 个月', '6 个月', '同口径回溯'],
        ['部署成本', '外购服务', '自建开源栈', '年度结算单'],
    ]
    table(s, 0.55, 1.78, 12.2, rows, [3.6, 2.8, 2.9, 2.9],
          row_h=0.56, head_h=0.50, size=16, head_size=16,
          red_cells={(1, 2), (2, 2), (3, 2)})

    box = textbox(s, 0.55, 4.90, 12.2, 1.30)
    bullets(box.text_frame, [rich('误差指标以【留出集】为准，不用训练集自评'),
                             rich('耗时与成本在【同一现场】对比，避免设备差异带来的偏差')],
            space_after=6)

    conclude(s, '指标写进任务书，中期检查按这张表逐项核对')


# ------------------------------------------------------------------ 5.1 预期成果
def s13(prs):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    header(s, '5.1', '预期成果与进度安排')
    judgment(s, '成果分【方法】与【平台】两类，都能独立交付')

    card(s, 0.42, 1.72, 6.05, 2.52, ribbon='方法与认识')
    box = textbox(s, 0.66, 2.36, 5.58, 1.80)
    bullets(box.text_frame, ['形成小样本迁移的标定流程',
                             '给出跨场景适配的参数建议',
                             '发表论文 1 篇、申请软著 1 项'], hang=0.24)

    card(s, 6.86, 1.72, 6.05, 2.52, ribbon='平台与工具')
    box = textbox(s, 7.10, 2.36, 5.58, 1.80)
    bullets(box.text_frame, [rich('一套可复用的【训练与评估脚本】'),
                             '一个现场可用的轻量推理服务',
                             '一份部署与运维手册'], hang=0.24)

    shape(s, MSO_SHAPE.ROUNDED_RECTANGLE, 0.42, 4.40, 12.48, 1.98,
          fill=BG_SOFT, line=BORDER, line_w=0.75, radius=0.05)
    box = textbox(s, 0.66, 4.48, 12.0, 1.84)
    para(box.text_frame, '进度安排', size=FS['sub'], color=BLUE, bold=True, first=True, space_after=6)
    bullets(box.text_frame, [
        [('第 1–4 月　', GRAY, True), ('完成数据治理与基线模型（里程碑一）', TXT, False)],
        [('第 5–8 月　', GRAY, True), ('完成迁移策略与内测（里程碑二）', TXT, False)],
        [('第 9–12 月　', GRAY, True), ('现场验证、结题与成果整理（里程碑三）', TXT, False)],
    ], marker='', first=False, space_after=4, line=1.3)

    conclude(s, '两个里程碑都可【单独验收】，不把风险压到最后一个月')
