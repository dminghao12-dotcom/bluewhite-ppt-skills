# -*- coding: utf-8 -*-
"""骨架入口：在 tpl2/t.pptx（= assets/default-template.pptx 的副本）上生成一份 7 页样张。

装好依赖后，在本目录的上一级执行：

    python -m deck.build2                 # 出 out/样张.pptx
    python -m deck.build2 out/我的稿.pptx  # 指定输出

改内容只需要动三处：本文件的 COVER / TOC / 页序，pages1.py 与 pages2.py 的正文。
"""
import copy
import os
import sys

from pptx import Presentation
from pptx.enum.text import PP_ALIGN
from pptx.oxml.ns import qn

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
sys.path.insert(0, ROOT)
os.chdir(ROOT)

from deck import pages1, pages2, tplcore  # noqa: E402
from deck.lib import FS, WHITE, textbox, para, sline  # noqa: E402
from pptx.enum.text import MSO_ANCHOR  # noqa: E402

TPL = 'tpl2/t.pptx'

# 封面：形状名 -> 新文字。名字取自模板，先用 tplcore.by_name 打印一遍再改
COVER = {
    '文本框 10': '大学生创新创业训练计划',          # 顶部计划名（72pt）
    '文本框 81': '小样本迁移的负荷预测标定方法',      # 项目名（28pt）
}
COVER_PILL = ('汇报人', '汇报人：张三　　指导教师：李四')   # 封面胶囊里的那行小字
UNIT_LINE = '示例大学　示例学院　·　2026 年 10 月'    # 封面与结束页底部的单位时间（整行重画，不继承模板字距）

# 目录：模板一共 8 条，条数对齐就不用改模板
TOC = [('一', '小组分工与指导老师'), ('二', '研究背景与问题'),
       ('三', '技术路线'), ('四', '关键指标与验收标准'),
       ('五', '预期成果与进度安排'), ('六', '项目资源'), ('七', '经费预算'), ('八', '风险与对策')]

END_THANKS = '感谢各位评审专家'
END_SUB = '敬请批评指正'

# 保留模板里的哪几页（0 基）：0 封面、1 目录、3 结束；第 2 页（成果）用自绘页替代
KEEP = [0, 1, 3]


def replace_containing(slide, needle, new_text):
    """把文本框里含 needle 的那段文字整行换掉（用于封面胶囊这类组合里的文字）。"""
    for shp in tplcore.iter_shapes(slide.shapes):
        if not shp.has_text_frame:
            continue
        if needle in shp.text_frame.text:
            tplcore.set_text(shp, new_text)
            return True
    print(f'  ! 封面里没找到含「{needle}」的文字，跳过')
    return False


def redraw_unit_line(slide):
    """单行短文案（单位时间）模板常带字符间距 spc，set_text 会继承成字间距拉开的怪行：删掉重画。"""
    for shp in list(slide.shapes):
        if shp.name == 'TextBox 113':
            shp._element.getparent().remove(shp._element)
    box = textbox(slide, 0.9, 6.42, 11.5, 0.42, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    para(box.text_frame, UNIT_LINE, size=FS['body'], color='003195', first=True,
         align=PP_ALIGN.CENTER, space_after=0, line=1.0)


def redraw_end(slide):
    """结束页一律整块重绘：删模板原文字框与汇报人胶囊，感谢语 44 / 副句 30 / 白色细线，落在深蓝区（0–5.0 in）内。"""
    for shp in list(slide.shapes):
        if shp.name == '文本框 6' or (shp.shape_type == 6 and '汇报人' in ''.join(
                x.text_frame.text for x in shp.shapes if x.has_text_frame)):
            shp._element.getparent().remove(shp._element)
    box = textbox(slide, 0.9, 1.35, 11.5, 1.0, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    para(box.text_frame, END_THANKS, size=44, color=WHITE, bold=True, first=True,
         align=PP_ALIGN.CENTER, space_after=0, line=1.0)
    sline(slide, 4.67, 2.62, 8.67, 2.62, color=WHITE, width=1.0)
    box = textbox(slide, 0.9, 2.85, 11.5, 0.75, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    para(box.text_frame, END_SUB, size=30, color=WHITE, first=True,
         align=PP_ALIGN.CENTER, space_after=0, line=1.0)
    redraw_unit_line(slide)


def center_all(slide):
    """set_text 只改首段对齐；封面多段文字逐段居中，避免一行居中、一行左偏。"""
    for shp in tplcore.iter_shapes(slide.shapes):
        if shp.has_text_frame:
            for p in shp.text_frame.paragraphs:
                p.alignment = PP_ALIGN.CENTER


def ensure_tpl():
    """模板缺失时自动从技能包的 assets/ 复制一份，避免包里重复带 8 MB 副本。"""
    if os.path.exists(TPL):
        return TPL
    import shutil
    cands = ['assets/default-template.pptx', '../assets/default-template.pptx']
    if os.environ.get('DACHUANG_SKILL_DIR'):
        cands.insert(0, os.path.join(os.environ['DACHUANG_SKILL_DIR'], 'assets', 'default-template.pptx'))
    home = os.path.expanduser('~')
    for base in ('.claude/skills', '.codex/skills', '.agents/skills', '.hermes/skills/productivity', '.hermes/skills'):
        cands.append(os.path.join(home, base, 'dachuang-ppt', 'assets', 'default-template.pptx'))
    for cand in cands:
        if os.path.exists(cand):
            os.makedirs(os.path.dirname(TPL), exist_ok=True)
            shutil.copyfile(cand, TPL)
            print(f'  模板缺失，已从 {cand} 复制到 {TPL}')
            return TPL
    raise SystemExit(
        f'缺模板：{TPL}。没有随包模板时，把你的模板存成 {TPL}；'
        f'或从技能包拷：cp <skill>/assets/default-template.pptx {TPL}；'
        f'也可设环境变量 DACHUANG_SKILL_DIR=<技能目录> 让脚本自动找')


def build(out_path='out/样张.pptx'):
    ensure_tpl()
    prs = Presentation(TPL)
    tplcore.keep_slides(prs, KEEP)

    cover, toc, end = list(prs.slides)
    tplcore.replace_all(cover, COVER)
    replace_containing(cover, *COVER_PILL)
    center_all(cover)
    redraw_unit_line(cover)
    tplcore.fill_toc(toc, TOC)
    redraw_end(end)

    # 自绘正文页：加在末尾，稍后重排
    pages1.s03(prs)
    pages1.s05(prs)
    pages2.s11(prs)
    pages2.s13(prs)

    # 当前顺序：0 封面 1 目录 2 结束 3..6 自绘 → 期望：封面、目录、4 页正文、结束
    tplcore.reorder(prs, [0, 1, 3, 4, 5, 6, 2])
    tplcore.renumber_partnames(prs)

    # 备注（讲稿）：与页数一一对应，删改页时同步改这里
    SCRIPT = [
        ('封面', '各位老师好，我汇报的题目是……，下面从背景、路线、指标、成果四部分说明。'),
        ('目录', '汇报分四部分，重点是技术路线和指标。'),
        ('研究背景与问题', '先说为什么做：需求在涨，现有方法在精度和成本上难以兼顾。'),
        ('技术路线', '四步闭环：数据治理、模型构建、迁移适配、现场验证。'),
        ('关键指标', '三项硬指标都可量化，验收方式写在表里。'),
        ('预期成果与进度', '成果分方法类和平台类，两个里程碑都能单独验收。'),
        ('结束', '以上是全部内容，请各位老师批评指正。'),
    ]
    slides = list(prs.slides)
    assert len(SCRIPT) == len(slides), f'讲稿 {len(SCRIPT)} 条，幻灯片 {len(slides)} 页'
    for slide, (head, body) in zip(slides, SCRIPT):
        slide.notes_slide.notes_text_frame.text = f'{head}\n{body}'

    os.makedirs(os.path.dirname(out_path) or '.', exist_ok=True)
    prs.save(out_path)
    tplcore.report(prs, out_path)
    return out_path


if __name__ == '__main__':
    build(*(sys.argv[1:] or []))
