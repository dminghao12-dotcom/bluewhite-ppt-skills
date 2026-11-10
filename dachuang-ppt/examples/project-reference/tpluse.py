"""把我们的内容灌进用户提供的蓝色模板（原版 20 页）自带的文字框与图片位。

原则：
- 直接复用模板页的版式、形状、图片位，只换文字与图片；
- 模板里属于“深海”主题的照片/图标一律换成我们资料里的油田图片；
- 字号沿用模板原值（形状自带），必要时只做轻微调整。
"""
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
import copy
import os

from pptx.enum.shapes import MSO_SHAPE

from deck.lib import shape, textbox, para, BORDER, GRAY, BLUE, WHITE, RED, BG_SOFT

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
IMG = os.path.join(ROOT, 'images') + os.sep
BRAND = os.path.join(ROOT, 'brand') + os.sep

WHITE = RGBColor(0xFF, 0xFF, 0xFF)
BLUE = RGBColor(0x00, 0x31, 0x95)
GRAY = RGBColor(0x5A, 0x6A, 0x85)


def shp(slide, idx):
    return slide.shapes[idx]


def set_text(shape, lines, size=None, color=None, bold=None, align=None, line_spacing=None):
    """保留第一段第一个 run 的字体属性，替换文本。lines 可以是 str 或 list。"""
    if isinstance(lines, str):
        lines = [lines]
    tf = shape.text_frame
    p0 = tf.paragraphs[0]
    r0 = p0.runs[0] if p0.runs else None
    ref = copy.deepcopy(r0._r) if r0 is not None else None
    # 清空
    for p in list(tf.paragraphs)[1:]:
        p._p.getparent().remove(p._p)
    for r in list(p0.runs):
        r._r.getparent().remove(r._r)

    for i, ln in enumerate(lines):
        p = p0 if i == 0 else tf.add_paragraph()
        if ref is not None:
            r = copy.deepcopy(ref)
            p._p.append(r)
            run = p.runs[0]
            run.text = ln
        else:
            run = p.add_run()
            run.text = ln
        if size is not None:
            run.font.size = Pt(size)
        if color is not None:
            run.font.color.rgb = color
        if bold is not None:
            run.font.bold = bold
        if align is not None:
            p.alignment = align
        if line_spacing is not None:
            p.line_spacing = line_spacing
    return shape


def set_pic(slide, shape, path):
    """把已有的图片形状换成另一张图，保留位置尺寸。"""
    image_part, rId = slide.part.get_or_add_image_part(path)
    blip = shape._element.blipFill.blip
    blip.rEmbed = rId
    return shape


def set_card_stat(shape, num, label):
    """把模板“21人 / 核心团队成员”这种两行统计卡改成我们的数字与标签。"""
    tf = shape.text_frame
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    p0 = tf.paragraphs[0]
    for p in list(tf.paragraphs)[1:]:
        p._p.getparent().remove(p._p)
    for r in list(p0.runs):
        r._r.getparent().remove(r._r)
    r0 = p0.add_run()
    r0.text = num
    r0.font.size = Pt(17)
    r0.font.bold = True
    r0.font.color.rgb = RGBColor(0xC0, 0x00, 0x00)
    p0.alignment = PP_ALIGN.CENTER
    p1 = tf.add_paragraph()
    r1 = p1.add_run()
    r1.text = label
    r1.font.size = Pt(11)
    r1.font.color.rgb = GRAY
    p1.alignment = PP_ALIGN.CENTER
    return shape


def fix_logo(slide, white=False):
    """模板内容页右上角是别人的 logo（10.95in, 0.17in 附近），换成长江大学校徽。"""
    target = BRAND + ('logo_white.png' if white else 'logo_blue.png')
    for sh in slide.shapes:
        if sh.shape_type == 13 and sh.left > 10.6 * 914400 and sh.top < 0.4 * 914400:
            try:
                set_pic(slide, sh, target)
            except Exception:
                pass


def drop(shape):
    shape._element.getparent().remove(shape._element)


def add_text(slide, x, y, w, h, text, size=18, color=BLUE, bold=False, align=PP_ALIGN.LEFT,
             anchor=MSO_ANCHOR.TOP, name=None):
    from pptx.util import Emu
    box = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = box.text_frame
    tf.word_wrap = True
    tf.margin_left = 0
    tf.margin_right = 0
    tf.margin_top = 0
    tf.margin_bottom = 0
    tf.vertical_anchor = anchor
    p = tf.paragraphs[0]
    r = p.add_run()
    r.text = text
    r.font.size = Pt(size)
    r.font.color.rgb = color
    r.font.bold = bold
    r.font.name = '微软雅黑'
    p.alignment = align
    return box


# ----------------------------------------------------------------- 1 封面
def fill_cover(slide):
    # 模板封面原件；按用户版：不要左上角那行小字副标题，只留大标题 + 项目名 + 单位时间
    big = shp(slide, 10)
    title = shp(slide, 6)
    logo = shp(slide, 9)
    sub = shp(slide, 8)
    pill = shp(slide, 3)
    grid = shp(slide, 2)

    set_text(big, '大学生创新创业训练计划', size=72)
    title.top = Inches(2.78)
    set_text(title, '中相微乳液驱油体系优化及机理研究', size=28)
    set_pic(slide, logo, BRAND + 'logo_white.png')
    for sub_shape in pill.shapes:
        if sub_shape.has_text_frame and sub_shape.text_frame.text.strip():
            set_text(sub_shape, '汇报人：焦妍　　指导教师：汪杰', size=13.5)
    for v in (sub, grid):
        try:
            drop(v)
        except Exception:
            pass
    add_text(slide, 0.9, 6.42, 11.5, 0.42,
             '长江大学　石油工程学院　·　2025 年 10 月', size=18, color=BLUE,
             align=PP_ALIGN.CENTER)


# ----------------------------------------------------------------- 2 目录
def fill_toc(slide):
    items = [('一', '小组分工及指导老师'), ('二', '研究背景及意义'),
             ('三', '国内外研究现状'), ('四', '创新点与拟解决关键问题'),
             ('五', '项目实施方案'), ('六', '项目资源介绍'),
             ('七', '预期成果'), ('八', '经费预算')]
    # 模板只有 5 行且只有第一行是亮的；八项改成单列竖排，序号与长条沿用模板样式
    circle = shp(slide, 6)
    number = shp(slide, 7)
    bar = shp(slide, 4)
    title = shp(slide, 5)
    proto = [copy.deepcopy(x._element) for x in (circle, number, bar, title)]
    for x in (circle, number, bar, title):
        drop(x)
    for idx in sorted([8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23],
                      reverse=True):
        try:
            drop(slide.shapes[idx])
        except IndexError:
            pass

    y0, pitch = 0.92, 0.66
    for i, (num, txt) in enumerate(items):
        y = y0 + i * pitch
        for e in proto:
            slide.shapes._spTree.append(copy.deepcopy(e))
        c, n, b, tt = list(slide.shapes)[-4:]
        for sp_, x, w, h, dy in ((c, 5.66, 0.58, 0.58, 0.0), (n, 5.66, 0.58, 0.58, 0.0),
                                 (b, 6.34, 6.86, 0.58, 0.0), (tt, 6.56, 6.5, 0.58, 0.0)):
            sp_.left = Inches(x)
            sp_.top = Inches(y + dy)
            sp_.width = Inches(w)
            sp_.height = Inches(h)
        c.name = 'toc-c'
        n.name = 'toc-n'
        b.name = 'toc-b'
        set_text(n, num, size=17, color=WHITE, bold=True, align=PP_ALIGN.CENTER)
        set_text(tt, txt, size=19, color=BLUE, bold=True, align=PP_ALIGN.LEFT)
        tt.name = 'toc-title'
        tt.left = Inches(6.56)

    # 底部两张波浪图与本页无关，删掉让八行更透
    for idx in (1, 0):
        try:
            drop(slide.shapes[idx])
        except IndexError:
            pass

    # 残留的空圆/空条（模板原五行留下的）
    for sh_ in list(slide.shapes):
        if (sh_.name or '').startswith('toc'):
            continue
        if getattr(sh_, 'has_text_frame', False) and sh_.text_frame.text.strip():
            continue
        if (sh_.shape_type == 1 and 5.2 < (sh_.left or 0) / 914400 < 6.6
                and (sh_.width or 0) < 0.9 * 914400
                and 0.7 < (sh_.top or 0) / 914400 < 5.6):
            drop(sh_)

    # 清掉模板原来五行的残留（文字、序号）
    old = {'项目背景与研究现状', '课题分解与研究内容', '任务分解与进度安排',
           '研发团队与工作基础', '预期成果与风险分析',
           '一', '二', '三', '四', '五'}
    for sh_ in list(slide.shapes):
        if (sh_.name or '') == 'toc-n':
            continue
        if sh_.has_text_frame and sh_.text_frame.text.strip() in old:
            drop(sh_)


# ----------------------------------------------------------------- 4 背景
def fill_bg(slide):
    fix_logo(slide)
    set_text(shp(slide, 1), '2.1', size=26)
    set_text(shp(slide, 2), '研究背景', size=26)
    set_text(shp(slide, 24), '能源安全与提采需求', size=13)
    # 左卡正文（GROUP 25 内两条）
    g25 = shp(slide, 25)
    lines = ['主力油田普遍进入高含水后期，水驱可采储量动用程度高，进一步提高采收率空间有限',
             '低渗—中低渗储量占比上升，孔喉细小、启动压力高，常规化学驱的注入性与经济性同时受限']
    k = 0
    for sub in g25.shapes:
        if sub.has_text_frame and sub.text_frame.text.strip():
            set_text(sub, lines[min(k, 1)], size=13.5, line_spacing=1.35,
                     align=PP_ALIGN.LEFT)
            k += 1
    # 右卡标题与正文
    set_text(shp(slide, 7), '高含水与低渗的双重约束', size=14)
    set_text(shp(slide, 8), '注入水沿高渗透条带突进，剩余油以油膜、孤立油滴与孔喉残余形式滞留；'
                            '低渗储层孔喉细小、毛管阻力大，常规体系难以兼顾波及体积与洗油效率。',
             size=12.5, line_spacing=1.4)
    # 通栏判断句
    g20 = shp(slide, 20)
    for sub in g20.shapes:
        if sub.has_text_frame and sub.text_frame.text.strip():
            set_text(sub, '高含水后期与低渗储量已成为提高采收率的主战场，'
                          '亟需兼顾洗油效率与注入性的驱油体系', size=20)
    # 三个小图标标题
    set_text(shp(slide, 17), '波及体积受限', size=12)
    set_text(shp(slide, 14), '洗油效率不足', size=12)
    # 三张图：剩余油形态、生产曲线、井场
    set_pic(slide, shp(slide, 9), IMG + 'gpt-remaining-oil.jpg')
    set_pic(slide, shp(slide, 10), IMG + 'prod-curve.png')
    set_pic(slide, shp(slide, 12), IMG + 'tubes.png')
    # 中上小图（原来是示意图）换相态图
    set_pic(slide, shp(slide, 19), IMG + 'ift-salt.png')
    # 左侧卡片里的“国务院新闻截图”换成井场照片
    for sub in g25.shapes:
        if sub.shape_type == 13:
            set_pic(slide, sub, IMG + 'wellsite.png')
    # 底部三个标签条
    g11 = shp(slide, 11)
    labels = ['高含水期', '低渗油藏', '三次采油']
    k = 0
    for sub in g11.shapes:
        if sub.has_text_frame and sub.text_frame.text.strip():
            set_text(sub, labels[min(k, 2)], size=12.5, align=PP_ALIGN.CENTER)
            sub.text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE
            k += 1


# ----------------------------------------------------------------- 11 指标
def fill_metric(slide):
    fix_logo(slide)
    set_text(shp(slide, 3), '5.4', size=26)
    set_text(shp(slide, 4), '岩心驱替与增油评价', size=26)
    set_text(shp(slide, 7), '同一流程同步记录采收率、含水、压差与注入体积，'
                            '区分流度调整与孔喉堵塞', size=16)
    set_text(shp(slide, 9), '采收率增幅以百分点表示，与含水下降、压差变化一并判读', size=12.5)
    set_text(shp(slide, 10), '对照水驱与化学驱两组岩心，重复性不达标时追加平行样', size=12.5)
    set_text(shp(slide, 23), '室内岩心驱替', size=11)
    for idx, txt in ((25, '水驱至含水 98%\n后转驱'), (27, '注入 0.3–0.5 PV\n段塞'),
                     (29, '压差与产出\n同步记录')):
        set_text(shp(slide, idx), txt, size=10.5)
    set_text(shp(slide, 33), '对照与重复性', size=11)
    for idx, txt in ((35, '平行样 2 块\n重复性核查'), (37, '同矿化度\n配液对照'),
                     (39, '产出液\n含油分析')):
        set_text(shp(slide, idx), txt, size=10.5)
    cap40 = shp(slide, 40)
    cap40.left = Inches(3.72)
    cap40.width = Inches(4.1)
    cap40.top = Inches(5.62)
    set_text(cap40, '采收率、含水与压差随注入体积的响应', size=13)
    pic_ = shp(slide, 41)
    pic_.width = Inches(3.86)
    set_pic(slide, pic_, IMG + 'recovery.png')
    # 中间第二张小图（注水驱对比曲线）留，第三张碎图去掉，避免挤
    try:
        p42 = shp(slide, 42)
        p42.left = Inches(8.0)
        p42.width = Inches(1.9)
    except Exception:
        pass
    set_pic(slide, shp(slide, 42), IMG + 'coreflood-rig.png')
    set_pic(slide, shp(slide, 43), IMG + 'ift-conc.png')
    # 顶部两条指标说明
    g1 = shp(slide, 1)
    inner = [x for x in g1.shapes if x.shape_type == 6]
    holder = inner[0] if inner else g1
    lines = ['采收率：以水驱后转化学驱的增幅计，同步给出含水与压差',
             '注入性：记录注入压力与视吸水指数的变化，判断是否存在堵塞']
    k = 0
    for sub in holder.shapes:
        if sub.has_text_frame and sub.text_frame.text.strip():
            set_text(sub, lines[min(k, 1)], size=12)
            k += 1
    # 竖排的“室内岩心驱替 / 对照与重复性”改成横排标题条（先取元素再删，避免索引错位）
    keep = {i: shp(slide, i) for i in (24, 25, 26, 27, 28, 29, 34, 35, 36, 37, 38, 39)}
    for idx in (33, 32, 23, 22):
        try:
            drop(shp(slide, idx))
        except Exception:
            pass
    for x0, txt in ((0.95, '室内岩心驱替'), (9.78, '对照与重复性')):
        bar = add_text(slide, x0, 3.30, 2.6, 0.3, txt, size=11, color=BLUE, bold=True)
        bar.fill.solid()
        bar.fill.fore_color.rgb = WHITE
    for base in (24, 34):
        for k in range(3):
            b = keep[base + 2 * k]
            b.top = Inches(3.66 + k * 0.76)
            b.height = Inches(0.66)
            tb = keep[base + 2 * k + 1]
            tb.top = Inches(3.68 + k * 0.76)
            tb.height = Inches(0.62)

    # 两侧“深海地图 / 三维地质”装饰图删掉（放最后，避免影响前面的索引）
    for idx in (19, 18):
        try:
            drop(shp(slide, idx))
        except Exception:
            pass

# ----------------------------------------------------------------- 15 资源
def fill_res(slide):
    fix_logo(slide)
    set_text(shp(slide, 2), '6.1', size=26)
    set_text(shp(slide, 3), '项目资源介绍', size=26)
    g5 = shp(slide, 5)
    for sub in g5.shapes:
        if sub.has_text_frame and sub.text_frame.text.strip():
            set_text(sub, '依托学院现有实验条件与前期中相微乳液资料，可支撑相态、界面与岩心驱替全流程',
                     size=20)
    set_text(shp(slide, 15), '实验仪器条件', size=13)
    set_text(shp(slide, 25), '材料与样品条件', size=13)
    set_text(shp(slide, 18), '驱替与配液装置', size=11)
    set_text(shp(slide, 21), '驱替后残余油分布', size=11)
    set_text(shp(slide, 28), '相态观察与配液', size=11)
    set_text(shp(slide, 31), '复配界面膜示意', size=11)
    set_pic(slide, shp(slide, 16), IMG + 'apparatus.png')
    set_pic(slide, shp(slide, 19), IMG + 'coreflood-rig.png')
    set_pic(slide, shp(slide, 26), IMG + 'tubes.png')
    set_pic(slide, shp(slide, 29), IMG + 'mixed-film.png')
    set_text(shp(slide, 11), '长江大学石油工程学院', size=12)
    # 底部两行说明统一左起点、等距
    band_lines = [s2 for s2 in slide.shapes
                  if s2.has_text_frame and 5.9 * 914400 < (s2.top or 0) < 7.4 * 914400
                  and s2.left > 3.5 * 914400]
    band_lines.sort(key=lambda s2: s2.top)
    for i, s2 in enumerate(band_lines[:2]):
        s2.left = Inches(3.9)
        s2.width = Inches(8.9)
        s2.top = Inches(6.22 + i * 0.5)
        for p_ in s2.text_frame.paragraphs:
            p_.alignment = PP_ALIGN.LEFT

    # 三张统计卡：模板那组是带图标的组合，缩放会带着图标一起变形，
    # 这里直接删掉，改用我们自己的三张等宽卡铺满整行。
    victims = [g for g in list(slide.shapes)
               if g.shape_type == 6 and 4.4 < (g.top or 0) / 914400 < 5.4
               and (g.left or 0) > 3.5 * 914400]
    for g in victims:
        drop(g)
    stats = [('4 人', '学生团队'), ('1 人', '指导教师'), ('4 类', '实验条件')]
    for i, (num, label) in enumerate(stats):
        x = 0.42 + i * 4.24
        c = shape(slide, MSO_SHAPE.ROUNDED_RECTANGLE, x, 4.74, 4.0, 0.96,
                  fill='FFFFFF', line=BORDER, radius=0.08)
        box = textbox(slide, x + 0.2, 4.78, 1.7, 0.88, anchor=MSO_ANCHOR.MIDDLE,
                      align=PP_ALIGN.RIGHT)
        para(box.text_frame, num, size=22, color='C00000', bold=True,
             first=True, align=PP_ALIGN.RIGHT, space_after=0)
        box = textbox(slide, x + 2.0, 4.78, 1.8, 0.88, anchor=MSO_ANCHOR.MIDDLE)
        para(box.text_frame, label, size=13, color='5A6A85', first=True, space_after=0)

    g8 = shp(slide, 8)
    lines = ['学院具备相态观察、界面张力、黏度与岩心驱替的完整实验条件',
             '前期已积累中相微乳液相态、界面与驱替资料，可直接用于方案设计']
    k = 0
    for sub in g8.shapes:
        if sub.has_text_frame and sub.text_frame.text.strip():
            set_text(sub, lines[min(k, 1)], size=11.5)
            k += 1
    # 图注条改成白底蓝字，避免压在图片上看不清
    for idx in (17, 20, 27, 30):
        try:
            bar_ = shp(slide, idx)
            bar_.fill.solid()
            bar_.fill.fore_color.rgb = WHITE
            bar_.line.color.rgb = WHITE
        except Exception:
            pass
    for idx in (18, 21, 28, 31):
        try:
            for p_ in shp(slide, idx).text_frame.paragraphs:
                for r_ in p_.runs:
                    r_.font.color.rgb = BLUE
                    r_.font.bold = True
        except Exception:
            pass

    # 左下角那张团队成员大图与底部蓝条重叠，删掉
    try:
        drop(shp(slide, 7))
    except Exception:
        pass



# ----------------------------------------------------------------- 17 成果
def fill_result(slide):
    fix_logo(slide)
    set_text(shp(slide, 1), '7.1', size=26)
    set_text(shp(slide, 2), '预期成果', size=26)
    g4 = shp(slide, 4)
    for sub in g4.shapes:
        if sub.has_text_frame and sub.text_frame.text.strip():
            set_text(sub, '形成配方、数据与报告三类成果，可支撑后续竞赛与论文',
                     size=20)
    set_text(shp(slide, 20), '配方与认识', size=13)
    set_text(shp(slide, 22), '数据与方法', size=13)

    set_text(shp(slide, 8), '总结与交流', size=12)
    set_text(shp(slide, 13), '形成相态—界面—驱替的关联认识', size=11)
    set_text(shp(slide, 11), '完成研究报告与结题材料，参加创新创业训练交流', size=11)
    set_text(shp(slide, 14), '建立相态图、界面张力与驱替曲线的数据表与评价方法', size=11)
    set_text(shp(slide, 16), '给出候选配方与段塞参数建议', size=11)

    # 底部那条原本是整条实心红块，太扎眼，改成浅底蓝字的说明条
    bar24 = shp(slide, 24)
    try:
        bar24.left = Inches(0.39)
        bar24.top = Inches(6.67)
    except Exception:
        pass
    try:
        bar24.fill.solid()
        bar24.fill.fore_color.rgb = RGBColor(0xF2, 0xF5, 0xFB)
        bar24.line.color.rgb = RGBColor(0xC1, 0xCC, 0xE4)
        bar24.line.width = Pt(0.75)
    except Exception:
        pass
    set_text(bar24, '贯通“配方筛选—相态评价—界面测试—岩心驱替”的完整链条，'
                    '形成可复用的实验流程与记录规范', size=13, color=BLUE, bold=True,
             align=PP_ALIGN.CENTER)
    set_pic(slide, shp(slide, 26), IMG + 'micro.png')
    set_pic(slide, shp(slide, 27), IMG + 'ift-conc.png')
    mech = shp(slide, 23)
    mech.left = Inches(4.42)
    mech.top = Inches(2.06)
    mech.width = Inches(4.5)
    mech.height = Inches(2.55)
    set_pic(slide, mech, IMG + 'gpt-mechanism.jpg')

    # 最后再删模板那套组织图（放在末尾，避免索引错位）
    victims = []
    for idx in (5, 6, 7, 8, 17, 18, 25):
        try:
            victims.append(shp(slide, idx))
        except IndexError:
            pass
    for v in victims:
        try:
            drop(v)
        except Exception:
            pass
    cap = add_text(slide, 5.16, 5.07, 2.95, 0.54,
                   '四重作用机理示意：界面张力、增溶、润湿反转、流度调整',
                   size=16, color=RGBColor(0x5A, 0x6A, 0x85), bold=False)
    cap.text_frame.paragraphs[0].alignment = PP_ALIGN.CENTER


# ----------------------------------------------------------------- 18 结束
def fill_end(slide):
    # 用户版：结束页只留“感谢…”大字 + 校徽 + 单位时间一行
    logo = shp(slide, 8)
    sub = shp(slide, 7)
    pill = shp(slide, 2)
    grid = shp(slide, 9)
    set_pic(slide, logo, BRAND + 'logo_white.png')
    for v in (sub, pill, grid):
        try:
            drop(v)
        except Exception:
            pass
    add_text(slide, 0.9, 6.42, 11.5, 0.42,
             '长江大学　石油工程学院　·　2025 年 10 月', size=18, color=BLUE,
             align=PP_ALIGN.CENTER)
