# -*- coding: utf-8 -*-
"""把 20 页压到 17 页，在用户那份 pptx 上增量改（不整份重建）。

- 3.2 + 4.2 → 一页（三列对应）
- 4.1 重画成五项创新
- 5.1 + 5.6 → 一页（技术路线 + 12 个月进度）
- 5.2 + 5.3 → 一页（配方筛选 + 界面性能）
- 5.4 → 5.3，5.5 → 5.4（只改编号）
"""
import os
import sys

from pptx import Presentation
from pptx.util import Inches

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
sys.path.insert(0, ROOT)
os.chdir(ROOT)

from deck import merge  # noqa: E402

SRC = '中相微乳液-大创申报答辩-20页-模板版.pptx'
DST = '中相微乳液-大创申报答辩-17页.pptx'

# 原稿里要删掉的页（0 基）：3.2、4.1、4.2、5.1、5.2、5.3、5.6
DROP = [3, 7, 8, 9, 10, 11, 12, 13, 15]

NEW_NOTES = {
    '2.1': ('2.1 研究背景',
            '先说背景。国内主力油田普遍进入高含水后期，水驱可采储量动用程度已经很高，'
            '继续靠水驱提采的空间有限；同时低渗、中低渗储量占比在上升，孔喉细、启动压力高。'
            '右边这张是典型井的生产曲线，日产油下来、含水上去，这是高含水期的常态。'
            '两处约束最后指向同一个问题：注入水沿高渗条带突进，剩余油以油膜、孤立油滴和'
            '孔喉残余的形式留在孔隙里；低渗这边想靠提压解决，又接近破裂压力。所以要同时'
            '考虑扩大波及体积和提高洗油效率，这也是后面配方和评价设计的出发点。'),
    '3.2': ('3.2 现有研究不足与拟解决的关键问题',
            '这一页把原来的“不足”和“关键问题”并成一页，三行一一对应：左边是不足，中间是本项目'
            '要回答的问题，右边是回答它的实验。第一行是窗口容错性，地层水混入、稀释和吸附之后'
            '相区会偏移，所以问题是怎么拿到一个有容错性的稳定相区，验证方式是浓度、比例、盐度'
            '扫描再加上混入稀释复测。第二行是注入性，增油和压差过去记录不同步，流度调整和孔喉'
            '堵塞分不开，验证是界面张力、增溶加上温度钙镁老化的复测。第三行是配方筛选靠试错，'
            '验证是岩心驱替里同步记录采收率、含水、压差和注入体积。三个问题是递进的：前一个'
            '不成立就不往下走。'),
    '4.1': ('4.1 项目创新点',
            '创新点按申请书列了五项，不是四项。前四项都能用实验直接检验：主剂与复配设计，'
            '看甜菜碱在目标油藏条件下的适用优势与限制；评价方法，把相态、界面和岩心驱替'
            '联合起来做优选；应用路线，探索低浓度注入液和地层油水接触后原位成相的条件；'
            '功能评价，把增油和注入风险放在同一套流程里判断。第五项是配方设计手段，属于'
            '探索性内容：建数据表，用回归或响应面看成相边界，再用新增实验检验，不承诺'
            '建通用模型。这一条原来是漏的，现在补上，和申报书一致。'),
    '5.1': ('5.1 实施方案 · 技术路线与进度安排',
            '技术路线和进度原来占两页，其实是同一件事，现在并成一页。上面五段是从准备到'
            '整合的推进顺序：文献与条件、配方与相态、界面与适应性、岩心驱替、数据整合，'
            '每段下面写清这一段要交出来的东西。下面这张是四阶段的时间轴，一共十二个月：'
            '前两个月做资料、条件和方案；三到六月做配方筛选和性能评价，这一段最长，因为'
            '变量最多；七到十月做驱替验证和机理分析；十一到十二月整理成果。阶段之间用'
            '输出和阶段评审衔接，每段留两周缓冲。'),
    '5.3': ('5.3 岩心驱替与增油评价',
            '静态筛出来的配方必须回到岩心里验证。流程是一条线走到底：饱和水、饱和油、水驱到'
            '含水百分之九十八，再注化学段塞零点三到零点五倍孔隙体积，最后后续水驱。整个'
            '过程同步记四项：采收率按转驱后的增幅算，用百分点；含水率看什么时候见效；'
            '驱替压差用来区分流度调整和孔喉堵塞；注入体积折算成孔隙体积倍数，比较药剂用量。'
            '判读上有两条约束：一是必须有同矿化度配液的水驱对照，二是至少两块平行样。'
            '压差升高不能直接说成调剖见效，要结合产出液和孔喉堵塞一起判断。'),
    '5.2': ('5.2 实施方案 · 配方筛选与界面性能评价',
            '配方筛选和界面评价原来是两页，实际是一条链，现在合成三步。第一步配方与盐度'
            '扫描，用磺基、羧基甜菜碱做主剂，三类体系做对照，浓度、复配比例、盐度都扫，'
            '记录相数、相体积和成相时间，画出鱼状相图，定出中相区间和最低有效浓度。'
            '第二步看界面与增溶，旋转滴法测动态和平衡界面张力，以十的负三次方毫牛每米'
            '量级或更低作为优选参考，同时算油和水的增溶参数。第三步是环境适应性：温度、'
            '钙镁、地层水混入、注入水稀释、老化和吸附之后都要复测，看窗口往哪边偏，'
            '偏移方向反过来指导前置段塞调盐。'),
}


def drop_slides(prs, idxs):
    lst = prs.slides._sldIdLst
    entries = list(lst)
    from pptx.oxml.ns import qn
    for i in sorted(idxs, reverse=True):
        e = entries[i]
        prs.part.drop_rel(e.get(qn('r:id')))
        lst.remove(e)


def reorder(prs, order):
    lst = prs.slides._sldIdLst
    entries = list(lst)
    new = [entries[i] for i in order]
    for e in entries:
        lst.remove(e)
    for e in new:
        lst.append(e)


def renumber(slide, old, new):
    for sh in slide.shapes:
        if sh.has_text_frame and sh.text_frame.text.strip() == old:
            for p in sh.text_frame.paragraphs:
                for r in p.runs:
                    r.text = new
            return True
    return False


def build(src=SRC, dst=DST):
    prs = Presentation(src)

    # 先记下要保留页的备注
    kept_notes = {}
    for i, s in enumerate(prs.slides):
        tf = s.notes_slide.notes_text_frame
        kept_notes[i] = tf.text if tf is not None else ''

    # 新页（追加在末尾）：3.2 → 4.1 → 5.1 → 5.2
    new = [merge.p21(prs), merge.p32(prs), merge.p41(prs), merge.p51(prs),
           merge.p52(prs), merge.p53(prs)]
    n0 = len(prs.slides) - len(new)

    drop_slides(prs, DROP)

    # 删除后旧页：0,1,2(册/目/分工), 3=2.2定位, 4=2.3意义, 5=3.1现状,
    # 6=5.5数据, 7=资源, 8=成果, 9=经费, 10=结束; 11..16=新页(2.1→5.3)
    order = [0, 1, 2, 11, 3, 4, 5, 12, 13, 14, 15, 16, 6, 7, 8, 9, 10]
    reorder(prs, order)

    # 编号：5.5 → 5.4
    assert renumber(prs.slides[12], '5.5', '5.4')

    # 备注
    from copy import deepcopy
    donor = None
    for s in prs.slides:
        if s.notes_slide.notes_text_frame is not None:
            donor = s.notes_slide.notes_text_frame._txBody
            break
    # 位置 → 原稿页号（保留页：0-6、14 数据、16 资源、17 成果、18 经费、19 结束）
    old_map = [0, 1, 2, 'new:2.1', 4, 5, 6, 'new:3.2', 'new:4.1', 'new:5.1', 'new:5.2',
               'new:5.3', 14, 16, 17, 18, 19]
    for s, key in zip(prs.slides, old_map):
        if isinstance(key, str):
            head, body = NEW_NOTES[key[4:]]
            text = head + '\n' + body
        else:
            text = kept_notes.get(key, '')
        ns = s.notes_slide
        if ns.notes_text_frame is None and donor is not None:
            ns.shapes._spTree.append(deepcopy(donor))
        ns.notes_text_frame.text = text

    # 备注首行统一成「第 N 页 · 页标题」
    for i, s_ in enumerate(prs.slides, 1):
        title = ''
        for sh in s_.shapes:
            if sh.has_text_frame and sh.top is not None and sh.top < 0.7 * 914400 \
                    and 1.0 < sh.left / 914400 < 3.0:
                title = sh.text_frame.text.strip()
                break
        tf = s_.notes_slide.notes_text_frame
        if tf is None:
            continue
        body = tf.text.split('\n', 1)[1] if '\n' in tf.text else tf.text
        fallback = {1: '封面', 2: '汇报提纲', 17: '结束'}
        tf.text = f'第 {i} 页 · {title or fallback.get(i, "内容")}\n' + body

    # partname 重编，避免重名
    from pptx.opc.packuri import PackURI
    parts = [s.part for s in prs.slides]
    for i, p in enumerate(parts):
        p.partname = PackURI(f'/ppt/slides/tmp{i}.xml')
    for i, p in enumerate(parts, 1):
        p.partname = PackURI(f'/ppt/slides/slide{i}.xml')

    prs.save(dst)
    print(f'saved {dst}  slides={len(prs.slides)}')
    return dst


if __name__ == '__main__':
    build(*(sys.argv[1:] or []))
