# 用用户自带的模板做整份 PPTX

先核实用户要求的是整份套用、只借单位VI、只取校徽，还是仅作视觉参考。用户明确“参考这里，但具体排版用我们的”时，沿用科研PPT skill的首尾页及正文版式，仅借鉴指定的小插图与视觉细节；不自动套封面、复制母版或移植校名。只有用户要求整份套用现成模板（医院答辩模板、单位讲课模板、会议模板，`.ppt` 或 `.pptx`）时，**交付件的视觉才是模板本身**：母版背景、页眉装饰、机构标识、比例全部原样保留。本技能的蓝白信息图风只在没有模板时用；有模板时改的是内容密度和排布，不是配色。

先问一句会场比例：模板是 4:3（10×7.5 in）就按 4:3 交，不要拉伸成 16:9，也不要因为“现在都用宽屏”私自换。模板比例就是交付比例，除非用户明确说会场是宽屏、要求重排。

## 零、只取校徽（用户说“这个模版只扣掉 logo 就行”时）

“扣 logo”“只要这个模版的 logo”“其他用你们的”都是在说：**只要校徽，版式全用本 skill**。这时不要克隆母版、几何底纹、章节页、校园照片，也不要按模板比例改页——照常 16:9 建稿，只把校徽接成 `meta.header_logo` / `meta.logo`。判错这一句的代价是整份返工，拿不准就先问“整份套模板，还是只把校徽拿来用”。

取 logo 不要信 `shape.image.blob`：模板右上角那个“图片”常是整条装饰底纹（实测 18000×3600、5:1），直接存下来是一块蓝底几何图，不是校徽。可靠做法是从渲图里截：

```python
import pymupdf, numpy as np
from PIL import Image

d = pymupdf.open('tpl.pdf')
# 只裁右上角标识区（页宽 10 in、高 5.6 in 的模板：x 7.0–9.95、y 0.05–0.85）
d[3].get_pixmap(dpi=300, clip=pymupdf.Rect(7.0*72, 0.05*72, 9.95*72, 0.85*72)).save('logo_zone.png')

im = Image.open('logo_zone.png').convert('RGB')
a = np.array(im); mask = a.sum(axis=2) < 720            # 非白像素
ys, xs = np.where(mask)
im.crop((xs.min()-4, ys.min()-4, xs.max()+5, ys.max()+5)).save('images/logo_header.png')

# 方形校徽：按列墨迹找“校徽 → 校名”之间的第一段空白（>4 px），切出来补白成正方形
cols = mask.any(axis=0); gaps=[]; run=None
for x, v in enumerate(cols):
    if not v and run is None: run = x
    elif v and run is not None:
        if x - run > 4: gaps.append((run, x))
        run = None
seal = Image.open('images/logo_header.png').crop((0, 0, gaps[0][0], im.size[1]))
side = max(seal.size); canvas = Image.new('RGB', (side, side), 'white')
canvas.paste(seal, ((side-seal.size[0])//2, (side-seal.size[1])//2)); canvas.save('images/logo_seal.png')
```

- 白底 logo 用在白页页眉上不需要抠透明；要贴到深色区先按像素阈值转 alpha，不要直接裁白边。
- **用户直接发一张 logo 截图（JPG / PNG）而不是模板时**：先量出它自带的“透明格子”。手机截图和微信转存的 JPG 会把棋盘格烤进像素（实测底色在 240–255、“格子” 244–254 不恒定），直接当 `header_logo` 贴到白页上会出现一层灰方块。做法：`a = np.asarray(im).astype(int)`，把 `(a.max(2)-a.min(2) <= 10) & (a.mean(2) > 236)` 的像素全置 255，再按非白 bbox 收紧；绿色字标与圆形院徽不受影响，边缘的抗锯齿会自然虚化。**图底那条横线是装饰线、不是 logo 的一部分**：按行墨迹量出它（实测满行墨迹在 y=217–221），裁切的 y 上限取在它上方；保留下来会变成页眉里一条很淡的横线，看着像做坏了。
- 同一张 logo 就是取色依据：墨绿类 VI 从笔画的深色像素统计（`max-min>25 且 sum<600` 的众数，实测 `#214332`–`#406150`，取 `#17493A` 做主色、章节带压到 `#0E3226`），不要凭肉眼估成“深绿”。
- **clip 的纵向区间要覆盖整个 lockup，不要按「校徽大约多高」设**：横版校徽的英文校名行压在最下方，y 上限给到 0.85 in 会把它切掉，渲出来是校徽底下一条 1–2 px 的灰线——封面和每页页眉同时出现，看着像模板自带的装饰线，其实是切剩下的半行字。y 上限给到 1.4–1.5 in 再靠墨迹 bbox 收紧；横向 lockup 的宽高比应在 3:1 左右（实测江南大学横版 894×286 = 3.13），比例明显偏离或图底出现细横线就是没裁全，重裁一次比在 PPT 里找原因快。
- 页眉仍按高 0.5 in、右缘 13.02 in 放；`spec_audit` 按形状名 `机构标识` 跳过尺寸检查，自绘封面里的 logo 也沿用这个命名。
- 借用的是别校（或上一个课题）的校徽时，交付里点一句“页眉沿用了上次的校徽，需要换/去掉说一声”——学科与单位不匹配时评委第一眼就会注意到。
- 用户后来改口“logo 空着即可 / 不要校徽”时，直接把 `meta.header_logo` 与 `meta.logo` 从内容 JSON 里删掉就行，不要留空白图片框，也不要自作主张换成别的单位标识；删完重渲一页确认页眉真的干净（`header_logo` 是页眉绘制入口，留着旧路径还会把上一份的 logo 贴回来）。

## 一、流程

```bash
# 1) .ppt 先转格式（python-pptx 读不了旧版 .ppt）
soffice --headless --convert-to pptx template.ppt
soffice --headless --convert-to pdf  template.ppt   # 同时渲出来看母版

# 2) 逐页渲成 png，量几何、取色
python -c "import pymupdf; d=pymupdf.open('template.pdf'); [d[i].get_pixmap(dpi=110).save(f'tpl-{i+1}.png') for i in range(d.page_count)]"
```

1. **先看模板有哪几页**：通常一页封面 + 一页空白内容页。用 python-pptx 把每页的 shape 名、位置、尺寸打出来，分清哪些在 slide 上、哪些在 layout 上。
2. **背景在 layout 的 PICTURE 里**（不在 master，也不在 slide）：所以 `add_slide(prs.slide_layouts[0])` 出来的新页会自动带上背景/页眉，不需要手工贴图。
3. **只留封面页，删掉示例内容页**：

```python
sldIdLst = prs.slides._sldIdLst
for sid in list(sldIdLst)[1:]:
    rId = sid.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id')
    prs.part.drop_rel(rId)
    sldIdLst.remove(sid)
```

4. **新页统一用 `new_slide()`**：`add_slide(layout)` 之后立刻删掉 layout 带来的占位符（`if sh.is_placeholder: sh._element.getparent().remove(sh._element)`），然后全部自己画——模板的占位符位置通常和正文区对不上。
5. **主色从渲图里取**：`PIL.Image.getpixel()` 采样色带、粗线、院徽，得到主色与浅底色（例：深红 `#8F1019`、浅底 `#F7EDEA`、正文近黑 `#221C1B`）。不要凭肉眼估色号。
6. **页脚与“结论条”**：页脚文字放 y≈6.94，结论条底边不要压过它——结论条统一在 6.40（高 0.42），底边 6.82 与页脚留 0.12 in。

## 二、4:3 的几何（10 × 7.5 in，先量页眉装饰的下沿）

模板的页眉通常有三层：云纹/纹理窄带、一条粗横线、右上角机构标识（有的还带左下角云纹角标）。**页标题必须落在这些东西下面**，否则字被压在红线上、或与角标重叠。

量法：把模板内容页渲成 png，看粗横线的像素 y 除以 DPI；实测一份院内模板是 0..1.16 in 装饰带、1.10–1.32 in 粗线、角标到 1.6 in。据此定的栅格：

| 元素 | y (in) | 说明 |
|---|---|---|
| 页标题 | 1.58 | 21–22 pt 粗体，主色 |
| 引导句 | 2.04 | 12.5 pt，一行 |
| 内容区 | 2.45 – 6.40 | 高约 3.95 in，两栏各 4.34 in（页宽 10、左右边距 0.55） |
| 结论条 | 6.40 | 高 0.42 |
| 页脚 | 6.94 | 8.5 pt 灰字 + 页码 |

字号比 16:9 的蓝白模板整体小一档：页标题 21–22、栏标题 12.2、要点 11–11.5、表格 10.3–10.8、说明块 10.8。4:3 的可用宽度只有 8.9 in，三栏会被压到 2.8 in 一栏——证型、条目类内容用**两栏 + 底部通栏说明**，不要硬套三栏。

## 三、四个必踩的 python-pptx 坑

1. **背景图盖住标题**：slide 上那张整页 PICTURE 在 spTree 里排在占位符之后（后画的在上），标题文字明明写进 XML 了，渲出来却完全看不到。修法是把 PICTURE 移到最底层、再把标题占位符挪到最上层：

```python
spTree = slide.shapes._spTree
for pic in [s for s in slide.shapes if s.shape_type == 13]:
    spTree.remove(pic._element); spTree.insert(2, pic._element)
# 写完文字后
for sh in [s for s in slide.shapes if s.is_placeholder]:
    spTree.remove(sh._element); spTree.append(sh._element)
```

2. **副标题前面冒出一个小圆点**：占位符从 layout 继承了 `buChar`。逐段删掉 `a:buChar` / `a:buAutoNum` 并补一个 `<a:buNone/>`，否则每条标题都带点。
3. **说明块第一行缩进 / 居中**：`add_shape()` 出来的 autoshape，其 `text_frame` 段落默认居中，直接用 `shape.text_frame` 写字会得到一段居中、下一段左对齐的怪排版。凡在形状里写字，逐段显式 `p.alignment = PP_ALIGN.LEFT`（或 `_rich()` 统一设）。
4. **表格单元格里的 `\n` 不换段落**：`run.text` 里带 `\n` 会变成段内换行，第二行对齐会跑偏（表头居中、第二行右对齐）。把单元格文本 `split('\n')` 成多个 paragraph，每段单独设 `alignment`。

## 四点半、只借 VI、不套版式（用户明确要求“排版还是用我们自己的 skill”时）

用户给了学校的形象识别模板（VI 讲义模板）但要求版面仍用本技能时，只取四样东西，其余全部沿用蓝白信息图（或对应主题）的栅格：

| 取 | 不取 |
|---|---|
| 主色（通常要把 VI 的亮色压深一档，否则正文看不清） | 模板的 4:3 栅格与占位符 |
| 校徽（从模板页 300–400 dpi 渲出后裁切，不改色不变形） | 模板的页眉装饰带 |
| 封面壳：白底 + 左上校徽 + 右下底纹/水印 + 左对齐标题块 | 模板的内容页版式 |
| 配色层次（深浅同色系） | 模板的字号体系 |

**只有用户要求沿用该单位VI或指定模板封面时，才套其封面壳，不用panel首页；仅作视觉参考的PPT不适用。** 确认套用后按以下步骤实现：

1. 从模板的 layout 背景图里取校徽与右下装饰（`.jpeg` 常是 CMYK，先 `convert('RGB')`）。
2. 装饰图左侧做 15%–20% 的 alpha 渐隐，否则裁出来的硬边会在白页上留下一道竖线。
3. 位置照模板占位符换算：标题左边缘约页宽 10%、顶端约 28%；副题在标题下 1.2 in，出处与日期再往下。
4. 装饰图贴右下角、右边与下边压齐页面边界，但**不要超出页面**——超出会报 `超出页面边界`。
5. 文字框宽度收到内容实际宽度，别通栏铺到装饰图上：lint 会报“与底纹重叠”，即使视觉上没压字。

主题色统一走 `scripts/theme.py`（`PPT_THEME=green` 等），不要在稿子里手写色号；新增一档主题就加一个 dict，同时更新 SKILL.md 的配色说明。

封面壳用独立脚本画（放 `custom/cover_*.py`），标题与副题从 `content.json` 读出来、不要在脚本里再抄一遍——两处抄写必然漂移；脚本写进 `build.sh`，位置在 `dense_deck` 之后（重建会整页重画封面，所以每次构建都要跑）。封面壳里每个文本框都要显式设左对齐与 `lang="zh-CN"`：模板式左对齐标题块若沿用文本框默认对齐，整块会跑偏。

## 五、内容与交付

- 模板里的**院徽、院名、院训印章原样保留**，不要重画、不要调深浅；封面落款只写“汇报人：＿＿＿＿”和日期，不要再重复一遍单位名（院名已经在封面美术里）。
- 模板是答辩/讲课用的，页脚写主题名 + 页码，页码用 `{i} / {total}` 脚本生成，不要手写。
- 讲稿写字幕备注（`slide.notes_slide.notes_text_frame.text`），导出 PPTX 后按 220–230 字/分核算时长；再导一份 PDF 给评委/会场电脑。
- 交付前逐页看渲染图：标题是否都在装饰带下方、有没有页被背景图压住、表格是否压到结论条、页码是否连续。模板类稿子的排版错误一眼就能看出来，别只看“生成成功”。

## 六、重建式套版（模板是成品页、源稿是另一份材料时）

典型委托：“这份是我们的模版（20 页密集商务/咨询风），尽可能套进去，内容只参考我发的这份稿子”。判据在**模板页的性质**：若模板页是“成组对象（`title` / `body` / `footer` 三个 group）+ 每行折行都单独一个 TEXT_BOX + 手绘柱条”，就**不要原地改字**——中文与英文的字宽、折行位置全不同，改一处整组错位，等于逐行重排还改不干净。正确做法是把它当**设计系统**：量出色号/字号/几何，取出角标、栏首条、面板这几个 donor 元素，然后用 python-pptx **逐页新建**。页数按源稿内容定（源稿 11 页就 11 页），不按模板的 20 页凑。

### 6.1 先量模板：把字号与几何读成数字

```python
from pptx import Presentation
p = Presentation('tpl.pptx')
def walk(shapes, d=0):
    for sh in shapes:
        t = ''
        if sh.has_text_frame:
            for pa in sh.text_frame.paragraphs:
                for r in pa.runs:
                    if r.text.strip():
                        t += f'[{r.font.size.pt if r.font.size else None},{r.font.color.rgb if r.font.color.type is not None else ""}]={r.text[:40]} '
        print('  '*d + f'{sh.shape_type} {sh.name} L{sh.left/914400:.2f} T{sh.top/914400:.2f} W{sh.width/914400:.2f} H{sh.height/914400:.2f} {t[:120]}')
        if sh.shape_type == 6: walk(sh.shapes, d+1)
```

再对目标形状 `print(sh.element.xml)` 拿渐变、圆角 `adj`、边框色与线宽。实测一套 16:9 密集咨询模板：页标题 20.2 pt `#151078`（前半句 bold + 后半句 regular）、栏首圆角矩形渐变 `#248FE8→#0700CC`（`adj 8163`）、面板白底 + `#D8D8DE` 0.75 pt、右上角标三角 `#39A9EA`（1.17×1.15 in）+ 白字页码、结论带 `#E9E9F6`、页脚 9.8 pt `#747474`、正文 12.8 pt、图表标题 11.2 pt `#252064`。**这些数值就是本份稿的字号体系，不要重新发明一套**。

### 6.2 donor 元素 + 页骨架

```python
DONOR_CORNER = deepcopy(find(slide_shapes, 'chapter-corner').element)   # 先在删除页之前取好
def add_donor(slide, el, x, y, w, h):
    slide.shapes._spTree.append(deepcopy(el))
    sh = slide.shapes[-1]
    sh.left, sh.top, sh.width, sh.height = Inches(x), Inches(y), Inches(w), Inches(h)
    return sh
```

固定骨架：`角标+页码` → `页标题（bold lead + regular rest）` → `面板(hbar + 白底细框)` → `结论带` → `页脚(来源 + 页码)`。基准栅格按模板量，别按默认 0.5 in 边距。

### 6.3 四个必踩的 python-pptx 坑（重建式）

1. **`prs.slide_layouts[6]` 可能不存在**：自制模板常只留 1 个 layout，取 `[6]` 抛 `IndexError: slide layout index out of range`。用 `prs.slide_layouts[0]` 再删掉 layout 带来的占位符。
2. **封面/标题文字在 GROUP 里**：`for sh in slide.shapes` 只拿到 group 本身，`sh.text_frame` 为空，替换逻辑静默失效（渲染出来还是模板英文）。写一个递归 `rewrite(shapes)`，进 `shape_type == 6` 继续往下找 TEXT_BOX。
3. **`from deck_lib import *` 不导入下划线开头的函数**（`_set_font` 之类）→ `NameError`。显式 `from deck_lib import _set_font`。
4. **`Chart` 没有 `.format`**：去掉图表区填充与边框要手写 XML，且 `c:spPr` 必须按 schema 顺序插在 `c:txPr` / `c:externalData` 之前：

```python
spPr = OxmlElement('c:spPr'); spPr.append(OxmlElement('a:noFill'))
ln = OxmlElement('a:ln'); ln.append(OxmlElement('a:noFill')); spPr.append(ln)
anchor = next((cs.find(qn(t)) for t in ('c:txPr','c:externalData','c:printSettings') if cs.find(qn(t)) is not None), None)
anchor.addprevious(spPr) if anchor is not None else cs.append(spPr)
```

### 6.4 原生图表：正负号、逐点标签、值轴步长

- **负值标签必须逐点写字**：`data_labels.number_format = '+0;-0;0'` 在 LibreOffice 里只认第一段，`-14` 会渲染成 `+14`（图上柱子朝左、标签写正数，最容易被当成数据错）。改成 `plot.has_data_labels = False`，再对每个 `series.points[i].data_label` 写 `text_frame.text = '+48' / '−14'`，颜色按正负给。
- **`plot.data_labels` 取值前必须先 `has_data_labels = True`**，否则抛 `ValueError: plot has no data labels`。
- **值轴步长是 `value_axis.major_unit`**：写成 `category_axis.major_unit` 不报错但无效，刻度还是每 10 一格，满屏数字。
- 双系列横条：`category_axis.reverse_order = True` 保持词典序从上往下；图例顺序会跟着翻，页内图注写清“青＝直播，蓝＝短视频”，不要靠图例位置表达从属。
- **负值柱的方向不要拿 LibreOffice 预览当验收**：删 `invertIfNegative`、把类目轴改回 `minMax`、取消隐藏数值轴三种改法都试过，预览里负值柱仍然画在零轴右边、长度也不对；同一份数据另存后在 PowerPoint 里是从零轴向左的。能改的只有标签（逐点写字），**不要为预览重画图表**；交付时说明一句“预览器会把负值柱画反，请用 PowerPoint 看一眼”，不要让用户以为数据错了。

### 6.5 中文排版：负号、窄栏标签、表格撑高

- **正文里的负数用半角 `-`**：`−`（U+2212）在 CJK 断行规则下会在负号后断行，渲染成“与 −\n14%”；插 U+2060 word joiner 能止断但会渲出可见空隙。半角 `-14%` 被当成一个数值 token，不会拆开。（图表数据标签没有断行问题，仍可用 `−`。）
- **窄栏里的条形图，类别标签会被裁**：3 栏（每栏约 4.4 in）放 5 类、标签 6–7 字时，左右两端都会被截半字。不要缩字号硬塞——这一页改成两栏（每栏 5.9 in），或把类别名压到 4–5 字（“详情页视频浮窗”→“详情页浮窗”仍然长，就换“其余入口”这类短名）。改完必须放大看标签有没有被切。
- **表格行高会被自动撑高**：单元格文字折成两行时，python-pptx 设的 `row.height` 只是下限，LibreOffice/PowerPoint 会把行抻高，压在表格下方的结论句就重叠。表格下方留“行数 × 每行可能两行”的余量，首列名压到 4 字（“组合贡献”而不是“组合项贡献”）。
- 结论带统一：浅紫 `#E9E9F6` 通栏（不加左端竖条，用户不接受任何竖条）+ 粗体判断 + 一行说明；栏首只用一层（渐变条），不要再叠 `region.head` 式小标题。

### 6.6 数据口径：只给端点就不画月度曲线

源稿若只披露基期=100 与同比（例如“观看 +35%、成交 −14%”），却在图上画了 13 个点的月度曲线，**不要照着重建折线**——那是替你编数据。改画“观看/成交/每千次观看成交”三项同比横条，另配一行派生指数（`100×(1+同比)`），页内注明“同比按披露的整数值换算”；原稿隐去的绝对值继续写“已隐去”。分解类图要在页内标注“按两期均值分配交互项的恒等分解，不等于追加流量的因果收益”。这些注脚写进交付说明的「数字口径」一节，评审追问时直接指过去。

### 6.7 交付前机检 + 逐页看图

```bash
soffice --headless --convert-to pdf --outdir out out/deck.pptx
python -c "import pymupdf; d=pymupdf.open('out/deck.pdf'); [d[i].get_pixmap(dpi=150).save(f'out/render/p{i+1:02d}.png') for i in range(d.page_count)]"
```

机检四项，逐页过：① 所有 shape 的 `left/top/left+width/top+height` 都在画布内；② 全文不含模板残留词（原公司的品牌名、`IMAGE` / `LOGO` 占位字样）；③ 每页 `notes` 非空；④ 图片数、图表数、表格数与计划一致（`shape_type==13`、`has_chart`、`has_table`）。然后**逐页看图**，重点看四处机检看不到的：栏内文字有没有压到分隔线、表格下方结论句有没有被抻高的行挤住、条形图两端标签有没有被裁、负号有没有单独跑到行首。

配图按**画面内容**映射图位，不按返回顺序：一次要 4 张时它可能把第 4 项先画出来（实测“每日发现/搜索/推荐三张卡 + 随机分组示意”回来时，分组示意排在第一位）。取回后逐张开 `vision` 认主题再落命名。

封面若模板留了 `IMAGE` 占位：白底插画贴进深蓝飘带会像贴上去的卡片，要么裁到图位比例贴满、要么干脆去掉留白（模板封面本来就是文字 + 飘带）。**不要为了“有张图”硬塞。**

### 6.8 接手别人（另一个 agent / 同事）改过的一版：只补它留的占位

用户把上一版交给别的 agent 改，再丢回来说一句“还有很多地方待生图”时，**不要重排整份**，也不要凭缩略图找空位——生成器留的占位经常不是“看上去空”，而是一个灰方块里写着 `ICON`。

**先枚举、再动手**：递归 group 打全稿 `name / left,top,w,h / fill / 文字`，再按 `IMG|SLOT|ICON|IMAGE|LOGO|占位|图位` 过滤。实测命名是 `IMG_SLOT ICON`、正文 `ICON`、0.55–0.72 in、填充 `E6E9EF` + 边 `BCC3CF`——在整页缩略图上就是几个噪点，肉眼看页会整个漏掉，报告出来就是“没有空图位”，而用户看到的是“还有很多地方待生图”。

**替换而不是重排**：删掉占位 shape，在**完全相同的 left/top/width/height** 放图（`add_picture(..., width=w, height=h)`），相邻文字框一个都不动。他排好的栅格、栏首、结论条都保留；只补他点名的那几处，多的图不要自己加页。

**能动的是渲染缺陷，不能动的是内容判断**：

- 直接修、并在交付时点名：`numFmt '+0;-0;0'` 导致的负值显示成 `+58`（改逐点文本标签）、缺 `lang`、图例/轴标签被裁。写清“补了哪几处、修了哪几处”。
- **不要默默重画他的图**：接手稿里若有没数据支撑的内容（实测：按端点插值出来的 13 个月曲线、标着精确到小数点的指数），指出来并给两个可选项（只留端点 / 图注写“月度路径为示意”），让他选；擅自改等于替人做决定，改完还会被问“谁让你改的”。
- 交付回复按三段写：**补了什么**（图位尺寸与来源）、**修了什么**（缺陷 + 原因）、**建议但没动什么**（问题 + 两个选项）。
