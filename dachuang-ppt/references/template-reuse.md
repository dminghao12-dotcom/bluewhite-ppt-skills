# 直接在用户模板上改：手法与坑

目标：让成品看起来是「用他们那份模板做的」，而不是照着风格仿的。做法是拿模板 PPTX 当底稿，**改它自己的文字框和图片位**，再把我们的自绘页插进去。

## 一、标准流程（build2.py 的顺序）

```python
prs = Presentation('tpl/模板.pptx')
drop_slides(prs, keep_idx)          # 删掉用不上的模板页
cover, toc, bg, metric, result, end = list(prs.slides)
tpluse.fill_cover(cover)            # 逐页填内容
...
for fn in [pages1.s03, ..., pages2.s19]:   # 加自绘页
    fn(prs)
reorder(prs, order)                 # 重排成目标页序
renumber_partnames(prs)             # 见坑 1
write_notes(prs, SCRIPT)            # 讲稿写进备注
prs.save(out)
```

## 二、八个坑（每个都踩过）

### 1. partname 冲突 → 保存出重名文件
删页之后 `add_slide()` 仍可能分到已被占用的 `slideNN.xml`，保存时 zip 里出现两份 `ppt/slides/slide11.xml`，PowerPoint 打开会报错或丢页。

修：
```python
from pptx.opc.packuri import PackURI
parts = [s.part for s in prs.slides]
for i, p in enumerate(parts):          # 先全部改成临时名
    p.partname = PackURI(f'/ppt/slides/tmp{i}.xml')
for i, p in enumerate(parts, 1):       # 再改成最终名
    p.partname = PackURI(f'/ppt/slides/slide{i}.xml')
```
保存后自检：`zipfile` 里不应有重名条目。

### 2. 删除形状会让索引全部错位
`slide.shapes[i]` 是位置索引。一旦 `drop()` 了某个形状，后面的下标全部前移。两种正确写法：

- 先把要用的形状对象全部取出来，再删：
```python
keep = {i: shp(slide, i) for i in (24, 25, 26)}
victims = [shp(slide, i) for i in (5, 6, 8)]
for v in victims: drop(v)
keep[24].top = Inches(3.66)            # 对象引用仍然有效
```
- 或者**把删除放到函数最末尾**，前面照常按原始索引操作。

### 3. group 不能随便改宽高
`grpSp` 有 `chOff/chExt`，改 group 的 `width/height` 会把子元素（图标、图形）整体缩放，出现「图标突然变巨大」。要改位置就只改 `left/top`；要改大小就逐个改子元素，或者删掉自绘。

### 4. 换图片要改 r:embed，不要 add_picture
```python
image_part, rId = slide.part.get_or_add_image_part(path)
shape._element.blipFill.blip.rEmbed = rId
```
这样位置与尺寸（含裁剪）全保留。`shape.shape_type == 13` 才是图片；`blipFill` 只存在于 `p:pic`。

### 5. 换文字要保留原字体属性
模板的 run 带字号/字体/字距，直接 `tf.text = '...'` 会全部丢掉。做法：取第一个 run 的元素 deepcopy 当模板，删掉原 run，再按行重建：
```python
ref = copy.deepcopy(tf.paragraphs[0].runs[0]._r)
...
r = copy.deepcopy(ref); p._p.append(r); p.runs[0].text = line
```
句子长一档就换行的事要在模板里留够宽；不够就减字，别缩字号。

### 6. 备注正文占位符可能不存在
模板页的 notesSlide 没有 body placeholder 时 `notes_text_frame` 是 `None`。从任意一张有备注的页 deepcopy 一个 `p:sp` 过来，再写文本。

### 7. 模板文本框可能自带「每段不同对齐」
结束页那种「感谢…／敬请批评指正」的双行框，常是一段左对齐一段居中：`set_text` 只换文字，对齐与缩进仍各自保留，渲染出来就是一行偏左、一行居中（用户会直接说「最后一页问题有点大」）。凡是整块重排的文字（封面抬头、结束页、目录行），**把原框删掉重画**，不在原框上修。

### 8. run 级字符间距会跟着 deepcopy 跑
模板的「单位时间」「汇报人」这类单行短文案有时带 `spc`（字符间距），拿首个 run 当模板就会连间距一起继承，渲染成「2026　 年　 10　 月」这种散开的行。单行短文案不要继承原 run，直接新建 textbox 写。

## 三、哪些模板页值得套

| 模板页形态 | 套做什么 | 注意 |
|---|---|---|
| 封面（大标题+波浪+校徽位） | 封面 | 只换 3 处文字 + logo，别加副标题小字 |
| 目录（左色块+若干行章节条） | 汇报提纲 | 行数不够就克隆「亮」的那一套样式，单列竖排 |
| 双卡 + 三个小指标 + 通栏判断句 | 背景、现状 | 卡片里的新闻截图换成我们的现场照 |
| 两侧指标柱 + 中间大图 | 驱替/指标页 | 竖排标签要改横排，指标块跟着下移 |
| 左右两卡 + 中间图 + 底部条 | 成果页 | 中间若带组织图/图标块，直接删，只留图 + 图注 |
| 结束页 | 结束 | 同封面，只留必要文字 |

**不要套**：带团队头像的团队页、四课题矩阵、装备清单页——都是对方业务版式，硬套会打架。
