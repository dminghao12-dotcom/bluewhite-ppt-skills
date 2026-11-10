---
name: research-ppt-skill
description: 科研场景 PPT——用脚本生成蓝白信息图式、原生可编辑的 16:9 PPTX，用于开题、中期、答辩、组会进展、文献精读/文献汇报、学术报告，也适用于竞选、活动策划与复盘。流程：先问主题与材料 → 逐页大纲（页型 + 大中小图位）→ 写 content.json → 生成 → lint/spec_audit 自检 → 渲染逐页看图 → 交 PPTX、PDF、Word 讲稿。用户提到开题PPT、答辩PPT、组会汇报、文献汇报、科研汇报、学术报告、蓝白答辩风格时使用；大创/大赛/项目申报答辩且给了自己的模板时改用 dachuang-ppt。
license: MIT
metadata:
  version: 2.0.0
  default_body_pt: 18
---

# 科研场景 PPT（蓝白信息图式）

目标：任何新主题都做出同一套蓝白信息图式、原生可编辑的 PPTX。**版式和视觉由脚本保证，内容永远来自用户材料。**

本文件只放骨架、硬规则和指针；每一节的完整细则（含实测参数与踩坑）在 `references/` 里，**做到哪一步就读哪一份，不要凭印象**。

## 〇、运行约定（任何 agent 通用）

- `$S` = 本技能目录；`$PY` = 装好依赖的 Python（见 `INSTALL.md`，最少 `python-pptx Pillow lxml python-docx pymupdf`，渲染要 LibreOffice）。下文命令都按这两个变量写。
- **提问**：有提问/澄清工具就用，没有就直接文字问；一次问齐，不挤牙膏。
- **看图**：文中的 “vision / 看图” 指用 agent 的读图能力打开 PNG 逐页看；没有读图能力时把渲染图发给用户确认。
- **生图通道**：用户指定的是**网页端 ChatGPT**（不用本机生图模型、不用 agent 内置生图工具）。agent 能操作浏览器就自己发需求、取图；不能就把每张图的提示词写进 `images/*.prompt.md` 交给用户，先用占位框排版。细则见 [visuals-rules.md](references/visuals-rules.md)。
- **交付文件**：用所在 agent 的发送文件/附件方式把文件直接给用户，不只报路径。
- **配套技能都是可选的**：文中提到的 `chatgpt-web-planning`、`humanizer`、`sci-ppt`、`biz-ppt-skill` 等未安装时，按本技能里的手工做法完成，不要去找。

## 一、先选对技能

| 场合 / 手上有什么 | 用哪套 |
|---|---|
| 科研汇报、开题、中期、组会、文献汇报、学术报告；或只有校徽 / VI 可取 | **本技能** |
| 大创、大赛、项目申报答辩，用户给了自己的成品模板且要 20 页内 | `dachuang-ppt` |
| 一张（或几页）参考版式图要 1:1 复刻成可编辑组件 | `sci-ppt`（若已安装） |
| 商业路演风视觉 | 商业风 PPT 技能（若已安装），不在本技能范围 |

科研答辩即使给了模板也走本技能（做法见 [custom-template-deck.md](references/custom-template-deck.md)）；只有大赛 / 申报答辩口径才交给 `dachuang-ppt`。

**规则冲突时的取舍顺序**：用户本次明确要求 > 内容真实（不编数据、不写材料不支持的结论）> 可读（字号、不溢出、图够大）> 本技能视觉规则 > 美观。

## 二、五条原则（每页都要过）

1. **一页只讲清一件事**。先写出这页的一句话判断（标题带）；引导句交代背景，总结条交代落脚点。出现第二个判断就拆页。
2. **信息密、留白少，但一眼能看清**。空白用图、小标题+要点、流程条、表格填，**不靠放大字号或装饰填空**。
3. **一句话说不清就分点**。有并列、先后、条件就拆成 ■ 要点，每点一行。
4. **图分大、中、小三级，大纲阶段就定**，并算出净图位宽×高与比例，按图位生图，不靠大幅裁切硬塞。
5. **图要具体**：论文原图 / 真实数据图 > 平台截图 > 设备、场景、样品 > GPT 画的具体对象 > 抽象概念图。每页至少一张图或一个图形组件。

文献汇报的整图保留、整篇框架图（必做）、SI 图取舍、结果图「怎么读 / 看到什么 / 说明什么」三段解读与落地版式 → [scenarios.md](references/scenarios.md)。

## 三、推荐页型（大纲每页从这里选）

样子看 `assets/page-types.jpg`（可直接发给网页端 GPT），骨架在 `assets/page-types.json`（按编号取 `skeleton`，只换文字和图）。

| 页型 | 用在哪 | 图位 |
|---|---|---|
| P01 背景需求 | 研究背景、现实需求 | 中图 2 + 小图 1 |
| P02 双栏对比 | 两个瓶颈、前后对比 | 中图 2 |
| P03 方案优缺点 | 两种方法的选择 | 小图 2 |
| P04 大图解读 | 论文框架、技术总图、关键 Fig | 大图 1（`role: big`） |
| P05 流程步骤 | 研究流程、实施过程 | 小图 4 |
| P06 技术路线矩阵 | 数据层 / 方法层 / 验证层 | 小图 9 |
| P07 分层架构 | 系统 / 平台架构 | 无 |
| P08 四环闭环 | 循环、迭代机制 | 小图 4 |
| P09 三栏并列 | 三个方向、三项任务 | 中图 3 |
| P10 对比表 | 多维对比、效益汇总 | 小图 4 |
| P11 时间轴计划 | 研究计划、阶段安排 | 无 |
| P12 结论行 | 主要结论、创新点、预期成果 | 小图 3 |
| P13 问题拆解 | 当前卡点、讨论重点 | 无 |
| P14 进展总表 | 组会进展 | 无 |
| P15 论文信息卡 | 文献汇报出发点 | 小图 1 |
| P16 分点结论 | 文献 / 结果三条结论 | 中图 3 |
| P17 年代流向图 | 领域十年脉络 | 小图 8（`scripts/timeline_cards.py` 自绘） |

典型骨架：
- **开题 / 答辩**：P01 → P02 → P05/P06 → P07/P08 → P04（核心结果）→ P09/P10 → P11 → P12。
- **文献汇报**：P15 → 背景（P02）→ P04（论文框架）→ P16 → P04（重点图解读）→ 创新与局限（`info_split`）→ 评价。精读页序见 [paper-close-reading.md](references/paper-close-reading.md)。
- **组会进展**：P14 → P09 → P13 → P11 → 结束页。
- **5–15 分钟短时开题（约 11 页）**、竞赛 / 创业答辩、参考稿只借页眉 → [scenarios.md](references/scenarios.md)。

## 四、工作流

### 1. 问主题、要材料（不跳过）

- 没说主题先问主题；拿到主题后**一次性**列出缺的：场合与听众、时长或页数、核心材料（论文、原图、数据、公式）、署名（汇报人、导师、单位 logo、日期）、必讲 / 必回避。缺的身份信息用明确空白占位，不编。
- 用户给了**模板文件**时，先按措辞判断用途（判错就整份返工，拿不准就问）：只取校徽 / 视觉参考 / 整份母版套用 / 只取 VI 配色。
- 用户给的是**他自己做完的 PPT** 要重排：先一次问清配色、图片范围、页数目标 → [restyle-existing-deck.md](references/restyle-existing-deck.md)。
- 主题涉及“今年 / 最新”的名单、奖项、政策、数字：先用官方源核实。
- 细则（开题 Word + 学校模板 PDF、分享链接取全文、四种模板用途的判据）→ [intake-outline.md](references/intake-outline.md)。

### 2. 逐页大纲与去重

- 总页数默认含封面、目录、结束页。**只给时长时按时长定页**：答辩 / 评审 10 分钟 15–16 页、15 分钟 18–20 页、30 分钟 28–30 页；讲课 / 组会分享 30 分钟 20–24 页。讲稿按 240–260 字/分折算。
- 用 [outline.md](references/outline.md) 的提示词向网页端 GPT 要大纲（附 `assets/page-types.jpg`）；没有网页端就按页型表自己写。
- 大纲表头：页码｜页型｜页标题｜标题带（一句话判断）｜引导句｜分点内容｜图位（级别 + 净宽×高 + 内容 + 来源）｜总结条。
- 拿回大纲先核对：页型存在、各页任务不重复、图位与密度匹配、大图只在 P04、小图不含要读的数字。同一证据链不要换页型讲两遍。
- 结构：封面 → 目录 / 汇报安排 → 内容页（编号 `1.1`…）→ 结束页。同一页型不连续超过 2 页，20 页级稿子至少 8 种页型。
- 删页 / 合并有歧义时给 2–4 个具体选项让用户选。第二次提要求先分清“改上一版”还是“重做”。细则 → [intake-outline.md](references/intake-outline.md)。

### 3. 配图

- 分工：概念 / 流程 / 机制 / 架构 → 网页端 GPT；定量对比 → 本机 matplotlib（图注写“示意数据”）；必须精确到分钟的时点链、≥6:1 页底超宽条 → 本机 PIL 自绘。**其余默认交 GPT，不要自己画。**
- 提示词写“这张图说明什么 + 画哪些对象和关系 + 净图位比例 + 逐字标签 + 字量上限”，**不替 GPT 定画风**。
- 取回后逐张核对主题、标签（对回源材料）、箭头语义、风格是否与全稿一致；错一处只发一条定点返工指令。
- 贴图前 `trim_margins.py` 裁白边（保留完整画布）；插画多时转 JPG 压体积。
- **GPT 图还没拿到时直接留空图位**（图位只写 `label` + `prompt`，或用 `prep_figs.py` 降级为占位框）：流程、结构、机理、架构这类示意图**本机一律不画**——PIL、matplotlib、原生形状拼图都不行，也不为了“没有空框”改成别的组件凑数。交付回复里列出仍空着的页码和对应的 `images/*.prompt.md`，由用户补图。
- 图注要解释图例；化学式、指标的下标不能省。完整规则 → [visuals-rules.md](references/visuals-rules.md)、[visuals.md](references/visuals.md)。

### 4. 生成 → 检查 → 渲染 → 修

```bash
$PY $S/scripts/prep_figs.py   content.json content.ready.json      # 缺图 → 占位框
$PY $S/scripts/dense_deck.py  --content content.ready.json --output out/slides.pptx
$PY $S/scripts/lint.py        --input out/slides.pptx
$PY $S/scripts/spec_audit.py  content.json out/slides.pptx "稿名"   # 必须“全部通过”
$PY $S/scripts/render.py      --input out/slides.pptx --directory out/render [--slides 3 5]
```

- **固定构建顺序**（写进 `build.sh`，每次重建全跑）：`prep_figs` → `dense_deck` → 自定义原生图形脚本 → `paper_cover.py` → `add_cover_logo.py` / `add_footer.py` → `spec_audit` → `lint` → `render` → `apply_notes` → `notes_docx` → 导出 PDF → 拷进交付目录。
- 生成日志里的 `over-full` / `has empty` / “太小看不清” / “高亮词不在正文里”先清零；lint 错误清零、警告逐条处理。
- **然后逐页看渲染图**——lint 查不出空块、图太小、标签被挤。改完只重渲改动页。
- **用户已在交付件上手动调过版**：不重跑生成脚本，用 `revision_audit.py` 比对后只改点名页。
- 细则（页脚、结束页字段、程序化填字、重建后哪些文件过期）→ [build-verify-deliver.md](references/build-verify-deliver.md)、[auto-generation.md](references/auto-generation.md)。

### 5. 交付

- **默认只交三份单独文件：PPTX、PDF、Word 讲稿。** 不默认附 ZIP、总览图、说明文档、审核报告、数据源码。
- 讲稿单一来源：写 `SCRIPT` → `apply_notes.py` 写进每页备注 → `notes_docx.py` 从备注导出 Word；不另写一份 docx 正文。规则见 [speaker-notes.md](references/speaker-notes.md)。
- 三份一起验收：`$PY $S/scripts/verify_delivery.py --pptx … --pdf … --docx …`。
- 完成标准是已生成并验证的文件，不是“将交付 N 页”的承诺。还缺的真实素材在对话里列出，不写在页面上。

## 五、怎么写一页

从 `assets/page-types.json` 取骨架，只换文字和图。页面骨架：`title` → `lead` → 可选 `sublead` / `band` / `process` → `layout` + `regions` → `summary`。组件与 22 种版式的参数见 [info-style.md](references/info-style.md)；槽位名、JSON 写法示例、文献页标题、机构标识与封面 logo → [page-writing.md](references/page-writing.md)。

## 六、视觉硬规则（速查）

- **字号：默认 18 档，已写进代码**——正文 / 要点 18、小标题 20、说明块 16、图注 14、表格 14 / 16、标题 32（不动）。用户要“大一号”用 `PPT_FONT_TIER=20`（或 21 / 22），不要只改 `para`；正文放不下就压条目或拆页，**不把字号退回 15**；课程作业硬性要求 ≥24 pt 时照 [course-assignment-24pt.md](references/course-assignment-24pt.md)。
- **配色**：深蓝 `#073DA5`、正文 `#111722`、红 `#DE1515`、细边框 `#9CB3D3`、浅蓝 `#EAF0FC`，白底。主色只当骨架，分区浅底可用 4–5 种浅色轮换。换校色加一档 `THEMES`（`PPT_THEME`），不手改色号。
- **强调只用红色加粗文字**：一页最多 4 处（引导句 + 表格高亮 + 总结条合计）；高亮词必须是正文原样子串。**不用红框块、浅红底块。**
- **不用任何竖条**：不加左侧竖条、不给某一行加粗竖条“表示重点”、不做“左竖条 + 大号高亮数字”的 KPI 卡（大数字块已在代码里停用）；关键数字写进句子或表格。
- **一个分区只一条小标题条**；三栏 / 四格栏首用 `section(head, head_align:"center", items:[])`，不和 `region.head` 重复。
- **字数**：页标题 ≤ 14 字、引导句与总结条各一行 ≤ 40 字、标题带 ≤ 26 字、窄栏要点一行说完。
- **表格**全部居中，单位写进表头；不让表格单独撑一页、下面空半页。
- **禁用**：大横框里放一句话的三条横栏版、纯文字三栏装饰堆叠、“三大浅蓝色块 + 空心标签框”。
- 全文与实测参数 → [visual-rules.md](references/visual-rules.md)。

## 七、语言与内容红线

- **学术语体、名词化判断**：页标题、小标题、引导句、结论条用学科通用表述；不写口播句（“说白了 / 要立的是 / 这就是”）、“不是 A 而是 B”、三项排比、口号式断句、段尾升华，不堆“真正 / 核心 / 闭环 / 赋能”，不用“证据链 / 边界”这类元话语，不在页面上写“见第 N 页”。改页面也要同步改讲稿。
- **去 AI 味之后要有故事线**：每页开头交代为什么讲这页，结尾落成一句带走的话，页与页之间一句话串起来。
- **成品稿不出现**：AI 免责（“AI 生成示意”）、来源追溯（“依据材料 A §五”）、改稿口气（“此处插入 / 待补”）、自我缩限。开题还没有数据时不放示意数据图表，预期成果写成计划口吻。
- **分组名不预设结论**（“过量组 / 最优组”→“较高添加量 / 候选区间”）；未测的指标写“拟测”。
- 细则 → [speaker-notes.md](references/speaker-notes.md)「去 AI 味」、[scenarios.md](references/scenarios.md)。

## 八、交付前自检（摘要，全文见 [qa-checklist.md](references/qa-checklist.md)）

- 生成日志无 `over-full`；lint 0 错误；`spec_audit` 全部通过；渲染图逐页看过（留白、孤行、图够不够大、红字落点）。
- 每页一句话判断清楚、说不清的已分点；大图实测 ≥7.5 in 宽或 ≥4 in 高（用 python-pptx 读图片 shape，不只信图位框）；需要读标签的插画 ≥3.5 in 宽；仍空着的图位已在交付回复里列清（页码 + 提示词文件），没有本机自绘的示意图。
- 结果图页三段解读齐全、数字对得上原文；无重复小标题、无竖条、无红框块、红字 ≤4。
- 页面与讲稿都过了一遍去 AI 味；改页后备注与 Word 讲稿已重导、三份文件同源。
- **自动检查只是结构门槛，不能代替按本技能逐项验收。** 每份新稿开工前重读 [visual-rules.md](references/visual-rules.md) 的「图的分级」「小标题去重与对齐」和 [page-writing.md](references/page-writing.md) 的「学术语体」。

## 九、文件索引

**references/**（按工作流顺序）

| 文件 | 什么时候读 |
|---|---|
| [intake-outline.md](references/intake-outline.md) | 问材料、判断模板用途、定页数、出大纲 |
| [outline.md](references/outline.md) | 给网页端 GPT 的大纲提示词、大纲核对 |
| [scenarios.md](references/scenarios.md) | 文献汇报总则、短时开题、竞赛、参考稿页眉、故事线 |
| [visuals-rules.md](references/visuals-rules.md) / [visuals.md](references/visuals.md) | 配图分工、提示词、图位、裁切、压图 |
| [page-writing.md](references/page-writing.md) | JSON 写法、文献页标题、机构标识、组件速查 |
| [info-style.md](references/info-style.md) | 全部组件参数、22 种版式、“要讲的事 → 版式”对照 |
| [visual-rules.md](references/visual-rules.md) | 视觉规则全文、字号档位、图的分级 |
| [build-verify-deliver.md](references/build-verify-deliver.md) | 构建顺序、用户改稿后的再修改、交付 |
| [qa-checklist.md](references/qa-checklist.md) | 交付前自检全文 |
| [pitfalls.md](references/pitfalls.md) | 踩坑清单（报错、溢出、版式问题先查这里） |
| [speaker-notes.md](references/speaker-notes.md) | 讲稿规则、去 AI 味 |
| [paper-close-reading.md](references/paper-close-reading.md) | 文献精读页序、PDF 取图、数字复算、扩页挖料 |
| [competition-pitch.md](references/competition-pitch.md) | 竞赛框架、评分映射、匿名规则、竞赛语体 |
| [custom-template-deck.md](references/custom-template-deck.md) | 整份套用户模板、只取校徽、重建式套版 |
| [restyle-existing-deck.md](references/restyle-existing-deck.md) | 重排用户自己做完的 PPT |
| [course-assignment-24pt.md](references/course-assignment-24pt.md) | 课程作业硬性 ≥24 pt |
| [auto-generation.md](references/auto-generation.md) | 程序 / 模型批量填字、槽位名对照 |
| [embedded-video.md](references/embedded-video.md) | 内嵌视频 |
| [templates-video-styles.md](references/templates-video-styles.md) | 出空白模板、配色档位、加新风格 |

**scripts/**：`dense_deck.py` 生成；`info_style.py` 版式、组件与字号档位；`theme.py` 配色档位；`lint.py` / `spec_audit.py` 检查；`render.py` 渲染；`prep_figs.py` 缺图占位；`trim_margins.py` 裁白边；`prep_gpt_figs.py` 整理 GPT 插画；`insert_image.py` 填图；`paper_cover.py` 文献封面；`add_cover_logo.py` / `add_footer.py` 后处理；`timeline_cards.py` 年代流向图；`apply_notes.py` / `notes_docx.py` 讲稿；`verify_delivery.py` 交付验收；`revision_audit.py` 改稿比对；`embed_video.py` 视频；`omml.py` / `fix_subscripts.py` 公式与下标；`set_font.py` 统一字体；`delivery_doc.py` md 转 docx。

**assets/**：`info-demo.json`（28 页开题 / 答辩示例）、`scenes-demo.json`（19 页组会 / 文献 / 竞选 / 活动示例）——两份示例在 18 档下生成与 lint 全部清零，但早于 `spec_audit`，图片尺寸下限、栏首写法等仍有多条不合规，**只当组件与版式的写法参考，不当验收样板**；新稿以 `page-types.json` 骨架起步、以 `spec_audit` 为准；`page-types.json` / `page-types.jpg`（页型骨架与总览）、`images/`（示例图，`make_demo_figs.py` 可重画）、`components/`（可复用图标）、`branding/`（示例机构标识）。
