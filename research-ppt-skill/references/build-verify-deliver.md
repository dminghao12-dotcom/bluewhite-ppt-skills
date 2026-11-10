# 生成、检查、交付细则

> 本文件由原 SKILL.md 正文迁出（内容未删减），主文件只保留摘要与指针。

#### 4. 写内容 JSON → 生成 → 检查 → 渲染 → 修
```bash
PY=python    # 装好依赖的环境（见 INSTALL.md）
S=<本技能目录>
$PY $S/scripts/prep_figs.py   内容.json 内容.ready.json          # 缺图 → 占位框
$PY $S/scripts/dense_deck.py  --content 内容.ready.json --output out/汇报.pptx
$PY $S/scripts/lint.py        --input out/汇报.pptx
$PY $S/scripts/render.py      --input out/汇报.pptx --directory out/render [--slides 3 5]
$PY $S/scripts/notes_docx.py  --input out/汇报.pptx --output out/讲稿.docx
```
- 生成日志里的“超出（over-full）”“内容下方空了（has empty）”“只分到…太小看不清”“高亮词不在正文里”先清零（日志标出第几页、哪个分区），再跑 lint；lint 错误必须清零，警告逐条处理。

**固定构建顺序（顺序错会互相覆盖）**：`prep_figs` → `dense_deck` → 自定义原生图形脚本（如 `custom_diagrams.py`）→ `paper_cover.py`（文献首页）→ `spec_audit` → `lint` → `render` → `apply_notes`（把 `SCRIPT` 写进每页备注）→ `notes_docx` → 导出 PDF → 拷进交付目录。自定义脚本与 `paper_cover` 都必须在 `dense_deck` 之后，每次重建都要重跑；把它们写进 `build.sh`，不要靠记忆。

**让程序或模型填字（网站生成、批量出稿、worker 跑任务）时先读 [auto-generation.md](auto-generation.md)**：版面归代码（骨架里写死 layout / 槽位 / 条目数），模型只填文字；槽位名、每栏密度、竖图该放哪个槽、推理模型的 token 预算都在那份文件里。人手写 `content.json` 的流程与本条无关，不必读。

**页脚与结束页的补充信息靠后处理脚本，不进内容 JSON**：`meta.info_footer: true` 只画一条深色带 + 页码（y=7.24、高 0.26），**不画任何文字**；要写“姓名｜单位｜会议名”这类页脚，用 `scripts/add_footer.py out/slides.pptx --text "姓名｜单位｜会议名"`（默认在 y=7.255 写 10.5 pt 白字，跳过封面与结束页；没开 `info_footer` 底带时加 `--color 111722`），它同样必须排在 `dense_deck` 之后并进 `build.sh`——重建会直接把它冲掉。结束页（`type: end`）只渲染 `title` 与 `points`，**`subtitle` / `footer` 字段被静默忽略**；写 `points` 时页面顶部那行小标题默认是“回到本次汇报的核心”，培训 / 讲课稿用 `"heading": "本课可以带走的三句话"` 覆盖它，否则一份讲义里会突然出现“汇报”二字；汇报人、邮箱电话要另加文本框，坐标得避开第三条要点的文字框（实测底边 5.39 in）与底部致谢条（自 6.13 in 起）：y=5.45、高 0.62、12.5 pt 两行正好落在两者之间，`lint` 才不会报“与 汇报人信息 重叠”。

**重建之后，交付目录里的三份文件全部过期**——`dense_deck` 只重写 `out/slides.pptx`，讲稿备注、Word 讲稿、PDF 都还是上一版的内容。改完最后一处内容，必须把 `apply_notes` → `notes_docx` → `soffice` 转 PDF → `cp` 进交付目录这一串**在同一条命令里跑完**再回复用户；只跑 `dense_deck` 就宣布“已更新”，用户拿到的是新旧混装的稿子。

**交付前必须跑 `scripts/spec_audit.py content.json out/slides.pptx "稿名"` 并且输出“全部通过”**。它把本文件里的硬规则（标题 ≤14 字、引导句 ≤40、标题带 ≤26、一页高亮 ≤4、版式 ≥8 种且不连续超 2 页、图片不被压到 2.6 in 以下、`info_three`／`info_four`／`info_quad` 等三／四栏页的图片宽度 ≥3.5 in、三／四栏页不得用 `region.head` 当栏首（脚本不读 `head_align`，栏首不会居中）、页面不得出现硬口播句（“要立的是／说白了／站不住／说到哪里为止”；“这就是／坑／伙伴／闭环／赋能”只作提示，避免误伤“坑塘”“闭环控制”这类术语）、每页有图或组件、红线词、编号连续性）逐条量化检查，并报出页号和实际数值。它统计“版式 ≥8 种”时只数非空 `layout`，封面与结束页的 `None` 不算——16 页短稿也必须真的有 8 种内容版式。**9 页左右的短稿最多只凑得出 7 种（内容页 = 总页数 − 封面 − 结束页），按实际页数取上限即可，不必为凑 8 种拆页**；这条只放宽版式数，其余检查仍须全部通过，并在交付说明里写清这一条为何不适用（`spec_audit.py` 已按内容页数取上限，短稿通过时会打印上限说明）。这是“严格执行”的证据：不要只凭“我觉得差不多了”就交付。**新增或收紧一条检查后，把最近已交付的稿子也重跑一遍**——旧稿通常本来就违规（实测两份都已交付的稿子一回扫就各出了十几条），而用户会拿新旧两份互相比对；连同规则一起修好再报“全部通过”。
- **然后逐页看渲染图**——lint 查不出“分区里空了一块”“图在格子里太小”“标签被挤”，只能看图发现。改完只重渲改动页。
- `render.py` 在 Windows 用 PowerPoint，其他系统用 LibreOffice（替代字体，断行略有差别）；终稿以 PowerPoint 为准。

**交付后的再修改（用户常自己在 PPT 里调图的大小、字号）**：不要重跑生成脚本，那会把用户的调整打回原样。

1. 找出用户改过的那份（按修改时间与图片坐标比对，或直接问他），拷成 `User_Edited_Base.pptx`。
2. `python scripts/revision_audit.py 底稿.pptx 用户改过的.pptx` 看清他改了什么页：`images` 图片坐标/尺寸、`tables` 表格字号、`texts` 文本框的文字/字号/颜色、`boxes` 文本框位置、`shapes` 增删的原生形状。只报 `changed` 不算看清：用户在目录页把标题 20→28 pt、把 ■ 换成 ✓、删提一条竖条，都属于这一类，必须逐项读出来再决定后稿要不要跟着改。状态为 `metadata-only` 的页只是 WPS/PowerPoint 重存的序列化差异，不算用户改动。
3. 用 python-pptx 只改他点名的页，另存新版本；不要覆盖用户底稿。
4. 改完跑 `python scripts/revision_audit.py 用户改过的.pptx 新版本.pptx --expect 3,16`：被点名页之外出现改动就报 FAIL。
5. 只重渲改动页再逐页看图；用户已放大的图不回退，lint 的名义重叠警告先看图确认无实际遮挡。

#### 5. 交付
完成标准是已生成并验证的文件，不是原图提取、大纲请求或讲稿计划。只完成准备工作时，明确报告未完成及实际阻塞，不把“将交付18页”等承诺当成交付；终稿回复必须附可打开的成果文件。

**默认对外交付仅三份文件（用户明确要求）：PDF、PPTX、Word 演讲稿。** 三份文件分别提供，PDF承担全稿预览，Word演讲稿与PPT逐页备注同源（每页写 `notes`，规则见 [speaker-notes.md](speaker-notes.md)：精炼、口语、该科普的讲清楚）。

**交付回复要把文件直接挂进聊天，不能只报路径**：用所在 agent 的发送文件方式（Hermes 桌面端是回复里一行一个 `MEDIA:/绝对路径`，Claude 等用发送文件／附件工具），用户点开就能拿到；只写“已放在某某目录”会被回一句“交付给我啊”。写之前 `ls` 确认文件存在、且是从生成原件拷出来的那一份。

**不要默认附加** ZIP压缩包、总览PNG、交付说明MD/DOCX、审核报告、数据JSON、原始材料、源码或其他文件；这些制作与验收材料可保存在项目目录供内部核查。只有用户本次明确索要某项额外材料时才增加交付。最终回复也围绕这三份成品，不把内部过程材料一并发送。

**讲课 / 汇报类要 Word 讲稿时，备注与 Word 只维护一份源**：写 `SCRIPT = [(“第 N 页 · 小标题”, “正文…”), …]`，先跑 `scripts/apply_notes.py out/slides.pptx` 把同一份写进每页备注（脚本自带 `assert len(prs.slides) == len(SCRIPT)`，并打印总字数与估算时长），再由 `notes_docx.py` 从 pptx 备注导出 Word——**不要另写一份 docx 正文**。两步都进 `build.sh`（`dense_deck` 之后、`lint` 之前）。分别手写两处必然对不上——实测同一份稿子备注 1320 字（≈6 分钟）、Word 2650 字（≈12 分钟），学员按备注讲会讲不完。docx 正文按空行 `split('\n\n')` 逐段加首行缩进，否则只有第一段缩进、后面贴左边。

还缺的真实素材在对话里列出，**不写在页面上**；示意图类图位拿不到 GPT 图时留空交付，并在回复里列出待补页码与提示词文件（不本机自绘、不换组件凑数）；数据图、论文原图这类真实素材缺失时同样留空并列清，不用示意数据冒充。**用户明确要求留给自己填充的人员、顾问或媒体页属于授权预留版**：按指定数量制作原生照片框和独立文本框，使用“姓名／职称”“报道标题”等中性字段，不补造人物、院士身份、合作、推荐或报道；交付时说明哪些页仍需用户填写。未提供视频原文件时只展示真实截图，不能把照片轮播称为测试视频。需要正式说明文档时，md 用 `scripts/delivery_doc.py` 转 docx。

**交付说明里单列一节「数字口径」**：写清哪些数字是从表里推算的（例如由百分比 × 单元格数反推的判断次数）、原文哪几处自相矛盾（图上的星号 vs 正文 p 值、index 与 SE 不自洽），以及据此在页面上做了哪种取舍。评审追问时直接指这一节回答，不用回头翻原文；这也是交付说明除了“文件清单”之外唯一值得写的正文。

PDF为默认必交件，用 `soffice --headless --convert-to pdf --outdir out out/slides.pptx` 从同一份 PPTX 导出，并用 `pymupdf` 核页数与页面尺寸；PDF 里中文会回退到 Hiragino Sans GB / ST Heiti，屏幕上看正常即可，公式必须写成 Unicode 线性式（见踩坑清单里 OMML 那条）。PPTX 与 PDF 都从生成原件拷一份，不要让 WPS / PowerPoint 重存后再交。

**三份文件一起验收，跑 `scripts/verify_delivery.py --pptx … --pdf … --docx …`**：它核 PPTX 页数 == PDF 页数、页面 16:9、每页备注非空、Word 里能按顺序找到每页备注首句（即“备注与讲稿同源”）、讲稿字数与估算时长，并打印三份文件的 sha256 前 12 位。报 FAIL 就回到 `apply_notes` → `notes_docx` → 转 PDF 那条链重跑，不要手改 Word 或备注去凑一致。

**不因“方便转发”自动打包或增加附件。** 默认仍只交PDF、PPTX、Word演讲稿三个单独文件。只有用户明确要求ZIP或资料包时才打包，并按其指定范围收录；压缩包内文件名用ASCII，避免Windows解压乱码。论文原文、联系表、交付说明和数据底稿不默认塞入资料包。
