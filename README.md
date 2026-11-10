<div align="center">

# 蓝白学术风 PPT 技能包

**把论文、开题报告、申报书丢给 AI，拿回一份能直接上台的 PPT。**

两套技能，一套管学术汇报，一套管比赛与项目答辩。

</div>

---

## 学术场景：`research-ppt-skill`

开题、中期、毕业答辩、组会进展、文献汇报、学术报告。

![research-ppt-skill 生成效果](docs/images/research-scenes-overview.jpg)

- **一页只讲一件事**：页顶一句判断，中间是图、表和分点，页底一句结论，评委扫一眼就知道这页想说什么。
- **页型丰富不重样**：进展总表、问题拆解、论文框架、结果大图解读、对比表、时间轴计划等 17 种推荐页型、22 种版式，按内容自动搭配。
- **图文比例拿捏好**：论文原图尽量完整放大，结果图配“怎么读 / 看到什么 / 说明什么”三段解读。
- **学术语体**：标题和结论用名词化表述，没有口号、排比和“AI 味”套话，讲稿同步按同一口径写。

## 比赛与项目答辩：`dachuang-ppt`

大创、挑战杯、互联网+、学科竞赛、校级 / 省级项目申报与结题答辩。

![dachuang-ppt 生成效果](docs/images/dachuang-overview.jpg)

- **套你的模板**：在你给的学校或赛事模板上改，封面、目录、结束页原样保留风格；没有模板就用自带的蓝白模板。
- **短时评审友好**：通常 20 页以内、10 分钟左右，页数按答辩时长定，申报书是唯一内容来源，不编数据。
- **版面规整**：蓝色判断句开头、深蓝结论条收尾，分点带 ➤ 箭头和悬挂缩进，表格、卡片、指标一目了然。

## 两套共同的默认规则

- **正文 18 pt**，投影坐在后排也看得清；放不下就压缩条目或拆页，不缩小字号硬塞。
- **每一页都是 PowerPoint 原生元素**（文字框、表格、形状），不是截图，打开就能改。
- **红色只点关键词**，一页最多 4 处；不用红框色块，也不用“左侧竖条 + 大号数字”的卡片。
- **示意图不乱画**：流程图、机理图没有合适的素材时留出空图位，并写好生图提示词，交给网页端 ChatGPT 等工具生成后再填进去。
- **一次交齐**：PPTX、PDF、逐页对应的 Word 讲稿。

## 怎么用

**1. 放进 AI 助手的技能目录**（两套可以只装一套）

```bash
git clone https://github.com/dminghao12-dotcom/bluewhite-ppt-skills.git
cp -R bluewhite-ppt-skills/research-ppt-skill ~/.claude/skills/
cp -R bluewhite-ppt-skills/dachuang-ppt      ~/.claude/skills/
```

| 你用的是 | 放到哪里 |
|---|---|
| Claude Code | `~/.claude/skills/`（或项目里的 `.claude/skills/`） |
| Claude 网页 / 桌面应用 | 把单个技能文件夹压成 zip，在设置的技能页上传 |
| Hermes Agent | `~/.hermes/skills/productivity/` |
| Codex 等其他支持 `SKILL.md` 的助手 | 放进它的技能目录，或让它先读 `SKILL.md` |

**2. 装依赖**：Python 3.10+，另需 LibreOffice（导出 PDF 和预览用）

```bash
python -m pip install python-pptx python-docx lxml Pillow pymupdf numpy matplotlib
```

**3. 直接跟 AI 助手说**

> 帮我做开题答辩 PPT，讲 15 分钟，开题报告在附件里，学校模板只要用校徽。

> 用这篇论文做一次组会文献汇报，12 页左右，重点讲方法和 Fig. 3。

> 帮我做大创中期答辩，模板和申报书都在附件里，10 分钟。

它会先问清场合、时长、署名和材料，给出逐页大纲（每页写明页型和图位大小），再生成、自检、逐页看渲染图，最后交给你 PPTX、PDF 和 Word 讲稿。

<details>
<summary>不用 AI 助手，直接跑脚本</summary>

```bash
# 学术场景：用自带示例生成
cd research-ppt-skill/assets
python ../scripts/dense_deck.py --content scenes-demo.json --output demo.pptx
python ../scripts/lint.py --input demo.pptx

# 比赛答辩：7 页骨架（回到仓库根目录再进）
cd ../../dachuang-ppt/deck-starter
python -m deck.build2          # 输出 out/样张.pptx
```

写自己的稿子时，复制 `research-ppt-skill/assets/page-types.json` 里的页型骨架，只替换文字和图片。

</details>

---

<div align="center">
<sub>示例页面均为技能包实际生成，内容与数据为演示用；示例中的校徽请换成你自己单位的。安装细节见各技能的 <code>INSTALL.md</code>，更新记录见 <a href="CHANGELOG.md">CHANGELOG.md</a>。MIT 许可。</sub>
</div>
