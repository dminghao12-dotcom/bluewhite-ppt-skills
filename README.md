# PPT 技能包 2.0（research-ppt-skill + dachuang-ppt）

两个独立的 SKILL.md 技能，任何支持技能目录的 agent 都能用（Claude Code、Claude 应用、Hermes、Codex 等）。安装与依赖见各自目录下的 `INSTALL.md`。

| 技能 | 用在哪 |
|---|---|
| `research-ppt-skill/` | 开题、中期、答辩、组会、文献汇报、学术报告：蓝白信息图式，写 content.json 由脚本生成 |
| `dachuang-ppt/` | 大创、挑战杯、互联网+、项目申报答辩：在用户模板上增量改，20 页内 |

## 2.0 改了什么

**通用**
- 去掉 Hermes 专属写法：`~/.venvs/ppt`、`~/.hermes/skills/...`、`~/Documents/Hermes成果`、`clarify` / `write_file` / `web_extract` / `MEDIA:` 等全部改成 `$PY` / `$S` / 通用描述；两份 SKILL.md 开头新增「运行约定」，说明提问、看图、生图通道、交付文件在不同 agent 下怎么做，以及配套技能缺失时的手工做法。
- frontmatter 规范成 `name / description / license / metadata`，description 写清触发场景与分流。
- 两个技能各补一份 `INSTALL.md`（原 README 引用但缺失）。
- **正文默认 18 pt**，已写进代码，不需要额外设置。

**research-ppt-skill**
- SKILL.md 从 605 行 / 159 KB 压到 188 行 / 19 KB；原正文**无删减**迁入 9 份新 references（scenarios、intake-outline、visuals-rules、build-verify-deliver、page-writing、visual-rules、pitfalls、qa-checklist、templates-video-styles），主文件只留骨架、硬规则和“什么时候读哪份”的索引。
- `info_style.py` 新增 `FONT_TIERS`：默认 18 档（正文 18 / 小标题 20 / 说明 16 / 图注 14），`PPT_FONT_TIER=20|21|22|legacy` 换档。时间轴阶段框高度随字号自适应。
- `note(tone="red")` 与 `callout` 不再画红框浅红底（规则早已否掉，代码没跟上），改为浅蓝底细边框 + 红色关键词。
- 补上缺失的 `scripts/add_footer.py`、README 预览图 `docs/`；示例里引用了却从未发布的 4 张流程示意图（`p21–p24.jpg`）改成留空图位，不本机自绘。
- 配图规则改为：GPT 图没拿到就留空图位并在交付回复里列出，流程 / 结构 / 机理示意图本机一律不画。
- 两份示例按 18 档调整措辞与组件：生成日志与 lint 都是 0 错误 0 警告；去掉已停用的 `head_style: "bar"` 与红框强调。
- 修正文档矛盾：自定义模板文档里的“左端竖条”、代码注释推荐红框、踩坑表两行粘连。

**dachuang-ppt**
- `lib.py` 新增字号阶梯 `FS`（12 / 14 / 18 / 19 / 20 / 21 / 26 / 72）；`bullets()` 默认按规则出 ➤ 蓝箭头、1.40 行距、0.30 in 悬挂缩进，支持 `{"label", "text"}` 标签自动大一号。
- `deck-starter` 拷到任意目录都能跑（`DACHUANG_SKILL_DIR` 或常见技能目录自动找模板）；结束页按规则整块重绘，封面单位行删掉重画；页面按 18 pt 重排。
- 默认模板里残留的原项目人名、项目名、单位年份清成占位。
- 写死原项目文字、无法运行的 `build2.py` / `tpluse.py` / `shrink.py` 移到 `examples/project-reference/`，与通用脚本分开。

## 已知限制
- `research-ppt-skill` 的两份示例早于 `spec_audit`，仍有图片尺寸下限、栏首写法等不合规项（原版即如此，本次未整体返工）。它们只当写法参考，新稿以 `page-types.json` 和 `spec_audit` 为准。
- 渲染以 LibreOffice 为准，字体回退时断行与 PowerPoint 略有差别。
