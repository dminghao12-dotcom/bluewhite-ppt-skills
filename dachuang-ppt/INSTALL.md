# 安装与依赖（任何 agent 通用）

本技能是一个**标准 SKILL.md 技能目录**：`SKILL.md`（给 agent 的操作手册）+ `references/`（按需读取的细则）+ `scripts/`（Python 脚本）+ `assets/`（示例与模板）。不绑定某个 agent。

## 1. 放到 agent 的技能目录

| 环境 | 放哪里 |
|---|---|
| Claude Code | `~/.claude/skills/<技能名>/`（个人），或项目里的 `.claude/skills/<技能名>/` |
| Claude 应用（claude.ai / 桌面端） | 把单个技能文件夹打成 zip，在设置的技能页上传 |
| Hermes Agent | `~/.hermes/skills/productivity/<技能名>/` |
| Codex 及其他支持 SKILL.md 的 agent | 放进它的技能目录（按所用版本的文档），或在会话里让它先读 `SKILL.md` |
| 不用 agent | 自己写 `content.json` / 页函数，直接调 `scripts/` 里的脚本 |

两个技能（`research-ppt-skill`、`dachuang-ppt`）是**两个独立目录**，各自放一份；互相只在“选哪套”时提到对方，缺了另一套也能用。

## 2. Python 依赖（3.10+）

```bash
python -m pip install python-pptx python-docx lxml Pillow pymupdf numpy matplotlib
# 可选
python -m pip install jieba        # lint 断词检查更准
```

- 脚本用 `import pymupdf`，请装 PyMuPDF 1.24 以上（更早的版本只有 `import fitz`）。
- 想隔离环境：`python -m venv .venv && .venv/bin/pip install …`，之后把 `$PY` 设成这个解释器。SKILL.md 里的命令都写成 `$PY $S/scripts/…`：`$PY` = 这个 Python，`$S` = 技能目录。

## 3. 渲染与导出

- **LibreOffice**（`soffice` 在 PATH 里）：导出 PDF、渲染预览图。macOS 也会在 `/Applications/LibreOffice.app` 里自动找。
- **poppler**（`pdftoppm`）：`render.py` 把 PDF 转 PNG 时用；也可以改用 pymupdf 自己转。
- **Windows**：`render.py` 优先调 PowerPoint（需要 `pywin32`）；PowerPoint COM 报 `0x80004005` 时手动另存 PDF。
- **字体**：版式按「微软雅黑」排。Linux / macOS 没有时 LibreOffice 会回退到思源黑体 / 苹方等，断行略有差别；终稿以 PowerPoint 打开为准。

## 4. 字号

两套技能默认**正文 18 pt**（已写进代码，不需要额外设置）：

- `research-ppt-skill`：`scripts/info_style.py` 的 `FONT_TIERS`。换档 `PPT_FONT_TIER=20`（或 21 / 22 / legacy），单项覆盖 `PPT_FONT_SIZES='{"para":18.5}'`。
- `dachuang-ppt`：`scripts/lib.py` 的 `FS` 字号阶梯（图注 12 / 标签 14 / 正文 18 / 分点小标题 19 / 首结论 20 / 尾结论 21 / 页眉 26 / 封面 72）。

## 5. 网页端 ChatGPT（可选）

大纲与概念插画默认交网页端 ChatGPT 做。agent 能操作浏览器（或装了 `chatgpt-web-planning` 之类的配套技能）就自动走；不能就把提示词写成 `images/*.prompt.md` 交给你，先用占位框排版，拿到图再替换。

## 6. 自检是否装好

```bash
# research-ppt-skill：在 assets/ 目录下生成 28 页示例
cd <技能目录>/research-ppt-skill/assets
python ../scripts/dense_deck.py --content info-demo.json --output /tmp/demo.pptx
python ../scripts/lint.py --input /tmp/demo.pptx        # 期望：0 个错误、0 个警告

# dachuang-ppt：7 页骨架
cd <技能目录>/dachuang-ppt/deck-starter
python -m deck.build2                                    # 期望：saved out/样张.pptx  slides=7
```
