# deck-starter —— 能直接跑起来的最小骨架

技能包里 `examples/project-reference/build2.py` 是**作者原项目的入口**，它 `from deck import pages1, pages2, tpluse`，
而原项目的 `pages1.py` / `pages2.py` / `notes2.py` 从没随包发布，所以它**不能直接运行**，只作写法参考。

这个目录就是补上的那一半：一套 7 页、能真跑出 pptx 的骨架（字号阶梯与 ➤ 分点已按 18pt 规则内置）。先跑通它，再把文字换成自己项目的。

---

## 跑起来

```bash
# 依赖（任选一种装法；见技能根目录 INSTALL.md）
python -m pip install python-pptx python-docx lxml Pillow pymupdf

# 方式一：直接在技能目录里跑（会自动从 ../assets/ 拷默认模板）
cd <技能目录>/dachuang-ppt/deck-starter
python -m deck.build2
# → out/样张.pptx，7 页，带讲稿备注

# 方式二：拷到自己的项目目录再跑（推荐，别在技能目录里堆项目文件）
cp -R <技能目录>/dachuang-ppt/deck-starter <项目目录>/my-deck
cd <项目目录>/my-deck
DACHUANG_SKILL_DIR=<技能目录>/dachuang-ppt python -m deck.build2
```

找不到模板时，`build2.py` 会依次找：`DACHUANG_SKILL_DIR`、`../assets/`、`~/.claude/skills`、`~/.codex/skills`、`~/.agents/skills`、`~/.hermes/skills[/productivity]` 下的 `dachuang-ppt/assets/default-template.pptx`；都没有就提示你手放 `tpl2/t.pptx`。

想要预览图：

```bash
soffice --headless --convert-to pdf --outdir out out/样张.pptx
# 或用 research-ppt-skill 的渲染脚本（若已安装）：
python <技能目录>/research-ppt-skill/scripts/render.py --input out/样张.pptx --directory out/render
```

实测（2.0 版）：7 页，`saved out/样张.pptx  slides=7`，渲染正常，没有溢出和裁切；结束页按规则整块重绘。

---

## 目录结构（和原项目入口期待的布局一致）

```
deck-starter/
├── deck/                 ← 代码：改内容只动这里
│   ├── __init__.py
│   ├── lib.py            ← 组件库（页眉/卡片/表格/结论条/富文本），从 scripts/ 拷来的
│   ├── tplcore.py        ← 套模板工具：按形状名改文字、保页、重排、修 partname
│   ├── pages1.py         ← 自绘页：研究背景(s03)、技术路线(s05)
│   ├── pages2.py         ← 自绘页：关键指标(s11)、预期成果(s13)
│   └── build2.py         ← 入口：套模板 + 拼页序 + 写讲稿备注
├── tpl2/                 ← 首次运行自动从 ../assets/default-template.pptx 复制出 t.pptx
│   └── t.pptx               （包里不带这份 8 MB 副本，省体积；换模板直接替换这个文件）
└── out/                  ← 生成物，不随包
```

**为什么模板放在 `tpl2/` 而不是 `deck/`**：`build2.py` 里 `os.chdir(ROOT)`（项目根），`TPL = 'tpl2/t.pptx'` 相对项目根解析。这是原作者的约定，照抄就不会踩路径坑。

`tpl2/t.pptx` 不在包里（避免重复带一份 6 MB 模板）：`build2.py` 的 `ensure_tpl()` 缺模板时按上面的顺序自动找来复制。你把 `deck-starter/` 单独拷到别的目录时，要么带上一份模板，要么手放一个 `tpl2/t.pptx`。

---

## 改成自己的稿子（四步）

1. **换模板**：把答辩方给的模板存成 `tpl2/t.pptx`。没有模板就用现在这份（长江大学蓝白壳）。
2. **改封面文字**：`deck/build2.py` 顶部的 `COVER` 字典，键是模板里的**形状名**。不知道名字就先打印：

   ```bash
   $PY -c "
   from pptx import Presentation
   p = Presentation('tpl2/t.pptx')
   for i, s in enumerate(p.slides[0].shapes): print(i, repr(s.name), repr(s.text_frame.text[:30]) if s.has_text_frame else '')
   "
   ```

   名字对不上会打印 `! 跳过：模板里没有形状 'xxx'`，不会崩。
   封面与结束页底部的单位时间改 `UNIT_LINE`（整行删掉重画，不继承模板字距）；结束页改 `END_THANKS` / `END_SUB`（`redraw_end()` 整块重绘：44 / 30 pt + 白色细线，并删掉模板的汇报人胶囊）。
3. **换章节**：`TOC` 列表。模板目录有 8 条，条目数不一致会提示，不会静默丢。
4. **写正文**：改 `pages1.py` / `pages2.py` 里的页函数。模板里的页用 `judgment`(首结论) + `card`/`bullet` + `conclude`(尾结论)，跟技能正文的版面规则一致。

调页数：`build2.py` 末尾的 `SCRIPT` 是讲稿，必须和页数一一对应（有 assert 拦着），加页就同步加一条。

---

## 三处最容易踩的

| 现象 | 原因 | 处理 |
|---|---|---|
| `ModuleNotFoundError: No module named 'deck'` | 在 `deck/` 目录里跑，或漏了 `__init__.py` | 回到 `deck-starter/` 这一级，用 `python -m deck.build2` |
| `缺模板：tpl2/t.pptx` | 模板没拷，或跑到了别的目录 | `cp <skill>/assets/default-template.pptx tpl2/t.pptx` |
| 改完文字版式没变 | `lib.header()` 会找 `brand/logo_blue.png`，找不到就不画校徽（不报错） | 要 logo 就把图片放 `brand/`，否则无视 |
| `verify_deck` 报「PDF 里残留 N 个方括号标记」 | 正文直接写 `【关键词】` 时，只有 `judgment`/`conclude` 会解析成红字；`bullets` 的纯字符串条目不会 | 把那条包一层 `rich('…【词】…')`（`lib.rich` 返回可给 `bullets` 用的 runs） |
| `verify_deck` 报「第 N 页正文红字 7 处，超过 4 处」 | 技能规则是每页红字 2–3 处，最多 4 | 少标几个词：一页只留最该强调的那两三个 |

`renumber_partnames` 在删页/加页后必须跑（`build2.py` 里已经调了），否则 PowerPoint 打开会提示"修复"。

