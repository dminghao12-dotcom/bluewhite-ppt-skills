# 程序化生成：让模型只填字，版面留在代码里

适用：把「材料 → 可编辑 pptx」做成一条自动链路（网站上点击生成、批量出稿、worker 跑任务），而不是人手写 `content.json`。

## 一、铁律：版面归代码，文字归模型

让模型同时决定「写什么」和「放几栏、每栏几条」必然不收敛：同一页会在「空 28–41%」和「超 1.99 in」之间来回横跳，三轮自修复都不收敛，因为每次改文字都会改变几何，几何又反过来要求改文字。

正确切法：

- `skeleton.py` 里把每页写死：`pid`、`layout`、标题字数上限、每个槽位由几个 `Section(key, head 字数, items 条数, item 字数)` 组成、哪个槽位放图、有没有 `summary`。
- 模型只输出文字：`{pid: {title, lead, summary, notes, sections: {key: {head, items[]}}, kv, table}}`。它看不到也改不了版式。
- 版面确定之后再谈微调——先有稳定几何，才有「这页图再大一点」这种话可说。

## 二、槽位名必须查，不能猜

`INFO_LAYOUTS` 的槽位名不是统一的 `left/right`，写错会在 `dense_deck.py` 里崩在 `layouts.stack(...)`，报错只给 `info_style.py` 的行号，看不出是槽位问题（极易误判成内容或尺寸问题）。常用对照：

| layout | 槽位 |
|---|---|
| `info_split` / `info_compare` / `info_table_fig` | `left`、`right` |
| `info_big` / `info_hero` | `main`、`side` |
| `info_hero_left` | `side`、`main` |
| `info_focus` | `main`、`bottom` |
| `info_stack` | `top`、`bottom` |
| `info_mosaic` | `left`、`top`、`bottom` |
| `info_three` | `a`、`b`、`c`（**不是** `c1/c2/c3`）|
| `info_quad` / `info_four` | `a`、`b`、`c`、`d`（**不是** `q1/q2`）|
| `info_rows3` | `r1`、`r2`、`r3` |
| `info_side3` | `main`、`s1`、`s2`、`s3` |
| `info_full` | `main` |

组件形状同理只以 [info-style.md](info-style.md) 为准。自己拼一个 `{"type": "table", "head": [...], "rows": [...]}` 塞进槽位，崩出来的堆栈落在库内部（`layouts.stack` → `info_style.py:14xx`），仍然看不出是组件问题——**崩在库内部行号时，逐个去掉可疑组件二分定位，不要改库**。

## 三、密度按实测标定（16–18 pt 正文）

- 一个 `section`（栏首 + 3 条要点）≈ 半个 12 行栏 → **一栏要两节**才填满；只放一节 3 条，下半截空 40% 以上。
- `info_compare` / `info_split` 半栏（12 列 ≈ 6 in）：两节各 3 条刚好；再塞一组 `kv` 就超约 0.05 in（一条要点的高度）。
- `info_big` / `info_hero` 侧栏（7 列 ≈ 3.5 in）：三节 × 2 条刚好，第三节压到 1 条才不越界；四块必超。
- `info_three` 三栏（4 in）：两节（3 条 + 2 条）刚好；三节必溢出。
- `info_quad` 四格：每格一节 2 条。

标定方法：先用短占位字跑一遍把几何跑通，再换成真文字；真文字变长后按同一比例减条数，不要靠缩字号。

## 四、先假字后真字

`dummy_fills(scene)` 生成一套占位文案 → `build_spec` → `render_deck`，几十秒就能知道几何是否成立（日志 `over-full` / `has empty` / `spec_audit` 的图片下限）。几何不过关时**不要花模型调用**；几何过关再换成真填字。每改一次骨架都要重跑这一轮。

## 五、版式多样性与审计

`spec_audit` 的「版式种类」按**内容页数取上限**（封面与结束页不算）：9 页稿 = 7 个内容页，要 7 种不同 `layout`。因为不能让模型自选版式，这一步必须在骨架里一次分配好——注意相冲突的两条：

- 同一 layout 不能连续超过 2 页；
- 每页都要有图或组件。

实际可行的一组（9 页文献汇报）：`info_big`（论文首页 3:2 裁片 + 三节侧栏）→ `info_compare`（背景两栏）→ `info_full`（框架三节）→ `info_three`（主要结果三栏）→ `info_focus`（关键图通栏 + 底部两节）→ `info_rows3`（贡献/局限/评价三横栏）→ `info_mosaic`（适用条件）。

## 六、图：先判断图源，再选槽位

1. **能抽位图就抽**（`page.get_images`）；抽不出来（论文图全是矢量路径）→ **让模型报「关键图在第几页」，worker 渲染那一页当图**（`render_page_png(data, page_no, dpi=150)`），不要为了图去改版式或硬裁。
2. **整页渲染是竖版（≈1:1.3）**：放进 `info_big` 的 17 列主区会被高度限制压成 0.7×0.9 in；改放 `info_split` 左栏（竖栏友好）或 `info_focus` 通栏 `main`（宽 ≈3.9 in，过 `spec_audit` 的 2.6 in 下限）。
3. **3:2 论文首页裁片**才适合 `info_big` 主区（宽图）。
4. 同一份稿里出现两张同一张图（首页既做概况页又做结果页）很难看：只剩一张可用图时，结果页改成无图版式（`info_rows3` / `info_full`），不要让它重复出现。

## 七、调模型填字的三个硬参数

- **推理模型要给足预算**：DeepSeek 这类一次填字烧 1–1.2 万 reasoning token，`max_tokens` 低于 8000 会 `finish_reason=length` 且**正文为空**；给 32000，遇空回复自动升到 64000 重试一次。空正文一定是预算问题，不是「模型不会」。
- **提示里给槽位清单**：把 `{pid: {title, lead, summary, sections:{key:{head, items}}}}` 的字数上限、条数上限逐页列出来，并明确「条目可以少、不能多」。
- **要 `missing` 字段**：材料里没有的东西（定理完整表述、参数取值、汇报人姓名）让模型单独列出来，而不是编进页面。这份清单直接进交付说明和前端提示。

## 八、渲染失败怎么回喂

`dense_deck` 的报错原文（含「超出」「内容下方空了」「只分到…太小」）连同当前填字 JSON 一起发回模型，只说一句「版面固定，不能改 layout 或增减页，把被点名分区里的要点缩短」，最多 3 轮。**不要在这一步允许它改版式**——那等于把几何又交回给它。三轮仍不过就是骨架的密度标定错了，回去查第三节。

## 九、交付前的两件必须做的事

1. `spec_audit` 必须打印「全部通过」；短稿的「版式种类」上限说明可以接受，其余不得豁免。
2. **转 PDF、逐页渲 PNG、拼一张 contact sheet 自己看**：图太小、底部横向条被压住、mosaic 左栏留白过多这类问题，只有看图能发现（本次实测就靠拼图发现了「关键图只剩 0.7 in」和「结果页底部文字被切」）。

## 十、服务化时的顺序

worker 取任务（`generation_jobs`）→ 解析（PDF/DOCX → 文本块 + 小标题 + 首页裁片）→ 填字（模型）→ `build_spec` → 渲染（失败回喂 ≤3 轮）→ 成品传对象存储 → 进度与结果回写。提交时扣积分、失败原批退回，`attempt` 记重试次数，不重复扣费。
