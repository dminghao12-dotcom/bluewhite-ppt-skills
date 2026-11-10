#!/usr/bin/env python3
"""比对两版 PPTX，为“交付后用户在 PPT 里手改过再让我局部修”这类轮次做审计。

用法：
    python revision_audit.py 底稿.pptx 新稿.pptx            # 只要报告
    python revision_audit.py 底稿.pptx 新稿.pptx --expect 3,5,16
        # --expect：本轮获授权改动的页（1 起）。其余页若变化会以 FAIL 列出。

输出（JSON）：
    pages[].status   unchanged / changed
    pages[].images   [{left,top,width,height}] 的差异（英寸）
    pages[].tables   表格字号差异（行, 列, 旧, 新）
    pages[].texts    文本框内的 run 变化（文字 / 字号 / 加粗 / 颜色）
    pages[].boxes    文本框位置或尺寸变化（英寸）
    pages[].shapes   按形状名统计的增删
    unauthorized     未授权却被改动的页码

为什么逐页比 str(slide._element.xml)：lxml 元素的 __eq__ 不是字符串比较，
直接 == 会抛 ValueError('line does not match pattern for an XML element')。

为什么还要单独比字号：用户常只把某几个文本框的标题调大（目录 20→28 pt），
XML 变了但图片、表格都没动，只报“changed”等于没报。

配对方式：同一页内按“形状名 + 该名第几次出现”配对，不用文字内容或坐标做键——
用户改写文字、挪动位置后仍能认成同一个形状，才能报出“把 ■ 换成 ✓、并左移 0.13 in”。
"""
import argparse
import json
from collections import Counter
from pathlib import Path

from pptx import Presentation

EMU = 914400.0


def page_xml(slide):
    return str(slide._element.xml)


def image_boxes(slide):
    out = []
    for sh in slide.shapes:
        if getattr(sh, "shape_type", None) == 13:  # PICTURE
            out.append({k: round(getattr(sh, k) / EMU, 3) for k in ("left", "top", "width", "height")})
    return out


def table_sizes(slide):
    out = []
    for sh in slide.shapes:
        if not getattr(sh, "has_table", False):
            continue
        for ri, row in enumerate(sh.table.rows):
            for ci, cell in enumerate(row.cells):
                sizes = set()
                for p in cell.text_frame.paragraphs:
                    for r in p.runs:
                        if r.font.size is not None:
                            sizes.add(round(r.font.size.pt, 1))
                if sizes:
                    out.append({"row": ri, "col": ci, "pt": sorted(sizes)})
    return out


def text_runs(sh):
    out = []
    if not getattr(sh, "has_text_frame", False):
        return out
    for p in sh.text_frame.paragraphs:
        for r in p.runs:
            color = None
            try:
                if r.font.color and r.font.color.type is not None:
                    color = str(r.font.color.rgb)
            except (AttributeError, ValueError):
                color = None
            out.append({"text": r.text,
                        "pt": round(r.font.size.pt, 1) if r.font.size else None,
                        "bold": bool(r.font.bold),
                        "color": color})
    return out


def shape_snapshot(slide):
    """{(name, 第几次出现): {'box': (l,t,w,h), 'runs': [...]}}"""
    snap, seen = {}, Counter()
    for sh in slide.shapes:
        seen[sh.name] += 1
        key = (sh.name, seen[sh.name])
        snap[key] = {
            "box": tuple(round(getattr(sh, k) / EMU, 3) for k in ("left", "top", "width", "height")),
            "runs": text_runs(sh),
        }
    return snap


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("base")
    ap.add_argument("new")
    ap.add_argument("--expect", default="", help="允许改动的页码，如 3,5,16")
    args = ap.parse_args()

    expect = {int(x) for x in args.expect.replace(" ", "").split(",") if x}
    a, b = Presentation(args.base), Presentation(args.new)
    if len(a.slides) != len(b.slides):
        raise SystemExit(f"页数不同：{len(a.slides)} vs {len(b.slides)}")

    report = {"base": str(Path(args.base)), "new": str(Path(args.new)),
              "slides": len(a.slides), "expect": sorted(expect),
              "pages": [], "unauthorized": []}

    for i, (sa, sb) in enumerate(zip(a.slides, b.slides), 1):
        same = page_xml(sa) == page_xml(sb)
        entry = {"page": i, "status": "unchanged" if same else "changed"}
        if not same:
            ia, ib = image_boxes(sa), image_boxes(sb)
            if ia != ib:
                entry["images"] = {"before": ia, "after": ib}

            ta = {(d["row"], d["col"]): d["pt"] for d in table_sizes(sa)}
            tb = {(d["row"], d["col"]): d["pt"] for d in table_sizes(sb)}
            diffs = [{"row": k[0], "col": k[1], "before": ta.get(k), "after": tb.get(k)}
                     for k in sorted(set(ta) | set(tb)) if ta.get(k) != tb.get(k)]
            if diffs:
                entry["tables"] = diffs

            ra, rb = shape_snapshot(sa), shape_snapshot(sb)
            texts, boxes = [], []
            for key in sorted(set(ra) & set(rb)):
                if ra[key]["runs"] != rb[key]["runs"]:
                    texts.append({"name": key[0], "index": key[1],
                                  "before": ra[key]["runs"], "after": rb[key]["runs"]})
                if ra[key]["box"] != rb[key]["box"]:
                    boxes.append({"name": key[0], "index": key[1],
                                  "before": ra[key]["box"], "after": rb[key]["box"]})
            if texts:
                entry["texts"] = texts
            if boxes:
                entry["boxes"] = boxes

            ca, cb = Counter(k[0] for k in ra), Counter(k[0] for k in rb)
            added = sorted((cb - ca).elements())
            removed = sorted((ca - cb).elements())
            if added or removed:
                entry["shapes"] = {"added": added, "removed": removed}

            if expect and i not in expect:
                report["unauthorized"].append(i)
            if not any(k in entry for k in ("images", "tables", "texts", "boxes", "shapes")):
                # XML 变了但结构、文字、字号、图片都没变：多半是编辑器重存时的序列化差异
                entry["status"] = "metadata-only"
        report["pages"].append(entry)

    print(json.dumps(report, ensure_ascii=False, indent=2))
    if report["unauthorized"]:
        raise SystemExit(f"FAIL: 未授权页被改动 {report['unauthorized']}")


if __name__ == "__main__":
    main()
