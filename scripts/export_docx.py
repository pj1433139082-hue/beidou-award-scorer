#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""逐项目意见 → Word（三段式）。

每份材料：一、总体判断　二、逐条修改建议（三段式）　三、答辩可能被追问

建议条目来自 advice_engine.build_items()，与 Excel「修改建议」表同一来源，
编号（M01、M02……）、顺序、条数完全一致。

每条建议三段，可直接复制：
    原文是什么 → 怎么改 → 修改之后的句子
"""
from __future__ import annotations
import argparse, json, re, sys
from pathlib import Path

try:
    from docx import Document
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn
    from docx.shared import Cm, Pt, RGBColor
except ImportError:
    print("缺少 python-docx。先运行：python scripts/setup.py --install", file=sys.stderr)
    raise SystemExit(2)

sys.path.insert(0, str(Path(__file__).resolve().parent))
from advice_engine import build_items  # noqa: E402

BLACK = RGBColor(0, 0, 0)
GREY = RGBColor(0x55, 0x55, 0x55)
ACCENT = RGBColor(0x0B, 0x6B, 0x57)
FIXED = RGBColor(0x0B, 0x4F, 0x3F)
FONT = "微软雅黑"


def strip_award(name: str) -> str:
    s = re.sub(r"^(金点子|金刚钻|金引擎|金智囊|金火炬|特别奖-[^\s\-－—]*)[-－—\s]*", "",
               str(name or "").strip())
    return re.sub(r"^[-－—\s]+", "", s).strip() or str(name or "")


def set_run(run, size=10.5, bold=False, color=BLACK, italic=False, name=FONT):
    run.font.name = name
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.italic = italic
    run.font.color.rgb = color
    rpr = run._element.get_or_add_rPr()
    rf = rpr.find(qn("w:rFonts"))
    if rf is None:
        rf = OxmlElement("w:rFonts"); rpr.append(rf)
    rf.set(qn("w:eastAsia"), name)


def para(doc, text="", size=10.5, bold=False, color=BLACK, italic=False,
         before=0, after=3, indent=0, keep=False):
    p = doc.add_paragraph()
    pf = p.paragraph_format
    pf.space_before = Pt(before)
    pf.space_after = Pt(after)
    pf.line_spacing = 1.15
    if indent:
        pf.left_indent = Cm(indent)
    if keep:
        pf.keep_with_next = True
    if text:
        set_run(p.add_run(text), size=size, bold=bold, color=color, italic=italic)
    return p


def h(doc, text, size=11, before=10):
    return para(doc, text, size=size, bold=True, color=ACCENT, before=before, after=4, keep=True)


def seg_label(doc, text):
    p = doc.add_paragraph()
    pf = p.paragraph_format
    pf.space_before = Pt(2)
    pf.space_after = Pt(0)
    pf.left_indent = Cm(0.5)
    pf.keep_with_next = True
    set_run(p.add_run(text), size=9, bold=True, color=GREY)
    return p


def body_text(doc, text, indent=0.5, size=10, color=BLACK):
    p = doc.add_paragraph()
    pf = p.paragraph_format
    pf.left_indent = Cm(indent)
    pf.space_before = Pt(0)
    pf.space_after = Pt(2)
    pf.line_spacing = 1.15
    set_run(p.add_run(text), size=size, color=color)
    return p


def build(data: dict, out_path, doc=None, mat_index: int = 0):
    if doc is None:
        doc = Document()
        sec = doc.sections[0]
        sec.page_width, sec.page_height = Cm(21), Cm(29.7)
        sec.left_margin = sec.right_margin = Cm(2.2)
        sec.top_margin = sec.bottom_margin = Cm(2.0)
        st = doc.styles["Normal"]
        st.font.name = FONT
        st.font.size = Pt(10.5)
        st.element.rPr.rFonts.set(qn("w:eastAsia"), FONT)

    title = strip_award(data["name"])
    if len(title) > 28 and "——" in title:
        main, sub = title.split("——", 1)
        main, sub = main.strip().rstrip("-－—").strip(), sub.strip()
        p = para(doc, "", size=12.5, bold=True, after=1)
        set_run(p.add_run(main), size=12.5, bold=True)
        p2 = para(doc, "", size=10.5, bold=True, color=GREY, after=3)
        set_run(p2.add_run("——" + sub), size=10.5, bold=True, color=GREY)
    else:
        para(doc, title, size=12.5, bold=True, after=3)

    para(doc, "奖项：%s　　得分：%s / 100　　置信度：%s" % (
        data["award"], data["ai_score"], data["confidence"]),
        size=9.5, color=GREY, after=1)
    para(doc, "区间：%s" % data["band"], size=9.5, color=GREY, after=2)
    if data.get("veto_pending"):
        para(doc, "否决项：待裁定（%s）" % (data.get("veto_note") or ""), size=9.5, after=2)

    h(doc, "一、总体判断")
    dims = data["dimensions"]
    by_ratio = sorted(dims, key=lambda d: d["score"] / d["max"] if d["max"] else 1)
    best = by_ratio[-1]
    worst = by_ratio[:2]
    para(doc, "最强项是%s（%s/%s）。最需要补的是%s。" % (
        best["name"], best["score"], best["max"],
        "、".join("%s（%s/%s）" % (d["name"], d["score"], d["max"]) for d in worst)), size=10.5)

    items = build_items(data)
    h(doc, "二、逐条修改建议")
    if not items:
        para(doc, "各维度均已达标，无需补强。", size=10.5)
    for it in items:
        p = doc.add_paragraph()
        pf = p.paragraph_format
        pf.space_before = Pt(8)
        pf.space_after = Pt(1)
        pf.keep_with_next = True
        no = ("%d-%s" % (mat_index, it["no"])) if mat_index else it["no"]
        set_run(p.add_run(no + "　"), size=10.5, bold=True, color=ACCENT)
        set_run(p.add_run("%s　%s" % (it["dim"], it["title"])), size=10.5, bold=True)
        if it["loss"] > 0.01:
            set_run(p.add_run("　（可追回 %s 分）" % it["loss"]), size=9, color=GREY)

        seg_label(doc, "原文是什么")
        body_text(doc, it["original"] or "（该项未定位到具体原文，见问题描述）")
        seg_label(doc, "怎么改")
        body_text(doc, it["how"])
        seg_label(doc, "修改之后的句子（可直接复制）")
        body_text(doc, it["revised"], color=FIXED)

    ds = data.get("disputes") or []
    if ds:
        h(doc, "三、答辩可能被追问")
        for i, x in enumerate(ds, 1):
            para(doc, "%d. %s" % (i, (x.get("point") or "").strip()),
                 size=10, indent=0.5, after=2)

    para(doc, "分数为自动机检 + 证据定级结果，需人工复核后方可作为评审结论。",
         size=8.5, color=GREY, italic=True, before=10)

    if out_path is not None:
        doc.save(str(out_path))
    return out_path


def load(src: Path):
    rows = []
    for sf in sorted(src.glob("*.score.json")):
        d = json.loads(sf.read_text(encoding="utf-8"))
        stem = sf.name.replace(".score.json", "")
        cf = src / (stem + ".checks.json")
        d["_checks"] = {x["id"]: x for x in
                        (json.loads(cf.read_text(encoding="utf-8")).get("checks", [])
                         if cf.is_file() else [])}
        df = src / (stem + ".doc.json")
        d["_plain"] = (json.loads(df.read_text(encoding="utf-8")).get("plain", "")
                       if df.is_file() else "")
        d["_stem"] = stem
        rows.append(d)
    return rows


def build_merged(rows, out_path: Path):
    doc = Document()
    sec = doc.sections[0]
    sec.page_width, sec.page_height = Cm(21), Cm(29.7)
    sec.left_margin = sec.right_margin = Cm(2.2)
    sec.top_margin = sec.bottom_margin = Cm(2.0)
    st = doc.styles["Normal"]
    st.font.name = FONT; st.font.size = Pt(10.5)
    st.element.rPr.rFonts.set(qn("w:eastAsia"), FONT)

    for _ in range(8):
        para(doc, "", after=0)
    p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    set_run(p.add_run("北斗奖申报材料评分意见"), size=22, bold=True)
    p2 = doc.add_paragraph(); p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
    set_run(p2.add_run("共 %d 份" % len(rows)), size=11, color=GREY)

    doc.add_page_break()
    h(doc, "目录", before=0)
    order = {}
    for r in rows:
        order.setdefault(r["award"], []).append(r)
    for award, items in order.items():
        para(doc, award, size=10.5, bold=True, before=8, after=2)
        for r in sorted(items, key=lambda x: -x["ai_score"]):
            nm = strip_award(r["name"])
            if len(nm) > 32:
                nm = nm[:31] + "…"
            idx = rows.index(r) + 1
            para(doc, "    %2d. %5.1f 分　%s" % (idx, r["ai_score"], nm), size=9.5, after=1)

    for mi, r in enumerate(rows, 1):
        doc.add_page_break()
        build(r, None, doc=doc, mat_index=mi)
    doc.save(str(out_path))
    return out_path


def main() -> int:
    ap = argparse.ArgumentParser(description="逐项目意见 → Word（三段式）")
    ap.add_argument("src", help="含 *.score.json 的目录")
    ap.add_argument("--out", default="意见Word")
    ap.add_argument("--merge", action="store_true", help="另存一份合并总册")
    a = ap.parse_args()
    src, out = Path(a.src), Path(a.out)
    rows = load(src)
    if not rows:
        print("没找到 score.json", file=sys.stderr)
        return 1
    # 排序必须与 export_xlsx.py 完全一致（ai_score 降序），否则材料序号对不上，
    # Word 与 Excel 的编号就会错位。
    rows.sort(key=lambda r: r.get("ai_score", 0), reverse=True)
    out.mkdir(parents=True, exist_ok=True)
    for mi, r in enumerate(rows, 1):
        build(r, out / (r["_stem"] + ".docx"), mat_index=mi)
    if a.merge:
        build_merged(rows, out / "北斗奖评分意见汇总.docx")
    print(json.dumps({"out": str(out), "count": len(rows), "merged": bool(a.merge)},
                     ensure_ascii=False))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
