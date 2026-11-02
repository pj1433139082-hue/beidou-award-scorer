#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""把逐项目意见导出成 Word（.docx），格式简洁明了。

用法：
  python scripts/export_docx.py 意见/ --out 意见word/
  python scripts/export_docx.py 意见/ --out 意见word/ --merge   # 合并成一本

排版原则（面向水务领导）：
  · A4，页边距 2cm，正文 10.5pt 宋体/微软雅黑
  · 分数用加粗大字，区间用颜色
  · 每个维度一个小标题 + 要点列表，不用密集表格
  · 引文用缩进斜体 + 左侧竖线
"""
from __future__ import annotations
import argparse, json, re
from pathlib import Path

try:
    from docx import Document
    from docx.enum.table import WD_TABLE_ALIGNMENT
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn
    from docx.shared import Cm, Pt, RGBColor
except ImportError:
    print("缺少 python-docx。先运行：python scripts/setup.py --install", file=__import__("sys").stderr)
    raise SystemExit(2)

INK = RGBColor(0x1A, 0x1A, 0x1A)
MUTED = RGBColor(0x5A, 0x64, 0x72)
ACCENT = RGBColor(0x0B, 0x6B, 0x57)
WARN = RGBColor(0xB4, 0x53, 0x09)
BAD = RGBColor(0xB9, 0x1C, 0x1C)
GOOD = RGBColor(0x15, 0x80, 0x3D)
BLUE = RGBColor(0x1D, 0x4E, 0xD8)
BAND_COLOR = {
    "一等奖": GOOD, "二等奖": BLUE, "三等奖": RGBColor(0xA1, 0x62, 0x07),
    "入围": MUTED, "待提升": WARN, "否决": BAD,
}
FONT_CN = "微软雅黑"


def set_font(run, size=10.5, bold=False, color=INK, italic=False, name=FONT_CN):
    run.font.name = name
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.italic = italic
    run.font.color.rgb = color
    # 中文字体必须单独设 eastAsia，否则 Word 里中文会退回宋体
    rpr = run._element.get_or_add_rPr()
    rf = rpr.find(qn("w:rFonts"))
    if rf is None:
        rf = OxmlElement("w:rFonts"); rpr.append(rf)
    rf.set(qn("w:eastAsia"), name)


def shade(cell, hex_color: str):
    tcPr = cell._tc.get_or_add_tcPr()
    sh = OxmlElement("w:shd")
    sh.set(qn("w:val"), "clear"); sh.set(qn("w:color"), "auto"); sh.set(qn("w:fill"), hex_color)
    tcPr.append(sh)


def cell_text(cell, text, size=9.5, bold=False, color=INK, align=WD_ALIGN_PARAGRAPH.LEFT):
    cell.text = ""
    p = cell.paragraphs[0]
    p.alignment = align
    p.paragraph_format.space_before = Pt(1)
    p.paragraph_format.space_after = Pt(1)
    set_font(p.add_run(str(text)), size=size, bold=bold, color=color)


def add_rule(doc):
    """一条细横线。"""
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(2)
    p.paragraph_format.space_after = Pt(6)
    pPr = p._p.get_or_add_pPr()
    pbdr = OxmlElement("w:pBdr")
    bottom = OxmlElement("w:bottom")
    bottom.set(qn("w:val"), "single"); bottom.set(qn("w:sz"), "6")
    bottom.set(qn("w:space"), "1"); bottom.set(qn("w:color"), "D9DEE5")
    pbdr.append(bottom); pPr.append(pbdr)


def add_heading(doc, text, level=1):
    sizes = {1: 15, 2: 12.5, 3: 11}
    colors = {1: ACCENT, 2: INK, 3: INK}
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(12 if level == 1 else 9)
    p.paragraph_format.space_after = Pt(4)
    p.paragraph_format.keep_with_next = True
    set_font(p.add_run(text), size=sizes.get(level, 11), bold=True, color=colors.get(level, INK))
    return p


def add_kv_table(doc, pairs, widths=(3.2, 12.5)):
    t = doc.add_table(rows=0, cols=2)
    t.alignment = WD_TABLE_ALIGNMENT.LEFT
    t.autofit = False
    for k, v in pairs:
        row = t.add_row()
        cell_text(row.cells[0], k, size=9.5, bold=True, color=MUTED)
        shade(row.cells[0], "F7F9FB")
        cell_text(row.cells[1], v, size=10)
        row.cells[0].width = Cm(widths[0])
        row.cells[1].width = Cm(widths[1])
    # 只留横向细线
    tbl = t._tbl
    borders = OxmlElement("w:tblBorders")
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        el = OxmlElement("w:" + edge)
        el.set(qn("w:val"), "single"); el.set(qn("w:sz"), "4")
        el.set(qn("w:color"), "E3E8EF")
        borders.append(el)
    tbl.tblPr.append(borders)
    return t


def add_quote(doc, text):
    """引文：缩进 + 左侧竖线 + 灰字。"""
    p = doc.add_paragraph()
    pf = p.paragraph_format
    pf.left_indent = Cm(0.6)
    pf.space_before = Pt(1); pf.space_after = Pt(1)
    pPr = p._p.get_or_add_pPr()
    pbdr = OxmlElement("w:pBdr")
    left = OxmlElement("w:left")
    left.set(qn("w:val"), "single"); left.set(qn("w:sz"), "12")
    left.set(qn("w:space"), "8"); left.set(qn("w:color"), "C9D6D1")
    pbdr.append(left); pPr.append(pbdr)
    set_font(p.add_run(text), size=9, color=MUTED, italic=True)


def add_bullet(doc, text, color=INK, bold_prefix=None):
    p = doc.add_paragraph(style="List Bullet")
    p.paragraph_format.space_before = Pt(0); p.paragraph_format.space_after = Pt(1)
    p.paragraph_format.left_indent = Cm(0.5)
    if bold_prefix:
        set_font(p.add_run(bold_prefix), size=10, bold=True, color=color)
    set_font(p.add_run(text), size=10, color=color)
    return p


# ---------- 解析意见 Markdown ----------
def parse_md(md: str) -> dict:
    """按 ## 段落切分再解析，避免不同段落里相同格式的行互相串。

    早期版本用行首模式猜，导致「争议点」里的编号列表被当成「优先补什么」。
    """
    d = {"title": "", "meta": [], "summary": "", "dims": [], "disputes": [],
         "deductions": [], "priority": []}
    # 先按 ## 切成段
    sections = {}
    cur = "_head"
    sections[cur] = []
    for ln in md.splitlines():
        if ln.startswith("## "):
            cur = ln[3:].strip()
            sections[cur] = []
            continue
        sections[cur].append(ln)
    # 维度内容在「各维度得分与意见」这一段里，必须一起解析
    dimsec = next((v for k, v in sections.items() if k.startswith("各维度")), [])
    body = sections.get("_head", []) + dimsec
    disp = next((v for k, v in sections.items() if k.startswith("争议点")), [])
    ded = next((v for k, v in sections.items() if k.startswith("固定扣分")), [])
    pri = next((v for k, v in sections.items() if k.startswith("优先补什么")), [])
    lines = body
    i = 0
    while i < len(lines):
        ln = lines[i]
        if ln.startswith("# "):
            d["title"] = ln[2:].strip()
        elif ln.startswith("| ") and "---" not in ln:
            cells = [c.strip() for c in ln.strip("|").split("|")]
            if len(cells) == 2 and cells[0] and cells[0] != "":
                d["meta"].append((cells[0], cells[1].replace("**", "")))
        elif ln.startswith("**一句话**："):
            d["summary"] = ln.replace("**一句话**：", "").strip()
        elif ln.startswith("### "):
            head = ln[4:].strip()
            m = re.match(r"(.+?)\u3000(\S+) / (\S+)（(\d+)%）", head)
            dim = {"name": head, "score": "", "max": "", "pct": "", "items": [], "quotes": [],
                   "rejected": [], "fixes": [], "cap": "", "grade": ""}
            if m:
                dim.update(name=m.group(1).strip(), score=m.group(2), max=m.group(3), pct=m.group(4))
            d["dims"].append(dim)
        elif d["dims"] and ln.startswith("- "):
            cur = d["dims"][-1]
            txt = ln[2:].strip()
            # 按明确模式提取，避免用 in 做模糊判断（会把 ** 一起留下）
            m_cap = re.match(r"\*\*封顶原因\*\*：(.+)$", txt)
            m_note = re.match(r"⚠\s*注意：(.+)$", txt)
            m_grade = re.match(r"证据等级：\*\*([ABCD])\*\*", txt)
            m_fix = re.match(r"（(C\d+) ([^）]+)）(.+)$", txt)
            if m_cap:
                cur["cap"] = m_cap.group(1).strip()
            elif m_note:
                cur["note"] = m_note.group(1).strip()
            elif m_grade:
                cur["grade"] = m_grade.group(1)
            elif m_fix:
                cur["fixes"].append((m_fix.group(1), m_fix.group(2), m_fix.group(3).strip()))
            else:
                cur["items"].append(txt)
        elif d["dims"] and ln.startswith("  > "):
            d["dims"][-1]["quotes"].append(ln.strip()[2:].strip())
        elif d["dims"] and ln.strip().startswith("> ") and "未通过核验" in "".join(lines[max(0,i-3):i]):
            d["dims"][-1]["rejected"].append(ln.strip()[2:].strip())
        elif ln.startswith("**怎么改**"):
            pass
        i += 1
    # 争议点
    for ln in disp:
        m = re.match(r"^\d+\. \*\*(.+?)\*\*：(.+)$", ln.strip())
        if m:
            d["disputes"].append((m.group(1), m.group(2)))
    # 固定扣分
    for ln in ded:
        if ln.strip().startswith("- "):
            d["deductions"].append(ln.strip()[2:].strip())
    # 优先补什么
    for ln in pri:
        m = re.match(r"^\d+\. (.+)$", ln.strip())
        if m:
            d["priority"].append(m.group(1).strip())
    return d


def build_doc(data: dict, out_path, doc=None):
    if doc is None:
        doc = Document()
        sec = doc.sections[0]
        sec.page_width, sec.page_height = Cm(21), Cm(29.7)
        for attr in ("left_margin", "right_margin"):
            setattr(sec, attr, Cm(2.0))
        sec.top_margin = Cm(1.8); sec.bottom_margin = Cm(1.8)
        st = doc.styles["Normal"]
        st.font.name = FONT_CN; st.font.size = Pt(10.5)
        st.element.rPr.rFonts.set(qn("w:eastAsia"), FONT_CN)

    # 标题
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(2)
    set_font(p.add_run(data["title"]), size=16, bold=True, color=INK)

    # 基本信息表
    meta = [(k, v) for k, v in data["meta"] if k]
    if meta:
        add_kv_table(doc, meta)

    # 一句话结论
    if data["summary"]:
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(8); p.paragraph_format.space_after = Pt(4)
        pPr = p._p.get_or_add_pPr()
        pbdr = OxmlElement("w:pBdr")
        for edge, sz in (("top", "4"), ("bottom", "4"), ("left", "18"), ("right", "4")):
            el = OxmlElement("w:" + edge)
            el.set(qn("w:val"), "single"); el.set(qn("w:sz"), sz)
            el.set(qn("w:space"), "6" if edge in ("left", "right") else "4")
            el.set(qn("w:color"), "0B6B57")
            pbdr.append(el)
        pPr.append(pbdr)
        set_font(p.add_run(data["summary"]), size=10.5, color=INK)

    # 各维度
    add_heading(doc, "各维度得分与意见", 1)
    for dim in data["dims"]:
        # 维度标题行：名称 …… 分数
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(10); p.paragraph_format.space_after = Pt(2)
        p.paragraph_format.keep_with_next = True
        set_font(p.add_run(dim["name"]), size=12, bold=True, color=INK)
        if dim["score"]:
            set_font(p.add_run("　　%s / %s" % (dim["score"], dim["max"])), size=12, bold=True, color=ACCENT)
            set_font(p.add_run("（%s%%）" % dim["pct"]), size=10, color=MUTED)

        if dim["grade"]:
            add_bullet(doc, "（这一项说的是引文能否在原文逐字核验，不等于业务达成度）",
                       color=MUTED, bold_prefix="证据等级 %s " % dim["grade"])
        if dim["cap"]:
            add_bullet(doc, dim["cap"], color=WARN, bold_prefix="封顶原因：")
        if dim.get("note"):
            add_bullet(doc, dim["note"].replace("⚠ 注意：", ""), color=WARN, bold_prefix="注意：")

        if dim["quotes"]:
            p = doc.add_paragraph()
            p.paragraph_format.space_before = Pt(4); p.paragraph_format.space_after = Pt(1)
            set_font(p.add_run("已核验原文引文"), size=9, bold=True, color=MUTED)
            for q in dim["quotes"][:3]:
                add_quote(doc, q)
        if dim["rejected"]:
            p = doc.add_paragraph()
            p.paragraph_format.space_before = Pt(4); p.paragraph_format.space_after = Pt(1)
            set_font(p.add_run("未通过核验的引文（原文中找不到，不计分）"), size=9, bold=True, color=BAD)
            for q in dim["rejected"][:3]:
                add_quote(doc, q)

        if dim["fixes"]:
            p = doc.add_paragraph()
            p.paragraph_format.space_before = Pt(6); p.paragraph_format.space_after = Pt(1)
            p.paragraph_format.keep_with_next = True
            set_font(p.add_run("怎么改"), size=9.5, bold=True, color=ACCENT)
            for item in dim["fixes"][:2]:
                if isinstance(item, tuple):
                    cid, cname, body = item
                    add_bullet(doc, body, bold_prefix="【%s %s】" % (cid, cname))
                else:
                    add_bullet(doc, re.sub(r"^（(C\d+) ([^）]+)）", r"【\1 \2】", item))
        add_rule(doc)

    # 争议点
    if data["disputes"]:
        add_heading(doc, "争议点（答辩会被追问）", 1)
        for i, (src, pt) in enumerate(data["disputes"], 1):
            p = doc.add_paragraph()
            p.paragraph_format.space_before = Pt(1); p.paragraph_format.space_after = Pt(1)
            p.paragraph_format.left_indent = Cm(0.5)
            set_font(p.add_run("%d. " % i), size=10, color=MUTED)
            set_font(p.add_run(src + "："), size=10, bold=True, color=BAD if "否决" in src else INK)
            set_font(p.add_run(pt), size=10, color=INK)

    # 扣分
    if data["deductions"]:
        add_heading(doc, "固定扣分", 1)
        for x in data["deductions"]:
            add_bullet(doc, x, color=WARN)

    # 优先补什么
    if data["priority"]:
        add_heading(doc, "优先补什么（按可追回分数排序）", 1)
        for i, x in enumerate(data["priority"], 1):
            m = re.match(r"\*\*(.+?)\*\*（可追回 (.+?) 分）：(.+)$", x)
            p = doc.add_paragraph()
            p.paragraph_format.space_before = Pt(2); p.paragraph_format.space_after = Pt(2)
            p.paragraph_format.left_indent = Cm(0.5)
            if m:
                set_font(p.add_run("%d. " % i), size=10.5, bold=True, color=ACCENT)
                set_font(p.add_run(m.group(1)), size=10.5, bold=True, color=INK)
                set_font(p.add_run("  可追回 %s 分" % m.group(2)), size=9.5, bold=True, color=GOOD)
                p2 = doc.add_paragraph()
                p2.paragraph_format.left_indent = Cm(1.0)
                p2.paragraph_format.space_before = Pt(0); p2.paragraph_format.space_after = Pt(3)
                set_font(p2.add_run(m.group(3)), size=10, color=MUTED)
            else:
                set_font(p.add_run("%d. %s" % (i, x)), size=10, color=INK)

    # 页脚说明
    add_rule(doc)
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(2)
    set_font(p.add_run("分数为自动机检 + 证据定级结果，需人工复核 judgments.json 后方可作为评审结论。"),
             size=8.5, color=MUTED, italic=True)

    if out_path is not None:
        doc.save(str(out_path))
    return out_path


def build_merged(files, out_path: Path):
    """把全部意见合成一本，按奖项分组，每份新起一页。"""
    doc = Document()
    sec = doc.sections[0]
    sec.page_width, sec.page_height = Cm(21), Cm(29.7)
    sec.left_margin = sec.right_margin = Cm(2.0)
    sec.top_margin = sec.bottom_margin = Cm(1.8)
    st = doc.styles["Normal"]
    st.font.name = FONT_CN; st.font.size = Pt(10.5)
    st.element.rPr.rFonts.set(qn("w:eastAsia"), FONT_CN)

    datas = [(f, parse_md(f.read_text(encoding="utf-8"))) for f in files]
    # 按奖项排序
    def award_of(d):
        for k, v in d["meta"]:
            if k == "奖项":
                return v
        return "其他"
    datas.sort(key=lambda x: (award_of(x[1]), -float(
        next((v.split("/")[0].strip() for k, v in x[1]["meta"] if k == "AI 评审分"), "0") or 0)))

    # 封面
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(160)
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    set_font(p.add_run("北斗奖申报材料"), size=26, bold=True, color=INK)
    p2 = doc.add_paragraph(); p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
    set_font(p2.add_run("评分意见汇总"), size=26, bold=True, color=ACCENT)
    p3 = doc.add_paragraph(); p3.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p3.paragraph_format.space_before = Pt(20)
    set_font(p3.add_run("共 %d 份材料" % len(datas)), size=12, color=MUTED)
    p4 = doc.add_paragraph(); p4.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p4.paragraph_format.space_before = Pt(6)
    set_font(p4.add_run("分数区间为相对刻度（85/75/65/55），非官方切线"),
             size=9.5, color=MUTED, italic=True)

    # 目录
    doc.add_page_break()
    add_heading(doc, "目录", 1)
    cur_award = None
    for f, d in datas:
        a = award_of(d)
        if a != cur_award:
            cur_award = a
            pp = doc.add_paragraph()
            pp.paragraph_format.space_before = Pt(8); pp.paragraph_format.space_after = Pt(2)
            set_font(pp.add_run(a), size=11, bold=True, color=ACCENT)
        score = next((v.split("/")[0].strip() for k, v in d["meta"] if k == "AI 评审分"), "")
        pp = doc.add_paragraph()
        pp.paragraph_format.space_before = Pt(0); pp.paragraph_format.space_after = Pt(1)
        pp.paragraph_format.left_indent = Cm(0.5)
        set_font(pp.add_run(d["title"]), size=9.5, color=INK)
        set_font(pp.add_run("　　%s 分" % score), size=9.5, bold=True, color=MUTED)

    # 正文：每份新起一页
    for f, d in datas:
        doc.add_page_break()
        build_doc(d, None, doc=doc)
    doc.save(str(out_path))
    return out_path


def main() -> int:
    ap = argparse.ArgumentParser(description="逐项目意见 → Word")
    ap.add_argument("src", help="含 *.意见.md 的目录")
    ap.add_argument("--out", default="意见word")
    ap.add_argument("--merge", action="store_true", help="另存一份合并总册（按奖项分组）")
    a = ap.parse_args()
    src, out = Path(a.src), Path(a.out)
    files = sorted(src.glob("*.意见.md"))
    if not files:
        print("没找到 *.意见.md", file=__import__("sys").stderr)
        return 1
    out.mkdir(parents=True, exist_ok=True)
    made = []
    for f in files:
        data = parse_md(f.read_text(encoding="utf-8"))
        p = out / (f.name.replace(".意见.md", ".docx"))
        build_doc(data, p)
        made.append(p.name)
    if a.merge:
        build_merged(files, out / "北斗奖评分意见汇总.docx")
        print(json.dumps({"out": str(out), "count": len(made), "merged": "北斗奖评分意见汇总.docx"},
                         ensure_ascii=False))
    else:
        print(json.dumps({"out": str(out), "count": len(made)}, ensure_ascii=False))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())