#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""逐项目意见 → Word（简洁版）。

每份材料：一、总体判断　二、逐项问题与修改（原文哪里不好 / 怎么改）　三、优先补什么

刻意不加配色、不加底纹、不加表格线——打印和投影都清楚。
数据直接读 score.json + checks.json，能拿到 checks[].evidence 里的原文定位。
"""
from __future__ import annotations
import argparse, json, re
from pathlib import Path

try:
    from docx import Document
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn
    from docx.shared import Cm, Pt, RGBColor
except ImportError:
    print("缺少 python-docx。先运行：python scripts/setup.py --install", file=__import__("sys").stderr)
    raise SystemExit(2)

BLACK = RGBColor(0, 0, 0)
GREY = RGBColor(0x55, 0x55, 0x55)
FONT = "微软雅黑"

# 机检编号 → (属于哪些维度, 怎么改)
CHECKS = {
    "C01": ({"_global"}, "补齐三张必备表：基本信息表、基线与效果对照表、验证清单。"),
    "C02": ({"value"}, "效益表分成两栏：「实际值【已实现】」与「远期目标【规划/在研】」。已兑现的填实际值并附台账截图，未兑现的只填远期目标。"),
    "C03": ({"value"}, "每项效益补上改造前的基准值和对比周期，否则变化率算不出来、别人无法复算。"),
    "C04": ({"value"}, "每个已实现效益补佐证材料：台账截图 / 运行日志 / 核算表，填进「佐证材料」列。"),
    "C05": ({"value", "economics"}, "正文里的百分比与表格数字对齐，偏差控制在 5 个百分点以内。"),
    "C06": ({"replication", "promotion"}, "补一份可点数的推广清单：大区 / 区域 / 水厂逐个列名，正文写清覆盖几个大区、几个区域、几个项目。"),
    "C07": ({"economics"}, "补「计算步骤」表：每一步写清公式、代入值、结果、单位、含义。只给总金额无法复算。"),
    "C08": ({"value", "economics"}, "同一金额在多处出现时，确认没有重复计入（例如吨水单耗与总费用各算一次）。"),
    "C09": ({"_global"}, "把没填的占位符（下划线、XXXX、（文本））全部补实。"),
    "C10": ({"_global"}, "补齐材料模板要求但当前缺失的段落。"),
    "C11": ({"intelligence"}, "补三件套：模型实验室线上部署记录、试点厂生产调用记录、不少于 90 天的运行验证数据。"),
    "C12": ({"readiness"}, "写清从方案到产品的转化程度：已产品化（有合同/订单/收入）→ 非常符合；仅有原型 → 符合；仍在方案 → 一般。"),
    "C13": ({"iteration"}, "补一段「相较上次参赛版本」的迭代说明：改了什么、指标提升多少。"),
    "C14": ({"novelty"}, "补与原有办法的量化对照：同一场景下新旧方案的效果差值（成本 / 精度 / 时效）。"),
    "C15": ({"usability"}, "补两段：与集团战略方向的对应关系；当前卡点的具体表现（现状数据 + 卡在哪 + 本方案如何解决）。"),
    "C16": ({"value", "usability"}, "把「预计 / 有望」换成实际发生的量，成效要落到具体金额或比例。"),
    "C17": ({"_global"}, "插入关键运行界面、台账、报告截图，并加文字说明。"),
    "C18": ({"replication", "promotion"}, "补齐推广案例佐证、经济测算过程、产品形态说明三块。"),
    "C19": ({"_global"}, "补齐专属模板六段：实施背景、模型应用、业务价值、ROI 核算、下一步计划、量化数据。"),
    "C20": ({"economics"}, "ROI 分投入与收益两段写，并配足可复算的量化数字。"),
}

PROBLEM = {
    "C01": "缺少必备表格",
    "C02": "效益里分不清哪些已实现、哪些还在规划",
    "C03": "有实际值但没有基准值，变化率算不出来",
    "C04": "已实现的效益没有佐证材料",
    "C05": "正文百分比与表格数字对不上",
    "C06": "没有可点数的推广清单",
    "C07": "没有计算步骤表，金额无法复算",
    "C08": "同一金额在多处出现，疑似重复计入",
    "C09": "还有未填的占位符",
    "C10": "缺少模板要求的段落",
    "C11": "模型实验室 / 试点厂 / 90 天验证的证据不全",
    "C12": "没写清产品转化到什么程度",
    "C13": "没有「相较上次参赛版本」的迭代说明",
    "C14": "没有与原有办法的量化对照",
    "C15": "战略对齐或卡点描述缺一项",
    "C16": "预计性表述偏多，实际数据不足",
    "C17": "缺少运行截图等图形证据",
    "C18": "推广 / 测算 / 产品形态证据不齐",
    "C19": "专属模板段落不全",
    "C20": "ROI 缺可复算的投入与收益",
}

FALLBACK = {
    "value": "把效益拆成「已实现」与「规划」两栏分别给数，每组数字都要能用一个公式复算出来。",
    "replication": "给出可点数的推广事实：已落地几个大区 / 区域 / 厂站，逐个列名，并说明可复制的原因。",
    "novelty": "点明与原有做法的差别：同一场景下新旧方案的指标差，并说明差异来自哪个技术点。",
    "intelligence": "补模型实验室部署记录、试点厂生产调用日志、以及不少于 90 天的运行表现。",
    "promotion": "补推广台账：按大区—区域—厂站三级列清单，标明部署时间与规模。",
    "economics": "补一张计算步骤表：输入参数、公式、中间结果、最终金额、单位与口径各占一列。",
    "readiness": "写清交付形态与转化程度：是否已有合同 / 订单 / 收入，是否有标准产品文档与运维手册。",
    "usability": "补战略对应关系与卡点具体表现（现状数据 + 卡在哪 + 怎么解决）。",
}


def set_run(run, size=10.5, bold=False, color=BLACK, italic=False):
    run.font.name = FONT
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.italic = italic
    run.font.color.rgb = color
    rpr = run._element.get_or_add_rPr()
    rf = rpr.find(qn("w:rFonts"))
    if rf is None:
        rf = OxmlElement("w:rFonts"); rpr.append(rf)
    rf.set(qn("w:eastAsia"), FONT)


def para(doc, text="", size=10.5, bold=False, color=BLACK, italic=False,
         before=0, after=3, indent=0, keep=False):
    p = doc.add_paragraph()
    pf = p.paragraph_format
    pf.space_before = Pt(before)
    pf.space_after = Pt(after)
    pf.line_spacing = 1.35
    if indent:
        pf.left_indent = Cm(indent)
    if keep:
        pf.keep_with_next = True
    if text:
        set_run(p.add_run(text), size=size, bold=bold, color=color, italic=italic)
    return p


def h(doc, text, size=11.5, before=14):
    return para(doc, text, size=size, bold=True, before=before, after=4, keep=True)


def bullet(doc, text, indent=0.55, size=10):
    p = doc.add_paragraph()
    pf = p.paragraph_format
    pf.left_indent = Cm(indent + 0.35)
    pf.first_line_indent = Cm(-0.35)
    pf.space_before = Pt(0)
    pf.space_after = Pt(2)
    pf.line_spacing = 1.3
    set_run(p.add_run("· "), size=size)
    set_run(p.add_run(text), size=size)
    return p


def label(doc, text, size=9.5):
    return para(doc, text, size=size, bold=True, color=GREY, before=6, after=2, keep=True)


def find_quote(plain: str, keywords: list[str], max_len: int = 70) -> str:
    """在原文里找一句能佐证问题的句子。

    要求至少命中 2 个线索词，且句子不是纯标题/纯奖项名——
    否则会挑出「参赛赛道：成果应用赛道 金火炬」这种对问题毫无说明力的句子。
    """
    if not plain:
        return ""
    import re as _re
    NOISE = ("参赛赛道", "申报单位", "项目名称", "填报", "目录")
    # 纯章节标题（[H3] 2.3.1 xxx）、纯名称串，对说明问题没有价值
    import re as _re2
    HEADING = _re2.compile(r"^\s*(\[H\d\]\s*)?(\d+(\.\d+)*\s*)?[^。；]{0,30}$")
    for s in _re.split(r"[。；\n]", plain):
        s = s.strip()
        if not (12 <= len(s) <= 200):
            continue
        if any(n in s for n in NOISE):
            continue
        # 没有句读、长度又短 → 多半是标题，跳过
        if HEADING.match(s) and "。" not in s and len(s) < 40:
            continue
        hit = sum(1 for k in keywords if k in s)
        if hit >= 2:
            return s if len(s) <= max_len else s[:max_len - 1] + "…"
    return ""


# 各机检问题对应的原文线索词（用于定位原文）
HINTS = {
    "C06": ["推广", "应用", "大区", "区域", "厂"],
    "C07": ["测算", "计算", "公式", "节约", "万元"],
    "C11": ["模型实验室", "试点", "上线", "运行"],
    "C12": ["产品", "交付", "合同", "原型"],
    "C13": ["迭代", "升级", "版本"],
    "C14": ["对比", "优于", "原有", "传统"],
    "C15": ["战略", "卡点", "痛点"],
    "C17": ["截图", "图"],
    "C02": ["效益", "实际值", "规划"],
    "C03": ["基准", "改造前"],
    "C04": ["佐证", "台账"],
}


def collect_issues(row: dict) -> list[dict]:
    """把 score.json + checks.json 整理成「每个维度有哪些原文问题 + 怎么改」。"""
    checks = row.get("_checks", {})
    out = []
    for dim in row["dimensions"]:
        did = dim["id"]
        problems = []

        if dim.get("cap_reason"):
            problems.append({"kind": "cap", "text": dim["cap_reason"], "ev": []})

        for cid, c in checks.items():
            if c.get("status") not in ("warn", "fail"):
                continue
            belongs = CHECKS.get(cid, (set(), ""))[0]
            if "_global" in belongs or did not in belongs:
                continue
            # 机检自己声明的 dims 优先：它说属于哪个维度就只挂哪个
            cdims = c.get("dims") or []
            if cdims and did not in cdims:
                continue
            ev = [str(x) for x in (c.get("evidence") or [])][:4]
            if not ev:
                # 机检没给定位时，从原文找一句线索句
                q = find_quote(row.get("_plain", ""), HINTS.get(cid, []))
                if q:
                    ev = [q]
            problems.append({
                "kind": "check", "cid": cid,
                "text": c.get("detail") or PROBLEM.get(cid, ""),
                "ev": ev,
            })

        for rj in dim.get("quotes_rejected") or []:
            q = rj.get("quote") if isinstance(rj, dict) else str(rj)
            problems.append({"kind": "quote", "text": "引文在原文中找不到：%s" % q, "ev": []})

        if dim.get("claimed_grade") and dim.get("claimed_grade") != dim.get("verified_grade"):
            problems.append({"kind": "grade",
                             "text": "证据等级从 %s 下调到 %s：部分引文未能在原文中逐字核验"
                                     % (dim.get("claimed_grade"), dim.get("verified_grade")),
                             "ev": []})

        fixes = []
        for pr in problems:
            cid = pr.get("cid")
            if cid:
                t = CHECKS.get(cid, (set(), ""))[1]
                if t and t not in fixes:
                    fixes.append(t)
        for cid, c in checks.items():
            if c.get("status") in ("warn", "fail"):
                belongs = CHECKS.get(cid, (set(), ""))[0]
                if did in belongs:
                    t = CHECKS.get(cid, (set(), ""))[1]
                    if t and t not in fixes:
                        fixes.append(t)
        if not fixes and FALLBACK.get(did):
            fixes.append(FALLBACK[did])

        # 去重：封顶原因与机检 detail 常常在说同一件事
        seen_txt, uniq = set(), []
        for pr in problems:
            # 去标点后取前 8 字做键：封顶原因与机检 detail 通常前半句一致
            key = re.sub(r"[\s　「」【】——，。、：；()（）]", "", pr["text"])[:8]
            if key in seen_txt:
                continue
            seen_txt.add(key)
            uniq.append(pr)

        # 没有具体问题但确实失分：说明原因是证据强度不足，保持「原文哪里不好 → 怎么改」格式完整
        if not uniq and round(dim["max"] - dim["score"], 1) > 0.01:
            g = dim.get("verified_grade") or ""
            why = {
                "A": "证据等级 A，但该维度仍因业务条件未完全达成而失分",
                "B": "证据等级 B：有论证但不够充分，缺少可核验的量化数据",
                "C": "证据等级 C：论证薄弱，多为定性描述，缺具体数字与佐证",
                "D": "证据等级 D：几乎没有可引用的证据",
            }.get(g, "该维度论证不足，失分但未定位到具体原文问题")
            uniq.append({"kind": "grade", "text": why, "ev": []})

        out.append({"dim": dim, "problems": uniq, "fixes": fixes[:2],
                    "loss": round(dim["max"] - dim["score"], 1)})
    return [x for x in out if x["loss"] > 0.01 or x["problems"]]


def build(data: dict, out_path, doc=None):
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

    para(doc, data["name"], size=15, bold=True, after=4)
    para(doc, "奖项：%s　　得分：%s / 100　　区间：%s　　置信度：%s" % (
        data["award"], data["ai_score"], data["band"], data["confidence"]),
        size=10, color=GREY, after=2)
    if data.get("veto_pending"):
        para(doc, "否决项：待裁定（%s）" % (data.get("veto_note") or ""), size=10, after=2)

    h(doc, "一、总体判断")
    dims = data["dimensions"]
    by_ratio = sorted(dims, key=lambda d: d["score"] / d["max"] if d["max"] else 1)
    best = by_ratio[-1]
    worst = by_ratio[:2]
    para(doc, "最强项是%s（%s/%s）。最需要补的是%s。" % (
        best["name"], best["score"], best["max"],
        "、".join("%s（%s/%s）" % (d["name"], d["score"], d["max"]) for d in worst)), size=10.5)

    issues = collect_issues(data)
    h(doc, "二、逐项问题与修改")
    if not issues:
        para(doc, "各维度均已达标，无需补强。", size=10.5)
    for i, it in enumerate(issues, 1):
        d = it["dim"]
        pct = round(d["score"] / d["max"] * 100) if d["max"] else 0
        para(doc, "%d. %s　%s / %s（%d%%）" % (i, d["name"], d["score"], d["max"], pct),
             size=10.5, bold=True, before=9, after=3, keep=True)
        if it["problems"]:
            label(doc, "原文哪里不好")
            for pr in it["problems"][:4]:
                bullet(doc, pr["text"])
                evs = pr.get("ev") or []
                for ev in evs[:2]:
                    # 原文定位单独一行，缩进区分
                    p = doc.add_paragraph()
                    pf = p.paragraph_format
                    pf.left_indent = Cm(1.5)
                    pf.space_before = Pt(0)
                    pf.space_after = Pt(1)
                    pf.line_spacing = 1.25
                    set_run(p.add_run("原文：" + ev), size=9, color=GREY)
        if it["fixes"]:
            label(doc, "怎么改")
            for fx in it["fixes"][:2]:
                bullet(doc, fx)

    ranked = sorted(issues, key=lambda x: -x["loss"])[:3]
    if ranked:
        h(doc, "三、优先补什么")
        for i, it in enumerate(ranked, 1):
            d = it["dim"]
            para(doc, "%d. %s（可追回 %s 分）" % (i, d["name"], it["loss"]),
                 size=10.5, bold=True, before=6, after=2, keep=True)
            para(doc, it["fixes"][0] if it["fixes"] else FALLBACK.get(d["id"], ""),
                 size=10, indent=0.55)

    ds = data.get("disputes") or []
    if ds:
        h(doc, "四、答辩可能被追问")
        for i, x in enumerate(ds[:4], 1):
            para(doc, "%d. %s" % (i, (x.get("point") or "").strip()), size=10, indent=0.55, after=2)
        if len(ds) > 4:
            para(doc, "（另有 %d 条，见 Excel 争议点表）" % (len(ds) - 4),
                 size=9, color=GREY, indent=0.55, after=2)

    para(doc, "分数为自动机检 + 证据定级结果，需人工复核后方可作为评审结论。",
         size=8.5, color=GREY, italic=True, before=12)

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
        # 带上原文，用于给"原文哪里不好"补充定位
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
            para(doc, "    %s　　　%s 分" % (r["name"], r["ai_score"]), size=9.5, after=1)

    for r in rows:
        doc.add_page_break()
        build(r, None, doc=doc)
    doc.save(str(out_path))
    return out_path


def main() -> int:
    ap = argparse.ArgumentParser(description="逐项目意见 → Word（简洁版）")
    ap.add_argument("src", help="含 *.score.json 的目录")
    ap.add_argument("--out", default="意见Word")
    ap.add_argument("--merge", action="store_true", help="另存一份合并总册")
    a = ap.parse_args()
    src, out = Path(a.src), Path(a.out)
    rows = load(src)
    if not rows:
        print("没找到 score.json", file=__import__("sys").stderr)
        return 1
    out.mkdir(parents=True, exist_ok=True)
    for r in rows:
        build(r, out / (r["_stem"] + ".docx"))
    if a.merge:
        build_merged(rows, out / "北斗奖评分意见汇总.docx")
    print(json.dumps({"out": str(out), "count": len(rows), "merged": bool(a.merge)},
                     ensure_ascii=False))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())