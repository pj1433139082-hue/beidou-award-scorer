#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""把评分结果导成 Excel（给领导看的精简版）。

输入：score.json（一或多份）+ leaderboard
输出：.xlsx，三个工作表
  Sheet1 汇总      —— 一行一份材料：分数 / 区间 / 置信度 / 各维度分 / 风险标记
  Sheet2 维度明细  —— 一行一个维度：得分、失分原因、封顶、引文
  Sheet3 争议点    —— 一行一个争议点：来源、问题、答辩要追问什么

用法：
  python scripts/export_xlsx.py out/ --out 排行榜.xlsx
  python scripts/export_xlsx.py out/材料.score.json --out 单份.xlsx
  python scripts/export_xlsx.py out/ --out 榜.xlsx --award 金火炬   # 只看某奖项
"""
from __future__ import annotations
import argparse, json, re, sys
from pathlib import Path

try:
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
    from openpyxl.utils import get_column_letter
except ImportError:
    print("缺少 openpyxl。先运行：python scripts/setup.py --install", file=sys.stderr)
    raise SystemExit(2)

# ---- 视觉规范（克制：只有表头有底色，数字列居中，长文本自动换行）----
INK = "1A1A1A"
MUTED = "5A6472"
HEAD_BG = "0B6B57"      # 深绿：表头
HEAD_FG = "FFFFFF"
BAND_BG = "F5F7F9"     # 斑马纹
WARN_BG = "FEF3C7"     # 需注意
BAD_BG = "FEE2E2"      # 否决/风险
GOOD_BG = "DCFCE7"     # 良好
LINE = "D9DEE5"

thin = Side(style="thin", color=LINE)
BORDER = Border(left=thin, right=thin, top=thin, bottom=thin)

F_TITLE = Font(name="Microsoft YaHei", size=14, bold=True, color=INK)
F_HEAD = Font(name="Microsoft YaHei", size=10.5, bold=True, color=HEAD_FG)
F_CELL = Font(name="Microsoft YaHei", size=10, color=INK)
F_SMALL = Font(name="Microsoft YaHei", size=9, color=MUTED)
F_BOLD = Font(name="Microsoft YaHei", size=10, bold=True, color=INK)

A_LEFT = Alignment(horizontal="left", vertical="center", wrap_text=True)
A_MID = Alignment(horizontal="center", vertical="center", wrap_text=True)
A_TOP = Alignment(horizontal="left", vertical="top", wrap_text=True)

BAND_ORDER = ["一等奖候选", "二等奖候选", "三等奖候选", "入围候选", "待提升"]

# ---------- 修改建议生成规则 ----------
# 键：维度 id + 失分原因关键词 -> 建议文案
# 面向申报人：说清"补什么"和"补了能加多少分"
ADVICE = {
    "value": [
        ("C02", "把效益表里的「实际值【已实现】」和「远期目标【规划/在研】」分开填：已兑现的填实际值并附台账截图，没兑现的只填远期目标。现在两者混填，价值性封顶 40%。"),
        ("C04", "每个已实现效益都要有佐证材料：台账截图、运行日志或核算表，填进「佐证材料（附截图）」列。覆盖率低于 80% 价值性封顶 70%。"),
        ("C03", "每项效益补上改造前的基准值和对比周期，否则变化率算不出来，别人无法复算。"),
        ("C05", "正文里的百分比要和表格里的数字对得上（偏差不超过 5 个百分点），否则会被判为数字不自洽。"),
        ("C16", "把「预计/有望」的表述换成实际发生的量，量化成效描述要能落到具体金额或比例。"),
    ],
    "replication": [
        ("C14", "补一段与原办法的量化对照：改造前后/两种方案的同一指标对比，给出百分比或倍数。"),
        ("C02", "补充可推广范围的实际数据：已应用厂站数、覆盖区域数，而不是只写定性描述。"),
    ],
    "novelty": [
        ("C14", "补与原有做法的量化对照：同一场景下新旧方案的效果差值（成本、精度、时效），并说明差异来自哪个创新点。"),
    ],
    "intelligence": [
        ("C11", "补模型实验室的线上部署记录、试点厂的生产调用记录，以及不少于 90 天的运行验证数据；三者齐备才能拿满智能化水平分。"),
    ],
    "promotion": [
        ("C06", "补一份推广清单：大区 / 区域 / 水厂逐个列出（要能数出个数），正文写清覆盖了几个大区、几个区域、几个项目。"),
    ],
    "economics": [
        ("C07", "补「计算步骤」表：每一步写清 Excel 公式、计算结果、单位和含义。只给一个总金额无法复算，经济效益封顶。"),
        ("C08", "同一个金额若在多处出现，确认不是重复计入（如吨水单耗和总费用各算一次）。"),
    ],
    "readiness": [
        ("C12", "写清成果从方案到产品的转化程度：已产品化（有合同/订单/收入）→ 非常符合；只有原型 → 符合；仍在方案 → 一般。补交付形态说明。"),
    ],
    "usability": [
        ("C15", "补集团战略对齐说明（对应十五五/集团战略方向）和当前卡点的具体描述，说明方案如何直接解决该卡点。"),
    ],
}

GLOBAL_ADVICE = [
    ("C09", "把没填的占位符（下划线、XXXX、（文本））全部补实，未填项会被判为材料未完成。"),
    ("C17", "插入关键运行界面、台账、报告的截图并加说明，只放图不放说明不构成证据。"),
]


# 按机检编号 → 建议文案。键是 checks.json 里的 Cxx 编号。
ADVICE_BY_CHECK = {
    "C02": "把效益表的「实际值【已实现】」与「远期目标【规划/在研】」分开填：已兑现的填实际值并附台账截图，未兑现的只填远期目标。两者混填，价值性最高只能拿 40%。",
    "C03": "每项效益补上改造前的基准值和对比周期。没有基线，变化率无法复算，等于没有说服力。",
    "C04": "每个已实现效益都要配佐证材料（台账截图 / 运行日志 / 核算表），填进「佐证材料（附截图）」列。覆盖率低于 80% 价值性封顶 70%。",
    "C05": "正文里的百分比要与表格数字一致（偏差不超过 5 个百分点），否则会被判为数字不自洽。",
    "C08": "同一金额若在多处出现，确认没有重复计入（如吨水单耗与总费用各算一次）。",
    "C09": "把没填的占位符（下划线、XXXX、（文本））全部补实，未填项会被判为材料未完成。",
    "C11": "补模型实验室线上部署记录、试点厂生产调用记录，以及不少于 90 天的运行验证；三者齐备才能拿满智能化水平分。",
    "C12": "写清从方案到产品的转化程度：已产品化（有合同/订单/收入）→非常符合；仅原型→符合；仍在方案→一般。",
    "C13": "补一段「相较上次参赛版本」的迭代说明：改了什么、指标提升多少。没有这段，成果迭代性这项会被判为否决项待裁定。",
    "C14": "补与原办法的量化对照：同一场景下新旧方案的效果差值（成本、精度、时效），并说明差异来自哪个创新点。",
    "C15": "补集团战略对齐说明（对应十五五 / 集团战略方向）与当前卡点的具体描述，说明方案如何直接解决该卡点。",
    "C16": "把「预计 / 有望」换成实际发生的量：量化成效要能落到具体金额或比例。",
    "C17": "插入关键运行界面、台账、报告截图并加说明；只放图不说明不构成证据。",
    "C06": "补一份推广清单：大区 / 区域 / 水厂逐个列出（要能数出个数），正文写清覆盖几个大区、几个区域、几个项目。",
    "C07": "补「计算步骤」表：每步写清公式、结果、单位、含义。只给总金额无法复算，经济效益拿不到高分。",
    "C18": "确认三个子项的证据齐全：推广案例佐证、经济测算过程、产品形态说明。",
    "C19": "补齐专属模板的六段：实施背景、模型应用、业务价值、ROI 核算、下一步计划、量化数据。",
    "C20": "ROI 要有可复算的投入与收益两段，并配足量化数字。",
}

# 维度 → 相关的机检编号（用于把检查结果映射到维度）
DIM_CHECKS = {
    "value": ["C02", "C03", "C04", "C05", "C08", "C16"],
    "replication": ["C06", "C14", "C09"],
    "novelty": ["C14", "C09"],
    "intelligence": ["C11", "C17"],
    "promotion": ["C06", "C18", "C07"],
    "economics": ["C07", "C08", "C05"],
    "readiness": ["C12"],
    "usability": ["C15"],
}


def build_advice(dim: dict, checks: list, gates: list, deductions: list) -> str:
    """按该维度命中的机检编号，给出可执行的修改建议（最多 3 条，按优先级）。

    编号来源（按可靠性排序）：
      1. caps_and_gates[].check —— 该维度的封顶由哪条机检造成（最准）
      2. deductions[].check     —— 固定扣分来自哪条机检
      3. checks.json 里 status 为 warn/fail 且 dims 含本维度的检查
    """
    dim_id = dim.get("id")
    hit = []
    for g in gates or []:
        if g.get("dim") == dim_id and g.get("check"):
            hit.append(g["check"])
    for d in deductions or []:
        if d.get("check"):
            hit.append(d["check"])
    for c in checks or []:
        if c.get("status") in ("warn", "fail"):
            cdims = c.get("dims") or []
            if dim_id in cdims:
                hit.append(c.get("id"))
            elif not cdims and c.get("id") in ("C09", "C17"):
                hit.append(c.get("id"))   # 全局项（占位符/截图）作为兜底提醒
    out = []
    for cid in hit:
        txt = ADVICE_BY_CHECK.get(cid)
        if txt and txt not in out:
            out.append(txt)
    # 证据等级被下调也算失分，给一条针对性建议
    if not out and dim.get("claimed_grade") != dim.get("verified_grade"):
        out.append("该维度引文没全部通过核验：把引文换成材料里的原句（逐字、不改写），"
                   "并补到 2 句以上，否则证据等级会被压到 C 及以下。")
    if not out:
        if dim.get("score", 0) < dim.get("max", 0) * 0.8:
            out.append("该维度论证偏弱：补具体数据与佐证，讲清「做了什么—产生了什么—怎么验证的」，"
                       "并给出一组可复算的数字。")
        else:
            out.append("该维度已基本达标，若要更稳，建议再补一条与原办法的量化对照。")
    return " ".join(out[:3])



def band_fill(band: str) -> str | None:
    """区间列不再大面积上色——只留很淡的提示色，避免整张表花掉。"""
    if band.startswith("否决"):
        return BAD_BG
    if "待提升" in band:
        return "FFF7ED"
    return None


def band_font(band: str):
    """区间用字体色区分，视觉更轻。"""
    if band.startswith("否决"):
        return Font(name="Microsoft YaHei", size=10, bold=True, color="B91C1C")
    if "一等奖" in band:
        return Font(name="Microsoft YaHei", size=10, bold=True, color="15803D")
    if "二等奖" in band:
        return Font(name="Microsoft YaHei", size=10, bold=True, color="1D4ED8")
    if "三等奖" in band:
        return Font(name="Microsoft YaHei", size=10, bold=True, color="A16207")
    if "待提升" in band:
        return Font(name="Microsoft YaHei", size=10, color="9A3412")
    return Font(name="Microsoft YaHei", size=10, color=INK)


def short(text: str, n: int = 60) -> str:
    t = re.sub(r"\s+", " ", str(text or "")).strip()
    return t if len(t) <= n else t[: n - 1] + "…"


def collect(src: Path) -> list[dict]:
    """收集 score.json。支持目录或单文件。"""
    files = sorted(src.glob("*.score.json")) if src.is_dir() else ([src] if src.is_file() else [])
    out = []
    for f in files:
        try:
            d = json.loads(f.read_text(encoding="utf-8"))
        except Exception:
            continue
        if "dimensions" not in d:
            continue
        d["_file"] = f
        cf = f.with_name(f.name.replace(".score.json", ".checks.json"))
        if cf.is_file():
            try:
                d["_checks"] = json.loads(cf.read_text(encoding="utf-8")).get("checks", [])
            except Exception:
                d["_checks"] = []
        else:
            d["_checks"] = []
        out.append(d)
    return out


def build_summary(wb, rows):
    """汇总表：固定 5 个维度槽位，所有奖项共用同一列结构。

    早期版本按各奖项维度数动态排布，导致金火炬（3 维）与金引擎（4 维）的建议列
    落在不同列上，列宽对不齐——看起来就是错位的。改为固定槽位后列宽恒定。
    """
    ws = wb.create_sheet("汇总")
    ws.sheet_view.showGridLines = False
    ws["A1"] = "北斗奖申报材料评分汇总"
    ws["A1"].font = F_TITLE
    ws["A2"] = "分数区间为相对刻度（85/75/65/55），非官方切线；机检结论需回原文确认"
    ws["A2"].font = F_SMALL

    MAXDIM = 5                      # 维度槽位上限（金引擎/金智囊 4 个，留 1 个余量）
    COL_NAME, COL_SCORE, COL_BAND, COL_CONF = 1, 2, 3, 4
    COL_DIM0 = 5                    # 维度从第 5 列开始
    COL_RISK = COL_DIM0 + MAXDIM    # 10
    COL_ADV = COL_RISK + 1          # 11
    LAST = COL_ADV

    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=LAST)
    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=LAST)

    # 固定列宽（与槽位一一对应，不随奖项变化）
    widths = {COL_NAME: 44, COL_SCORE: 8, COL_BAND: 12, COL_CONF: 7}
    for i in range(MAXDIM):
        widths[COL_DIM0 + i] = 12
    widths[COL_RISK] = 14
    widths[COL_ADV] = 64
    for c, w in widths.items():
        ws.column_dimensions[get_column_letter(c)].width = w

    # 全局表头
    hr = 4
    for c, h in ((COL_NAME, "申报材料"), (COL_SCORE, "总分"), (COL_BAND, "奖项区间"),
                 (COL_CONF, "置信度"), (COL_RISK, "需关注"), (COL_ADV, "修改建议")):
        cell = ws.cell(hr, c, h)
        cell.font = F_HEAD; cell.fill = PatternFill("solid", fgColor=HEAD_BG)
        cell.alignment = A_MID; cell.border = BORDER
    ws.row_dimensions[hr].height = 26

    # 按奖项分组
    groups = []
    for r in rows:
        a = r["award"]
        g = next((x for x in groups if x["award"] == a), None)
        if g is None:
            g = {"award": a, "dims": [d["name"] for d in r["dimensions"]], "rows": []}
            groups.append(g)
        g["rows"].append(r)

    rw = hr + 1
    for gi, g in enumerate(groups):
        # 分组标题行：奖项名占前 4 列，维度名从第 5 列开始
        ws.merge_cells(start_row=rw, start_column=COL_NAME, end_row=rw, end_column=COL_CONF)
        gc = ws.cell(rw, COL_NAME, "%s　·　%d 份" % (g["award"], len(g["rows"])))
        gc.font = Font(name="Microsoft YaHei", size=10.5, bold=True, color="FFFFFF")
        gc.alignment = Alignment(horizontal="left", vertical="center", indent=1)
        for c in range(COL_NAME, LAST + 1):
            ws.cell(rw, c).fill = PatternFill("solid", fgColor=HEAD_BG)
            ws.cell(rw, c).border = BORDER
        # 该组的维度名
        for i, dname in enumerate(g["dims"][:MAXDIM]):
            cell = ws.cell(rw, COL_DIM0 + i, dname)
            cell.font = Font(name="Microsoft YaHei", size=9.5, bold=True, color="FFFFFF")
            cell.alignment = A_MID
        ws.row_dimensions[rw].height = 22
        rw += 1

        for r in g["rows"]:
            # 材料名
            # 分组标题已写明奖项，材料名去掉前缀，避免重复占宽
            disp = re.sub(r"^(特别奖-AI融创[^-]*|金点子|金刚钻|金引擎|金智囊|金火炬)[-－—]?", "", r["name"])
            c1 = ws.cell(rw, COL_NAME, short(disp, 40))
            c1.font = F_CELL; c1.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True, indent=1)
            # 总分
            c2 = ws.cell(rw, COL_SCORE, r["ai_score"])
            c2.font = Font(name="Microsoft YaHei", size=11, bold=True, color=INK)
            c2.alignment = A_MID
            # 区间
            band_txt = r["band"]
            if "暂列" in band_txt:          # 否决待裁定的长文案，压成短标签
                band_txt = "待裁定"
            elif len(band_txt) > 8:
                band_txt = band_txt[:7] + "…"
            c3 = ws.cell(rw, COL_BAND, band_txt)
            c3.font = band_font(r["band"])
            c3.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            # 置信度
            c4 = ws.cell(rw, COL_CONF, r["confidence"])
            c4.font = F_CELL; c4.alignment = A_MID
            # 维度分（固定槽位）
            for i, d in enumerate(r["dimensions"][:MAXDIM]):
                cell = ws.cell(rw, COL_DIM0 + i, "%s/%s" % (d["score"], d["max"]))
                cell.font = F_CELL; cell.alignment = A_MID
                ratio = d["score"] / d["max"] if d["max"] else 1
                if ratio < 0.6:
                    cell.fill = PatternFill("solid", fgColor=WARN_BG)
                elif ratio >= 0.85:
                    cell.fill = PatternFill("solid", fgColor="E8F5E9")
            # 空槽位留白
            for i in range(len(r["dimensions"]), MAXDIM):
                ws.cell(rw, COL_DIM0 + i, "").border = BORDER
            # 需关注
            risks = []
            if r.get("veto"):
                risks.append("否决")
            if r.get("veto_pending"):
                risks.append("否决待裁定")
            ncap = sum(1 for x in r.get("caps_and_gates", []) if x.get("type") == "cap")
            if ncap:
                risks.append("封顶%d" % ncap)
            if r.get("deductions"):
                risks.append("扣%d" % sum(x["points"] for x in r["deductions"]))
            rc = ws.cell(rw, COL_RISK, "、".join(risks) or "—")
            rc.font = F_CELL; rc.alignment = A_MID
            if risks:
                rc.fill = PatternFill("solid", fgColor=WARN_BG)
            # 修改建议（失分最多的两个维度）
            items = []
            for d in sorted(r["dimensions"], key=lambda x: (x["max"] - x["score"]), reverse=True):
                if d["max"] - d["score"] <= 0.01:
                    continue
                a_txt = build_advice(d, r.get("_checks") or [], r.get("caps_and_gates") or [],
                                     r.get("deductions") or [])
                if not a_txt:
                    continue
                parts = re.split(r"(?<=。)", a_txt)
                first = next((p.strip() for p in parts if p.strip()), "")
                if len(first) > 38:
                    first = first[:37] + "…"
                items.append("[%s] %s" % (d["name"], first))
                if len(items) >= 1:
                    break
            ac = ws.cell(rw, COL_ADV, chr(10).join(items) or "—")
            ac.font = Font(name="Microsoft YaHei", size=9, color=MUTED)
            ac.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True, indent=1)

            for c in range(COL_NAME, LAST + 1):
                ws.cell(rw, c).border = BORDER
            # 行高按材料名长度估算（中文按 2 字宽计）
            est = len(disp) * 1.0
            ws.row_dimensions[rw].height = 24 if est <= 26 else (34 if est <= 44 else 44)
            rw += 1

        rw += 1   # 组间留空

    ws.freeze_panes = ws.cell(hr + 1, COL_NAME)
    return ws


def build_detail(wb, rows):
    ws = wb.create_sheet("维度明细")
    ws.sheet_view.showGridLines = False
    ws["A1"] = "维度得分明细"
    ws["A1"].font = F_TITLE
    ws["A2"] = "「失分原因」= 封顶（材料证据不足）或等级下调；引文为原文逐字片段，未通过核验的会标出"
    ws["A2"].font = F_SMALL
    ws.merge_cells("A1:H1"); ws.merge_cells("A2:H2")

    head = ["申报材料", "维度", "满分", "得分", "得分率", "证据等级", "失分原因", "原文引文", "修改建议"]
    hr = 4
    for c, h in enumerate(head, 1):
        cell = ws.cell(hr, c, h)
        cell.font = F_HEAD; cell.fill = PatternFill("solid", fgColor=HEAD_BG)
        cell.alignment = A_MID; cell.border = BORDER

    rw = hr + 1
    for r in rows:
        for d in r["dimensions"]:
            reasons = []
            if d.get("cap_reason"):
                reasons.append(d["cap_reason"])
            # 该维度命中的机检（即使没造成封顶，也要显示）
            for c in r.get("_checks") or []:
                if c.get("status") in ("warn", "fail") and d["id"] in (c.get("dims") or []):
                    reasons.append("%s %s" % (c.get("id"), c.get("name")))
            if d.get("claimed_grade") != d.get("verified_grade"):
                reasons.append("证据等级 %s→%s（引文未全部核验）" % (d.get("claimed_grade"), d.get("verified_grade")))
            quotes = "；".join(d.get("quotes_verified") or [])
            rejected = d.get("quotes_rejected") or []
            if rejected:
                # quotes_rejected 是 [{quote, why}] 结构，不是字符串列表
                rej_txt = "；".join(
                    ("%s（%s）" % (x.get("quote", ""), x.get("why", ""))) if isinstance(x, dict) else str(x)
                    for x in rejected)
                quotes = (quotes + "；" if quotes else "") + "【未通过核验】" + rej_txt
            advice_txt = build_advice(d, r.get("_checks") or [], r.get("caps_and_gates") or [], r.get("deductions") or [])
            disp_name = re.sub(r"^(特别奖-AI融创[^-]*|金点子|金刚钻|金引擎|金智囊|金火炬)[-－—]?", "", r["name"])
            vals = [
                short(disp_name, 26), d["name"], d["max"], d["score"],
                ("%.0f%%" % (d["score"] / d["max"] * 100)) if d["max"] else "—",
                d.get("verified_grade", "—"),
                short("；".join(reasons), 46) or "—",
                short(quotes, 55) if quotes else "—",
                short(advice_txt, 150) or "—",
            ]
            for c, v in enumerate(vals, 1):
                cell = ws.cell(rw, c, v)
                cell.font = F_CELL; cell.border = BORDER
                cell.alignment = A_MID if c in (3, 4, 5, 6) else A_LEFT
            rate = d["score"] / d["max"] if d["max"] else 1
            if rate < 0.6:
                ws.cell(rw, 4).fill = PatternFill("solid", fgColor=WARN_BG)
            if d.get("quotes_rejected"):
                ws.cell(rw, 8).fill = PatternFill("solid", fgColor=BAD_BG)
            # 行高：按每列"字数 ÷ 列宽"估换行数，取最大值（中文按 1 字≈1 单位宽）
            col_w = [30, 10, 6, 6, 7, 8, 32, 40, 62]
            max_lines = 1
            for vi, v in enumerate(vals):
                w = col_w[vi] if vi < len(col_w) else 20
                lines_needed = max(1, int(len(str(v)) / max(w - 1, 4)) + 1)
                max_lines = max(max_lines, lines_needed)
            ws.row_dimensions[rw].height = min(96, max(20, max_lines * 13))
            rw += 1

    for c, w in enumerate([30, 10, 6, 6, 7, 8, 32, 40, 62], 1):
        ws.column_dimensions[get_column_letter(c)].width = w
    ws.row_dimensions[hr].height = 26
    ws.freeze_panes = ws.cell(hr + 1, 1)
    return ws


def build_disputes(wb, rows):
    ws = wb.create_sheet("争议点")
    ws.sheet_view.showGridLines = False
    ws["A1"] = "争议点清单（答辩追问用）"
    ws["A1"].font = F_TITLE
    ws["A2"] = "这些是材料里含糊、缺证据或前后不一致的地方，答辩时按此逐条追问"
    ws["A2"].font = F_SMALL
    ws.merge_cells("A1:D1"); ws.merge_cells("A2:D2")

    head = ["申报材料", "来源", "争议内容", "佐证 / 定位"]
    hr = 4
    for c, h in enumerate(head, 1):
        cell = ws.cell(hr, c, h)
        cell.font = F_HEAD; cell.fill = PatternFill("solid", fgColor=HEAD_BG)
        cell.alignment = A_MID; cell.border = BORDER

    rw = hr + 1
    n = 0
    for r in rows:
        for dsp in r.get("disputes") or []:
            ev = "；".join(str(x) for x in (dsp.get("evidence") or []))
            disp_name = re.sub(r"^(特别奖-AI融创[^-]*|金点子|金刚钻|金引擎|金智囊|金火炬)[-－—]?", "", r["name"])
            vals = [short(disp_name, 26), dsp.get("source", "—"), short(dsp.get("point", ""), 90), short(ev, 60) or "—"]
            for c, v in enumerate(vals, 1):
                cell = ws.cell(rw, c, v)
                cell.font = F_CELL; cell.border = BORDER
                cell.alignment = A_MID if c == 2 else A_LEFT
            if "否决" in str(dsp.get("source", "")):
                for c in range(1, 5):
                    ws.cell(rw, c).fill = PatternFill("solid", fgColor=BAD_BG)
            elif n % 2 == 1:
                for c in range(1, 5):
                    ws.cell(rw, c).fill = PatternFill("solid", fgColor=BAND_BG)
            rw += 1; n += 1

    if not n:
        ws.cell(hr + 1, 1, "无争议点。").font = F_CELL
        ws.merge_cells(start_row=hr + 1, start_column=1, end_row=hr + 1, end_column=4)

    for c, w in enumerate([30, 14, 62, 40], 1):
        ws.column_dimensions[get_column_letter(c)].width = w
    ws.row_dimensions[hr].height = 26
    ws.freeze_panes = ws.cell(hr + 1, 1)
    return ws


def main() -> int:
    ap = argparse.ArgumentParser(description="评分结果 → Excel")
    ap.add_argument("src", help="score.json 文件或含多个 score.json 的目录")
    ap.add_argument("--out", default="评分汇总.xlsx")
    ap.add_argument("--award", help="只导出该奖项")
    a = ap.parse_args()

    rows = collect(Path(a.src))
    if a.award:
        rows = [r for r in rows if r.get("award") == a.award]
    if not rows:
        print("没找到 score.json。用法：先跑 score.py 或 batch.py。", file=sys.stderr)
        return 1
    rows.sort(key=lambda r: r.get("ai_score", 0), reverse=True)

    wb = Workbook()
    wb.remove(wb.active)
    build_summary(wb, rows)
    build_detail(wb, rows)
    build_disputes(wb, rows)
    wb.save(a.out)
    print(json.dumps({"out": a.out, "sheets": ["汇总", "维度明细", "争议点"],
                      "materials": len(rows)}, ensure_ascii=False))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())