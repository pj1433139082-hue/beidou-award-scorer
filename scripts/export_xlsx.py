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


def band_fill(band: str) -> str | None:
    if band.startswith("否决"):
        return BAD_BG
    if "一等奖" in band:
        return GOOD_BG
    if "二等奖" in band:
        return "E0F2FE"
    if "三等奖" in band:
        return "FEF9C3"
    if "待提升" in band:
        return BAD_BG
    return None


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
        out.append(d)
    return out


def build_summary(wb, rows):
    ws = wb.create_sheet("汇总")
    ws.sheet_view.showGridLines = False
    ws["A1"] = "北斗奖申报材料评分汇总"
    ws["A1"].font = F_TITLE
    ws["A2"] = "分数区间为相对刻度（85/75/65/55），非官方切线；机检结论需回原文确认"
    ws["A2"].font = F_SMALL
    ws.merge_cells("A1:H1"); ws.merge_cells("A2:H2")

    # 按奖项分组：同一奖项用同一套维度列，行间不再出现大片"—"
    groups = []
    for r in rows:
        a = r["award"]
        g = next((x for x in groups if x["award"] == a), None)
        if g is None:
            g = {"award": a, "dims": [d["name"] for d in r["dimensions"]], "rows": []}
            groups.append(g)
        g["rows"].append(r)

    hr = 4
    head = ["申报材料", "奖项", "总分", "奖项区间", "置信度"] + [""] * 0
    for c, h in enumerate(head, 1):
        cell = ws.cell(hr, c, h)
        cell.font = F_HEAD; cell.fill = PatternFill("solid", fgColor=HEAD_BG)
        cell.alignment = A_MID; cell.border = BORDER

    rw = hr + 1
    for gi, g in enumerate(groups):
        # 奖项分组的维度表头行
        for c, dname in enumerate(g["dims"], 6):
            cell = ws.cell(rw, c, dname)
            cell.font = F_HEAD; cell.fill = PatternFill("solid", fgColor=HEAD_BG)
            cell.alignment = A_MID; cell.border = BORDER
        ws.cell(rw, 1, "%s（%d 份）" % (g["award"], len(g["rows"])))
        ws.cell(rw, 1).font = F_BOLD
        ws.cell(rw, 1).alignment = A_LEFT
        ws.cell(rw, 1).fill = PatternFill("solid", fgColor=BAND_BG)
        for c in range(2, 6):
            ws.cell(rw, c).fill = PatternFill("solid", fgColor=BAND_BG)
            ws.cell(rw, c).border = BORDER
        ws.row_dimensions[rw].height = 22
        rw += 1

        for ri, r in enumerate(g["rows"]):
            ws.cell(rw, 1, short(r["name"], 30)).alignment = A_LEFT
            ws.cell(rw, 2, "").alignment = A_MID      # 奖项已在分组标题行给出
            ws.cell(rw, 3, r["ai_score"]).alignment = A_MID
            ws.cell(rw, 4, short(r["band"], 18)).alignment = A_MID
            ws.cell(rw, 5, r["confidence"]).alignment = A_MID
            for c, dname in enumerate(g["dims"], 6):
                d = next((x for x in r["dimensions"] if x["name"] == dname), None)
                cell = ws.cell(rw, c, ("%s/%s" % (d["score"], d["max"])) if d else "—")
                cell.alignment = A_MID
                if d and d["score"] < d["max"] * 0.6:
                    cell.fill = PatternFill("solid", fgColor=WARN_BG)
            risks = []
            if r.get("veto"):
                risks.append("否决")
            if r.get("veto_pending"):
                risks.append("否决待裁定")
            ncap = sum(1 for c in r.get("caps_and_gates", []) if c.get("type") == "cap")
            if ncap:
                risks.append("封顶%d" % ncap)
            if r.get("deductions"):
                risks.append("扣%d" % sum(x["points"] for x in r["deductions"]))
            cell = ws.cell(rw, 6 + len(g["dims"]), "、".join(risks) or "—")
            cell.alignment = A_MID
            bf = band_fill(r["band"])
            if bf:
                ws.cell(rw, 4).fill = PatternFill("solid", fgColor=bf)
            if risks:
                cell.fill = PatternFill("solid", fgColor=WARN_BG)
            for c in range(1, 7 + len(g["dims"])):
                ws.cell(rw, c).font = F_CELL
                ws.cell(rw, c).border = BORDER
            rw += 1
        rw += 1   # 组间留空

    maxdim = max((len(g["dims"]) for g in groups), default=0)
    widths = [32, 10, 8, 20, 9] + [13] * maxdim + [18]
    for c, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(c)].width = w
    ws.row_dimensions[hr].height = 24
    return ws


def build_detail(wb, rows):
    ws = wb.create_sheet("维度明细")
    ws.sheet_view.showGridLines = False
    ws["A1"] = "维度得分明细"
    ws["A1"].font = F_TITLE
    ws["A2"] = "「失分原因」= 封顶（材料证据不足）或等级下调；引文为原文逐字片段，未通过核验的会标出"
    ws["A2"].font = F_SMALL
    ws.merge_cells("A1:H1"); ws.merge_cells("A2:H2")

    head = ["申报材料", "维度", "满分", "得分", "得分率", "证据等级", "失分原因", "原文引文"]
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
            vals = [
                short(r["name"], 28), d["name"], d["max"], d["score"],
                ("%.0f%%" % (d["score"] / d["max"] * 100)) if d["max"] else "—",
                d.get("verified_grade", "—"),
                "；".join(reasons) or "—",
                short(quotes, 80) if quotes else "—",
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
            rw += 1

    for c, w in enumerate([28, 12, 7, 7, 8, 9, 42, 60], 1):
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
            vals = [short(r["name"], 28), dsp.get("source", "—"), short(dsp.get("point", ""), 90), short(ev, 50) or "—"]
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

    for c, w in enumerate([28, 16, 62, 36], 1):
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