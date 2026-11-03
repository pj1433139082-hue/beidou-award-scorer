#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""评分结果 → Excel。

只有两张表：
  汇总       一行一份材料：总分、区间、各维度得分、首要修改建议；含小计与总计
  修改建议   维度明细 + 争议点 + 逐条改法合并成这一张：每条含
             编号 / 材料 / 维度 / 问题 / 原文是什么 / 怎么改 / 修改之后的句子

建议条目来自 advice_engine.build_items()，与 Word 意见书同一来源，
因此两边编号、顺序、条数完全一致。
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

sys.path.insert(0, str(Path(__file__).resolve().parent))
from advice_engine import build_items  # noqa: E402

INK = "1A1A1A"
MUTED = "5A6472"
HEAD_BG = "0B6B57"
BAND_BG = "F5F7F9"
WARN_BG = "FEF3C7"
LINE = "D9DEE5"

thin = Side(style="thin", color=LINE)
BORDER = Border(left=thin, right=thin, top=thin, bottom=thin)

F_TITLE = Font(name="Microsoft YaHei", size=14, bold=True, color=INK)
F_HEAD = Font(name="Microsoft YaHei", size=10.5, bold=True, color="FFFFFF")
F_CELL = Font(name="Microsoft YaHei", size=10, color=INK)
F_SMALL = Font(name="Microsoft YaHei", size=9, color=MUTED)
F_MONO = Font(name="Consolas", size=9, color=INK)

A_LEFT = Alignment(horizontal="left", vertical="top", wrap_text=True)
A_MID = Alignment(horizontal="center", vertical="center", wrap_text=True)


def band_font(band: str):
    if band.startswith("否决"):
        return Font(name="Microsoft YaHei", size=10, bold=True, color="B91C1C")
    for k, color in {"一等奖": "15803D", "二等奖": "1D4ED8", "三等奖": "A16207"}.items():
        if k in band:
            return Font(name="Microsoft YaHei", size=10, bold=True, color=color)
    if "待提升" in band:
        return Font(name="Microsoft YaHei", size=10, color="9A3412")
    return Font(name="Microsoft YaHei", size=10, color=INK)


def short(t, n=60):
    t = re.sub(r"\s+", " ", str(t or "")).strip()
    return t if len(t) <= n else t[: n - 1] + "…"


def strip_award(name: str) -> str:
    s = re.sub(r"^(金点子|金刚钻|金引擎|金智囊|金火炬|特别奖-[^\s\-－—]*)[-－—\s]*", "",
               str(name or "").strip())
    return re.sub(r"^[-－—\s]+", "", s).strip() or str(name or "")


def collect(src: Path) -> list[dict]:
    out = []
    for sf in sorted(src.glob("*.score.json")):
        try:
            d = json.loads(sf.read_text(encoding="utf-8"))
        except Exception:
            continue
        if "dimensions" not in d:
            continue
        stem = sf.name.replace(".score.json", "")
        cf = src / (stem + ".checks.json")
        d["_checks"] = {c["id"]: c for c in
                        (json.loads(cf.read_text(encoding="utf-8")).get("checks", [])
                         if cf.is_file() else [])}
        df = src / (stem + ".doc.json")
        d["_plain"] = (json.loads(df.read_text(encoding="utf-8")).get("plain", "")
                       if df.is_file() else "")
        d["_stem"] = stem
        try:
            d["_items"] = build_items(d)
        except Exception:
            d["_items"] = []
        out.append(d)
    return out


def build_summary(wb, rows):
    """汇总表：材料 / 总分 / 区间 / 维度…… / 首要修改建议。"""
    ws = wb.create_sheet("汇总")
    ws.sheet_view.showGridLines = False
    ws["A1"] = "北斗奖申报材料评分汇总"
    ws["A1"].font = F_TITLE
    ws["A2"] = "分数区间为相对刻度（85/75/65/55），非官方切线；机检结论需回原文确认"
    ws["A2"].font = F_SMALL

    COL_NAME, COL_SCORE, COL_BAND = 1, 2, 3
    COL_DIM0 = 4
    MAXDIM = max((len(d["dimensions"]) for d in rows), default=3)
    COL_ADV = COL_DIM0 + MAXDIM
    LAST = COL_ADV

    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=LAST)
    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=LAST)

    widths = {COL_NAME: 44, COL_SCORE: 8, COL_BAND: 12}
    for i in range(MAXDIM):
        widths[COL_DIM0 + i] = 11
    widths[COL_ADV] = 54
    for c, w in widths.items():
        ws.column_dimensions[get_column_letter(c)].width = w

    hr = 4
    for c, h in ((COL_NAME, "申报材料"), (COL_SCORE, "总分"), (COL_BAND, "奖项区间"),
                 (COL_ADV, "首要修改建议（编号见「修改建议」表）")):
        cell = ws.cell(hr, c, h)
        cell.font = F_HEAD; cell.fill = PatternFill("solid", fgColor=HEAD_BG)
        cell.alignment = A_MID; cell.border = BORDER
    if MAXDIM >= 1:
        ws.merge_cells(start_row=hr, start_column=COL_DIM0, end_row=hr,
                       end_column=COL_DIM0 + MAXDIM - 1)
        dc = ws.cell(hr, COL_DIM0, "维度得分（满分）")
        dc.font = F_HEAD; dc.alignment = A_MID
    for c in range(COL_DIM0, COL_DIM0 + MAXDIM):
        ws.cell(hr, c).fill = PatternFill("solid", fgColor=HEAD_BG)
        ws.cell(hr, c).border = BORDER
    ws.row_dimensions[hr].height = 26

    groups = []
    for r in rows:
        a = r["award"]
        g = next((x for x in groups if x["award"] == a), None)
        if g is None:
            g = {"award": a, "dims": [d["name"] for d in r["dimensions"]], "rows": []}
            groups.append(g)
        g["rows"].append(r)

    rw = hr + 1
    all_scores = []
    for g in groups:
        ws.merge_cells(start_row=rw, start_column=COL_NAME, end_row=rw, end_column=COL_BAND)
        gc = ws.cell(rw, COL_NAME, "%s　·　%d 份" % (g["award"], len(g["rows"])))
        gc.font = Font(name="Microsoft YaHei", size=10.5, bold=True, color="FFFFFF")
        gc.alignment = Alignment(horizontal="left", vertical="center", indent=1)
        for c in range(COL_NAME, LAST + 1):
            ws.cell(rw, c).fill = PatternFill("solid", fgColor=HEAD_BG)
            ws.cell(rw, c).border = BORDER
        for i, dname in enumerate(g["dims"][:MAXDIM]):
            cell = ws.cell(rw, COL_DIM0 + i, dname)
            cell.font = Font(name="Microsoft YaHei", size=9.5, bold=True, color="FFFFFF")
            cell.alignment = A_MID
        for i in range(len(g["dims"]), MAXDIM):
            ws.cell(rw, COL_DIM0 + i, "")
        ws.row_dimensions[rw].height = 22
        rw += 1

        for r in g["rows"]:
            disp = strip_award(r["name"])
            c1 = ws.cell(rw, COL_NAME, short(disp, 50))
            c1.font = F_CELL
            c1.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True, indent=1)
            c2 = ws.cell(rw, COL_SCORE, r["ai_score"])
            c2.font = Font(name="Microsoft YaHei", size=11, bold=True, color=INK)
            c2.alignment = A_MID
            band_txt = r["band"]
            if "暂列" in band_txt:
                band_txt = "待裁定"
            elif len(band_txt) > 8:
                band_txt = band_txt[:7] + "…"
            c3 = ws.cell(rw, COL_BAND, band_txt)
            c3.font = band_font(r["band"]); c3.alignment = A_MID

            for i, d in enumerate(r["dimensions"][:MAXDIM]):
                cell = ws.cell(rw, COL_DIM0 + i, "%s/%s" % (d["score"], d["max"]))
                cell.font = F_CELL; cell.alignment = A_MID
                ratio = d["score"] / d["max"] if d["max"] else 1
                if ratio < 0.6:
                    cell.fill = PatternFill("solid", fgColor=WARN_BG)
                elif ratio >= 0.85:
                    cell.fill = PatternFill("solid", fgColor="E8F5E9")
            for i in range(len(r["dimensions"]), MAXDIM):
                blank = ws.cell(rw, COL_DIM0 + i, "—")
                blank.font = Font(name="Microsoft YaHei", size=9, color=MUTED)
                blank.alignment = A_MID

            items = r.get("_items") or []
            if items:
                first = items[0]
                mi = rows.index(r) + 1
                adv = "%d-%s [%s] %s" % (mi, first["no"], first["dim"], short(first["how"], 36))
            else:
                adv = "—"
            ac = ws.cell(rw, COL_ADV, adv)
            ac.font = Font(name="Microsoft YaHei", size=9, color=MUTED)
            ac.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True, indent=1)

            for c in range(COL_NAME, LAST + 1):
                ws.cell(rw, c).border = BORDER
            est = len(disp)
            ws.row_dimensions[rw].height = 24 if est <= 26 else (34 if est <= 44 else 44)
            all_scores.append(r["ai_score"])
            rw += 1

        sc = [x["ai_score"] for x in g["rows"]]
        if sc:
            ws.merge_cells(start_row=rw, start_column=COL_NAME, end_row=rw, end_column=COL_BAND)
            label = ws.cell(rw, COL_NAME, "%s　小计：%d 份，均分 %.1f" % (
                g["award"], len(sc), sum(sc) / len(sc)))
            label.font = Font(name="Microsoft YaHei", size=9.5, bold=True, color=INK)
            label.alignment = Alignment(horizontal="right", vertical="center", indent=1)
            for c in range(COL_NAME, LAST + 1):
                ws.cell(rw, c).fill = PatternFill("solid", fgColor=BAND_BG)
                ws.cell(rw, c).border = BORDER
            ws.row_dimensions[rw].height = 20
            rw += 1
        rw += 1

    if all_scores:
        ws.merge_cells(start_row=rw, start_column=COL_NAME, end_row=rw, end_column=COL_BAND)
        tc = ws.cell(rw, COL_NAME, "总计　%d 份材料　平均 %.1f 分　最高 %.1f　最低 %.1f" % (
            len(all_scores), sum(all_scores) / len(all_scores),
            max(all_scores), min(all_scores)))
        tc.font = Font(name="Microsoft YaHei", size=10, bold=True, color="FFFFFF")
        tc.alignment = Alignment(horizontal="right", vertical="center", indent=1)
        for c in range(COL_NAME, LAST + 1):
            ws.cell(rw, c).fill = PatternFill("solid", fgColor=HEAD_BG)
            ws.cell(rw, c).border = BORDER
        ws.row_dimensions[rw].height = 22

    ws.freeze_panes = ws.cell(hr + 1, COL_NAME)
    return ws


def build_advice_sheet(wb, rows):
    """维度明细 + 争议点 + 逐条改法，合并成一张表。"""
    ws = wb.create_sheet("修改建议")
    ws.sheet_view.showGridLines = False
    ws["A1"] = "修改建议（维度明细 + 争议点 + 逐条改法）"
    ws["A1"].font = F_TITLE
    ws["A2"] = "每条按「原文是什么 → 怎么改 → 修改之后的句子」给出，可直接复制；编号与 Word 意见书一一对应"
    ws["A2"].font = F_SMALL

    head = ["编号", "申报材料", "维度", "问题", "原文是什么", "怎么改", "修改之后的句子"]
    widths = [9, 34, 12, 28, 40, 44, 46]
    LAST = len(head)
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=LAST)
    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=LAST)
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w

    hr = 4
    for c, h in enumerate(head, 1):
        cell = ws.cell(hr, c, h)
        cell.font = F_HEAD; cell.fill = PatternFill("solid", fgColor=HEAD_BG)
        cell.alignment = A_MID; cell.border = BORDER
    ws.row_dimensions[hr].height = 26

    rw = hr + 1
    # 编号加材料序号前缀（1-M01），保证整张表里唯一；Word 里用同一套编号
    for mi, r in enumerate(rows, 1):
        mat = strip_award(r["name"])
        items = r.get("_items") or []
        for it in items:
            no = "%d-%s" % (mi, it.get("no", ""))
            vals = [no, mat if len(mat) <= 32 else mat[:31] + "…", it.get("dim", ""),
                    it.get("title", ""), it.get("original", ""),
                    it.get("how", ""), it.get("revised", "")]
            for c, v in enumerate(vals, 1):
                cell = ws.cell(rw, c, v)
                cell.font = F_MONO if c == 1 else F_CELL
                cell.alignment = A_MID if c in (1, 3) else A_LEFT
                cell.border = BORDER
            lines = max(
                max(1, int(len(str(vals[3])) / 12)),
                max(1, int(len(str(vals[4])) / 18)),
                max(1, int(len(str(vals[5])) / 20)),
                max(1, int(len(str(vals[6])) / 21)),
            )
            ws.row_dimensions[rw].height = min(150, max(30, lines * 15))
            rw += 1

        ds = r.get("disputes") or []
        if ds:
            ws.merge_cells(start_row=rw, start_column=1, end_row=rw, end_column=LAST)
            bc = ws.cell(rw, 1, "%s　答辩可能被追问（争议点 %d 条）" % (short(mat, 34), len(ds)))
            bc.font = Font(name="Microsoft YaHei", size=9.5, bold=True, color=MUTED)
            bc.alignment = Alignment(horizontal="left", vertical="center", indent=1)
            for c in range(1, LAST + 1):
                ws.cell(rw, c).fill = PatternFill("solid", fgColor=BAND_BG)
                ws.cell(rw, c).border = BORDER
            ws.row_dimensions[rw].height = 18
            rw += 1
            for i, x in enumerate(ds, 1):
                point = str(x.get("point") or "").strip()
                src_ = str(x.get("source") or "").strip()
                vals = ["%d-Q%02d" % (mi, i), mat if len(mat) <= 32 else mat[:31] + "…",
                        src_, point, "—", "答辩时按此追问", "—"]
                for c, v in enumerate(vals, 1):
                    cell = ws.cell(rw, c, v)
                    cell.font = F_MONO if c == 1 else F_SMALL
                    cell.alignment = A_MID if c in (1, 3) else A_LEFT
                    cell.border = BORDER
                ws.row_dimensions[rw].height = min(80, max(24, int(len(point) / 30) * 15 + 15))
                rw += 1

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
    build_advice_sheet(wb, rows)
    wb.save(a.out)
    total_items = sum(len(r.get("_items") or []) for r in rows)
    print(json.dumps({"out": a.out, "sheets": ["汇总", "修改建议"],
                      "materials": len(rows), "advice_items": total_items},
                     ensure_ascii=False))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
