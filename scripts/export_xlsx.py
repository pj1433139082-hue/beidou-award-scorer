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
from advice_engine import build_items, total_gain  # noqa: E402

INK = "1A1A1A"
MUTED = "5A6472"
HEAD_BG = "0B6B57"
BAND_BG = "F5F7F9"
WARN_BG = "FEF3C7"
LINE = "D9DEE5"
ACCENT = "0B6B57"

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
    """读一批评分结果，连机检、原文、建议一起算好。

    迁移到 MaterialRecord：键的选择、缺失字段语义、私有字段都归它管，
    这里只负责把结果摊平成导出器要的 dict（见 as_row）。
    """
    from material_record import load_dir
    rows = []
    for rec in load_dir(src, with_context=True):
        try:
            rec.advice = build_items(rec.as_row())
        except Exception:
            rec.advice = []
        rows.append(rec.as_row())
    return rows


def build_sheet(wb, rows):
    """唯一的一张表：每份材料一行主行，其下跟着该材料的建议子行。

    版式（同一张表，不另开 sheet）：
        主行   申报材料 | 总分 | 区间 | 维度得分×N | 预计改完提升 | 改完后总分 | 修改建议
        子行   （合并前 3 列，写编号+维度+问题）… 原文是什么 / 怎么改 / 改后句子
        分隔行 合并整行写"答辩可能被追问"块
        小计行 各奖项组末
        总计行 全表末
    """
    ws = wb.create_sheet("评分汇总")
    ws.sheet_view.showGridLines = False
    ws["A1"] = "北斗奖申报材料评分汇总"
    ws["A1"].font = F_TITLE
    ws["A2"] = ("分数区间为相对刻度（85/75/65/55），非官方切线；机检结论需回原文确认。"
                "「改完提升」= 各维度失分 × 证据等级回收集（A 90% / B 60% / C 50% / D 30%），"
                "同一维度只计一次，补齐材料不代表满分。")
    ws["A2"].font = F_SMALL

    COL_NAME, COL_SCORE, COL_BAND = 1, 2, 3
    COL_DIM0 = 4
    MAXDIM = max((len(d["dimensions"]) for d in rows), default=3)
    COL_GAIN = COL_DIM0 + MAXDIM
    COL_AFTER = COL_GAIN + 1
    # 建议区 3 列：原文 / 怎么改 / 改后句子
    COL_ORIG = COL_AFTER + 1
    COL_HOW = COL_ORIG + 1
    COL_FIX = COL_HOW + 1
    LAST = COL_FIX

    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=LAST)
    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=LAST)

    widths = {COL_NAME: 38, COL_SCORE: 7, COL_BAND: 10}
    for i in range(MAXDIM):
        widths[COL_DIM0 + i] = 9.5
    widths[COL_GAIN] = 13
    widths[COL_AFTER] = 11
    widths[COL_ORIG] = 38
    widths[COL_HOW] = 40
    widths[COL_FIX] = 42
    for c, w in widths.items():
        ws.column_dimensions[get_column_letter(c)].width = w

    hr = 4
    for c, h in ((COL_NAME, "申报材料 / 修改建议"), (COL_SCORE, "总分"),
                 (COL_BAND, "奖项区间"), (COL_GAIN, "改完提升"),
                 (COL_AFTER, "改完总分"),
                 (COL_ORIG, "原文是什么"), (COL_HOW, "怎么改"),
                 (COL_FIX, "修改之后的句子")):
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
    ws.row_dimensions[hr].height = 34
    # 表头统一垂直居中、允许换行，避免长表头被压成三行
    for c in range(1, LAST + 1):
        cell = ws.cell(hr, c)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    groups = []
    for r in rows:
        a = r["award"]
        g = next((x for x in groups if x["award"] == a), None)
        if g is None:
            g = {"award": a, "dims": [d["name"] for d in r["dimensions"]], "rows": []}
            groups.append(g)
        g["rows"].append(r)

    rw = hr + 1
    all_scores, all_gains = [], []
    for g in groups:
        ws.merge_cells(start_row=rw, start_column=COL_NAME, end_row=rw, end_column=LAST)
        gc = ws.cell(rw, COL_NAME, "%s　·　%d 份" % (g["award"], len(g["rows"])))
        gc.font = Font(name="Microsoft YaHei", size=10.5, bold=True, color="FFFFFF")
        gc.alignment = Alignment(horizontal="left", vertical="center", indent=1)
        for c in range(COL_NAME, LAST + 1):
            ws.cell(rw, c).fill = PatternFill("solid", fgColor=HEAD_BG)
            ws.cell(rw, c).border = BORDER
        ws.row_dimensions[rw].height = 20
        rw += 1

        for r in g["rows"]:
            mi = rows.index(r) + 1
            items = r.get("_items") or []
            gain = total_gain(items, r["ai_score"]) if items else 0.0
            after = round(min(100.0, r["ai_score"] + gain), 1)

            # ---- 主行 ----
            disp = strip_award(r["name"])
            c1 = ws.cell(rw, COL_NAME, short(disp, 46))
            c1.font = Font(name="Microsoft YaHei", size=10, bold=True, color=INK)
            c1.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True,
                                     indent=1)
            c2 = ws.cell(rw, COL_SCORE, r["ai_score"])
            c2.font = Font(name="Microsoft YaHei", size=11, bold=True, color=INK)
            c2.alignment = A_MID
            band_txt = "待裁定" if "暂列" in r["band"] else r["band"]
            c3 = ws.cell(rw, COL_BAND, band_txt if len(band_txt) <= 8 else band_txt[:7] + "…")
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

            cg = ws.cell(rw, COL_GAIN, ("+%.1f" % gain) if gain else "—")
            cg.font = Font(name="Microsoft YaHei", size=10, bold=bool(gain),
                           color="15803D" if gain >= 10 else (INK if gain else MUTED))
            cg.alignment = A_MID
            ca = ws.cell(rw, COL_AFTER, after if gain else "—")
            ca.font = Font(name="Microsoft YaHei", size=10, color=INK if gain else MUTED)
            ca.alignment = A_MID
            # 主行右侧写一句总览
            note = ("建议 %d 条，见下方" % len(items)) if items else "各维度达标"
            nc = ws.cell(rw, COL_ORIG, note)
            nc.font = Font(name="Microsoft YaHei", size=9, color=MUTED, italic=True)
            nc.alignment = Alignment(horizontal="left", vertical="center", indent=1)
            for c in range(COL_HOW, LAST + 1):
                ws.cell(rw, c, "")
            # 主行：浅底 + 加粗下边框，与下方的建议子行明确分开
            med = Side(style="medium", color="9CC5B8")
            for c in range(COL_NAME, LAST + 1):
                cell = ws.cell(rw, c)
                cell.border = Border(left=thin, right=thin, top=med, bottom=med)
                cell.fill = PatternFill("solid", fgColor="F0FDF4")
            # A 列宽 38（约 19 个中文字/行），材料名要按实际换行数给高度，
            # 否则名字被裁掉、上面的字看不见。
            name_lines = max(1, -(-len(disp) // 18))
            ws.row_dimensions[rw].height = max(26, name_lines * 15 + 10)
            rw += 1

            # ---- 建议子行：原文 / 怎么改 / 改后 ----
            for it in items:
                no_cell = ws.cell(rw, COL_NAME, "　%s　%s　%s" % (
                    "%d-%s" % (mi, it["no"]), it["dim"], it["title"]))
                no_cell.font = Font(name="Microsoft YaHei", size=9.5, bold=True, color=ACCENT)
                no_cell.alignment = Alignment(horizontal="left", vertical="top",
                                              wrap_text=True, indent=1)
                ws.merge_cells(start_row=rw, start_column=COL_NAME, end_row=rw,
                               end_column=COL_BAND)
                ws.merge_cells(start_row=rw, start_column=COL_DIM0, end_row=rw,
                               end_column=COL_AFTER)
                gain_txt = (("＋%.1f" % it["gain"]) if it["gain"]
                            else ("合并计入" if "合并" in (it.get("gain_note") or "") else "—"))
                gl = ws.cell(rw, COL_DIM0, gain_txt)
                gl.font = Font(name="Microsoft YaHei", size=8.5, bold=bool(it["gain"]),
                               color="15803D" if it["gain"] else MUTED)
                gl.alignment = Alignment(horizontal="center", vertical="top")

                for c, key in ((COL_ORIG, "original"), (COL_HOW, "how"), (COL_FIX, "revised")):
                    cell = ws.cell(rw, c, it.get(key, ""))
                    cell.font = (Font(name="Microsoft YaHei", size=9,
                                      color="0B4F3F" if c == COL_FIX else INK))
                    cell.alignment = A_LEFT
                    cell.border = BORDER

                for c in range(COL_NAME, LAST + 1):
                    ws.cell(rw, c).border = BORDER
                # 按 9pt 字号估算：列宽 38/40/42 ≈ 每行 30/32/34 个中文，
                # 每行 12.5pt 足够（12.5 × 1.15 ≈ 14pt 实际行距）。
                lbl = "　%d-%s　%s　%s" % (mi, it["no"], it["dim"], it["title"])
                lines = max(
                    max(1, -(-len(lbl) // 30)),                          # A:C 合并
                    max(1, -(-len(str(it.get("original", ""))) // 30)),
                    max(1, -(-len(str(it.get("how", ""))) // 32)),
                    max(1, -(-len(str(it.get("revised", ""))) // 34)),
                )
                ws.row_dimensions[rw].height = min(120, max(24, lines * 12.5 + 6))
                rw += 1

            # ---- 争议点 ----
            ds = r.get("disputes") or []
            if ds:
                ws.merge_cells(start_row=rw, start_column=COL_NAME, end_row=rw,
                               end_column=LAST)
                bc = ws.cell(rw, COL_NAME, "　　答辩可能被追问（%d 条）" % len(ds))
                bc.font = Font(name="Microsoft YaHei", size=9, bold=True, color=MUTED)
                bc.alignment = Alignment(horizontal="left", vertical="center", indent=1)
                for c in range(COL_NAME, LAST + 1):
                    ws.cell(rw, c).fill = PatternFill("solid", fgColor=BAND_BG)
                    ws.cell(rw, c).border = BORDER
                ws.row_dimensions[rw].height = 16
                rw += 1
                for i, x in enumerate(ds, 1):
                    pt = (x.get("point") or "").strip()
                    src_ = (x.get("source") or "").strip()
                    ws.merge_cells(start_row=rw, start_column=COL_NAME, end_row=rw,
                                   end_column=COL_DIM0 - 1)
                    qc = ws.cell(rw, COL_NAME, "　　Q%02d" % i)
                    qc.font = Font(name="Microsoft YaHei", size=9, bold=True, color=MUTED)
                    qc.alignment = Alignment(horizontal="left", vertical="top", indent=1)
                    ws.merge_cells(start_row=rw, start_column=COL_DIM0, end_row=rw,
                                   end_column=COL_FIX)
                    qc2 = ws.cell(rw, COL_DIM0, pt + ("　（来源：%s）" % src_ if src_ else ""))
                    qc2.font = Font(name="Microsoft YaHei", size=9, color=INK)
                    qc2.alignment = A_LEFT
                    for c in range(COL_NAME, LAST + 1):
                        ws.cell(rw, c).border = BORDER
                    ws.row_dimensions[rw].height = min(60, max(20, int(len(str(x.get("point") or "")) / 40) * 14 + 16))
                    rw += 1

            all_scores.append(r["ai_score"])
            all_gains.append(gain)
            # 不留空行：靠主行浅绿底色分隔，空行只会让表格显得断裂

        sc = [x["ai_score"] for x in g["rows"]]
        gg = [total_gain(x.get("_items") or [], x["ai_score"]) for x in g["rows"]]
        ws.merge_cells(start_row=rw, start_column=COL_NAME, end_row=rw, end_column=LAST)
        label = ws.cell(rw, COL_NAME,
                        "%s　小计：%d 份　均分 %.1f　平均可提升 %.1f　改完均分 %.1f　" % (
                            g["award"], len(sc), sum(sc) / len(sc),
                            sum(gg) / len(gg) if gg else 0,
                            sum(min(100.0, x["ai_score"] + total_gain(x.get("_items") or [],
                                                                       x["ai_score"]))
                                for x in g["rows"]) / len(sc)))
        label.font = Font(name="Microsoft YaHei", size=9.5, bold=True, color=INK)
        label.alignment = Alignment(horizontal="right", vertical="center", indent=1)
        for c in range(COL_NAME, LAST + 1):
            ws.cell(rw, c).fill = PatternFill("solid", fgColor=BAND_BG)
            ws.cell(rw, c).border = BORDER
        ws.row_dimensions[rw].height = 20
        rw += 1   # 不额外留空行：分组靠深绿标题行 + 小计灰底分隔，空行会让表格显得断裂

    if all_scores:
        ws.merge_cells(start_row=rw, start_column=COL_NAME, end_row=rw, end_column=LAST)
        tc = ws.cell(rw, COL_NAME,
                     "总计　%d 份材料　平均 %.1f 分　最高 %.1f　最低 %.1f　"
                     "平均可提升 %.1f　改完平均 %.1f" % (
                         len(all_scores), sum(all_scores) / len(all_scores),
                         max(all_scores), min(all_scores),
                         sum(all_gains) / len(all_gains),
                         sum(min(100.0, x["ai_score"] + total_gain(x.get("_items") or [],
                                                                   x["ai_score"]))
                             for x in rows) / len(rows)))
        tc.font = Font(name="Microsoft YaHei", size=10, bold=True, color="FFFFFF")
        tc.alignment = Alignment(horizontal="right", vertical="center", indent=1)
        for c in range(COL_NAME, LAST + 1):
            ws.cell(rw, c).fill = PatternFill("solid", fgColor=HEAD_BG)
            ws.cell(rw, c).border = BORDER
        ws.row_dimensions[rw].height = 24

    ws.freeze_panes = ws.cell(hr + 1, COL_NAME)
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
    build_sheet(wb, rows)
    wb.save(a.out)
    total_items = sum(len(r.get("_items") or []) for r in rows)
    print(json.dumps({"out": a.out, "sheets": wb.sheetnames,
                      "materials": len(rows), "advice_items": total_items},
                     ensure_ascii=False))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())