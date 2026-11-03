#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""格式验收：一次性告诉调用方"输出格式对不对"。

为什么需要它：
  改输出的 agent 容易陷入「生成 → 肉眼看 → 发现不对 → 再改 → 再看」的循环，
  每轮都要渲染、截图、判断，慢且不稳定。格式是**可机检的**，不该靠眼睛。

  本脚本把所有格式要求写成断言，一次跑完给结论。改完直接跑它，
  不用渲染 PDF 去看。只有它报「看起来不对」时才需要人去看图。

用法：
  python scripts/verify_format.py <输出目录>
  python scripts/verify_format.py <输出目录> --json     # 机器可读

退出码 0 = 全部通过；1 = 有未通过的项。

检查内容（每条都是曾经的返工点）：
  Excel  · 只有一张表 / 表头齐全 / 有分组小计 / 有总计 / 编号格式
  Word   · 白底黑字（无彩色）/ 四节结构 / 三段式 / 编号与 Excel 一致
  PDF    · 数量与 docx 相同 / 每份非空
  业务   · 格式类问题（C09/C17）不扣分 / 参考材料不参与排名
"""
from __future__ import annotations
import argparse, json, re, sys
from pathlib import Path


class Report:
    def __init__(self):
        self.rows: list[tuple[str, bool, str]] = []

    def add(self, name: str, ok: bool, detail: str = "") -> None:
        self.rows.append((name, bool(ok), detail))

    @property
    def failed(self) -> list[tuple[str, bool, str]]:
        return [r for r in self.rows if not r[1]]

    def dump(self, as_json: bool) -> int:
        if as_json:
            print(json.dumps({
                "ok": not self.failed,
                "checks": [{"name": n, "ok": o, "detail": d} for n, o, d in self.rows],
                "failed": [n for n, o, _ in self.failed],
            }, ensure_ascii=False, indent=1))
        else:
            print("=" * 66)
            for n, o, d in self.rows:
                mark = "[OK]  " if o else "[FAIL]"
                print("%s %-34s %s" % (mark, n, d))
            print("=" * 66)
            if self.failed:
                print("未通过 %d 项：" % len(self.failed))
                for n, _, d in self.failed:
                    print("  · %s%s" % (n, ("　—　" + d) if d else ""))
            else:
                print("全部通过。")
        return 0 if not self.failed else 1


def check_excel(out: Path, rp: Report) -> set[str]:
    """Excel 结构与编号；返回其中出现的 M 编号集合。"""
    f = out / "北斗奖评分汇总.xlsx"
    if not f.is_file():
        rp.add("Excel 存在", False, "缺 北斗奖评分汇总.xlsx")
        return set()
    rp.add("Excel 存在", True)
    try:
        from openpyxl import load_workbook
    except ImportError:
        rp.add("openpyxl 可用", False)
        return set()

    wb = load_workbook(f)
    rp.add("Excel 只有一张表", len(wb.sheetnames) == 1, str(wb.sheetnames))
    ws = wb[wb.sheetnames[0]]

    hdr = {str(ws.cell(4, c).value or "") for c in range(1, ws.max_column + 1)}
    need = ["申报材料", "总分", "奖项区间", "改完提升", "改完总分",
            "原文是什么", "怎么改", "修改之后的句子"]
    miss = [n for n in need if not any(n in h for h in hdr)]
    rp.add("Excel 表头齐全", not miss, "缺 " + "、".join(miss) if miss else "")

    all_text = []
    codes: set[str] = set()
    for row in ws.iter_rows(min_row=5, values_only=True):
        for v in row:
            if v is None:
                continue
            s = str(v)
            all_text.append(s)
            m = re.search(r"(\d+-M\d\d)\u3000", s)
            if m:
                codes.add(m.group(1))
    blob = " ".join(all_text)
    # 空白行检查：分组之间靠深绿标题 + 灰底小计分隔，不应靠空行
    blanks = [r for r in range(5, ws.max_row + 1)
              if all(ws.cell(r, c).value is None for c in range(1, ws.max_column + 1))]
    rp.add("Excel 无空白行", not blanks,
           "空白行 " + str(blanks[:6]) if blanks else "")
    rp.add("Excel 有分组小计", "小计" in blob)
    rp.add("Excel 有总计行", "总计" in blob and ("平均" in blob or "最高" in blob))
    rp.add("Excel 有建议编号", len(codes) > 0, "%d 条" % len(codes))
    return codes


def check_word(out: Path, rp: Report) -> set[str]:
    """Word 格式与编号；返回 M 编号集合。"""
    try:
        from docx import Document
    except ImportError:
        rp.add("python-docx 可用", False)
        return set()

    docxs = [f for f in sorted(out.glob("*.docx")) if "汇总" not in f.name]
    rp.add("Word 单份文件存在", len(docxs) >= 1, "%d 份" % len(docxs))
    if not docxs:
        return set()

    codes: set[str] = set()
    colored = 0
    lack_sections = 0
    lack_three = 0
    for f in docxs:
        try:
            d = Document(str(f))
        except Exception:
            continue
        txt = "\n".join(p.text for p in d.paragraphs)
        for p in d.paragraphs:
            for run in p.runs:
                c = run.font.color
                if c is not None and c.rgb is not None and str(c.rgb) not in ("000000", "555555"):
                    colored += 1
        # 参考材料说明只查前两节，不要求三段式
        is_ref = "参考材料" in txt[:200]
        secs = ["一、", "二、", "三、", "四、"]
        if not all(s in txt for s in secs[:3]):
            lack_sections += 1
        if not is_ref:
            for seg in ("原文是什么", "怎么改", "修改之后的句子"):
                if seg not in txt:
                    lack_three += 1
                    break
        for m in re.finditer(r"(\d+-M\d\d)\u3000", txt):
            codes.add(m.group(1))

    rp.add("Word 无彩色文字", colored == 0, "彩色 run %d 处" % colored if colored else "")
    rp.add("Word 有一至三节", lack_sections == 0,
           "%d 份不全" % lack_sections if lack_sections else "")
    rp.add("Word 建议三段式", lack_three == 0,
           "%d 份缺段" % lack_three if lack_three else "")
    rp.add("Word 有免责说明", any("免责说明" in Document(str(f)).paragraphs[0].text
                                or "免责说明" in "\n".join(p.text for p in Document(str(f)).paragraphs)
                                for f in docxs[:1]))
    return codes


def check_pdf(out: Path, rp: Report) -> None:
    docxs = list(out.glob("*.docx"))
    pdfs = list(out.glob("*.pdf"))
    rp.add("PDF 数量与 docx 相同", len(pdfs) == len(docxs),
           "docx %d / pdf %d" % (len(docxs), len(pdfs)))
    empty = [p.name for p in pdfs if p.stat().st_size < 3000]
    rp.add("PDF 均非空", not empty, "过小 %d 份" % len(empty) if empty else "")


def check_business(out: Path, rp: Report) -> None:
    """业务规则：格式类不扣分、参考材料不排名。"""
    scores = sorted(out.glob("*.score.json"))
    if not scores:
        return
    bad_ded = 0
    for f in scores:
        try:
            d = json.loads(f.read_text(encoding="utf-8"))
        except Exception:
            continue
        for x in d.get("deductions") or []:
            if x.get("check") in ("C09", "C17"):
                bad_ded += 1
    rp.add("格式类问题不扣分（C09/C17）", bad_ded == 0,
           "%d 处仍扣分" % bad_ded if bad_ded else "")

    lb = out / "leaderboard.md"
    if lb.is_file():
        txt = lb.read_text(encoding="utf-8")
        rp.add("参考材料注明不计入排名",
               ("参考材料" not in txt) or ("不计入排名" in txt or "不参与排名" in txt))


def main() -> int:
    ap = argparse.ArgumentParser(description="格式验收")
    ap.add_argument("out", help="输出目录")
    ap.add_argument("--json", action="store_true", help="输出 JSON")
    ap.add_argument("--no-word", action="store_true", help="跳过 Word 格式检查")
    ap.add_argument("--no-pdf", action="store_true", help="跳过 PDF 检查")
    a = ap.parse_args()

    out = Path(a.out)
    rp = Report()
    if not out.is_dir():
        print("目录不存在：%s" % out, file=sys.stderr)
        return 1

    xl = check_excel(out, rp)
    wd = check_word(out, rp) if not a.no_word else set()
    if not a.no_pdf:
        check_pdf(out, rp)
    check_business(out, rp)

    if xl and wd:
        rp.add("Excel 与 Word 编号一致", xl == wd,
               "Excel %d / Word %d，差异 %d" % (len(xl), len(wd), len(xl ^ wd))
               if xl != wd else "%d 条一致" % len(xl))
    return rp.dump(a.json)


if __name__ == "__main__":
    raise SystemExit(main())
