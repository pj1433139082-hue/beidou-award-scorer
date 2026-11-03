#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""核验输出目录是否产出齐全。

这是"流程硬不硬"的判据：不看日志怎么写的，只看文件在不在、数量对不对。

用法：
  python scripts/verify_outputs.py out/
  python scripts/verify_outputs.py out/ --expect-docx --expect-pdf
  python scripts/verify_outputs.py out/ --strict        # 缺任何一项退出码 2

退出码：
  0  核心产物齐全（PDF/Word 未要求时不检查）
  1  有缺失（非严格模式也返回 1，便于 CI 判断）
  2  目录不可用或没有可核验的材料
"""
from __future__ import annotations
import argparse, json, sys
from pathlib import Path

CORE_SUFFIX = [".doc.json", ".checks.json", ".score.json", ".report.md"]
BATCH_FILES = ["leaderboard.md", "北斗奖评分汇总.xlsx"]
MERGED_DOCX = "北斗奖评分意见汇总.docx"
MERGED_PDF = "北斗奖评分意见汇总.pdf"


def main() -> int:
    ap = argparse.ArgumentParser(description="核验输出产物是否齐全")
    ap.add_argument("out", help="batch 的输出目录")
    ap.add_argument("--expect-docx", action="store_true", help="要求每份材料和汇总册都有 docx")
    ap.add_argument("--expect-pdf", action="store_true", help="要求每份材料和汇总册都有 pdf")
    ap.add_argument("--strict", action="store_true", help="缺失时退出码 2")
    a = ap.parse_args()

    out = Path(a.out)
    if not out.is_dir():
        print("错误：目录不存在：%s" % out, file=sys.stderr)
        return 2

    # 以 score.json 为该目录"有材料"的依据
    stems = sorted(p.name[:-len(".score.json")] for p in out.glob("*.score.json"))
    if not stems:
        # 注意：missing / missing_total 必须始终存在——调用方靠它们判断，
        # 早期版本这里只给 reason，batch 兜底成 0，于是把"什么都没产出"报成"缺 0 项"。
        print(json.dumps({
            "ok": False, "materials": 0,
            "reason": "目录里没有 *.score.json，没有可核验的材料",
            "missing": ["目录里没有任何 *.score.json（0 份材料完成评分）"],
            "missing_total": 1,
            "counts": {p.suffix: 1 for p in out.iterdir() if p.is_file()},
        }, ensure_ascii=False, indent=1))
        return 2

    missing: list[str] = []

    # ① 每份材料的四个核心文件
    for s in stems:
        for suf in CORE_SUFFIX:
            if not (out / (s + suf)).is_file():
                missing.append(s + suf)

    # ② 每批的固定文件
    for f in BATCH_FILES:
        if not (out / f).is_file():
            missing.append(f)

    # ③ 可选：Word
    if a.expect_docx:
        for s in stems:
            if not (out / (s + ".docx")).is_file():
                missing.append(s + ".docx")
        if not (out / MERGED_DOCX).is_file():
            missing.append(MERGED_DOCX)

    # ④ 可选：PDF
    if a.expect_pdf:
        for s in stems:
            if not (out / (s + ".pdf")).is_file():
                missing.append(s + ".pdf")
        if not (out / MERGED_PDF).is_file():
            missing.append(MERGED_PDF)

    # ⑤ 若有 manifest，核对记录的文件是否还在、以及"有几份材料没出分"
    mf = out / "manifest.json"
    if mf.is_file():
        try:
            m = json.loads(mf.read_text(encoding="utf-8"))
            for item in m.get("files", []):
                if not (out / item["name"]).is_file():
                    missing.append("manifest 记录但已丢失：" + item["name"])
            # 本轮看过几份材料 vs 实际出了几分——差额就是打分失败的
            rows = m.get("rows")
            if isinstance(rows, int) and rows > len(stems):
                missing.append("有 %d 份材料未产出 score.json（打分失败或格式错误）"
                               % (rows - len(stems)))
        except Exception as e:
            missing.append("manifest.json 无法解析：%s" % e)

    # 按"产物类型"统计，而不是按后缀（.json 会把 doc/checks/score/judgments 混在一起）
    KEY = {".doc.json": "doc", ".checks.json": "checks", ".score.json": "score",
           ".report.md": "report", ".docx": "docx", ".pdf": "pdf",
           ".judgments.json": "judgments"}
    counts: dict[str, int] = {}
    for p in out.iterdir():
        if not p.is_file():
            continue
        kind = next((v for k, v in KEY.items() if p.name.endswith(k)), None)
        if kind is None:
            kind = {"xlsx": "xlsx", "md": "md", "json": "json"}.get(p.suffix[1:], "other")
        counts[kind] = counts.get(kind, 0) + 1

    ok = not missing
    print(json.dumps({
        "ok": ok, "materials": len(stems), "missing": missing[:20],
        "missing_total": len(missing), "counts": counts,
    }, ensure_ascii=False, indent=1))
    if ok:
        return 0
    return 2 if a.strict else 1


if __name__ == "__main__":
    raise SystemExit(main())