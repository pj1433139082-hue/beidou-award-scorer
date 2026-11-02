#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""批量跑：目录里所有 .docx/.pdf -> doc.json -> checks.json -> 排行榜。

用法：
  python scripts/batch.py "<材料目录>" --out ./out
  python scripts/batch.py "...\材料" --out ./out --award 金引擎     # 全部同一奖项
  python scripts/batch.py "...\材料" --out ./out --glob "金火炬-*.docx"
奖项默认从文件名前缀推断（金点子/金刚钻/金引擎/金智囊/金火炬）。
没有 judgments 的报告只出机检分（见 --machine-only），不臆造评审分。
"""
from __future__ import annotations
import argparse, json, subprocess, sys
from pathlib import Path

def load_awards(rubric_path):
    """奖项名与文件名别名都从 rubric 读，不在脚本里硬编码。"""
    data = json.loads(Path(rubric_path).read_text(encoding="utf-8"))
    names = list((data.get("awards") or {}).keys())
    aliases = []
    for name, spec in (data.get("awards") or {}).items():
        for prefix in (spec.get("filename_prefixes") or []):
            aliases.append((prefix, name))
    return names, aliases


DEFAULT_RUBRIC = Path(__file__).resolve().parents[1] / "assets" / "rubric.example.json"
HERE = Path(__file__).resolve().parent

def award_of(path: Path, forced, awards, aliases):
    if forced: return forced
    for prefix, a in aliases:
        if path.name.startswith(prefix): return a
    for a in awards:
        if path.name.startswith(a): return a
    return None

def main():
    ap = argparse.ArgumentParser(description="批量抽取 + 机检 (+ 已有 judgments 时打分)")
    ap.add_argument("src")
    ap.add_argument("--out", default="./out")
    ap.add_argument("--award")
    ap.add_argument("--glob", default="*.docx,*.pdf")
    ap.add_argument("--machine-only", action="store_true", help="只出机检，不打分")
    ap.add_argument("--rubric", default=str(DEFAULT_RUBRIC), help="评分卡路径，奖品名与别名从这里读")
    a = ap.parse_args()
    src, out = Path(a.src), Path(a.out)
    awards, aliases = load_awards(a.rubric)
    out.mkdir(parents=True, exist_ok=True)
    files = []
    for pat in a.glob.split(","):
        files += sorted(src.glob(pat.strip()))
    if not files:
        print("没有匹配到材料文件", file=sys.stderr); return 2
    rows = []
    for f in files:
        award = award_of(f, a.award, awards, aliases)
        if not award:
            print("[skip] %s：文件名无法判断奖项，请用 --award 指定" % f.name, file=sys.stderr); continue
        subprocess.run([sys.executable, str(HERE / "extract_report.py"), str(f), "--out", str(out)], check=True,
                       stdout=subprocess.DEVNULL)
        doc = out / (f.stem + ".doc.json")
        chk = out / (f.stem + ".checks.json")
        subprocess.run([sys.executable, str(HERE / "checks.py"), str(doc), "--award", award,
                        "--rubric", a.rubric, "--out", str(chk)], check=True, stdout=subprocess.DEVNULL)
        c = json.loads(chk.read_text(encoding="utf-8"))
        row = {"name": f.stem, "award": award, "checks": c["summary"]}
        jp = out / (f.stem + ".judgments.json")
        if jp.exists() and not a.machine_only:
            subprocess.run([sys.executable, str(HERE / "score.py"), str(doc), "--checks", str(chk),
                            "--judgments", str(jp), "--award", award, "--rubric", a.rubric,
                            "--out", str(out)], check=True,
                           stdout=subprocess.DEVNULL)
            s = json.loads((out / (f.stem + ".score.json")).read_text(encoding="utf-8"))
            row.update({"ai_score": s["ai_score"], "band": s["band"], "confidence": s["confidence"],
                        "veto": s["veto"]})
        else:
            row["note"] = "未提供 judgments，仅机检"
        rows.append(row)
    # 排行榜
    scored = [r for r in rows if "ai_score" in r]
    lines = ["# 北斗奖申报材料批量评分表", "",
             "| 排名 | 材料 | 奖项 | AI 评审分 | 奖项区间 | 置信度 | 机检 pass/warn/fail |",
             "|---|---|---|---|---|---|---|"]
    scored.sort(key=lambda r: r["ai_score"], reverse=True)
    for i, r in enumerate(scored, 1):
        s = r["checks"]
        lines.append("| %d | %s | %s | **%s** | %s | %s | %d/%d/%d |" % (
            i, r["name"][:40], r["award"], r["ai_score"], r["band"], r["confidence"],
            s["pass"], s["warn"], s["fail"]))
    others = [r for r in rows if "ai_score" not in r]
    if others:
        lines += ["", "## 仅机检（缺 judgments）", "",
                  "| 材料 | 奖项 | pass/warn/fail |", "|---|---|---|"]
        for r in others:
            s = r["checks"]
            lines.append("| %s | %s | %d/%d/%d |" % (r["name"][:40], r["award"], s["pass"], s["warn"], s["fail"]))
    (out / "leaderboard.md").write_text("\n".join(lines) + "\n", encoding="utf-8")
    print("写出 %s（%d 份材料，%d 份有评分）" % (out / "leaderboard.md", len(rows), len(scored)))
    return 0

if __name__ == "__main__":
    raise SystemExit(main())