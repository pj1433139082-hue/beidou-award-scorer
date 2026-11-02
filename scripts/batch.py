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


def default_rubric() -> Path:
    """默认评分卡：优先 rubric.json（真实赛事卡），没有就用 rubric.example.json（脱敏示例）。

    两份副本共用同一份脚本，靠这个顺序让各自都能直接跑。
    """
    assets = Path(__file__).resolve().parents[1] / "assets"
    for name in ("rubric.json", "rubric.example.json"):
        p = assets / name
        if p.is_file():
            return p
    return assets / "rubric.example.json"


DEFAULT_RUBRIC = default_rubric()
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
    ap.add_argument("--from-json", action="store_true",
                    help="src 里已是抽好的 *.doc.json（不再抽原文），直接机检与打分")
    ap.add_argument("--from-scores", action="store_true",
                    help="src 里已有 *.score.json，只重出排行榜与 Excel（不再机检打分）")
    ap.add_argument("--machine-only", action="store_true", help="只出机检，不打分")
    ap.add_argument("--rubric", default=str(DEFAULT_RUBRIC), help="评分卡路径，奖品名与别名从这里读")
    a = ap.parse_args()
    src, out = Path(a.src), Path(a.out)
    awards, aliases = load_awards(a.rubric)
    out.mkdir(parents=True, exist_ok=True)
    if a.from_scores:
        # 只重出排行榜与 Excel
        scores = sorted(src.glob("*.score.json"))
        if not scores:
            print("没找到 *.score.json。", file=sys.stderr); return 2
        rows = []
        for s in scores:
            try:
                d = json.loads(s.read_text(encoding="utf-8"))
            except Exception:
                continue
            if "dimensions" not in d: continue
            rows.append({"name": d.get("name", s.stem), "award": d.get("award", ""),
                         "ai_score": d.get("ai_score", 0), "band": d.get("band", ""),
                         "confidence": d.get("confidence", ""),
                         "checks": {"pass": 0, "warn": 0, "fail": 0}})
        rows.sort(key=lambda r: r["ai_score"], reverse=True)
        lines = ["# 北斗奖申报材料批量评分表", "",
                 "| 排名 | 材料 | 奖项 | AI 评审分 | 奖项区间 | 置信度 |",
                 "|---|---|---|---|---|---|"]
        for i, r in enumerate(rows, 1):
            lines.append("| %d | %s | %s | **%s** | %s | %s |" % (
                i, r["name"][:40], r["award"], r["ai_score"], r["band"], r["confidence"]))
        (out / "leaderboard.md").write_text("\n".join(lines) + "\n", encoding="utf-8")
        xlsx = out / "北斗奖评分汇总.xlsx"
        r2 = subprocess.run([sys.executable, str(HERE / "export_xlsx.py"), str(src), "--out", str(xlsx)],
                            capture_output=True, text=True)
        print("写出 %s" % (out / "leaderboard.md"))
        if r2.returncode == 0:
            print("写出 %s" % xlsx)
        else:
            print("Excel 导出未完成：" + r2.stderr.strip()[-160:], file=sys.stderr)
        return 0

    files = []
    if a.from_json:
        files = sorted(src.glob("*.doc.json"))
    else:
        for pat in a.glob.split(","):
            files += sorted(src.glob(pat.strip()))
    if not files:
        print("没有匹配到材料文件。\n"
              "  · 放 docx/pdf 后用默认模式；\n"
              "  · 若 src 里已是抽好的 *.doc.json，加 --from-json。", file=sys.stderr)
        return 2
    rows = []
    for f in files:
        award = award_of(Path(f.name.replace(".doc.json", "")), a.award, awards, aliases)
        if not award:
            print("[skip] %s：文件名无法判断奖项，请用 --award 指定" % f.name, file=sys.stderr); continue
        # --from-json 时 f 是 xxx.doc.json，stem 要去掉 .doc 后缀才是材料基名
        stem = f.name.replace(".doc.json", "") if a.from_json else f.stem
        doc = f if a.from_json else out / (stem + ".doc.json")
        chk = out / (stem + ".checks.json")
        if not a.from_json and not doc.is_file():   # 没抽过才抽
            subprocess.run([sys.executable, str(HERE / "extract_report.py"), str(f), "--out", str(out)],
                           check=True, stdout=subprocess.DEVNULL)
        if not chk.is_file():          # 没检过才检
            subprocess.run([sys.executable, str(HERE / "checks.py"), str(doc), "--award", award,
                            "--rubric", a.rubric, "--out", str(chk)], check=True, stdout=subprocess.DEVNULL)
        c = json.loads(chk.read_text(encoding="utf-8"))
        row = {"name": stem, "award": award, "checks": c["summary"]}
        # judgments 可能就在源目录（from-json 场景），也可能在输出目录
        jp = (f.parent / (stem + ".judgments.json")) if a.from_json else (out / (stem + ".judgments.json"))
        if jp.exists() and not a.machine_only:
            subprocess.run([sys.executable, str(HERE / "score.py"), str(doc), "--checks", str(chk),
                            "--judgments", str(jp), "--award", award, "--rubric", a.rubric,
                            "--out", str(out)], check=True,
                           stdout=subprocess.DEVNULL)
            s = json.loads((out / (stem + ".score.json")).read_text(encoding="utf-8"))
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

    # 有评分才导 Excel（没评分时导出没有意义）
    if scored:
        xlsx = out / "北斗奖评分汇总.xlsx"
        r = subprocess.run([sys.executable, str(HERE / "export_xlsx.py"), str(out), "--out", str(xlsx)],
                           capture_output=True, text=True)
        if r.returncode == 0:
            print("写出 %s" % xlsx)
        else:
            print("Excel 导出未完成（可稍后手动跑 scripts/export_xlsx.py）：" + r.stderr.strip()[-160:], file=sys.stderr)
    return 0

if __name__ == "__main__":
    raise SystemExit(main())