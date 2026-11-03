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

sys.path.insert(0, str(Path(__file__).resolve().parent))
from material_role import is_reference, declared_track, track_matches  # noqa: E402


class RubricError(Exception):
    """评分卡不可用。"""


def load_awards(rubric_path):
    """奖项名与文件名别名都从 rubric 读，不在脚本里硬编码。"""
    rp = Path(rubric_path)
    if not rp.is_file():
        raise RubricError(
            "找不到评分卡：%s\n"
            "  公开副本用 assets/rubric.example.json；"
            "你自己的赛事用 assets/rubric.json 或 --rubric 指定。" % rp)
    try:
        data = json.loads(rp.read_text(encoding="utf-8"))
    except json.JSONDecodeError as e:
        raise RubricError("评分卡不是合法 JSON（%s）：\n  %s" % (e.msg, rp))
    if not (data.get("awards") or {}):
        raise RubricError("评分卡里没有 awards：%s" % rp)
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
    ap.add_argument("--with-docx", action="store_true",
                    help="同时导出 Word（每份一个 + 一本汇总）")
    ap.add_argument("--with-pdf", action="store_true",
                    help="再把 Word 转成 PDF（需要 LibreOffice，找不到则跳过）")
    ap.add_argument("--strict", action="store_true",
                    help="严格模式：任何预期产物缺失、或有材料缺 judgments，都以退出码 2 结束")
    ap.add_argument("--all", action="store_true",
                    help="等同 --with-docx --with-pdf，一条命令出全套交付物")
    ap.add_argument("--from-scores", action="store_true",
                    help="src 里已有 *.score.json，只重出排行榜与 Excel（不再机检打分）")
    ap.add_argument("--machine-only", action="store_true", help="只出机检，不打分")
    ap.add_argument("--rubric", default=str(DEFAULT_RUBRIC), help="评分卡路径，奖品名与别名从这里读")
    a = ap.parse_args()
    if getattr(a, "all", False):
        a.with_docx = True
        a.with_pdf = True
    src, out = Path(a.src), Path(a.out)
    try:
        awards, aliases = load_awards(a.rubric)
    except RubricError as e:
        print("错误：" + str(e), file=sys.stderr)
        return 2
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

        # 参考材料（过往版本等）不是参赛项目：不算分、不排名
        row["role"] = "reference" if is_reference(stem) else "submission"

        # 正文声明的赛道与文件名奖项是否一致——改赛道后忘了改名会用错评分维度
        try:
            _doc = json.loads(doc.read_text(encoding="utf-8"))
            _track = declared_track(_doc.get("plain", ""))
        except Exception:
            _track = None
        row["track"] = _track
        if track_matches(award, _track) is False:
            row["track_mismatch"] = True
            print("[注意] %s：文件名判为「%s」，但正文声明「%s」"
                  % (stem[:36], award, _track), file=sys.stderr)
        # judgments 可能放输出目录，也可能跟材料放在一起——两处都找，避免"放了却没用上"
        jp = out / (stem + ".judgments.json")
        if not jp.is_file() and f.parent != out:
            alt = f.parent / (stem + ".judgments.json")
            if alt.is_file():
                jp = alt
        if row["role"] == "reference":
            row["note"] = "参考材料（过往版本），不参与排名"
        elif jp.exists() and not a.machine_only:
            # 单份 judgments 坏了不能拖垮整批：报出来、继续下一份
            r = subprocess.run([sys.executable, str(HERE / "score.py"), str(doc), "--checks", str(chk),
                                "--judgments", str(jp), "--award", award, "--rubric", a.rubric,
                                "--out", str(out)],
                               capture_output=True, text=True)
            if r.returncode != 0:
                err = (r.stderr or "").strip().splitlines()
                row["score_error"] = err[0][:120] if err else "打分失败"
                print("[skip] %s：%s" % (stem[:32], row["score_error"]), file=sys.stderr)
                rows.append(row)
                continue
            s = json.loads((out / (stem + ".score.json")).read_text(encoding="utf-8"))
            row.update({"ai_score": s["ai_score"], "band": s["band"], "confidence": s["confidence"],
                        "veto": s["veto"]})
        else:
            row["note"] = "未提供 judgments，仅机检"
        rows.append(row)
    # 排行榜
    submissions = [r for r in rows if r.get("role") != "reference"]
    references = [r for r in rows if r.get("role") == "reference"]
    scored = [r for r in submissions if "ai_score" in r]
    lines = ["# 北斗奖申报材料批量评分表", "",
             "| 排名 | 材料 | 奖项 | AI 评审分 | 奖项区间 | 置信度 | 机检 pass/warn/fail |",
             "|---|---|---|---|---|---|---|"]
    scored.sort(key=lambda r: r["ai_score"], reverse=True)
    for i, r in enumerate(scored, 1):
        s = r["checks"]
        lines.append("| %d | %s | %s | **%s** | %s | %s | %d/%d/%d |" % (
            i, r["name"][:40], r["award"], r["ai_score"], r["band"], r["confidence"],
            s["pass"], s["warn"], s["fail"]))
    others = [r for r in submissions if "ai_score" not in r]
    if others:
        lines += ["", "## 未评分", "", "| 材料 | 奖项 | 原因 |", "|---|---|---|"]
        for r in others:
            why = r.get("score_error") or r.get("note") or "未提供 judgments"
            lines.append("| %s | %s | %s |" % (r["name"][:40], r["award"], why))
    if references:
        lines += ["", "## 参考材料（不计入排名）", "",
                  "| 材料 | 名义奖项 | 正文声明赛道 |", "|---|---|---|"]
        for r in references:
            lines.append("| %s | %s | %s |" % (
                r["name"][:40], r["award"], r.get("track") or "—"))
    (out / "leaderboard.md").write_text("\n".join(lines) + "\n", encoding="utf-8")
    print("写出 %s（%d 份参赛项目，%d 份有评分；另有 %d 份参考材料）"
          % (out / "leaderboard.md", len(submissions), len(scored), len(references)))

    # 有评分才导 Excel（没评分时导出没有意义）
    if scored:
        xlsx = out / "北斗奖评分汇总.xlsx"
        r = subprocess.run([sys.executable, str(HERE / "export_xlsx.py"), str(out), "--out", str(xlsx)],
                           capture_output=True, text=True)
        if r.returncode == 0:
            print("写出 %s" % xlsx)
        else:
            print("Excel 导出未完成（可稍后手动跑 scripts/export_xlsx.py）：" + r.stderr.strip()[-160:], file=sys.stderr)

        # Word：每份一份 + 一本汇总册
        if getattr(a, "with_docx", False):
            r = subprocess.run([sys.executable, str(HERE / "export_docx.py"), str(out),
                                "--out", str(out), "--merge"],
                               capture_output=True, text=True)
            if r.returncode == 0:
                n_docx = len(list(out.glob("*.docx")))
                print("写出 %d 个 Word（含汇总册）" % n_docx)
            else:
                print("Word 导出未完成（可手动跑 scripts/export_docx.py）："
                      + r.stderr.strip()[-160:], file=sys.stderr)

        # PDF：可选，找不到 LibreOffice 就跳过
        if getattr(a, "with_pdf", False):
            r = subprocess.run([sys.executable, str(HERE / "export_pdf.py"), str(out),
                                "--out", str(out)],
                               capture_output=True, text=True)
            if r.returncode == 0:
                if '"skipped": true' in r.stdout or '"skipped":true' in r.stdout:
                    print("PDF 跳过（未找到 LibreOffice，Word 版可正常使用）")
                else:
                    print("写出 %d 个 PDF" % len(list(out.glob("*.pdf"))))
            else:
                print("PDF 转换未全部完成：" + r.stdout.strip()[-160:], file=sys.stderr)
        elif getattr(a, "strict", False):
            # 严格模式下，要求 PDF 却没转出来，就是失败
            print("严格模式：要求 PDF 但未执行转换（缺 LibreOffice）", file=sys.stderr)
            return 2

    # ---- 后置核验：不看日志，只看文件在不在 ----
    manifest = out / "manifest.json"
    files = []
    for p in sorted(out.iterdir()):
        if p.is_file() and p.name != "manifest.json":
            files.append({"name": p.name, "size": p.stat().st_size})
    counts = {}
    for it in files:
        suf = Path(it["name"]).suffix
        counts[suf] = counts.get(suf, 0) + 1
    manifest.write_text(json.dumps({
        "materials": len(scored),
        "submissions": len(submissions),
        "references": len(references),
        "rows": len(rows),
        "counts": counts, "files": files,
    }, ensure_ascii=False, indent=1), encoding="utf-8")

    # 缺 judgments 的材料：严格模式下算失败
    mismatch = [r for r in rows if r.get("track_mismatch")]
    if mismatch:
        msg = "有 %d 份材料的文件名奖项与正文赛道不一致" % len(mismatch)
        if getattr(a, "strict", False):
            print("严格模式失败：" + msg, file=sys.stderr)
            for r in mismatch[:3]:
                print("  - %s：文件名为「%s」，正文为「%s」"
                      % (r["name"][:30], r["award"], r.get("track")), file=sys.stderr)
            return 2
        print("提示：" + msg)
    scored_fail = [r for r in rows if r.get("score_error")]
    if scored_fail and getattr(a, "strict", False):
        print("严格模式失败：%d 份材料打分失败" % len(scored_fail), file=sys.stderr)
        return 2
    no_judg = [r["name"] for r in rows
               if r.get("role") != "reference"
               and "ai_score" not in r and not r.get("score_error")]
    if no_judg:
        msg = "有 %d 份材料缺 judgments，只出了机检：%s" % (
            len(no_judg), "、".join(n[:20] for n in no_judg[:3]))
        if getattr(a, "strict", False):
            print("严格模式失败：" + msg, file=sys.stderr)
            return 2
        print("提示：" + msg)

    # 调独立核验脚本，逐项确认产物齐全
    vargs = [sys.executable, str(HERE / "verify_outputs.py"), str(out)]
    if getattr(a, "with_docx", False):
        vargs.append("--expect-docx")
    if getattr(a, "with_pdf", False):
        vargs.append("--expect-pdf")
    if getattr(a, "strict", False):
        vargs.append("--strict")
    v = subprocess.run(vargs, capture_output=True, text=True)
    try:
        vs = json.loads(v.stdout)
    except Exception:
        vs = {"ok": False, "missing": ["核验脚本输出无法解析"], "missing_total": 1}
    if vs.get("ok"):
        print("核验通过：%d 份材料，产物齐全" % vs.get("materials", 0))
    else:
        strict = getattr(a, "strict", False)
        head = "严格模式失败：缺 %d 项" if strict else "提示：非严格模式，有 %d 项未生成"
        print(head % vs.get("missing_total", 0), file=sys.stderr)
        for m in vs.get("missing", [])[:5]:
            print("  - " + m, file=sys.stderr)
        if vs.get("missing_total", 0) > 5:
            print("  …共 %d 项" % vs["missing_total"], file=sys.stderr)
        if strict:
            return 2
    print("清单：%s" % manifest)
    # 一份都没评分：这不是成功。要么奖项识别不了，要么 judgments 没对上。
    if not scored and submissions:
        print("错误：没有任何材料完成评分（共识别 %d 份参赛项目）。\n"
              "  常见原因：文件名没带奖项前缀、judgments 文件名与材料不匹配、或 --award 与评分卡不符。"
              % len(rows), file=sys.stderr)
        for r in rows[:5]:
            if r.get("score_error"):
                print("  - %s：%s" % (r["name"][:30], r["score_error"]), file=sys.stderr)
        return 2
    return 0

if __name__ == "__main__":
    raise SystemExit(main())