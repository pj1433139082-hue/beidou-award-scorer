#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""同一批材料多轮独立评分，看判断在哪些地方摇摆。

## 解决什么

「不同机器分数不同」有两种完全不同的原因，必须分开：

    ① 判断本身不稳    同一份材料，换个时间或换个 agent 给出不同分数
    ② 材料本身模糊    证据就在中间地带，任何人都可能给两个相邻等级

① 是工具的毛病，② 是材料的问题。混在一起就分不清该修哪个。

## 前提（已验证）

score.py 是确定性的：同一份 judgments 跑两次，19 份 score.json 字节相同。
所以多轮之间的差异只可能来自 judgments 不同，也就是判断不同。
这正是一致性检验要测的东西。

## 用法

    python scripts/consistency.py 轮1/ 轮2/ 轮3/ --out 输出/
    python scripts/consistency.py rounds/ --out 输出/   # rounds 下有子目录也行
"""
from __future__ import annotations
import argparse, json, statistics, sys
from pathlib import Path


def load_round(d):
    """读一轮结果，键是材料 name（ADR-0003）。"""
    from material_record import load_index
    return {name: rec.raw for name, rec in load_index(d).items()}


def analyze(rounds):
    names = set()
    for r in rounds:
        names |= set(r)

    per_material = []
    for name in sorted(names):
        totals, missing = [], 0
        dims = {}
        for r in rounds:
            sc = r.get(name)
            if not sc:
                missing += 1
                continue
            totals.append(sc.get("ai_score", 0))
            for d in (sc.get("dimensions") or []):
                dims.setdefault(d.get("name"), []).append(
                    (d.get("score", 0), d.get("verified_grade")))
        if not totals:
            continue

        dim_rows = []
        for dim, vals in dims.items():
            scores = [v[0] for v in vals]
            grades = [v[1] for v in vals]
            spread = round(max(scores) - min(scores), 1)
            grade_varies = len(set(grades)) > 1
            dim_rows.append({
                "dim": dim,
                "scores": scores,
                "grades": grades,
                "mean": round(statistics.fmean(scores), 1),
                "spread": spread,
                "grade_varies": grade_varies,
                "stable": spread < 0.05 and not grade_varies,
            })

        per_material.append({
            "name": name,
            "totals": totals,
            "mean": round(statistics.fmean(totals), 1),
            "spread": round(max(totals) - min(totals), 1),
            "rounds_present": len(totals),
            "rounds_missing": missing,
            "dims": dim_rows,
        })

    dim_overall = {}
    for m in per_material:
        for d in m["dims"]:
            k = d["dim"]
            e = dim_overall.setdefault(k, {"n": 0, "unstable": 0, "max_spread": 0.0})
            e["n"] += 1
            if not d["stable"]:
                e["unstable"] += 1
            e["max_spread"] = max(e["max_spread"], d["spread"])

    return {"materials": per_material, "dim_overall": dim_overall}


def render(res, rounds_desc):
    ms = res["materials"]
    lines = ["# 多轮评分一致性", "", "参与对比的轮次：%d" % len(rounds_desc), ""]
    for i, d in enumerate(rounds_desc, 1):
        lines.append("%d. %s" % (i, d))
    lines.append("")

    if not ms:
        lines.append("没有可对比的材料（各轮的材料名对不上）。")
        return chr(10).join(lines)

    spreads = [m["spread"] for m in ms]
    full = [m for m in ms if m["rounds_missing"] == 0]
    lines.append("## 总体")
    lines.append("")
    lines.append("材料数：%d（其中 %d 份在所有轮次都有结果）" % (len(ms), len(full)))
    lines.append("")
    lines.append("| 指标 | 值 |")
    lines.append("|---|---|")
    lines.append("| 平均分数跨度 | %.1f 分 |" % statistics.fmean(spreads))
    lines.append("| 最大跨度 | %.1f 分 |" % max(spreads))
    lines.append("| 完全一致的材料 | %d / %d |" % (
        sum(1 for m in ms if m["spread"] < 0.05), len(ms)))
    lines.append("")

    lines.append("## 分数最不稳的材料")
    lines.append("")
    lines.append("| 材料 | 均分 | 跨度 | 各轮 |")
    lines.append("|---|---|---|---|")
    for m in sorted(ms, key=lambda x: -x["spread"])[:10]:
        lines.append("| %s | %.1f | %.1f | %s |" % (
            m["name"][:36], m["mean"], m["spread"],
            " / ".join("%.1f" % t for t in m["totals"])))
    lines.append("")

    do = res["dim_overall"]
    if do:
        lines.append("## 各维度的稳定性")
        lines.append("")
        lines.append("| 维度 | 出现 | 摇摆 | 最大跨度 | 判断 |")
        lines.append("|---|---|---|---|---|")
        for dim, e in sorted(do.items(), key=lambda x: -x[1]["unstable"]):
            rate = e["unstable"] / e["n"] if e["n"] else 0
            if rate >= 0.5:
                verdict = "**材料普遍模糊**（证据在等级交界）"
            elif rate >= 0.2:
                verdict = "偶有摇摆"
            else:
                verdict = "稳定"
            lines.append("| %s | %d | %d | %.1f | %s |" % (
                dim, e["n"], e["unstable"], e["max_spread"], verdict))
        lines.append("")

    pairs = []
    for m in ms:
        for d in m["dims"]:
            if not d["stable"]:
                pairs.append((m["name"], d))
    if pairs:
        lines.append("## 优先复核这些判断")
        lines.append("")
        lines.append("| 材料 | 维度 | 各轮分数 | 等级 | 跨度 |")
        lines.append("|---|---|---|---|---|")
        for name, d in sorted(pairs, key=lambda x: -x[1]["spread"])[:15]:
            lines.append("| %s | %s | %s | %s | %.1f |" % (
                name[:30], d["dim"],
                " / ".join("%.1f" % s for s in d["scores"]),
                "→".join(sorted(set(d["grades"]), key=str)), d["spread"]))
        lines.append("")

    lines.append("## 怎么读这份报告")
    lines.append("")
    lines.append("- **跨度大但等级一致**：分数在等级区间内浮动，可接受，但说明该维度判据不够硬。")
    lines.append("- **等级也在变**：材料卡在两个等级交界，需补证据才能定下来，不算判断错。")
    lines.append("- **某些维度普遍摇摆**：说明评分卡对该维度描述不清，应该改评分卡，而不是怪评审。")
    lines.append("")
    return chr(10).join(lines)


def main():
    ap = argparse.ArgumentParser(description="多轮评分一致性")
    ap.add_argument("rounds", nargs="+")
    ap.add_argument("--out", default=".")
    a = ap.parse_args()

    dirs = []
    for p in a.rounds:
        d = Path(p)
        if not d.is_dir():
            print("目录不存在：%s" % d, file=sys.stderr)
            return 2
        if list(d.glob("*.score.json")):
            dirs.append(d)
        else:
            subs = [x for x in sorted(d.iterdir())
                    if x.is_dir() and list(x.glob("*.score.json"))]
            if not subs:
                print("%s 下没有含 score.json 的子目录" % d, file=sys.stderr)
                return 2
            dirs.extend(subs)

    if len(dirs) < 2:
        print("至少需要两轮结果（当前 %d 轮）" % len(dirs), file=sys.stderr)
        return 2

    res = analyze([load_round(d) for d in dirs])
    md = render(res, [str(d) for d in dirs])

    out = Path(a.out)
    out.mkdir(parents=True, exist_ok=True)
    p = out / "一致性检验.md"
    p.write_text(md, encoding="utf-8")

    ms = res["materials"]
    print(json.dumps({
        "out": str(p),
        "rounds": len(dirs),
        "materials": len(ms),
        "mean_spread": round(statistics.fmean([m["spread"] for m in ms]), 2) if ms else 0,
        "max_spread": max((m["spread"] for m in ms), default=0),
        "fully_consistent": sum(1 for m in ms if m["spread"] < 0.05),
    }, ensure_ascii=False, indent=1))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())