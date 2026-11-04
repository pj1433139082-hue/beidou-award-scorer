#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""三种运行模式：init / refine / compare。

## 为什么要区分模式

「给材料打分」其实有三种不同意图，混在一起结果就不可信：

    init     第一次评，干净隔离，一次定分，不受历史影响
    refine   评过了想再找改进点。分数可以一样，建议必须重新想
    compare  同一项目改过一版，要看改前改后差在哪

## init

不读任何历史结果；judgments 当场写；输出绑定材料指纹。

## refine

这里有个容易做错的地方：如果「优化建议」还是从机检查表得来的，
跑一百次建议都一样，等于没优化。

所以 refine 的产出是**给 agent 的思考任务**，不是成品建议。
技能负责把材料摆好、把上次结论摊开、把「必须回答哪些问题」写清楚。
真的重新思考由调用它的 agent 完成——技能保持零外部依赖，换任何模型都能用。

## compare

同项目新旧两版逐维度对比，看分数变化，判断真改进还是换个说法。

用法：
    python scripts/modes.py init    <材料目录> --out <输出>
    python scripts/modes.py refine  <材料目录> --baseline <上次输出> --out <输出>
    python scripts/modes.py compare <新版目录> --baseline <旧版输出> --out <输出>
"""
from __future__ import annotations
import argparse, json, re, sys
from pathlib import Path


def load_scores(d):
    """读一批评分结果，键是材料 name（ADR-0003）。"""
    from material_record import load_index
    return {name: rec.raw for name, rec in load_index(d).items()}


def _same_material(full_name: str, xlsx_name: str) -> bool:
    """xlsx 的短名是否对应这份材料。规则见 material_record.same_material。"""
    from material_record import same_material
    return same_material(full_name, xlsx_name)


def load_items(d):
    """从汇总表里读上次的建议条目。布局与解析都在 advice_sheet 里定义。"""
    from advice_sheet import read as read_advice
    return read_advice(d)


def build_refine_pack(mat_dir, baseline, out):
    scores = load_scores(baseline)
    items = load_items(baseline)
    docs = {}
    for f in sorted(Path(mat_dir).glob("*.doc.json")):
        try:
            dd = json.loads(f.read_text(encoding="utf-8"))
            docs[dd.get("name") or f.stem] = dd
        except Exception:
            pass

    lines = [
        "# 反思输入包",
        "",
        "上次的分数和建议都在下面。你的任务是**重新思考**，不是复述。",
        "",
        "## 必须回答的问题",
        "",
        "对每份材料的每个维度：",
        "",
        "1. 上次给的分数，按现在的理解是否仍然成立？变了就说明变多少、为什么。",
        "2. 上次那条建议现在还能提吗？材料没改就依然成立，但要用你自己的话重说，",
        "   并说明你重新核对了原文哪一段。上次说错了要明确指出来。",
        "3. 上次**漏掉**了什么？这是重点。找出上次没提、但确实值得改的点。",
        "4. 有没有哪条是套话（换成任何材料都成立）？有就删掉，别沿用。",
        "",
        "**分数可以完全不变，但第 3、4 项必须给出实质内容。**",
        "一条新发现都没有，说明你没真的重读材料。",
        "",
        "---",
        "",
    ]

    n = 0
    for name, doc in docs.items():
        sc = scores.get(name)
        if not sc:
            continue
        n += 1
        lines.append("## %d. %s" % (n, name))
        lines.append("")
        lines.append("总分 %.1f　区间 %s" % (sc.get("ai_score", 0), sc.get("band", "")))
        lines.append("")
        lines.append("### 上次的维度判断")
        lines.append("")
        lines.append("| 维度 | 得分 | 等级 | 上次判据 |")
        lines.append("|---|---|---|---|")
        for dm in sc.get("dimensions") or []:
            lines.append("| %s | %.1f/%d | %s | %s |" % (
                dm.get("name", ""), dm.get("score", 0), dm.get("max", 0),
                dm.get("verified_grade", ""), (dm.get("rationale") or "")[:60]))
        lines.append("")

        # 材料身份统一用 name（ADR-0003）。
        # xlsx 第一列存的是去掉奖项前缀的短名（给人看），而 doc.json 的 name
        # 带前缀。所以这里按「短名 + 已声明的奖项前缀」精确还原，不做模糊匹配：
        # 模糊匹配会把「名字相近的两份材料」搞混，这正是要消灭的补丁。
        its = items.get(name) or []
        if not its:
            for k, v in items.items():
                if _same_material(name, k):
                    its = v
                    break
        if its:
            lines.append("### 上次给过的建议（%d 条）" % len(its))
            lines.append("")
            for i, it in enumerate(its, 1):
                lines.append("%d. **%s　%s**" % (i, it["dim"], it["title"]))
                lines.append("   - 上次说原文：%s" % ((it["original"] or "（无）")[:120]))
                lines.append("   - 上次怎么改：%s" % ((it["how"] or "（无）")[:140]))
            lines.append("")

        plain = doc.get("plain") or ""
        lines.append("### 材料原文（供重新核对）")
        lines.append("")
        lines.append("```")
        lines.append(plain[:6000])
        if len(plain) > 6000:
            lines.append("...（原文共 %d 字，此处截断）" % len(plain))
        lines.append("```")
        lines.append("")

    out = Path(out)
    out.mkdir(parents=True, exist_ok=True)
    p = out / "反思输入包.md"
    p.write_text("\n".join(lines), encoding="utf-8")
    return {"out": str(p), "materials": n, "baseline": str(baseline),
            "next": "交给 agent 重新判断，产出新 judgments，再跑 batch.py --from-json 重出结果"}


def build_compare(new_dir, baseline, out):
    new = load_scores(new_dir)
    old = load_scores(baseline)
    rows = []
    for name in sorted(set(new) | set(old)):
        a, b = old.get(name), new.get(name)
        if not a or not b:
            rows.append({"name": name, "status": "仅%s有" % ("新版" if b else "旧版")})
            continue
        dims = []
        om = {d["id"]: d for d in (a.get("dimensions") or [])}
        for d in (b.get("dimensions") or []):
            o = om.get(d["id"])
            if not o:
                continue
            dims.append({
                "dim": d.get("name"),
                "old": o.get("score", 0),
                "new": d.get("score", 0),
                "delta": round(d.get("score", 0) - o.get("score", 0), 1),
                "old_grade": o.get("verified_grade"),
                "new_grade": d.get("verified_grade"),
            })
        rows.append({
            "name": name,
            "old_total": a.get("ai_score", 0),
            "new_total": b.get("ai_score", 0),
            "delta": round(b.get("ai_score", 0) - a.get("ai_score", 0), 1),
            "dims": dims,
        })

    lines = ["# 版本对比", "", "旧版：%s" % baseline, "新版：%s" % new_dir, ""]
    for r in rows:
        if "status" in r:
            lines.append("- %s（%s）" % (r["name"][:40], r["status"]))
            continue
        arrow = "↑" if r["delta"] > 0 else ("↓" if r["delta"] < 0 else "→")
        lines.append("## %s" % r["name"][:44])
        lines.append("")
        lines.append("总分 %.1f → %.1f　%s %+.1f" % (
            r["old_total"], r["new_total"], arrow, r["delta"]))
        lines.append("")
        lines.append("| 维度 | 旧 | 新 | 变化 | 等级 |")
        lines.append("|---|---|---|---|---|")
        for d in r["dims"]:
            mark = "%+.1f" % d["delta"] if abs(d["delta"]) > 0.01 else "0"
            lines.append("| %s | %.1f | %.1f | %s | %s→%s |" % (
                d["dim"], d["old"], d["new"], mark,
                d["old_grade"], d["new_grade"]))
        lines.append("")

    out = Path(out)
    out.mkdir(parents=True, exist_ok=True)
    p = out / "版本对比.md"
    p.write_text("\n".join(lines), encoding="utf-8")
    regress = [r for r in rows if r.get("delta", 0) < -0.01]
    improve = [r for r in rows if r.get("delta", 0) > 0.01]
    return {"out": str(p), "materials": len(rows),
            "improved": len(improve), "regressed": len(regress),
            "note": "退化的重点看：改完反而低了，可能是方向改错了"}


def main():
    ap = argparse.ArgumentParser(description="init / refine / compare")
    ap.add_argument("mode", choices=["init", "refine", "compare"])
    ap.add_argument("src")
    ap.add_argument("--baseline", help="上次输出目录（refine/compare 必填）")
    ap.add_argument("--out", default=".")
    a = ap.parse_args()

    src = Path(a.src)
    if not src.is_dir():
        print("目录不存在：%s" % src, file=sys.stderr)
        return 2
    if a.mode in ("refine", "compare") and not a.baseline:
        print("%s 模式必须给 --baseline" % a.mode, file=sys.stderr)
        return 2
    base = Path(a.baseline) if a.baseline else None
    if base and not base.is_dir():
        print("baseline 不存在：%s" % base, file=sys.stderr)
        return 2

    if a.mode == "init":
        print(json.dumps({
            "mode": "init",
            "steps": [
                "1. python scripts/isolate.py prepare <新目录> --materials %s ..." % src,
                "2. 由 agent 自己读材料写 judgments，不要复用任何历史",
                "3. python scripts/batch.py <隔离目录> --out <输出> --all --bind",
                "4. python scripts/verify_checks.py <输出> --sample 3",
            ],
            "note": "judgments 必须当场写，--bind 写指纹，换材料就用不了",
        }, ensure_ascii=False, indent=1))
        return 0

    if a.mode == "refine":
        print(json.dumps(build_refine_pack(src, base, a.out),
                         ensure_ascii=False, indent=1))
        return 0

    print(json.dumps(build_compare(src, base, a.out),
                     ensure_ascii=False, indent=1))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())