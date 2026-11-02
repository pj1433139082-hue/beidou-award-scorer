#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""真实材料回归测试：用真实评分卡 + 真实申报材料验证行为稳定。

需要环境变量或默认路径能找到真实材料。找不到就跳过（不算失败）。
跑法： python scripts/selftest_real.py
"""
from __future__ import annotations
import json, os, sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
R = []

def rec(name, ok, note=""):
    R.append((name, ok, note))
    print("  %s %-44s %s" % ("[OK]  " if ok else "[SKIP]" if ok is None else "[FAIL]", name, note))

def main() -> int:
    print("=" * 70)
    print("北斗奖评分技能 · 真实材料回归")
    print("=" * 70)

    # 真实评分卡：优先私有副本，其次环境变量
    rub_path = Path(os.environ.get("BD_RUBRIC", "")) if os.environ.get("BD_RUBRIC") else None
    if not rub_path or not rub_path.is_file():
        cand = Path(r"D:\obsidian_work\cangku\paper_work\工作台\项目\北控水务集团\北斗奖评分技能\assets\rubric.json")
        rub_path = cand if cand.is_file() else None
    if not rub_path:
        print("找不到真实评分卡，跳过本组测试。")
        return 0
    rub = json.loads(rub_path.read_text(encoding="utf-8"))
    awards = list(rub["awards"].keys())
    rec("真实评分卡可加载", True, "%d 个奖项：%s" % (len(awards), "、".join(awards)))

    # 每个奖项的维度满分必须合计 100
    bad_sum = []
    for a, spec in rub["awards"].items():
        total = sum(d["max"] for d in spec["dimensions"])
        if total != 100:
            bad_sum.append("%s=%s" % (a, total))
    rec("各奖项维度满分合计 100", not bad_sum, "、".join(bad_sum) if bad_sum else "")

    # 迭代性只能是闸门，不能是计分维度
    has_iter_dim = [a for a, s in rub["awards"].items()
                    if any(d["id"] == "iteration" for d in s["dimensions"])]
    rec("iteration 不占分值（闸门非维度）", not has_iter_dim, "、".join(has_iter_dim) if has_iter_dim else "")

    # 金火炬必须声明不设现场评审
    hj = rub["awards"].get("金火炬") or {}
    rec("金火炬标记 no_onsite", bool(hj.get("no_onsite")))

    # 区间不重叠（复用 score.py 的校验）
    sys.path.insert(0, str(ROOT / "scripts"))
    import score as S
    try:
        S.validate_rubric_bands(rub)
        rec("证据等级区间无重叠", True)
    except Exception as e:
        rec("证据等级区间无重叠", False, str(e)[:60])

    # 真实材料回归（若存在基线）
    baseline = Path(os.environ.get("BD_BASELINE", "")) if os.environ.get("BD_BASELINE") else None
    docs = Path(os.environ.get("BD_DOCS", "")) if os.environ.get("BD_DOCS") else None
    if baseline and baseline.is_file() and docs and docs.is_dir():
        import subprocess
        base = json.loads(baseline.read_text(encoding="utf-8"))
        diff = []
        for f in sorted(docs.glob("*.doc.json")):
            n = f.name
            b = base.get(n)
            if not b:
                continue
            r = subprocess.run([sys.executable, str(ROOT / "scripts" / "checks.py"), str(f),
                                "--award", b["award"], "--rubric", str(rub_path)],
                               capture_output=True, text=True, encoding="utf-8")
            if r.returncode != 0:
                diff.append("%s: 机检失败" % n[:20]); continue
            got = json.loads(r.stdout.splitlines()[0])["summary"]
            if got != b["summary"]:
                diff.append("%s: %s -> %s" % (n[:20], b["summary"], got))
        rec("真实材料机检与基线一致", not diff, "；".join(diff[:2]) if diff else "%d 份" % len(base))
    else:
        rec("真实材料机检与基线一致", None, "未提供 BD_DOCS/BD_BASELINE")

    print("\n" + "=" * 70)
    hard_fail = [n for n, ok, _ in R if ok is False]
    passed = sum(1 for _, ok, _ in R if ok is True)
    print("结果：%d 通过 / %d 失败 / %d 跳过" % (
        passed, len(hard_fail), sum(1 for _, ok, _ in R if ok is None)))
    if hard_fail:
        print("失败项：" + "、".join(hard_fail))
        return 1
    return 0

if __name__ == "__main__":
    raise SystemExit(main())
