#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""两份副本的脚本必须逐字节一致。

本地私有副本（真实评分卡）与公开仓库（脱敏示例卡）共用同一份 scripts/。
任一被单边修改都会让两边行为分叉——那个问题在真实材料上表现为机检结果不一致。

用法：
  python scripts/sync_check.py <另一份技能目录>
目录里没有 scripts/ 时视为配置错误，退出码 2。
"""
from __future__ import annotations
import hashlib, sys
from pathlib import Path

SHARED = ["scripts/extract_report.py", "scripts/checks.py", "scripts/score.py", "scripts/batch.py",
          "references/scoring-method.md", "references/extending.md",
          "assets/judgments-template.md"]

def digest(p: Path) -> str:
    return hashlib.sha256(p.read_bytes()).hexdigest()

def main() -> int:
    if len(sys.argv) != 2:
        print("用法：python scripts/sync_check.py <另一份技能目录>", file=sys.stderr)
        return 2
    here = Path(__file__).resolve().parents[1]
    other = Path(sys.argv[1]).resolve()
    if not (other / "scripts").is_dir():
        print(f"不是技能目录：{other}", file=sys.stderr)
        return 2
    bad = []
    for rel in SHARED:
        a, b = here / rel, other / rel
        if not a.is_file() or not b.is_file():
            bad.append(f"{rel}: 缺失（{a if not a.is_file() else b}）")
            continue
        if digest(a) != digest(b):
            bad.append(f"{rel}: 内容不一致")
    if bad:
        print("两份副本已分叉：")
        for line in bad:
            print("  - " + line)
        print("\n修法：在一处改完，复制到另一处，再跑一次本检查。")
        return 1
    print(f"一致：{len(SHARED)} 个共享文件无差异")
    return 0

if __name__ == "__main__":
    raise SystemExit(main())
