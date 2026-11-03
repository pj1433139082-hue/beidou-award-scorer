#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""判定一份材料的角色，以及它正文声明的赛道。

为什么需要：
  1. 申报材料里可能夹带"过往版本"/"上一版"这类**参考材料**——它是用来对照迭代的，
     不是独立参赛项目。把它当参赛项目排名，会凭空多出一个 47 分的"弱项目"。
  2. 文件名里的奖项前缀可能与正文声明的赛道不一致（改赛道时忘了改名）。
     只看文件名就会用错评分维度。

用法（命令行自检）：
  python scripts/material_role.py "金火炬-xxx-过往版本.docx"
"""
from __future__ import annotations
import re, sys

# 文件名里出现这些词，说明是参考材料而不是参赛项目
REFERENCE_MARKERS = (
    "过往版本", "过往版", "上一版本", "上一版", "旧版", "原版本", "此前版本",
    "历史版本", "往届版本", "往届", "前一版", "老版本",
)

# 正文里声明赛道的写法
TRACK_PAT = re.compile(r"参赛赛道\s*[：:]\s*([^\n\r]{2,40})")


def is_reference(name: str) -> bool:
    """这份材料是不是参考材料（过往版本等），而非参赛项目。"""
    n = str(name or "")
    return any(m in n for m in REFERENCE_MARKERS)


def declared_track(plain: str) -> str | None:
    """从正文前部取「参赛赛道：xxx」。取不到返回 None。"""
    if not plain:
        return None
    head = plain[:3000]
    m = TRACK_PAT.search(head)
    if not m:
        return None
    t = m.group(1).strip()
    # 「参赛赛道：」后面紧跟换行时，正则可能抓到下一行的标签
    if t.endswith(("：", ":")) or len(t) < 3:
        return None
    return t


def track_matches(award: str, track: str | None) -> bool | None:
    """正文声明的赛道是否与该奖项一致。

    None = 无法判定（正文没写赛道），True/False = 明确一致/不一致。
    """
    if not track:
        return None
    if award in track:
        return True
    # AI 融创的写法不统一（特别奖-AI融创 / AI 融创奖 / AI融创（生产））
    if award == "AI融创" and re.search(r"AI\s*融创", track):
        return True
    return False


def main() -> int:
    if len(sys.argv) < 2:
        print("用法：python scripts/material_role.py <材料文件名>", file=sys.stderr)
        return 2
    name = sys.argv[1]
    print("参考材料" if is_reference(name) else "参赛项目")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
