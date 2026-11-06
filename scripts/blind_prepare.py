#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""盲化工作区：为「无上下文子智能体独立评审」准备输入。

## 为什么需要它

`isolate.py` 隔离的是文件系统：把历史 judgments/score 挪走，让后跑的 agent
在磁盘上找不到先跑的答案。但它挡不住另一条泄露通道 —— **写 judgments 的
那个 agent 自己的上下文**。如果主会话刚改完稿、刚给过 89.3 分，再由它来写
独立判断，文件系统藏得再干净也没用：它记得自己给过什么。

真正的隔离要两层：
  1. 换一个**全新上下文**的 agent（subagent，不带主会话记忆）来写 judgments；
  2. 给它的输入**连文件名与元数据都不含历史分数**。

实测踩过的坑：直接把材料复制进盲区，文件名仍叫「（89.3终版）.docx」；
更隐蔽的是 checks.json / doc.json 的 `name` 与 `source` 字段也带着该名字，
子智能体读机检文件第一眼就看到分数。所以本脚本会：
  · 材料改名为中性名（默认 申报主报告）
  · 同步改写 *.checks.json / *.doc.json 的 name、source
  · 复扫全目录，确认不含任何历史分数与版本词
  · 拒绝在输出区留下 judgments / score

## 用法

    python scripts/blind_prepare.py <盲区目录> --material <材料.docx> \
        --checks <机检目录> --skill <技能目录>

    python scripts/blind_prepare.py <盲区目录> --verify

## 配套流程（见 SKILL.md「Isolated review with a blind subagent」）

    1. blind_prepare.py 准备盲区
    2. 派 subagent，prompt 只给盲区路径，不给任何分数/判断
    3. 子智能体在盲区输出 judgments.json
    4. 主会话用 score.py --require-bound 打分（不修改子智能体的判断）
    5. blind_prepare.py --verify 复核盲区未被污染
"""
from __future__ import annotations
import argparse, json, os, re, shutil, sys
from pathlib import Path

# 历史分数与版本标记：出现即视为泄露
LEAK_PATTERNS = (
    # 只匹配「分数 + 版本/终版」这种历史评分痕迹；材料正文里的
    # 「合计 79.6 分」「暂估 74.1 分」是业务计算，不能被当成泄露。
    re.compile(r"\b\d{2,3}\.\d\s*(?:分\s*)?(?:终版|版本|迭代版)"),
    re.compile(r"[（(]\s*\d{2,3}\.\d\s*[^)）]*?(?:终版|版|优化)[^)）]*[)）]"),
    re.compile(r"\b\d{2,3}\.\d\s*终版"),
    re.compile(r"(?:终版|优化版|补全版|迭代版)[（(]?\s*\d{2,3}\.\d"),
)
DEFAULT_NAME = "申报主报告"


def leaks_in(text):
    return [p.pattern for p in LEAK_PATTERNS if p.search(text)]


def scrub_json(path, old_names, new_base):
    """把机检/抽取产物里的 name、source 换成中性名。"""
    try:
        d = json.loads(path.read_text(encoding="utf-8"))
    except Exception:
        return 0
    changed = 0
    if isinstance(d, dict):
        for k in ("name", "source", "material", "material_name"):
            v = d.get(k)
            if isinstance(v, str):
                nv = v
                for old in old_names:
                    nv = nv.replace(old, new_base)
                if nv != v:
                    d[k] = nv
                    changed += 1
    if changed:
        path.write_text(json.dumps(d, ensure_ascii=False, indent=2), encoding="utf-8")
    return changed


def prepare(ws, material, checks_dir, skill_dir, name):
    ws = Path(ws)
    if ws.exists():
        print("目录已存在，先删除或换名：%s" % ws, file=sys.stderr)
        return 2
    for sub in ("材料", "机检", "输出"):
        (ws / sub).mkdir(parents=True)

    # 1. 材料改中性名
    src = Path(material)
    if not src.is_file():
        print("找不到材料：%s" % src, file=sys.stderr)
        return 2
    old_stem = src.stem
    dst = ws / "材料" / (name + src.suffix.lower())
    shutil.copy2(src, dst)

    # 2. 机检产物复制 + 洗元数据
    old_names = [old_stem, src.name]
    n_json = 0
    cd = Path(checks_dir)
    if cd.is_dir():
        for p in cd.iterdir():
            if p.suffix.lower() == ".json" and old_stem in p.name:
                suffix = p.name[len(old_stem):]
                np = ws / "机检" / (name + suffix)
                shutil.copy2(p, np)
                scrub_json(np, old_names, name)
                n_json += 1

    # 3. 技能副本：只复制运行与判断必需的目录。
    # 关键：不要复制 docs/。实测 docs/judgments-rules.md 里写着
    # 「同一批材料在这台电脑打 59.4 分，另一台打 41.5 分」——历史分数
    # 会顺着文档泄露给子智能体。docs/ 也不是打分必需。
    sd = Path(skill_dir)
    if sd.is_dir():
        (ws / "skill").mkdir(exist_ok=True)
        for sub in ("scripts", "assets", "references"):
            s = sd / sub
            if s.is_dir():
                shutil.copytree(s, ws / "skill" / sub,
                                ignore=shutil.ignore_patterns(
                                    "__pycache__", "*.pyc",
                                    "*.judgments.json", "*.score.json",
                                    "*.checks.json", "*.doc.json"))
        for fn in ("SKILL.md", "README.md"):
            if (sd / fn).is_file():
                shutil.copy2(sd / fn, ws / "skill" / fn)

    # 4. 复扫：文件名 + 可读文件内容
    problems = []
    for root, dirs, files in os.walk(ws):
        for fn in files:
            p = Path(root) / fn
            rel = str(p.relative_to(ws))
            if leaks_in(fn):
                problems.append("文件名: " + rel)
            if p.suffix.lower() in (".json", ".md", ".yaml", ".yml", ".txt"):
                try:
                    c = p.read_text(encoding="utf-8")
                except Exception:
                    continue
                hits = leaks_in(c)
                if hits:
                    # 报告首次命中上下文，便于定位
                    m = LEAK_PATTERNS[0].search(c) or LEAK_PATTERNS[2].search(c)
                    ctx = c[max(0, m.start()-30):m.end()+30].replace("\n", " ") if m else ""
                    problems.append("内容: %s  ...%s..." % (rel, ctx[:60]))

    # 5. 输出区不得有历史判断
    for pat in ("*.judgments.json", "*.score.json"):
        if list((ws / "输出").glob(pat)):
            problems.append("输出区残留 " + pat)

    result = {
        "workspace": str(ws),
        "material": dst.name,
        "json_copies": n_json,
        "blind": not problems,
        "leaks": problems[:20],
        "note": "blind=true 才可派子智能体；leaks 非空必须先修掉。",
    }
    print(json.dumps(result, ensure_ascii=False, indent=2))
    return 0 if not problems else 1


def verify(ws):
    ws = Path(ws)
    problems = []
    for root, dirs, files in os.walk(ws):
        for fn in files:
            p = Path(root) / fn
            if leaks_in(fn):
                problems.append("文件名: " + str(p.relative_to(ws)))
            if p.suffix.lower() in (".json", ".md", ".yaml", ".yml", ".txt"):
                try:
                    c = p.read_text(encoding="utf-8")
                except Exception:
                    continue
                if leaks_in(c):
                    problems.append("内容: " + str(p.relative_to(ws)))
    n_out = len(list((ws / "输出").glob("*.judgments.json")))
    out = {
        "blind": not problems,
        "leaks": problems[:20],
        "judgments_written": n_out,
        "verdict": "盲区干净" if not problems else "存在泄露，不可视为独立判断",
    }
    print(json.dumps(out, ensure_ascii=False, indent=2))
    return 0 if not problems else 1


def main() -> int:
    ap = argparse.ArgumentParser(description="盲化工作区（供独立子智能体评审）")
    ap.add_argument("workspace")
    ap.add_argument("--material")
    ap.add_argument("--checks")
    ap.add_argument("--skill")
    ap.add_argument("--name", default=DEFAULT_NAME)
    ap.add_argument("--verify", action="store_true")
    a = ap.parse_args()
    if a.verify:
        return verify(Path(a.workspace))
    if not (a.material and a.checks):
        print("prepare 需要 --material 与 --checks", file=sys.stderr)
        return 2
    return prepare(Path(a.workspace), a.material, a.checks, a.skill, a.name)


if __name__ == "__main__":
    raise SystemExit(main())