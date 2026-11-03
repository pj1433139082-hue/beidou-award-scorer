#!/usr/bin/env python3
# -*- coding: utf-8 -*-
r"""抽查机检结论对不对——补上验证体系的最后一个盲区。

## 为什么需要它

现有三层验证各有盲区：

    selftest_pipeline   验代码逻辑
    verify_outputs      验文件在不在
    verify_format       验格式对不对
    ──────────────────────────────
    ✗ 没人验「机检说 A 材料的 C08 有问题」这个**判断**对不对

这个盲区有真实代价。C08 曾用 \d{2,6} 匹配金额，把 0.50万元 读成 50.0、
6.00万元 读成 0.0——代码没报错、格式没错，只是结论错了，于是活了很久，
直到一个外部独立评审回原文逐条核对才发现。

## 它做什么

随机抽几份材料，把每条 fail/warn 连**原文上下文**摆出来，让人一眼判断
这个结论站不站得住。不自动判定——机器判断不了"机器判断对不对"。

用法：
    python scripts/verify_checks.py <目录> --sample 3
    python scripts/verify_checks.py <目录> --sample 5 --only C08,C06
    python scripts/verify_checks.py <目录> --material 水厂设备卫士
"""
from __future__ import annotations
import argparse, json, random, re, sys
from pathlib import Path


def load(p: Path):
    try:
        return json.loads(p.read_text(encoding="utf-8"))
    except Exception:
        return None


def context_of(plain: str, needle: str, width: int = 40) -> str:
    """在原文里找 needle，返回带上下文的片段。"""
    if not needle or len(needle) < 2:
        return ""
    i = plain.find(str(needle)[:30])
    if i < 0:
        i = plain.find(str(needle)[:10])
        if i < 0:
            return ""
    s = max(0, i - width)
    e = min(len(plain), i + len(str(needle)) + width)
    return plain[s:e].replace("\n", " / ")


def probe_patterns(cid: str, plain: str) -> list:
    """针对不同检查项，从原文捞它该看的证据。

    目的：让人能自己判断机检说的是不是真的，而不是只看机检一家之言。
    """
    out = []
    if cid == "C08":
        seen = []
        for m in re.finditer(r"[\d.]+\s*万元", plain):
            seg = plain[max(0, m.start() - 20):m.end() + 10].replace("\n", " ").strip()
            if seg not in seen:
                seen.append(seg)
        out = ["金额：%s" % s for s in seen[:6]]
    elif cid == "C07":
        for kw in ("公式", "计算步骤", "计算过程", "测算过程", "计算方法"):
            if kw in plain:
                out.append("原文含「%s」" % kw)
    elif cid == "C06":
        for m in list(re.finditer(
                r"\d{1,3}\s*(?:家|个)[\u4e00-\u9fa5]{0,4}(?:项目|厂站|厂|水厂)", plain))[:5]:
            out.append("推广：%s" % plain[max(0, m.start() - 15):m.end() + 10].replace("\n", " "))
    elif cid == "C13":
        for m in list(re.finditer(r"(迭代|升级|改造|上一届|上届|往届|原版本)", plain))[:4]:
            out.append("迭代：%s" % plain[max(0, m.start() - 25):m.end() + 25].replace("\n", " "))
    elif cid in ("C09", "C17"):
        out.append("（格式类，只提醒不扣分）")
    return out


def main() -> int:
    ap = argparse.ArgumentParser(description="抽查机检结论")
    ap.add_argument("src", help="含 *.checks.json 与 *.doc.json 的目录")
    ap.add_argument("--sample", type=int, default=3, help="随机抽几份")
    ap.add_argument("--material", help="指定材料名片段")
    ap.add_argument("--only", help="只看这些检查项，逗号分隔")
    ap.add_argument("--seed", type=int, default=0, help="随机种子")
    a = ap.parse_args()

    src = Path(a.src)
    if not src.is_dir():
        print("目录不存在：%s" % src, file=sys.stderr)
        return 2
    checks = sorted(src.glob("*.checks.json"))
    if not checks:
        print("没有 *.checks.json。先跑 batch.py。", file=sys.stderr)
        return 2

    only = {x.strip().upper() for x in a.only.split(",")} if a.only else None

    if a.material:
        picked = [c for c in checks if a.material in c.name]
        if not picked:
            print("没有匹配「%s」的材料" % a.material, file=sys.stderr)
            return 2
    else:
        random.seed(a.seed)
        picked = random.sample(checks, min(a.sample, len(checks)))

    print("=" * 70)
    print("机检结论抽查：%d / %d 份材料" % (len(picked), len(checks)))
    print("=" * 70)

    flagged = 0
    for cf in picked:
        c = load(cf)
        if not c:
            continue
        name = cf.name.replace(".checks.json", "")
        doc = load(src / (name + ".doc.json")) or {}
        plain = doc.get("plain", "")

        items = [x for x in (c.get("checks") or [])
                 if x.get("status") in ("fail", "warn")
                 and (not only or x.get("id", "").upper() in only)]

        print()
        print("-- %s" % name[:60])
        if not items:
            print("   （无 fail/warn）")
            continue

        for x in items:
            flagged += 1
            print()
            print("  [%s] %-6s %s" % (x.get("id"), x.get("status"), x.get("name", "")))
            print("     机检说：%s" % str(x.get("detail", ""))[:110])

            ev = x.get("evidence") or []
            if ev:
                print("     机检证据：")
                for e in ev[:3]:
                    print("       . %s" % str(e)[:95])

            probes = probe_patterns(x.get("id", ""), plain)
            if probes:
                print("     原文实际：")
                for p in probes[:4]:
                    print("       %s" % p[:100])

            hits = 0
            for e in ev[:3]:
                ctx = context_of(plain, str(e))
                if ctx:
                    hits += 1
                    if hits <= 2:
                        print("     上下文：…%s…" % ctx[:110])
            if ev and hits == 0:
                print("     [!] 证据在原文找不到 —— 这条结论可能有问题")

    print()
    print("=" * 70)
    print("共 %d 条 fail/warn 待人工判断" % flagged)
    print("标准：机检给的证据，原文里找得到吗？找到了支持它的结论吗？")
    print("=" * 70)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())