#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""生成 / 校验仓库根目录的 SHA256SUMS。

## 它解决什么

install.py 从 GitHub 下载 zip 后，会读仓库根目录的 SHA256SUMS 做完整性校验。
没有这个文件时，安装器只能打印「未发布清单，跳过校验」——TLS 保证内容来自
github.com，但不保证来自仓库作者。

## 为什么能做到「自我维持」

GitHub 的 archive 接口（refs/heads/<branch>.zip）对同一个 commit 是**字节稳定**的，
实测两次下载 SHA256 完全相同。所以：

    提交 SHA256SUMS（记录提交前那一刻的 zip 哈希）
      → 这次提交改变了 zip
      → 清单与当前 zip 不再匹配
      → CI 立刻发现并报错
      → 重新生成，形成新的提交
      → 但新提交又会改变 zip ……

看起来是死循环。解法是：**清单只保证「install.py 拿到的包与发布时一致」**，
而发布走的是 Release 附件（版本化 tag），不是分支归档。

分支归档的哈希无法自指（记录它就会改变它），所以本脚本给出两种模式：

    --write-release   为 Release 附件生成清单（真正可校验的那条路径）
    --check           校验当前清单是否与指定 zip 匹配（供 CI 用）

## 实际生效的校验路径

install.py 优先用历史版本 tag 的归档 → 哈希稳定 → 清单可校验。
分支归档只能给出「参考哈希」，读不到就跳过并提示。
"""
from __future__ import annotations
import argparse, hashlib, subprocess, sys, urllib.request
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
SUMS = ROOT / "SHA256SUMS"


def sha256_bytes(data: bytes) -> str:
    return hashlib.sha256(data).hexdigest()


def sha256_file(p: Path) -> str:
    h = hashlib.sha256()
    with open(p, "rb") as f:
        for chunk in iter(lambda: f.read(1 << 20), b""):
            h.update(chunk)
    return h.hexdigest()


def github_archive(repo: str, ref: str) -> bytes:
    url = "https://github.com/%s/archive/refs/%s.zip" % (repo, ref)
    with urllib.request.urlopen(url, timeout=180) as r:
        return r.read()


def read_sums() -> dict:
    """读 SHA256SUMS，返回 {文件名: 哈希}。"""
    out = {}
    if not SUMS.is_file():
        return out
    for line in SUMS.read_text(encoding="utf-8").splitlines():
        line = line.strip()
        if not line or line.startswith("#"):
            continue
        parts = line.split()
        if len(parts) >= 2 and len(parts[0]) == 64:
            out[parts[-1].lstrip("*")] = parts[0].lower()
    return out


def main() -> int:
    ap = argparse.ArgumentParser(description="生成/校验 SHA256SUMS")
    ap.add_argument("--repo", default="pj1433139082-hue/beidou-award-scorer")
    ap.add_argument("--tag", help="为某个 tag 生成（如 v3.5.0）")
    ap.add_argument("--zip", help="校验这个本地 zip")
    ap.add_argument("--check", action="store_true", help="只校验，不写")
    ap.add_argument("--print-branch-hash", action="store_true",
                    help="打印当前分支归档的哈希（仅供参考，写入即失效）")
    a = ap.parse_args()

    sums = read_sums()

    if a.print_branch_hash:
        try:
            data = github_archive(a.repo, "heads/main")
        except Exception as e:
            print("取不到分支归档：%s" % e, file=sys.stderr)
            return 2
        h = sha256_bytes(data)
        print("当前 main 分支归档 SHA256（写入清单后会立刻失效，仅供对照）：")
        print("  %s  beidou-award-scorer-main.zip" % h)
        print("  大小 %.1f KB" % (len(data) / 1024))
        return 0

    if a.tag:
        ref = "tags/" + a.tag.lstrip("tags/")
        try:
            data = github_archive(a.repo, ref)
        except Exception as e:
            print("取不到 tag %s 的归档：%s" % (a.tag, e), file=sys.stderr)
            return 2
        h = sha256_bytes(data)
        name = "beidou-award-scorer-%s.zip" % a.tag.lstrip("tags/")
        entry = "%s  %s" % (h, name)
        if a.check:
            cur = sums.get(name)
            if cur == h:
                print("校验通过：%s" % name)
                return 0
            print("校验失败：清单里是 %s，实际 %s" % (cur or "(缺失)", h), file=sys.stderr)
            return 1
        lines = SUMS.read_text(encoding="utf-8").splitlines() if SUMS.is_file() else []
        lines = [l for l in lines if name not in l]
        lines.append(entry)
        header = ["# 北斗奖评分技能 · 发布包校验和",
                  "# 校验方式：sha256sum -c SHA256SUMS（Linux/macOS）",
                  "# 这些是 Release 附件的哈希，install.py 读取后比对下载内容。"]
        body = [l for l in lines if not l.startswith("#") and l.strip()]
        SUMS.write_text("\n".join(header + [""] + body).rstrip() + "\n", encoding="utf-8")
        print("已写入 %s：" % SUMS.name)
        print("  " + entry)
        return 0

    if a.zip:
        p = Path(a.zip)
        if not p.is_file():
            print("找不到 %s" % p, file=sys.stderr)
            return 2
        h = sha256_file(p)
        name = p.name
        cur = sums.get(name)
        if a.check:
            if cur == h:
                print("校验通过：%s" % name)
                return 0
            print("校验失败：清单里是 %s，实际 %s" % (cur or "(缺失)", h), file=sys.stderr)
            return 1
        print("%s  %s" % (h, name))
        return 0

    print("清单内容：")
    if not sums:
        print("  （空）")
    for name, h in sorted(sums.items()):
        print("  %s  %s" % (h, name))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())