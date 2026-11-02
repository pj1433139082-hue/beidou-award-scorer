#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""一键安装北斗奖评分技能。

三种来源，任选其一：

  1) GitHub（推荐，自动拿最新版）
     python install.py --from github

  2) 本地 zip 包
     python install.py --from zip --zip ./beidou-award-scorer.zip

  3) 本地目录
     python install.py --from dir --dir ./beidou-award-scorer

装到哪里（默认装两处，让 agent 和 qwen 都能用）：
     ~/.agents/skills/beidou-award-scorer
     ~/.qwenworkcn/skills/beidou-award-scorer

装完自动做依赖体检与自测。装不上的位置会跳过并报告，不中断。

参数：
  --from {github,zip,dir}   来源（默认 github）
  --github <repo>          GitHub 仓库（默认 pj1433139082-hue/beidou-award-scorer）
  --branch <branch>        分支（默认 main）
  --zip <path>             zip 包路径
  --dir <path>             目录路径
  --target <path>          只装到这一处（可重复）
  --no-setup               不跑依赖体检与安装
  --no-self-test           跳过自测
"""
from __future__ import annotations
import argparse, io, os, shutil, subprocess, sys, tempfile, zipfile
from pathlib import Path

SKILL_NAME = "beidou-award-scorer"
DEFAULT_REPO = "pj1433139082-hue/beidou-award-scorer"

HOME = Path.home()
DEFAULT_TARGETS = [
    HOME / ".agents" / "skills" / SKILL_NAME,
    HOME / ".qwenworkcn" / "skills" / SKILL_NAME,
]


def say(msg: str) -> None:
    print(msg, flush=True)


def default_targets() -> list[Path]:
    """只保留真实存在的技能根目录。"""
    out = []
    for t in DEFAULT_TARGETS:
        if t.parent.is_dir():
            out.append(t)
    if not out:                       # 都没有就退回用户主目录
        out.append(HOME / f".{SKILL_NAME}")
    return out


def fetch_github(repo: str, branch: str) -> Path:
    """优先用 git clone；没有 git 就退回下载 zip。"""
    tmp = Path(tempfile.mkdtemp(prefix="beidou-install-"))
    say(f"[1/4] 从 GitHub 拉取 {repo}@{branch} …")
    if shutil.which("git"):
        dest = tmp / SKILL_NAME
        r = subprocess.run(["git", "clone", "--depth", "1", "--branch", branch,
                            "https://github.com/%s.git" % repo, str(dest)],
                           capture_output=True, text=True)
        if r.returncode == 0:
            say("      git clone 完成")
            return dest
        say("      git clone 失败，改用 zip 下载：" + r.stderr.strip().splitlines()[-1][:120])
    url = "https://github.com/%s/archive/refs/heads/%s.zip" % (repo, branch)
    try:
        import urllib.request
        say("      下载 " + url)
        with urllib.request.urlopen(url, timeout=180) as resp:
            data = resp.read()
    except Exception as e:
        raise SystemExit("下载失败：%s\n"
                         "  可改用本地包：python install.py --from zip --zip <文件.zip>" % e)
    zpath = tmp / "src.zip"
    zpath.write_bytes(data)
    return unpack_zip(zpath, tmp)


def unpack_zip(zpath: Path, workdir: Path) -> Path:
    say(f"[1/4] 解压 {zpath} …")
    with zipfile.ZipFile(zpath) as z:
        z.extractall(workdir)
    # GitHub 的 zip 会带一层 <repo>-<branch>/，找含 SKILL.md 的那一层
    for p in sorted(workdir.rglob("SKILL.md")):
        return p.parent
    raise SystemExit("zip 里没找到 SKILL.md，确认这是技能包而不是别的压缩文件。")


def copy_skill(src: Path, dest: Path) -> bool:
    """复制到 dest。dest 已存在且不是软链时，先备份。"""
    dest.parent.mkdir(parents=True, exist_ok=True)
    if dest.exists() or dest.is_symlink():
        if dest.is_symlink() or dest.is_junction():
            try:
                dest.unlink()
            except OSError:
                shutil.rmtree(dest, ignore_errors=True)
        else:
            backup = dest.with_name(dest.name + ".bak")
            if backup.exists():
                shutil.rmtree(backup, ignore_errors=True)
            shutil.move(str(dest), str(backup))
            say("      已存在，旧版备份到 %s" % backup.name)
    shutil.copytree(src, dest, ignore=shutil.ignore_patterns(".git", "__pycache__", "*.pyc"))
    return True


def run_setup(target: Path, do_self_test: bool) -> bool:
    """在目标目录里跑依赖体检与安装。"""
    setup = target / "scripts" / "setup.py"
    if not setup.is_file():
        return False
    args = [sys.executable, str(setup), "--install"]
    if not do_self_test:
        args.append("--no-self-test")
    say("      依赖体检与安装 …")
    r = subprocess.run(args, capture_output=True, text=True, timeout=900)
    tail = [l for l in r.stdout.splitlines() if l.strip()][-6:]
    for l in tail:
        say("        " + l)
    if r.returncode != 0:
        say("        依赖准备未完全成功（可稍后手动跑 python scripts/setup.py --install）")
    return r.returncode == 0


def main() -> int:
    ap = argparse.ArgumentParser(description="一键安装北斗奖评分技能")
    ap.add_argument("--from", dest="src", choices=["github", "zip", "dir"], default="github")
    ap.add_argument("--github", default=DEFAULT_REPO)
    ap.add_argument("--branch", default="main")
    ap.add_argument("--zip", dest="zip_path")
    ap.add_argument("--dir", dest="dir_path")
    ap.add_argument("--target", action="append", help="安装目标目录（可重复）")
    ap.add_argument("--no-setup", action="store_true")
    ap.add_argument("--no-self-test", action="store_true")
    a = ap.parse_args()

    say("=" * 64)
    say("北斗奖申报材料评分技能 · 安装")
    say("=" * 64)

    tmp = Path(tempfile.mkdtemp(prefix="beidou-inst-"))
    try:
        if a.src == "github":
            src = fetch_github(a.github, a.branch)
        elif a.src == "zip":
            if not a.zip_path:
                raise SystemExit("--from zip 需要 --zip <文件.zip>")
            src = unpack_zip(Path(a.zip_path).expanduser(), tmp)
        else:
            if not a.dir_path:
                raise SystemExit("--from dir 需要 --dir <目录>")
            src = Path(a.dir_path).expanduser().resolve()
            if not (src / "SKILL.md").is_file():
                raise SystemExit("%s 里没有 SKILL.md，不是技能目录" % src)
        say("      技能源：%s" % src)

        if not (src / "SKILL.md").is_file():
            raise SystemExit("源目录里没有 SKILL.md")

        targets = [Path(t).expanduser().resolve() for t in a.target] if a.target else default_targets()
        say("[2/4] 安装到 %d 个位置：" % len(targets))
        ok_any = False
        for t in targets:
            try:
                copy_skill(src, t)
                say("      [OK] %s" % t)
                ok_any = True
            except Exception as e:
                say("      [跳过] %s —— %s" % (t, e))

        if not ok_any:
            raise SystemExit("没有任何位置安装成功。")

        if not a.no_setup:
            say("[3/4] 准备依赖")
            for t in targets:
                if (t / "scripts" / "setup.py").is_file():
                    run_setup(t, not a.no_self_test)
        else:
            say("[3/4] 跳过依赖准备（--no-setup）")

        say("[4/4] 校验")
        for t in targets:
            sk = t / "SKILL.md"
            sc = t / "scripts" / "score.py"
            say("      %s  SKILL.md:%s  score.py:%s" % (
                t, "有" if sk.is_file() else "缺", "有" if sc.is_file() else "缺"))

        say("\n" + "=" * 64)
        say("安装完成。现在可以直接用了：")
        say("  · 在对话里说「给这份申报材料打分」并附上材料（docx 或 pdf）")
        say("  · 批量：python %s/scripts/batch.py \"<材料目录>\" --out out/" % (targets[0] if targets else ""))
        say("  · 体检：python %s/scripts/setup.py" % (targets[0] if targets else ""))
        say("=" * 64)
        return 0
    finally:
        shutil.rmtree(tmp, ignore_errors=True)


if __name__ == "__main__":
    raise SystemExit(main())
