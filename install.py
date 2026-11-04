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


def sha256_of(data: bytes) -> str:
    import hashlib
    return hashlib.sha256(data).hexdigest()


def fetch_expected_sha256(repo: str, branch: str, archive_name: str = "",
                          sums_ref: str = "main") -> str | None:
    """尝试取官方发布的 SHA256 清单。取不到就返回 None（不阻断安装）。

    为什么需要它：TLS 只保证「内容来自 github.com」，不保证「来自作者」。
    如果仓库或账号被入侵，下载到的就是攻击者的代码——而安装器随后会执行
    其中的 scripts/setup.py。清单把「信任来源」从「域名」收紧到「发布内容」。
    """
    import urllib.request
    # 清单固定从默认分支读，不跟着要安装的 ref 走。
    # 原因：raw.githubusercontent 按 ref 取文件，用 v3.6.0 当 ref 时拿到的
    # 是那个 tag 时刻的旧清单（甚至 200 但内容不含新版本），于是明明发布了
    # 校验和却读不到。清单是「仓库当前的发布记录」，属于 main，不属于某个 tag。
    # 两个源都试：raw.githubusercontent 有约 5 分钟强制缓存，刚发版时读到的
    # 是旧内容（实测：GitHub API 已能读到 v3.6.1，raw 仍返回 300 字节旧版）。
    # jsDelivr 镜像更新更快，作为第二源。
    sources = []
    for name in ("SHA256SUMS", "sha256sums.txt"):
        sources.append("https://raw.githubusercontent.com/%s/%s/%s"
                       % (repo, sums_ref, name))
        sources.append("https://cdn.jsdelivr.net/gh/%s@%s/%s"
                       % (repo, sums_ref, name))
    for url in sources:
        try:
            with urllib.request.urlopen(url, timeout=20) as resp:
                text = resp.read().decode("utf-8", "replace")
        except Exception:
            continue
        entries = {}
        for line in text.splitlines():
            line = line.strip()
            if not line or line.startswith("#"):
                continue
            parts = line.split()
            if len(parts) >= 2 and len(parts[0]) == 64:
                entries[parts[-1].lstrip("*")] = parts[0].lower()
            elif len(parts) == 1 and len(parts[0]) == 64:
                entries[""] = parts[0].lower()
        if not entries:
            continue
        # 只按归档文件名精确匹配。
        # 不做「只有一条就返回它」的宽松回退——那会让「清单里没有这个包」
        # 退化成「用别的包的哈希来比对」，要么误报失败，要么在校验本身被
        # 削弱时误报成功。取不到就返回 None，让调用方明确提示未校验。
        if archive_name and archive_name in entries:
            return entries[archive_name]
        # 清单里只有裸哈希（没有文件名）时，视为对「唯一这个包」的校验
        if "" in entries and len(entries) == 1:
            return entries[""]
        # 这个源没有目标归档 —— 可能是 CDN 缓存还没刷新（实测 raw 会滞后数分钟，
        # 明明 GitHub API 已能看到新版本，raw 仍返回旧内容）。继续试下一个源，
        # 不要在这里 return None。
    return None


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
    # 归档文件名：GitHub 的格式是 <仓库名>-<ref>.zip
    repo_name = repo.rstrip("/").split("/")[-1]
    tag_like = branch.startswith("v") and branch[1:2].isdigit()
    if tag_like:
        url = "https://github.com/%s/archive/refs/tags/%s.zip" % (repo, branch)
        archive_name = "%s-%s.zip" % (repo_name, branch)
    else:
        url = "https://github.com/%s/archive/refs/heads/%s.zip" % (repo, branch)
        archive_name = "%s-%s.zip" % (repo_name, branch)
    try:
        import urllib.request
        say("      下载 " + url)
        with urllib.request.urlopen(url, timeout=180) as resp:
            data = resp.read()
    except Exception as e:
        raise SystemExit("下载失败：%s\n"
                         "  可改用本地包：python install.py --from zip --zip <文件.zip>" % e)
    # 完整性校验：有官方清单就必须匹配；没有则明确告知风险而不是静默放行
    expected = fetch_expected_sha256(repo, branch, archive_name)
    actual = sha256_of(data)
    if expected:
        if actual != expected:
            raise SystemExit(
                "下载内容校验失败，已中止。\n"
                "  期望 SHA256：%s\n"
                "  实际 SHA256：%s\n"
                "  可能是分支被改过，或下载被中间人替换。请核对后再装。" % (expected, actual))
        say("      内容校验通过（SHA256 %s…）" % actual[:12])
    else:
        say("      仓库未发布 SHA256SUMS，跳过内容校验")
        say("      提示：TLS 只保证内容来自 github.com，不保证来自仓库作者")

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


def confirm_trust(source_kind: str, yes: bool) -> bool:
    """执行下载来的代码前先确认。--yes 可跳过（CI 用）。

    为什么需要：安装器会跑目标目录里的 scripts/setup.py --install。
    那是**从网络下载的代码**，而下载只靠 TLS 保证来源域名，不保证作者身份。
    在动手执行前把这件事说清楚，是最后一道用户可控的关口。
    """
    if yes or source_kind == "dir":
        return True          # 本地目录是用户自己的东西，不用问
    if not sys.stdin.isatty():
        say("      非交互环境，跳过确认（如需显式确认请加 --yes）")
        return True
    say("")
    say("  即将执行：<安装目标>/scripts/setup.py --install")
    say("  来源：%s" % ("GitHub 下载" if source_kind == "github" else "本地 zip 包"))
    say("  该脚本会检查并安装 Python 依赖（python-docx / PyMuPDF / openpyxl）。")
    try:
        ans = input("  继续？[y/N] ").strip().lower()
    except (EOFError, KeyboardInterrupt):
        return False
    return ans in ("y", "yes", "是")


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
    ap.add_argument("--yes", action="store_true",
                    help="跳过执行下载代码前的确认（CI / 自动化用）")
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
            # 执行下载来的代码前先确认。这是最后一道用户可控的关口：
            # 下载只靠 TLS 保证「来自 github.com」，不保证「来自作者」。
            if any((t / "scripts" / "setup.py").is_file() for t in targets):
                if not confirm_trust(a.src, a.yes):
                    say("      已取消。技能文件已就位，依赖请稍后手动准备：")
                    say("        python <安装目标>/scripts/setup.py --install")
                    return 0
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