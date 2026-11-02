#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""把技能打成可分发的 zip 包。

产物：beidou-award-scorer-<version>.zip
  · 内含技能本体（SKILL.md / scripts / assets / references / docs）
  · 内含 VERSION、CHECKSUMS.sha256、INSTALL.txt

用法：
  python scripts/make_package.py                    # 打到 dist/
  python scripts/make_package.py --out dist/ --version 1.0.0
  python scripts/make_package.py --out dist/ --release    # 顺带建 GitHub Release
"""
from __future__ import annotations
import argparse, hashlib, json, os, shutil, subprocess, sys, zipfile
from pathlib import Path

SKILL_NAME = "beidou-award-scorer"
ROOT = Path(__file__).resolve().parents[1]

# 打包时排除的东西
EXCLUDE_DIRS = {".git", "__pycache__", ".pytest_cache", ".mypy_cache", ".ruff_cache", "node_modules"}
# Office 打开文件时留下的锁文件（~$xxx.xlsx）、系统缩略图等，绝不能进包
EXCLUDE_FILES = {".DS_Store", "Thumbs.db", "desktop.ini"}
EXCLUDE_PREFIX = ("~$", "._")
EXCLUDE_SUFFIX = {".pyc", ".pyo"}

# 必须包含（缺了就说明技能不完整）
REQUIRED = [
    "SKILL.md",
    "install.py",
    "scripts/setup.py",
    "scripts/extract_report.py",
    "scripts/checks.py",
    "scripts/score.py",
    "scripts/batch.py",
    "scripts/export_xlsx.py",
    "scripts/sync_check.py",
    # 评分卡：公开副本是 rubric.example.json，私有副本是 rubric.json

    "assets/judgments-template.md",
    "references/scoring-method.md",
    "references/extending.md",
]

INSTALL_TXT = """北斗奖申报材料评分技能 · 安装说明
================================================

方式一（推荐，一条命令）：

    python install.py --from zip --zip ./%(name)s.zip

  从本包安装到 ~/.agents/skills/ 与 ~/.qwenworkcn/skills/，
  自动装依赖（python-docx、PyMuPDF）并做自测。

方式二（GitHub 自动拉最新版）：

    python install.py --from github

方式三（直接解压使用）：

    解压后把 %(name)s 目录放进你的 agent 技能目录，例如
    ~/.agents/skills/%(name)s

依赖：
  Python 3.10+ ，python-docx （读 .docx） ，PyMuPDF （读 .pdf）
  体检命令： python scripts/setup.py
  自助安装： python scripts/setup.py --install

校验：
  sha256sum -c CHECKSUMS.sha256     （Linux/macOS）
  certutil -hashfile %(name)s.zip SHA256   （Windows，看下方 SHA256 值）

%(sha)s  %(name)s.zip
"""


def version_of(explicit: str | None) -> str:
    if explicit:
        return explicit
    # 优先读 git 描述
    try:
        r = subprocess.run(["git", "describe", "--tags", "--abbrev=0"], cwd=ROOT,
                           capture_output=True, text=True, timeout=20)
        if r.returncode == 0 and r.stdout.strip():
            return r.stdout.strip().lstrip("v")
    except Exception:
        pass
    return "0.0.0-dev"


def collect_files() -> list[Path]:
    out = []
    for p in sorted(ROOT.rglob("*")):
        if not p.is_file():
            continue
        rel = p.relative_to(ROOT)
        if any(part in EXCLUDE_DIRS for part in rel.parts):
            continue
        if rel.name in EXCLUDE_FILES or rel.suffix in EXCLUDE_SUFFIX:
            continue
        if rel.name.startswith(EXCLUDE_PREFIX):
            continue
        # 打分产物目录不进分发包（体积大、且是用户自己的数据）
        if rel.parts and rel.parts[0] in ("dist", "out", "评分结果", "意见", ".scratch"):
            continue
        out.append(rel)
    return out


def sha256_of(path: Path) -> str:
    h = hashlib.sha256()
    with path.open("rb") as f:
        for chunk in iter(lambda: f.read(1 << 16), b""):
            h.update(chunk)
    return h.hexdigest()


def main() -> int:
    ap = argparse.ArgumentParser(description="打包技能为 zip")
    ap.add_argument("--out", default="dist", help="输出目录（默认 dist/）")
    ap.add_argument("--version", help="版本号（默认取 git tag）")
    ap.add_argument("--release", action="store_true", help="顺带创建 GitHub Release")
    ap.add_argument("--repo", default="pj1433139082-hue/beidou-award-scorer")
    a = ap.parse_args()

    # 完整性检查
    # 评分卡二选一即可
    has_rubric = any((ROOT / "assets" / n).is_file() for n in ("rubric.json", "rubric.example.json"))
    missing = [r for r in REQUIRED if not (ROOT / r).is_file()]
    if not has_rubric:
        missing.append("assets/rubric.json 或 assets/rubric.example.json")
    if missing:
        print("技能不完整，缺少：", file=sys.stderr)
        for m in missing:
            print("  - " + m, file=sys.stderr)
        return 2

    ver = version_of(a.version)
    outdir = Path(a.out).resolve()
    outdir.mkdir(parents=True, exist_ok=True)
    pkgname = "%s-%s" % (SKILL_NAME, ver)
    zpath = outdir / (pkgname + ".zip")

    files = collect_files()
    print("打包 %s v%s，%d 个文件" % (SKILL_NAME, ver, len(files)))

    # 校验和清单（包内每个文件的 sha256）
    checks = []
    for rel in files:
        checks.append("%s  %s" % (sha256_of(ROOT / rel), rel.as_posix()))
    checks_text = "\n".join(checks) + "\n"

    pkg_sha = ""   # 打包后才知道

    with zipfile.ZipFile(zpath, "w", zipfile.ZIP_DEFLATED, compresslevel=9) as z:
        for rel in files:
            z.write(ROOT / rel, arcname="%s/%s" % (pkgname, rel.as_posix()))
        z.writestr("%s/VERSION" % pkgname, ver + "\n")
        z.writestr("%s/CHECKSUMS.sha256" % pkgname, checks_text)
        z.writestr("%s/INSTALL.txt" % pkgname, INSTALL_TXT % {
            "name": pkgname, "sha": "(见下方说明)"})
        # 把包自身的 sha 写进 INSTALL.txt（占位后重写一次）
    pkg_sha = sha256_of(zpath)
    with zipfile.ZipFile(zpath, "w", zipfile.ZIP_DEFLATED, compresslevel=9) as z:
        for rel in files:
            z.write(ROOT / rel, arcname="%s/%s" % (pkgname, rel.as_posix()))
        z.writestr("%s/VERSION" % pkgname, ver + "\n")
        z.writestr("%s/CHECKSUMS.sha256" % pkgname, checks_text)
        z.writestr("%s/INSTALL.txt" % pkgname, INSTALL_TXT % {"name": pkgname, "sha": pkg_sha})

    size = zpath.stat().st_size
    print("写出 %s（%.1f KB）" % (zpath, size / 1024))
    print("包 SHA256: %s" % pkg_sha)

    # 自检：解出来能用
    import tempfile
    tmp = Path(tempfile.mkdtemp())
    try:
        with zipfile.ZipFile(zpath) as z:
            z.extractall(tmp)
        root = tmp / pkgname
        missing2 = [r for r in REQUIRED if not (root / r).is_file()]
        if not any((root / "assets" / n).is_file() for n in ("rubric.json", "rubric.example.json")):
            missing2.append("assets 评分卡")
        if missing2:
            print("包内缺文件：", missing2, file=sys.stderr)
            return 2
        r = subprocess.run([sys.executable, str(root / "scripts" / "setup.py"), "--no-self-test"],
                           capture_output=True, text=True, timeout=300)
        if r.returncode != 0:
            print("包内 setup.py 体检失败：\n" + r.stdout[-500:] + r.stderr[-300:], file=sys.stderr)
            return 2
        print("包内自检通过")
    finally:
        shutil.rmtree(tmp, ignore_errors=True)

    if a.release:
        tag = "v%s" % ver
        print("\n创建 GitHub Release %s …" % tag)
        r = subprocess.run(["gh", "release", "create", tag, str(zpath),
                            "--repo", a.repo,
                            "--title", "%s v%s" % (SKILL_NAME, ver),
                            "--notes", "北斗奖申报材料评分技能 %s\n\n"
                                       "安装：\n\n"
                                       "    python install.py --from zip --zip ./%s\n\n"
                                       "或从 GitHub 直接装：\n\n"
                                       "    python install.py --from github\n\n"
                                       "SHA256: %s" % (ver, zpath.name, pkg_sha)],
                           capture_output=True, text=True)
        print(r.stdout.strip() or r.stderr.strip()[-400:])
        if r.returncode != 0:
            return r.returncode
    return 0


if __name__ == "__main__":
    raise SystemExit(main())