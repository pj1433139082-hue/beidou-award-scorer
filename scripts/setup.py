#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""依赖体检 + 一键安装。

用法：
  python scripts/setup.py            # 只体检，缺什么报什么
  python scripts/setup.py --install  # 体检 + 自动安装缺失依赖

依赖只有两个（都在 PyPI）：
  python-docx  读 .docx
  PyMuPDF      读 .pdf

不联网也能用吗？不能——但本脚本在**没有**这些包时依然能跑（它只用标准库），
所以体检本身永远可用。安装走 pip，需要网络。
"""
from __future__ import annotations
import argparse, importlib, importlib.util, shutil, subprocess, sys
from pathlib import Path

REQUIRED = [
    ("docx", "python-docx", "读 .docx 申报材料"),
    ("fitz", "PyMuPDF", "读 .pdf 申报材料"),
]
OPTIONAL = [
    ("yaml", "PyYAML", "仅调试用，非必需"),
]

MIN_PY = (3, 10)


def find_interpreter() -> str | None:
    """找一个可用的 Python：当前解释器 -> python3 -> python。"""
    if sys.version_info >= MIN_PY:
        return sys.executable
    for name in ("python3", "python", "py"):
        p = shutil.which(name)
        if p:
            try:
                out = subprocess.run([p, "-c", "import sys;print('%d.%d' % sys.version_info[:2])"],
                                     capture_output=True, text=True, timeout=20)
                ver = tuple(int(x) for x in out.stdout.strip().split("."))
                if ver >= MIN_PY:
                    return p
            except Exception:
                continue
    return None


def check(interpreter: str) -> tuple[list, list, list]:
    """返回 (已装, 缺失, 装了但版本旧/可选)。"""
    have, missing, optional = [], [], []
    code = (
        "import importlib.util,json,sys\n"
        "mods=%r\n"
        "res={}\n"
        "for entry in mods:\n" +
        "    name = entry[0]\n"
        "    spec=importlib.util.find_spec(name)\n"
        "    res[name]=bool(spec)\n"
        "    if spec:\n"
        "        try:\n"
        "            m=importlib.import_module(name)\n"
        "            res[name+'_ver']=getattr(m,'__version__','?')\n"
        "        except Exception as e:\n"
        "            res[name+'_ver']='import-error:'+str(e)[:60]\n"
        "print(json.dumps(res))\n"
    ) % (REQUIRED + OPTIONAL,)
    out = subprocess.run([interpreter, "-c", code], capture_output=True, text=True, timeout=120)
    if out.returncode != 0:
        raise RuntimeError("体检脚本执行失败：\n" + out.stderr[-500:])
    import json
    res = json.loads(out.stdout.strip().splitlines()[-1])
    for mod, pkg, why in REQUIRED:
        if res.get(mod):
            have.append((pkg, res.get(mod + "_ver", "?"), why))
        else:
            missing.append((pkg, why))
    for mod, pkg, why in OPTIONAL:
        if res.get(mod):
            optional.append((pkg, res.get(mod + "_ver", "?"), why))
    return have, missing, optional


def pip_install(interpreter: str, packages: list[str]) -> bool:
    cmd = [interpreter, "-m", "pip", "install", "--disable-pip-version-check", "-q"] + packages
    print("  执行：", " ".join(cmd))
    try:
        r = subprocess.run(cmd, timeout=600)
        return r.returncode == 0
    except subprocess.TimeoutExpired:
        print("  安装超时（10 分钟）")
        return False


def self_test(interpreter: str) -> bool:
    """装完必须自测：造一份最小 docx，跑通抽取→机检。"""
    print("\n[自测] 造一份最小 docx，跑 抽取 → 机检 …")
    probe = (
        "import sys, tempfile, os, subprocess, json\n"
        "sys.path.insert(0, %r)\n"
        "import docx\n"
        "tmp = tempfile.mkdtemp()\n"
        "d = docx.Document()\n"
        "d.add_paragraph('自测文档：曝气系统能耗改造，电单耗由 0.1069 降至 0.0784。')\n"
        "p = os.path.join(tmp, 'selftest.docx')\n"
        "d.save(p)\n"
        "r = subprocess.run([sys.executable, os.path.join(%r, 'extract_report.py'), p, '--out', tmp],\n"
        "                   capture_output=True, text=True)\n"
        "assert r.returncode == 0, r.stderr[-400:]\n"
        "doc = json.load(open(os.path.join(tmp, 'selftest.doc.json'), encoding='utf-8'))\n"
        "assert doc['stats']['chars'] > 10, doc['stats']\n"
        "r2 = subprocess.run([sys.executable, os.path.join(%r, 'checks.py'),\n"
        "                     os.path.join(tmp, 'selftest.doc.json'), '--award', %r],\n"
        "                    capture_output=True, text=True)\n"
        "assert r2.returncode == 0, r2.stderr[-400:]\n"
        "print('SELFTEST_OK', doc['stats']['chars'])\n"
    ) % (str(Path(__file__).resolve().parents[1]), str(Path(__file__).resolve().parent),
         str(Path(__file__).resolve().parent), "示例奖A")
    out = subprocess.run([interpreter, "-c", probe], capture_output=True, text=True, timeout=180)
    ok = "SELFTEST_OK" in out.stdout
    print("  " + ("通过：" + out.stdout.strip().split("SELFTEST_OK")[-1].strip() + " 字" if ok else "失败\n" + out.stderr[-600:]))
    return ok


def main() -> int:
    ap = argparse.ArgumentParser(description="北斗奖评分技能：依赖体检与安装")
    ap.add_argument("--install", action="store_true", help="自动安装缺失依赖")
    ap.add_argument("--no-self-test", action="store_true", help="跳过安装后的自测")
    a = ap.parse_args()

    print("=" * 62)
    print("北斗奖申报材料评分技能 · 依赖体检")
    print("=" * 62)

    interp = find_interpreter()
    if not interp:
        print(f"\n错误：找不到 Python {MIN_PY[0]}.{MIN_PY[1]} 及以上。\n"
              f"  请先安装 Python：https://www.python.org/downloads/\n"
              f"  安装时勾选「Add Python to PATH」。")
        return 2
    ver = subprocess.run([interp, "-c", "import sys;print('%d.%d.%d' % sys.version_info[:3])"],
                         capture_output=True, text=True).stdout.strip()
    print(f"\nPython: {ver}   ({interp})")

    try:
        have, missing, optional = check(interp)
    except Exception as e:
        print("体检失败：" + str(e))
        return 2

    print("\n已安装：")
    for pkg, v, why in have:
        print(f"  [OK]   {pkg:<14} {v:<10} {why}")
    if optional:
        print("\n可选（未装不影响使用）：")
        for pkg, v, why in optional:
            print(f"  [--]   {pkg:<14} {v:<10} {why}")
    if not missing:
        print("\n依赖齐全，无需安装。")
    else:
        print("\n缺失：")
        for pkg, why in missing:
            print(f"  [MISS] {pkg:<14} {why}")

    if missing and not a.install:
        print("\n下一步：运行  python scripts/setup.py --install  自动安装")
        return 1

    if missing:
        print("\n开始安装 …")
        ok = pip_install(interp, [p for p, _ in missing])
        if not ok:
            print("\n安装失败。可手动执行：")
            print("  " + interp + " -m pip install " + " ".join(p for p, _ in missing))
            return 2
        print("  安装完成，重新体检 …")
        have, missing, _ = check(interp)
        if missing:
            print("  仍有缺失：" + ", ".join(p for p, _ in missing))
            return 2
        print("  依赖齐了。")

    if not a.no_self_test:
        if not self_test(interp):
            print("\n自测未通过——技能可能无法正常工作，请检查上面的错误信息。")
            return 2

    print("\n" + "=" * 62)
    print("准备就绪。现在可以直接用了：")
    print("  在对话里说「给这份申报材料打分」并附上材料；")
    print("  或手动跑：python scripts/batch.py \"<材料目录>\" --out out/")
    print("=" * 62)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())