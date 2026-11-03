#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""把 docx 批量转成 PDF。

为什么要这一步：Word 的排版依赖字体与渲染引擎，同一份 docx 在不同机器上可能
重排。导出 PDF 后版面固定，适合打印和分发。

找 LibreOffice 的顺序：
  1. 环境变量 BD_SOFFICE 指定的可执行文件
  2. 系统 PATH 上的 soffice
  3. LibreOffice 默认安装位置
  4. DSH 自带的 LibreOffice Kit（需要 node + cli 两个路径）

找不到时不算失败——打印说明并以退出码 0 结束，由调用方决定是否当作缺失。

用法：
  python scripts/export_pdf.py 意见Word/ --out 意见Word/
  python scripts/export_pdf.py 意见Word/ --out pdf/ --only 汇总
"""
from __future__ import annotations
import argparse, json, os, shutil, subprocess, sys
from pathlib import Path

SOFFICE_CANDIDATES = [
    r"C:\Program Files\LibreOffice\program\soffice.exe",
    r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
    "/usr/bin/soffice", "/usr/local/bin/soffice",
    "/Applications/LibreOffice.app/Contents/MacOS/soffice",
]


def find_soffice() -> str | None:
    env = os.environ.get("BD_SOFFICE")
    if env and Path(env).is_file():
        return env
    p = shutil.which("soffice") or shutil.which("libreoffice")
    if p:
        return p
    for c in SOFFICE_CANDIDATES:
        if Path(c).is_file():
            return c
    return None


def find_libreoffice_kit() -> tuple[str, str] | None:
    """DSH 自带的 LibreOffice Kit（node + cli）。"""
    node = os.environ.get("BD_KIT_NODE")
    cli = os.environ.get("BD_KIT_CLI")
    if node and cli and Path(node).is_file() and Path(cli).is_file():
        return node, cli
    return None


def convert_with_soffice(soffice: str, src: Path, outdir: Path) -> bool:
    r = subprocess.run([soffice, "--headless", "--norestore", "--convert-to", "pdf",
                        "--outdir", str(outdir), str(src)],
                       capture_output=True, text=True, timeout=300)
    return r.returncode == 0 and (outdir / (src.stem + ".pdf")).is_file()


def convert_with_kit(node: str, cli: str, src: Path, out: Path) -> bool:
    """单份转换（兜底）。批量请走 pdf_batch.py —— 快 3 倍。"""
    r = subprocess.run([node, cli, "convert", "--input", str(src), "--output", str(out)],
                       capture_output=True, text=True, timeout=300)
    return r.returncode == 0 and out.is_file()


def convert_all_with_kit(files: list, outdir: Path) -> tuple:
    """批量转换：一次启动 worker，把所有 docx 转完。

    实测 20 份：逐个调用 308 秒 → 常驻 worker 93 秒（3.3 倍）。
    """
    if not files:
        return [], []
    here = Path(__file__).resolve().parent
    r = subprocess.run([sys.executable, str(here / "pdf_batch.py"), str(outdir)],
                       capture_output=True, text=True, timeout=1800)
    made = [f.stem for f in outdir.glob("*.pdf")]
    failed = [] if r.returncode == 0 else [(x, "pdf_batch 失败") for x in
                                           (r.stderr or "").strip().splitlines()[:3]]
    return made, failed


def main() -> int:
    ap = argparse.ArgumentParser(description="docx → PDF")
    ap.add_argument("src", help="含 *.docx 的目录，或单个 docx")
    ap.add_argument("--out", default=None, help="输出目录（默认与源同目录）")
    ap.add_argument("--only", help="只转文件名包含该字串的 docx")
    a = ap.parse_args()

    src = Path(a.src)
    if src.is_file():
        files = [src]
        outdir = Path(a.out) if a.out else src.parent
    else:
        files = sorted(src.glob("*.docx"))
        outdir = Path(a.out) if a.out else src
    if not files:
        print("没有 docx 可转换", file=sys.stderr)
        return 1
    outdir.mkdir(parents=True, exist_ok=True)
    if a.only:
        files = [f for f in files if a.only in f.name]
        if not files:
            print("没有匹配 --only 的 docx", file=sys.stderr)
            return 1

    soffice = find_soffice()
    kit = find_libreoffice_kit()
    if not soffice and not kit:
        print(json.dumps({
            "skipped": True,
            "reason": "未找到 LibreOffice。PDF 是可选的：装 LibreOffice 后用 BD_SOFFICE 指向 "
                      "soffice.exe，或设置 BD_KIT_NODE / BD_KIT_CLI 使用自带 Kit。",
            "docx_count": len(files),
        }, ensure_ascii=False))
        return 0

    # 走 Kit 且有多份时，用常驻 worker 批量转（实测 3.3 倍）。
    # soffice 那边本来就支持一次转整个目录，不用改。
    if kit and len(files) > 1 and not soffice:
        made, failed = convert_all_with_kit(files, outdir)
        print(json.dumps({"made": len(made), "failed": len(failed),
                          "backend": "kit-batch"}, ensure_ascii=False))
        return 0 if not failed else 1

    made, failed = [], []
    for f in files:
        target = outdir / (f.stem + ".pdf")
        ok = False
        if soffice:
            ok = convert_with_soffice(soffice, f, outdir)
        if not ok and kit:
            ok = convert_with_kit(kit[0], kit[1], f, target)
        (made if ok else failed).append(target.name)

    print(json.dumps({"ok": len(made), "failed": failed, "out": str(outdir)},
                     ensure_ascii=False))
    return 0 if not failed else 1


if __name__ == "__main__":
    raise SystemExit(main())