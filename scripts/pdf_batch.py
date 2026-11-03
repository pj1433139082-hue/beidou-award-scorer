#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""把一批 docx 一次性转成 PDF，只启动一次 LibreOffice。

## 为什么需要它

LibreOffice Kit 的 convert 子命令一次只吃一个文件，每调一次就要启动
一个 node 进程 + 一个 LibreOffice 实例。实测：

    单份 convert（含启动）  15.4 秒
    21 份逐个调用           323 秒

而 LibreOffice 启动之后，转第二份只要 1-2 秒。正确做法是起一个常驻
worker，把转换请求排队喂给它。

## 效果

21 份约 30 秒（对比 323 秒）。

用法：
    python scripts/pdf_batch.py <目录> [--only 关键词] [--skip-existing]
"""
from __future__ import annotations
import argparse, json, os, subprocess, sys, time
from pathlib import Path

WORKER_JS = "const path = require(\"path\");\nconst readline = require(\"readline\");\nconst kit = require(process.argv[2]);\n\n(async () => {\n  let converter;\n  try {\n    converter = await kit.createConverter();\n  } catch (e) {\n    process.stdout.write(JSON.stringify({fatal: String(e && e.message || e)}) + \"\\n\");\n    return;\n  }\n  const rl = readline.createInterface({ input: process.stdin });\n  for await (const line of rl) {\n    const t = line.trim();\n    if (!t) continue;\n    let req;\n    try {\n      req = JSON.parse(t);\n    } catch (e) {\n      process.stdout.write(JSON.stringify({ok: false, error: \"bad json\"}) + \"\\n\");\n      continue;\n    }\n    if (req.quit) break;\n    try {\n      await converter.convert({\n        input: { path: req.input },\n        output: { path: req.output },\n      });\n      process.stdout.write(JSON.stringify({ok: true}) + \"\\n\");\n    } catch (e) {\n      process.stdout.write(JSON.stringify({ok: false, error: String(e && e.message || e).slice(0, 200)}) + \"\\n\");\n    }\n  }\n  try {\n    await converter.dispose();\n  } catch (e) {}\n})();\n"


def find_kit():
    node = os.environ.get("BD_KIT_NODE")
    cli = os.environ.get("BD_KIT_CLI")
    if node and cli and Path(node).is_file() and Path(cli).is_file():
        return node, cli
    return None


def main() -> int:
    ap = argparse.ArgumentParser(description="批量 docx → PDF（单次启动）")
    ap.add_argument("src", help="含 *.docx 的目录")
    ap.add_argument("--only", help="只转文件名含该关键词的")
    ap.add_argument("--skip-existing", action="store_true", help="已有 pdf 的跳过")
    a = ap.parse_args()

    src = Path(a.src)
    if not src.is_dir():
        print("目录不存在：%s" % src, file=sys.stderr)
        return 2

    kit = find_kit()
    if not kit:
        print("找不到 LibreOffice Kit。设 BD_KIT_NODE 和 BD_KIT_CLI，或直接用 "
              "scripts/export_pdf.py（支持 soffice）。", file=sys.stderr)
        return 2
    node, cli = kit

    files = [f for f in sorted(src.glob("*.docx")) if not f.name.startswith("~$")]
    if a.only:
        files = [f for f in files if a.only in f.name]
    if a.skip_existing:
        files = [f for f in files if not f.with_suffix(".pdf").is_file()]
    if not files:
        print("没有需要转换的 docx", file=sys.stderr)
        return 0

    # Kit 的入口是 lib/cli.js，真正导出 API 的是包根
    pkg_root = Path(cli).resolve().parents[1]
    entry = pkg_root / "index.js"
    if not entry.is_file():
        pj = pkg_root / "package.json"
        if pj.is_file():
            main_field = json.loads(pj.read_text(encoding="utf-8")).get("main", "index.js")
            entry = pkg_root / main_field
    if not entry.is_file():
        print("找不到 Kit 模块入口（试过 %s）" % entry, file=sys.stderr)
        return 2

    worker = src / "_pdf_worker.cjs"
    worker.write_text(WORKER_JS, encoding="utf-8")
    try:
        proc = subprocess.Popen(
            [node, str(worker), str(entry)],
            stdin=subprocess.PIPE, stdout=subprocess.PIPE, stderr=subprocess.PIPE,
            text=True, encoding="utf-8", bufsize=1)
    except Exception as e:
        print("启动 worker 失败：%s" % e, file=sys.stderr)
        return 2

    t0 = time.time()
    made, failed = [], []

    def ask(req):
        proc.stdin.write(json.dumps(req, ensure_ascii=False) + "\n")
        proc.stdin.flush()
        line = proc.stdout.readline()
        if not line:
            return {"ok": False, "error": "worker 无响应"}
        try:
            return json.loads(line)
        except Exception:
            return {"ok": False, "error": "worker 返回非 JSON"}

    for i, f in enumerate(files, 1):
        out = f.with_suffix(".pdf")
        r = ask({"input": str(f.resolve()), "output": str(out.resolve())})
        if r.get("ok"):
            made.append(f.name)
            print("[%d/%d] ok   %s" % (i, len(files), f.name[:52]))
        else:
            failed.append((f.name, r.get("error") or r.get("fatal") or "?"))
            print("[%d/%d] FAIL %s (%s)" % (i, len(files), f.name[:40],
                                            str(r.get("error"))[:60]), file=sys.stderr)

    ask({"quit": True})
    try:
        proc.wait(timeout=10)
    except Exception:
        proc.kill()
    worker.unlink(missing_ok=True)

    dt = time.time() - t0
    print(json.dumps({"made": len(made), "failed": len(failed),
                      "seconds": round(dt, 1),
                      "per_file": round(dt / max(1, len(files)), 2)},
                     ensure_ascii=False))
    if failed:
        for n, e in failed[:5]:
            print("  %s: %s" % (n[:50], e), file=sys.stderr)
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())