#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""把一批 docx 一次性转成 PDF，只启动一次 LibreOffice。

## 为什么需要它

LibreOffice Kit 的 convert 子命令一次只吃一个文件，每调一次要启动一个
node 进程 + 一个 LibreOffice 实例。实测：

    单份 convert（含启动）  15.4 秒
    21 份逐个调用           323 秒

起一个常驻 worker 之后，每份只要 3-6 秒（省掉启动与模型加载）。

## Kit 的正确用法（踩过的坑）

1. 模块入口是 package.json 的 main（lib/index.js），不是包根的 index.js
2. createConverter() 返回的对象有 convert 方法
3. convert 的参数是 { inputPath: <绝对路径字符串>, outputPath: <绝对路径字符串> }
   —— 不是 { input: { path } }，传错会报 "inputPath and outputPath must be
   absolute filesystem paths."

用法：
    python scripts/pdf_batch.py <目录> [--only 关键词] [--skip-existing]
"""
from __future__ import annotations
import argparse, json, os, subprocess, sys, time
from pathlib import Path

WORKER_JS = "const readline = require(\"readline\");\nconst kit = require(process.argv[2]);\n\n(async () => {\n  let converter;\n  try {\n    converter = await kit.createConverter();\n  } catch (e) {\n    process.stdout.write(JSON.stringify({fatal: String(e && e.message || e)}) + \"\\n\");\n    return;\n  }\n  // 握手：告诉 Python 端\"我准备好了\"。\n  // 少了这一行，Python 会在第一份文件上死等到超时——\n  // createConverter() 要几秒，期间 stdout 什么都没有。\n  process.stdout.write(JSON.stringify({ready: true}) + \"\\n\");\n  const rl = readline.createInterface({ input: process.stdin });\n  for await (const line of rl) {\n    const t = line.trim();\n    if (!t) continue;\n    let req;\n    try { req = JSON.parse(t); }\n    catch (e) {\n      process.stdout.write(JSON.stringify({ok: false, error: \"bad json\"}) + \"\\n\");\n      continue;\n    }\n    if (req.quit) break;\n    try {\n      await converter.convert({ inputPath: req.input, outputPath: req.output });\n      process.stdout.write(JSON.stringify({ok: true}) + \"\\n\");\n    } catch (e) {\n      process.stdout.write(JSON.stringify({ok: false, error: String(e && e.message || e).slice(0, 200)}) + \"\\n\");\n    }\n  }\n  try { await converter.dispose(); } catch (e) {}\n})();\n"


def find_kit():
    node = os.environ.get("BD_KIT_NODE")
    cli = os.environ.get("BD_KIT_CLI")
    if node and cli and Path(node).is_file() and Path(cli).is_file():
        return node, cli
    return None


def resolve_entry(cli: str):
    """从 package.json 读模块入口，不猜路径。"""
    pkg_root = Path(cli).resolve().parents[1]
    pj = pkg_root / "package.json"
    if not pj.is_file():
        return None
    meta = json.loads(pj.read_text(encoding="utf-8"))
    entry = None
    if meta.get("main"):
        entry = pkg_root / meta["main"]
    if (not entry or not entry.is_file()):
        exp = (meta.get("exports") or {}).get(".")
        if isinstance(exp, dict):
            d = exp.get("default") or exp.get("require") or exp.get("import")
            if d:
                entry = pkg_root / d.lstrip("./")
    return entry if entry and entry.is_file() else None


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
        print("找不到 LibreOffice Kit。设 BD_KIT_NODE 和 BD_KIT_CLI。", file=sys.stderr)
        return 2
    node, cli = kit
    entry = resolve_entry(cli)
    if not entry:
        print("找不到 Kit 模块入口（cli=%s）" % cli, file=sys.stderr)
        return 2

    files = [f for f in sorted(src.glob("*.docx")) if not f.name.startswith("~$")]
    if a.only:
        files = [f for f in files if a.only in f.name]
    if a.skip_existing:
        files = [f for f in files if not f.with_suffix(".pdf").is_file()]
    if not files:
        print("没有需要转换的 docx", file=sys.stderr)
        return 0

    worker = src / "_pdf_worker.cjs"
    worker.write_text(WORKER_JS, encoding="utf-8")
    try:
        # stderr 必须导到文件，不能接 PIPE——
        # LibreOffice 会往 stderr 打大量日志，PIPE 没人读很快写满缓冲区，
        # 子进程阻塞在 write() 上，父进程阻塞在 readline() 上，两边死锁。
        # （实测：接 PIPE 时永恒卡死，导到文件立刻正常。）
        errlog = src / "_pdf_worker.err"
        errfh = open(errlog, "w", encoding="utf-8")
        proc = subprocess.Popen(
            [node, str(worker), str(entry)],
            stdin=subprocess.PIPE, stdout=subprocess.PIPE, stderr=errfh,
            text=True, encoding="utf-8", bufsize=1)
    except Exception as e:
        print("启动 worker 失败：%s" % e, file=sys.stderr)
        return 2

    t0 = time.time()
    made, failed = [], []
    # 读握手：createConverter() 要几秒，不等这一行会卡在第一次 readline
    hello = proc.stdout.readline()
    if not hello:
        try:
            err = errlog.read_text(encoding="utf-8", errors="replace")[-400:]
        except Exception:
            err = ""
        print("worker 启动失败：%s" % err, file=sys.stderr)
        proc.kill()
        worker.unlink(missing_ok=True)
        return 2
    try:
        h = json.loads(hello)
    except Exception:
        h = {}
    if h.get("fatal"):
        print("worker 启动失败：%s" % str(h["fatal"])[:200], file=sys.stderr)
        proc.kill()
        worker.unlink(missing_ok=True)
        return 2
    print("worker 就绪（%.1fs）" % (time.time() - t0))

    def ask(req):
        try:
            proc.stdin.write(json.dumps(req, ensure_ascii=False) + "\n")
            proc.stdin.flush()
        except Exception as e:
            return {"ok": False, "error": "写管道失败：%s" % e}
        line = proc.stdout.readline()
        if not line:
            return {"ok": False, "error": "worker 无响应"}
        try:
            return json.loads(line)
        except Exception:
            return {"ok": False, "error": "worker 返回非 JSON"}

    for i, f in enumerate(files, 1):
        out = f.with_suffix(".pdf")
        # Kit 的 convert 在目标已存在时直接抛 EEXIST，不会覆盖。
        # 所以先删旧的——否则重跑永远失败。
        try:
            if out.exists():
                out.unlink()
        except Exception:
            pass
        r = ask({"input": str(f.resolve()), "output": str(out.resolve())})
        if r.get("ok"):
            made.append(f.name)
            print("[%d/%d] ok   %s" % (i, len(files), f.name[:50]))
        elif r.get("fatal"):
            print("worker 启动失败：%s" % str(r["fatal"])[:160], file=sys.stderr)
            break
        else:
            failed.append((f.name, r.get("error") or "?"))
            print("[%d/%d] FAIL %s (%s)" % (i, len(files), f.name[:38],
                                            str(r.get("error"))[:60]), file=sys.stderr)

    # 通知退出；等不到就 kill。
    # 之前这里 proc.wait(timeout=15) 会把整个流程拖住——
    # worker 收到 quit 后还要 dispose()，而 dispose 有时不返回，
    # 结果转换全做完了、PDF 也都在，进程却挂着。
    try:
        proc.stdin.write(json.dumps({"quit": True}) + "\n")
        proc.stdin.flush()
        proc.stdin.close()
    except Exception:
        pass
    try:
        proc.wait(timeout=5)
    except Exception:
        proc.kill()
        try:
            proc.wait(timeout=5)
        except Exception:
            pass
    worker.unlink(missing_ok=True)
    try:
        errfh.close()
        errlog.unlink(missing_ok=True)
    except Exception:
        pass

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