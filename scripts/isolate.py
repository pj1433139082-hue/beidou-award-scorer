#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""隔离评分：让两次评分真正互不可见。

要解决什么
----------
多 agent 并行评同一批材料时，最容易出的错不是"评得不准"，而是**评得一样**——
因为后跑的 agent 读到了先跑的判断。一旦如此，比对就失去意义。

光靠提示词说"别看"没用。真正的隔离必须是结构性的：
  · 判断文件不在它找得到的地方
  · 找得到的地方也不含判断
  · 结束后能证明"确实没读到"

设计
----
    独立分析-<标签>/
    ├── 材料/        原始材料（唯一业务输入，只读）
    ├── 机检/        确定性机检（两 agent 必然相同，不算泄露）
    ├── skill/       代码 + 评分卡（无历史产物）
    ├── 输出/        这个 agent 自己的产物
    └── SEALED.json  隔离声明：什么被藏了、藏到哪、什么时候

用法::

    python scripts/isolate.py prepare <工作目录> --materials M --skill S --rubric R --label 千问
    python scripts/isolate.py check <工作目录>
    python scripts/isolate.py end <工作目录>
    python scripts/isolate.py audit <目录A> <目录B>

为什么"藏"而不是"删"
--------------------
隐藏比删除安全：出问题能恢复。但隐藏也必须留痕，否则下次谁也说不清哪些文件被动过。
SEALED.json 就是这本账。
"""
from __future__ import annotations
import argparse, json, os, shutil, subprocess, sys, time
from pathlib import Path

HOLD_ROOT = Path(__file__).resolve().parents[2] / ".dsh-isolation-hold"
PRODUCT_PATTERNS = ("*.judgments.json", "*.score.json", "*.checks.json")


def scan_leaks(roots):
    found = {}
    for root in roots:
        if not root.is_dir():
            continue
        for pat in PRODUCT_PATTERNS:
            try:
                for p in root.rglob(pat):
                    s = str(p)
                    if ".dsh-isolation-hold" in s or "node_modules" in s:
                        continue
                    found.setdefault(str(p.parent), []).append(p)
            except Exception:
                pass
    return found


def seal(workspace, label, extra_roots):
    roots = [workspace.parent, Path(os.environ.get("TEMP", "/tmp")),
             Path.home() / ".agents" / "skills",
             Path.home() / ".qwenworkcn" / "skills"]
    roots.extend(extra_roots)
    roots = [r for r in roots if r.resolve() != workspace.resolve()]

    # 关键：工作目录**及其所有子目录**都不算泄露源。
    # 之前只判断 d == workspace，漏掉了 workspace 的子目录
    # （比如 workspace/机检 会被误判为泄露并被藏走）。
    ws_resolved = workspace.resolve()
    groups = scan_leaks(roots)
    groups = {k: v for k, v in groups.items()
              if not (Path(k).resolve() == ws_resolved
                      or ws_resolved in Path(k).resolve().parents)}
    stamp = time.strftime("%Y%m%d-%H%M%S")
    hold = HOLD_ROOT / ("%s-%s" % (label, stamp))
    hold.mkdir(parents=True, exist_ok=True)

    moved = []
    for i, (srcdir, files) in enumerate(sorted(groups.items())):
        d = Path(srcdir)
        if workspace in d.parents or d == workspace:
            continue
        dst = hold / ("%03d_%s" % (i, d.name.replace(os.sep, "_")))
        if dst.exists():
            continue
        try:
            shutil.move(str(d), str(dst))
            moved.append({"from": srcdir, "to": str(dst), "files": len(files)})
        except Exception as e:
            print("  跳过 %s（%s）" % (srcdir, str(e)[:50]), file=sys.stderr)

    manifest = {
        "label": label,
        "sealed_at": time.strftime("%Y-%m-%d %H:%M:%S"),
        "workspace": str(workspace),
        "hold": str(hold),
        "directories_moved": len(moved),
        "details": moved[:300],
        "note": "恢复用 isolate.py end。这些目录含隔离前的判断，end 之前不要读。",
    }
    (workspace / "SEALED.json").write_text(
        json.dumps(manifest, ensure_ascii=False, indent=1), encoding="utf-8")
    return len(moved)


def prepare(workspace, materials, skill_src, rubric, label, extra_roots):
    if workspace.exists():
        print("目录已存在，先删除或换个名字：%s" % workspace, file=sys.stderr)
        return {}
    for sub in ("材料", "机检", "输出"):
        (workspace / sub).mkdir(parents=True)

    n_mat = 0
    for p in materials.iterdir():
        if p.suffix.lower() in (".docx", ".pdf") and not p.name.startswith("~$"):
            shutil.copy2(p, workspace / "材料" / p.name)
            n_mat += 1

    dst_skill = workspace / "skill"
    for sub in ("scripts", "assets", "references", "agents"):
        s = skill_src / sub
        if s.is_dir():
            shutil.copytree(s, dst_skill / sub,
                            ignore=shutil.ignore_patterns("__pycache__", "*.pyc",
                                                          "*.judgments.json", "*.score.json"))
    for f in ("SKILL.md", "README.md"):
        if (skill_src / f).is_file():
            shutil.copy2(skill_src / f, dst_skill / f)
    (dst_skill / "assets").mkdir(exist_ok=True)
    shutil.copy2(rubric, dst_skill / "assets" / "rubric.json")

    r = subprocess.run([sys.executable, str(dst_skill / "scripts" / "batch.py"),
                        str(workspace / "材料"), "--out", str(workspace / "机检"),
                        "--rubric", str(dst_skill / "assets" / "rubric.json"),
                        "--machine-only"], capture_output=True, text=True)
    # --machine-only 模式下"零评分"会返回 2（这是对的，因为一份都没打分），
    # 但机检本身可能已经成功。用实际产出判断，不看退出码。
    n_checks = len(list((workspace / "机检").glob("*.checks.json")))
    n_moved = seal(workspace, label, extra_roots)
    return {"materials": n_mat, "machine_checks": n_checks,
            "machine_ok": n_checks >= n_mat,
            "directories_sealed": n_moved, "workspace": str(workspace)}


def check(workspace):
    bad = []
    for pat in PRODUCT_PATTERNS:
        for p in workspace.rglob(pat):
            if p.parent.name == "机检" and p.name.endswith(".checks.json"):
                continue
            bad.append(str(p))
    return {"clean": not bad, "stray": bad[:20], "stray_count": len(bad),
            "sealed": (workspace / "SEALED.json").is_file()}


def unseal(workspace):
    mf = workspace / "SEALED.json"
    if not mf.is_file():
        print("没有 SEALED.json，无法恢复", file=sys.stderr)
        return 0
    m = json.loads(mf.read_text(encoding="utf-8"))
    n = 0
    for d in m.get("details") or []:
        src, dst = Path(d["to"]), Path(d["from"])
        if not src.exists() or dst.exists():
            continue
        dst.parent.mkdir(parents=True, exist_ok=True)
        try:
            shutil.move(str(src), str(dst))
            n += 1
        except Exception as e:
            print("  恢复失败 %s（%s）" % (dst, str(e)[:50]), file=sys.stderr)
    mf.rename(mf.with_suffix(".json.done"))
    return n


def audit(a, b):
    def load(d):
        """读一棵目录树下的全部评分结果，键是材料 name（ADR-0003）。"""
        from material_record import load_index
        return {name: rec.total for name, rec in load_index(d, recursive=True).items()}

    sa, sb = load(a), load(b)
    common = set(sa) & set(sb)
    if not common:
        return {"error": "没有共同材料", "a": len(sa), "b": len(sb)}
    diffs = [(k, sa[k], sb[k], round(sb[k] - sa[k], 1)) for k in common]
    identical = sum(1 for _, x, y, _ in diffs if abs(y - x) < 0.01)
    ratio = identical / len(common)
    if ratio >= 0.95:
        verdict = "可疑：%.0f%% 分数完全相同，几乎不可能是独立判断" % (ratio * 100)
    elif ratio >= 0.70:
        verdict = "部分一致（%.0f%%）：可能共享了部分输入" % (ratio * 100)
    else:
        verdict = "独立（仅 %.0f%% 相同）" % (ratio * 100)
    return {"materials": len(common), "identical": identical,
            "identical_ratio": round(ratio, 3), "verdict": verdict,
            "differences": [{"name": k[:40], "a": x, "b": y, "diff": d}
                            for k, x, y, d in sorted(diffs, key=lambda t: -abs(t[3]))[:20]]}


def main():
    ap = argparse.ArgumentParser(description="隔离评分")
    sub = ap.add_subparsers(dest="cmd", required=True)
    p1 = sub.add_parser("prepare")
    p1.add_argument("workspace")
    p1.add_argument("--materials", required=True)
    p1.add_argument("--skill", required=True)
    p1.add_argument("--rubric", required=True)
    p1.add_argument("--label", default="agent")
    p1.add_argument("--extra-root", action="append", default=[])
    p2 = sub.add_parser("check"); p2.add_argument("workspace")
    p3 = sub.add_parser("end");   p3.add_argument("workspace")
    p4 = sub.add_parser("audit"); p4.add_argument("a"); p4.add_argument("b")

    a = ap.parse_args()
    if a.cmd == "prepare":
        res = prepare(Path(a.workspace), Path(a.materials), Path(a.skill),
                      Path(a.rubric), a.label, [Path(x) for x in a.extra_root])
    elif a.cmd == "check":
        res = check(Path(a.workspace))
    elif a.cmd == "end":
        res = {"restored": unseal(Path(a.workspace))}
    else:
        res = audit(Path(a.a), Path(a.b))
    print(json.dumps(res, ensure_ascii=False, indent=1))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())