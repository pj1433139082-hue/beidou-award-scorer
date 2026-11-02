#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""端到端流程测试：把技能的每条路径都真跑一遍，逐项断言。

跑法： python scripts/selftest_pipeline.py
全部通过退出码 0，任一失败退出码 1 并列出失败项。
"""
from __future__ import annotations
import json, os, re, shutil, subprocess, sys, tempfile, zipfile
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
PY = sys.executable
TMP = Path(tempfile.mkdtemp(prefix="bd-selftest-"))
R = []          # (环节, 是否通过, 说明)

def rec(name, ok, note=""):
    R.append((name, ok, note))
    print("  %s %-42s %s" % ("[OK]  " if ok else "[FAIL]", name, note))

def run(args, **kw):
    return subprocess.run([PY] + [str(a) for a in args], capture_output=True, text=True,
                          encoding="utf-8", errors="replace", **kw)

def make_docx(path: Path, text: str) -> None:
    import docx
    d = docx.Document()
    for line in text.split("\n"):
        d.add_paragraph(line)
    d.save(path)

def main() -> int:
    print("=" * 70)
    print("北斗奖评分技能 · 端到端流程测试")
    print("=" * 70)
    print("临时目录:", TMP)

    # ---------- 1. 依赖体检 ----------
    print("\n[1] 依赖体检")
    r = run([ROOT / "scripts" / "setup.py", "--no-self-test"])
    rec("setup.py 体检可执行", r.returncode == 0, (r.stdout.strip().splitlines() or [""])[-1][:50])

    # ---------- 2. 抽取 ----------
    print("\n[2] 抽取 docx → doc.json")
    src = TMP / "src"; src.mkdir()
    body = """示例奖A-测试材料

一、项目背景
某污水厂曝气系统能耗偏高，改造前电单耗 0.1069 kWh/吨，改造后降至 0.0784 kWh/吨。
累计已实现节约电费 202.745 万元。

二、技术创新
相比上一届参赛版本，本次新增一表多点采样，控制算法进行了优化升级。

三、推广情况
已在东部大区 3 个区域 13 家水厂推广应用。

四、经济效益
计算步骤：吨水电耗差 × 处理水量 × 电价。
年化收益 82.5 万元。
"""
    docx_path = src / "示例奖A-测试材料.docx"
    make_docx(docx_path, body)
    out = TMP / "out"
    r = run([ROOT / "scripts" / "extract_report.py", docx_path, "--out", out])
    docjson = out / "示例奖A-测试材料.doc.json"
    ok = r.returncode == 0 and docjson.is_file()
    chars = 0
    if ok:
        d = json.loads(docjson.read_text(encoding="utf-8"))
        chars = d["stats"]["chars"]
        ok = chars > 100 and "0.1069" in d["plain"]
    rec("抽取产出 doc.json 且内容完整", ok, "%d 字" % chars)

    # ---------- 3. 机检 ----------
    print("\n[3] 机检 20 项")
    chkjson = out / "示例奖A-测试材料.checks.json"
    r = run([ROOT / "scripts" / "checks.py", docjson, "--award", "示例奖A",
             "--rubric", ROOT / "assets" / "rubric.example.json", "--out", chkjson])
    ok = r.returncode == 0 and chkjson.is_file()
    summary = {}
    if ok:
        c = json.loads(chkjson.read_text(encoding="utf-8"))
        summary = c["summary"]
        n = len(c["checks"])
        ok = n >= 20
        rec("机检产出 checks.json（20 项）", ok, "共 %d 项 %s" % (n, summary))
    else:
        rec("机检产出 checks.json（20 项）", False, r.stderr[-120:])

    # ---------- 4. 打分 ----------
    print("\n[4] 打分（judgments → score.json）")
    judg = {"award": "示例奖A", "judgments": {
        "value": {"score": 30, "grade": "A", "rationale": "有实测电单耗与金额",
                  "quotes": ["改造前电单耗 0.1069 kWh/吨", "累计已实现节约电费 202.745 万元"]},
        "replication": {"score": 28, "grade": "B", "rationale": "有推广数量",
                        "quotes": ["已在东部大区 3 个区域 13 家水厂推广应用"]},
        "novelty": {"score": 24, "grade": "B", "rationale": "有迭代说明",
                    "quotes": ["本次新增一表多点采样"]},
    }, "disputes": ["ROI 缺投入侧测算"]}
    jp = out / "示例奖A-测试材料.judgments.json"
    jp.write_text(json.dumps(judg, ensure_ascii=False, indent=1), encoding="utf-8")
    r = run([ROOT / "scripts" / "score.py", docjson, "--checks", chkjson, "--judgments", jp,
             "--award", "示例奖A", "--rubric", ROOT / "assets" / "rubric.example.json", "--out", out])
    scorejson = out / "示例奖A-测试材料.score.json"
    ok = r.returncode == 0 and scorejson.is_file()
    sc = {}
    if ok:
        sc = json.loads(scorejson.read_text(encoding="utf-8"))
        ok = 0 < sc["ai_score"] <= 100 and len(sc["dimensions"]) == 3
        rec("打分产出 score.json", ok, "%s 分 / %s" % (sc.get("ai_score"), sc.get("band")))
        # 引文核验必须生效
        verified = sum(len(d.get("quotes_verified") or []) for d in sc["dimensions"])
        rejected = sum(len(d.get("quotes_rejected") or []) for d in sc["dimensions"])
        rec("引文逐字核验生效", verified >= 4 and rejected == 0,
            "通过 %d 条 / 驳回 %d 条" % (verified, rejected))
        rec("report.md 已生成", (out / "示例奖A-测试材料.report.md").is_file())
    else:
        rec("打分产出 score.json", False, r.stderr[-160:])

    # ---------- 5. 假引文必须被拒 ----------
    print("\n[5] 假引文必须被拒（没引文就没有分）")
    bad = {"award": "示例奖A", "judgments": {
        "value": {"score": 35, "grade": "A", "rationale": "编的",
                  "quotes": ["这句话材料里根本不存在"]},
        "replication": {"score": 35, "grade": "A", "rationale": "编的", "quotes": ["也不存在"]},
        "novelty": {"score": 30, "grade": "A", "rationale": "编的", "quotes": ["还是不存在"]},
    }, "disputes": []}
    bp = out / "示例奖A-测试材料.bad.judgments.json"
    bp.write_text(json.dumps(bad, ensure_ascii=False), encoding="utf-8")
    out2 = TMP / "out2"; out2.mkdir()
    r = run([ROOT / "scripts" / "score.py", docjson, "--checks", chkjson, "--judgments", bp,
             "--award", "示例奖A", "--rubric", ROOT / "assets" / "rubric.example.json", "--out", out2])
    ok = r.returncode == 0
    if ok:
        b = json.loads((out2 / "示例奖A-测试材料.score.json").read_text(encoding="utf-8"))
        allrej = all((d.get("quotes_verified") or []) == [] for d in b["dimensions"])
        capped = all(d.get("verified_grade") in ("C", "D") for d in b["dimensions"])
        rec("假引文被全部驳回", allrej and capped, "分数降到 %s" % b.get("ai_score"))
    else:
        rec("假引文被全部驳回", False, r.stderr[-160:])

    # ---------- 5b. 未知奖项要给可读错误 ----------
    print("\n[5b] 未知奖项给可读错误（不是 traceback）")
    r = run([ROOT / "scripts" / "score.py", docjson, "--checks", chkjson, "--judgments", jp,
             "--award", "不存在的奖项", "--rubric", ROOT / "assets" / "rubric.example.json",
             "--out", out2])
    rec("未知奖项报错可读", r.returncode == 2 and "评分卡里没有" in r.stderr
        and "Traceback" not in r.stderr, r.stderr.strip().splitlines()[0][:40] if r.stderr else "")

    # ---------- 6. Excel 导出 ----------
    print("\n[6] Excel 导出")
    xl = TMP / "汇总.xlsx"
    r = run([ROOT / "scripts" / "export_xlsx.py", out, "--out", xl])
    ok = r.returncode == 0 and xl.is_file()
    if ok:
        from openpyxl import load_workbook
        wb = load_workbook(xl)
        sheets = wb.sheetnames
        ws = wb["汇总"]
        # 关键列必须存在
        hdr = {str(ws.cell(4, c).value).strip() for c in range(1, ws.max_column + 1)}
        need = {"申报材料", "总分", "奖项区间", "需关注", "修改建议"}
        rec("Excel 三个工作表齐全", sheets == ["汇总", "维度明细", "争议点"], str(sheets))
        rec("汇总表关键列存在", need <= hdr, "缺 %s" % (need - hdr) if not need <= hdr else "")
        # 建议列必须有内容（不能是空列）
        adv_col = next(c for c in range(1, ws.max_column + 1) if str(ws.cell(4, c).value).strip() == "修改建议")
        filled = sum(1 for rw in range(5, ws.max_row + 1)
                     if ws.cell(rw, adv_col).value and str(ws.cell(rw, adv_col).value).strip() not in ("", "—"))
        rec("修改建议列有实际内容", filled >= 1, "%d 行有建议" % filled)
    else:
        rec("Excel 导出", False, r.stderr[-160:])

    # ---------- 7. batch 全链路 ----------
    print("\n[7] batch 全链路（docx → Excel）")
    bdir = TMP / "batch"; bdir.mkdir()
    shutil.copy(docx_path, bdir / docx_path.name)
    # batch 需要 judgments 才打分；先放好
    shutil.copy(jp, bdir / "示例奖A-测试材料.judgments.json")
    r = run([ROOT / "scripts" / "batch.py", bdir, "--out", bdir,
             "--rubric", ROOT / "assets" / "rubric.example.json"])
    ok = r.returncode == 0
    has_lb = (bdir / "leaderboard.md").is_file()
    has_xl = (bdir / "北斗奖评分汇总.xlsx").is_file()
    scored = "1 份有评分" in r.stdout
    rec("batch 产出 leaderboard.md", has_lb)
    rec("batch 产出 Excel", has_xl)
    rec("batch 完成打分（非仅机检）", scored, (r.stdout.strip().splitlines() or [""])[-1][:60])

    # ---------- 8. 打包 ----------
    print("\n[8] 打包 + 包内自检")
    dist = TMP / "dist"
    r = run([ROOT / "scripts" / "make_package.py", "--out", dist, "--version", "0.0.0-test"])
    zips = list(dist.glob("*.zip"))
    ok = r.returncode == 0 and len(zips) == 1
    if ok:
        z = zips[0]
        with zipfile.ZipFile(z) as zf:
            names = zf.namelist()
        has_ver = any(n.endswith("/VERSION") for n in names)
        has_ck = any(n.endswith("/CHECKSUMS.sha256") for n in names)
        has_inst = any(n.endswith("/INSTALL.txt") for n in names)
        rec("打包产出 zip", ok, "%.0f KB" % (z.stat().st_size / 1024))
        rec("包内含 VERSION/CHECKSUMS/INSTALL", has_ver and has_ck and has_inst)
        rec("包内自检通过", "包内自检通过" in r.stdout)
    else:
        rec("打包", False, r.stderr[-160:])

    # ---------- 9. 漂移检测 ----------
    print("\n[9] 漂移检测（行尾容错 + 真实改动必报）")
    mirror = TMP / "mirror"
    shutil.copytree(ROOT / "scripts", mirror / "scripts")
    shutil.copytree(ROOT / "references", mirror / "references")
    shutil.copytree(ROOT / "assets", mirror / "assets")
    r = run([ROOT / "scripts" / "sync_check.py", mirror])
    rec("完全一致时通过", r.returncode == 0)
    # 转 CRLF 不应报分叉
    p = mirror / "scripts" / "checks.py"
    data = p.read_bytes().replace(b"\r\n", b"\n").replace(b"\n", b"\r\n")
    p.write_bytes(data)
    r = run([ROOT / "scripts" / "sync_check.py", mirror])
    rec("CRLF 差异不误报", r.returncode == 0)
    # 真实改动必须报
    p.write_bytes(p.read_bytes() + b"\n# drift\n")
    r = run([ROOT / "scripts" / "sync_check.py", mirror])
    rec("真实改动被检出", r.returncode == 1)

    # ---------- 10. 迭代性判定 ----------
    print("\n[10] 迭代性判定（锚点 + 变化）")
    sys.path.insert(0, str(ROOT / "scripts"))
    import checks as C
    rub = json.loads((ROOT / "assets" / "rubric.example.json").read_text(encoding="utf-8"))
    cases = [
        ("相比第七届参赛版本，本届新增一表多点采样。", "pass"),
        ("相较于过去获奖版本，本次在部署架构上做了重构。", "pass"),
        ("本项目进行了迭代。", "warn"),
        ("体现了平台对人员工作安排优化的支持。", "fail"),
        ("推广范围由项目应用拓展至跨大区应用。", "fail"),
    ]
    bad_case = []
    for txt, exp in cases:
        d = {"name": "t", "source": "t", "plain": txt, "stats": {}, "blocks": []}
        got = C.c13_iteration(d, "示例奖C（推广类）", rub)["status"]
        if got != exp:
            bad_case.append((txt, got, exp))
    rec("5 个迭代性用例全判对", not bad_case, str(bad_case) if bad_case else "")

    # ---------- 11. 异常输入必须给可读错误 ----------
    print("\n[11] 异常输入给可读错误（不是 traceback）")
    edge = TMP / "edge"; edge.mkdir()
    # 不存在的文件
    r = run([ROOT / "scripts" / "checks.py", edge / "nofile.doc.json",
             "--award", "示例奖A", "--rubric", ROOT / "assets" / "rubric.example.json"])
    rec("不存在的 doc.json 报错可读", r.returncode == 2 and "找不到" in r.stderr
        and "Traceback" not in r.stderr)
    # 损坏 JSON
    badjson = edge / "bad.doc.json"; badjson.write_text("not json", encoding="utf-8")
    r = run([ROOT / "scripts" / "checks.py", badjson,
             "--award", "示例奖A", "--rubric", ROOT / "assets" / "rubric.example.json"])
    rec("损坏 JSON 报错可读", r.returncode == 2 and "不是合法 JSON" in r.stderr
        and "Traceback" not in r.stderr)
    # 不存在的 rubric
    r = run([ROOT / "scripts" / "checks.py", docjson, "--award", "示例奖A",
             "--rubric", edge / "norubric.json"])
    rec("不存在的评分卡报错可读", r.returncode == 2 and "找不到" in r.stderr
        and "Traceback" not in r.stderr)
    # 不存在的 judgments
    r = run([ROOT / "scripts" / "score.py", docjson, "--checks", chkjson,
             "--judgments", edge / "nojudg.json", "--award", "示例奖A",
             "--rubric", ROOT / "assets" / "rubric.example.json", "--out", edge])
    rec("不存在的 judgments 报错可读", r.returncode == 2 and "找不到" in r.stderr
        and "Traceback" not in r.stderr)
    # 空目录 batch
    r = run([ROOT / "scripts" / "batch.py", edge, "--out", edge])
    rec("空目录 batch 报错可读", r.returncode == 2 and "没有匹配到材料文件" in r.stderr)
    # 空目录 export
    r = run([ROOT / "scripts" / "export_xlsx.py", edge, "--out", edge / "x.xlsx"])
    rec("无评分数据时 Excel 导出报错可读", r.returncode == 1 and "没找到 score.json" in r.stderr)

    # ---------- 汇总 ----------
    shutil.rmtree(TMP, ignore_errors=True)
    print("\n" + "=" * 70)
    passed = sum(1 for _, ok, _ in R if ok)
    failed = [n for n, ok, _ in R if not ok]
    print("结果：%d/%d 通过" % (passed, len(R)))
    if failed:
        print("失败项：")
        for f in failed:
            print("  - " + f)
        return 1
    print("全部通过。")
    return 0

if __name__ == "__main__":
    raise SystemExit(main())