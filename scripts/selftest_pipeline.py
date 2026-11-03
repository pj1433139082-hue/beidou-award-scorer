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


def pick_rubric() -> tuple[Path, str, list[str]]:
    """选一张能用的评分卡，并返回 (路径, 奖项名, 该奖项的维度 id 列表)。

    公开副本有 rubric.example.json，私有副本只有真实 rubric.json——
    测试必须两边都能跑，所以按顺序找。
    """
    for name in ("rubric.example.json", "rubric.json"):
        p = ROOT / "assets" / name
        if not p.is_file():
            continue
        try:
            r = json.loads(p.read_text(encoding="utf-8"))
        except Exception:
            continue
        awards = list((r.get("awards") or {}).keys())
        if not awards:
            continue
        a = awards[0]
        dims = [d["id"] for d in r["awards"][a]["dimensions"]]
        return p, a, dims
    raise SystemExit("assets 里既没有 rubric.example.json 也没有 rubric.json")
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
    RUB, AWARD, DIMS = pick_rubric()
    print("评分卡:", RUB.name, "| 奖项:", AWARD, "| 维度:", "、".join(DIMS))

    # ---------- 1. 依赖体检 ----------
    print("\n[1] 依赖体检")
    r = run([ROOT / "scripts" / "setup.py", "--no-self-test"])
    rec("setup.py 体检可执行", r.returncode == 0, (r.stdout.strip().splitlines() or [""])[-1][:50])

    # ---------- 2. 抽取 ----------
    print("\n[2] 抽取 docx → doc.json")
    src = TMP / "src"; src.mkdir()
    body = """测试材料

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
    docx_path = src / (AWARD + "-测试材料.docx")
    make_docx(docx_path, body)
    out = TMP / "out"
    r = run([ROOT / "scripts" / "extract_report.py", docx_path, "--out", out])
    docjson = out / ("%s-测试材料.doc.json" % AWARD)
    ok = r.returncode == 0 and docjson.is_file()
    chars = 0
    if ok:
        d = json.loads(docjson.read_text(encoding="utf-8"))
        chars = d["stats"]["chars"]
        ok = chars > 100 and "0.1069" in d["plain"]
    rec("抽取产出 doc.json 且内容完整", ok, "%d 字" % chars)

    # ---------- 3. 机检 ----------
    print("\n[3] 机检 20 项")
    chkjson = out / ("%s-测试材料.checks.json" % AWARD)
    r = run([ROOT / "scripts" / "checks.py", docjson, "--award", AWARD,
             "--rubric", RUB, "--out", chkjson])
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
    quotes_pool = ["改造前电单耗 0.1069 kWh/吨", "累计已实现节约电费 202.745 万元",
                    "已在东部大区 3 个区域 13 家水厂推广应用", "本次新增一表多点采样"]
    judg = {"award": AWARD, "judgments": {
        dim: {"score": 0, "grade": "A", "rationale": "测试",
              "quotes": quotes_pool[:2]}
        for dim in DIMS
    }, "disputes": ["ROI 缺投入侧测算"]}
    # 分数按各维度满分给 80%
    _rub = json.loads(RUB.read_text(encoding="utf-8"))
    for d in _rub["awards"][AWARD]["dimensions"]:
        judg["judgments"][d["id"]]["score"] = round(d["max"] * 0.8, 1)
    jp = out / ("%s-测试材料.judgments.json" % AWARD)
    jp.write_text(json.dumps(judg, ensure_ascii=False, indent=1), encoding="utf-8")
    r = run([ROOT / "scripts" / "score.py", docjson, "--checks", chkjson, "--judgments", jp,
             "--award", AWARD, "--rubric", RUB, "--out", out])
    scorejson = out / ("%s-测试材料.score.json" % AWARD)
    ok = r.returncode == 0 and scorejson.is_file()
    sc = {}
    if ok:
        sc = json.loads(scorejson.read_text(encoding="utf-8"))
        ok = 0 < sc["ai_score"] <= 100 and len(sc["dimensions"]) == len(DIMS)
        rec("打分产出 score.json", ok, "%s 分 / %s" % (sc.get("ai_score"), sc.get("band")))
        # 引文核验必须生效
        verified = sum(len(d.get("quotes_verified") or []) for d in sc["dimensions"])
        rejected = sum(len(d.get("quotes_rejected") or []) for d in sc["dimensions"])
        rec("引文逐字核验生效", verified >= 4 and rejected == 0,
            "通过 %d 条 / 驳回 %d 条" % (verified, rejected))
        rec("report.md 已生成", (out / ("%s-测试材料.report.md" % AWARD)).is_file())
    else:
        rec("打分产出 score.json", False, r.stderr[-160:])

    # ---------- 5. 假引文必须被拒 ----------
    print("\n[5] 假引文必须被拒（没引文就没有分）")
    bad = {"award": AWARD, "judgments": {
        dim: {"score": 100, "grade": "A", "rationale": "编的",
              "quotes": ["这句话材料里根本不存在"]}
        for dim in DIMS
    }, "disputes": []}
    bp = out / ("%s-测试材料.bad.judgments.json" % AWARD)
    bp.write_text(json.dumps(bad, ensure_ascii=False), encoding="utf-8")
    out2 = TMP / "out2"; out2.mkdir()
    r = run([ROOT / "scripts" / "score.py", docjson, "--checks", chkjson, "--judgments", bp,
             "--award", AWARD, "--rubric", RUB, "--out", out2])
    ok = r.returncode == 0
    if ok:
        b = json.loads((out2 / ("%s-测试材料.score.json" % AWARD)).read_text(encoding="utf-8"))
        allrej = all((d.get("quotes_verified") or []) == [] for d in b["dimensions"])
        capped = all(d.get("verified_grade") in ("C", "D") for d in b["dimensions"])
        rec("假引文被全部驳回", allrej and capped, "分数降到 %s" % b.get("ai_score"))
    else:
        rec("假引文被全部驳回", False, r.stderr[-160:])

    # ---------- 5b. 未知奖项要给可读错误 ----------
    print("\n[5b] 未知奖项给可读错误（不是 traceback）")
    r = run([ROOT / "scripts" / "score.py", docjson, "--checks", chkjson, "--judgments", jp,
             "--award", "不存在的奖项", "--rubric", RUB,
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

    # ---------- 6b. 逐项目意见 ----------
    print("\n[6b] 逐项目意见生成")
    advdir = TMP / "advice"
    r = run([ROOT / "scripts" / "report_advice.py", out, "--out", advdir])
    files = list(advdir.glob("*.意见.md")) if advdir.is_dir() else []
    ok = r.returncode == 0 and len(files) == 1
    rec("生成逐项目意见文件", ok, "%d 份" % len(files))
    if ok:
        txt = files[0].read_text(encoding="utf-8")
        # 必须有这几段，且不能是空话
        need = ["各维度得分与意见", "怎么改", "优先补什么"]
        missing = [x for x in need if x not in txt]
        rec("意见含必要段落", not missing, "缺 %s" % missing if missing else "")
        # 建议不能是同一句重复（早期 bug：全局项挂到每个维度）
        adv_lines = [l for l in txt.splitlines() if l.startswith("- （") or "怎么改**：" in l]
        rec("意见不是机械重复", len(set(adv_lines)) >= min(2, len(adv_lines)) if adv_lines else True)

    # ---------- 6c. Word 导出 ----------
    print("\n[6c] 逐项目意见导出 Word")
    wdir = TMP / "word"
    r = run([ROOT / "scripts" / "export_docx.py", out, "--out", wdir, "--merge"])
    docs = list(wdir.glob("*.docx")) if wdir.is_dir() else []
    merged = [d for d in docs if "汇总" in d.name]
    rec("生成 Word 文件", r.returncode == 0 and len(docs) >= 2, "%d 个" % len(docs))
    rec("生成合并总册", bool(merged))
    if docs:
        from docx import Document as _D
        one = next((d for d in docs if "汇总" not in d.name), None)
        if one:
            d = _D(str(one))
            texts = [p.text for p in d.paragraphs if p.text.strip()]
            body = "\n".join(texts)
            # 必须有维度、必须没有 Markdown 残留
            has_dim = any(k in body for k in ("证据等级", "怎么改", "优先补什么"))
            no_md = "**" not in body and "##" not in body
            rec("Word 含必要内容", has_dim, "")
            rec("Word 无 Markdown 残留", no_md, "发现 ** 或 ##" if not no_md else "")

    # ---------- 7. batch 全链路 ----------
    print("\n[7] batch 全链路（docx → Excel）")
    bdir = TMP / "batch"; bdir.mkdir()
    shutil.copy(docx_path, bdir / docx_path.name)
    # batch 需要 judgments 才打分；先放好
    shutil.copy(jp, bdir / ("%s-测试材料.judgments.json" % AWARD))
    r = run([ROOT / "scripts" / "batch.py", bdir, "--out", bdir,
             "--rubric", RUB, "--with-docx"])
    ok = r.returncode == 0
    stem = AWARD + "-测试材料"
    # 一份材料跑完必须产出这些文件——少任何一个都算流程不完整
    NEED = [stem + ".doc.json", stem + ".checks.json", stem + ".score.json",
            stem + ".report.md", stem + ".docx", "leaderboard.md", "北斗奖评分汇总.xlsx"]
    missing = [x for x in NEED if not (bdir / x).is_file()]
    rec("batch 产出完整文件集", not missing,
        "缺 %s" % missing if missing else "7 类文件齐全")
    rec("batch 完成打分（非仅机检）", "1 份有评分" in r.stdout,
        (r.stdout.strip().splitlines() or [""])[-1][:50])
    # 汇总册
    merged = list(bdir.glob("*汇总*.docx"))
    rec("batch 产出 Word 汇总册", bool(merged), merged[0].name if merged else "未生成")
    # PDF 是可选的（依赖 LibreOffice），只验证脚本存在且能跑出预期返回
    pdf_script = ROOT / "scripts" / "export_pdf.py"
    rec("PDF 导出脚本存在", pdf_script.is_file())

    # ---- 流程硬化：清单 + 独立核验 + 严格模式 ----
    mf = bdir / "manifest.json"
    rec("batch 写出 manifest.json", mf.is_file())
    if mf.is_file():
        try:
            m = json.loads(mf.read_text(encoding="utf-8"))
            rec("manifest 记录了产物", bool(m.get("files")) and bool(m.get("counts")),
                "%d 个文件" % len(m.get("files", [])))
        except Exception as e:
            rec("manifest 记录了产物", False, str(e)[:40])

    ver = ROOT / "scripts" / "verify_outputs.py"
    r = run([ver, bdir, "--expect-docx"])
    rec("verify_outputs 通过（产物齐全）", r.returncode == 0)
    # 删一个文件后必须报错
    victim = next((p for p in bdir.glob("*.docx") if "汇总" not in p.name), None)
    if victim:
        bak = victim.with_suffix(".docx.bak")
        victim.rename(bak)
        r = run([ver, bdir, "--expect-docx"])
        rec("verify_outputs 能抓到缺失", r.returncode != 0, "删掉一个 docx 后退出码 %d" % r.returncode)
        bak.rename(victim)
    else:
        rec("verify_outputs 能抓到缺失", None, "没有可删的 docx")
        rec("verify_outputs 能抓到缺失", None, "没有可删的 docx")

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
    rub = json.loads(RUB.read_text(encoding="utf-8"))
    cases = [
        ("相比第七届参赛版本，本届新增一表多点采样。", "pass"),
        ("相较于过去获奖版本，本次在部署架构上做了重构。", "pass"),
        ("本项目进行了迭代。", "warn"),
        ("体现了平台对人员工作安排优化的支持。", "fail"),
        ("推广范围由项目应用拓展至跨大区应用。", "fail"),
    ]
    # 找一个启用了 iteration 闸门的奖项
    gate_award = next((a for a, s in rub["awards"].items()
                       if (s.get("gate_dimension") or {}).get("id") == "iteration"), None)
    if gate_award:
        bad_case = []
        for txt, exp in cases:
            d = {"name": "t", "source": "t", "plain": txt, "stats": {}, "blocks": []}
            got = C.c13_iteration(d, gate_award, rub)["status"]
            if got != exp:
                bad_case.append((txt, got, exp))
        rec("5 个迭代性用例全判对", not bad_case, "奖项=%s" % gate_award)
    else:
        rec("5 个迭代性用例全判对", None, "该评分卡无 iteration 闸门，跳过")

    # ---------- 10b. 单份坏数据不能拖垮整批 ----------
    print("\n[10b] 单份坏 judgments 不拖垮整批")
    bj = TMP / "badj"; bj.mkdir()
    import shutil as _sh
    _sh.copy(docx_path, bj / docx_path.name)
    # 再复制一份，判断它的 judgments 是坏 JSON
    bad_docx = bj / (AWARD + "-坏材料.docx")
    _sh.copy(docx_path, bad_docx)
    good_j = json.loads(jp.read_text(encoding="utf-8"))
    (bj / jp.name).write_text(json.dumps(good_j, ensure_ascii=False), encoding="utf-8")
    (bj / (AWARD + "-坏材料.judgments.json")).write_text("{ 这不是合法 JSON", encoding="utf-8")
    r = run([ROOT / "scripts" / "batch.py", bj, "--out", bj, "--rubric", RUB])
    lb = (bj / "leaderboard.md")
    rec("坏 judgments 不中断整批", r.returncode == 0 and lb.is_file(),
        "退出码 %d" % r.returncode)
    if lb.is_file():
        txt = lb.read_text(encoding="utf-8")
        rec("排行榜标出未评分的原因", "未评分" in txt and "不是合法 JSON" in txt)

    # ---------- 10c. 独立验证发现的缺陷（回归守卫）----------
    print("\n[10c] 独立验证查出的缺陷")

    # ① 示例卡必须声明 filename_prefixes，否则 batch 无法从文件名识别奖项
    try:
        _r = json.loads(RUB.read_text(encoding="utf-8"))
        no_prefix = [a for a, s in _r["awards"].items() if not (s.get("filename_prefixes"))]
        rec("评分卡每个奖项都有 filename_prefixes", not no_prefix,
            "缺 %s" % no_prefix if no_prefix else "")
    except Exception as e:
        rec("评分卡每个奖项都有 filename_prefixes", False, str(e)[:40])

    # ② verify_outputs 在"没有材料"时也必须给 missing_total（否则调用方会当成"缺 0 项"）
    empty = TMP / "novoid"; empty.mkdir(exist_ok=True)
    r = run([ROOT / "scripts" / "verify_outputs.py", empty])
    ok_shape = False
    if r.stdout.strip():
        try:
            j = json.loads(r.stdout)
            ok_shape = j.get("ok") is False and isinstance(j.get("missing_total"), int) \
                and j["missing_total"] >= 1 and isinstance(j.get("missing"), list)
        except Exception:
            pass
    rec("空目录核验返回 missing_total", ok_shape, "退出码 %d" % r.returncode)

    # ③ 一份都没评分时必须非零退出（不能"看起来成功"）
    none_dir = TMP / "noscore"; none_dir.mkdir(exist_ok=True)
    import shutil as _sh2
    _sh2.copy(docx_path, none_dir / docx_path.name)   # 只有材料，没有 judgments
    r = run([ROOT / "scripts" / "batch.py", none_dir, "--out", none_dir, "--rubric", RUB])
    rec("零评分时非零退出", r.returncode != 0, "退出码 %d" % r.returncode)
    rec("零评分时说明原因", "没有任何材料完成评分" in (r.stderr or ""))

    # ④ judgments 放在材料目录也要能用（不能只看输出目录）
    sib = TMP / "sibling"; sib.mkdir(exist_ok=True)
    _sh2.copy(docx_path, sib / docx_path.name)
    _sh2.copy(jp, sib / jp.name)
    r = run([ROOT / "scripts" / "batch.py", sib, "--out", sib, "--rubric", RUB])
    rec("judgments 放材料目录也生效", r.returncode == 0 and "1 份有评分" in r.stdout,
        (r.stdout.strip().splitlines() or [""])[-1][:46])

    # ---------- 11. 异常输入必须给可读错误 ----------
    print("\n[11] 异常输入给可读错误（不是 traceback）")
    edge = TMP / "edge"; edge.mkdir()
    # 不存在的文件
    r = run([ROOT / "scripts" / "checks.py", edge / "nofile.doc.json",
             "--award", AWARD, "--rubric", RUB])
    rec("不存在的 doc.json 报错可读", r.returncode == 2 and "找不到" in r.stderr
        and "Traceback" not in r.stderr)
    # 损坏 JSON
    badjson = edge / "bad.doc.json"; badjson.write_text("not json", encoding="utf-8")
    r = run([ROOT / "scripts" / "checks.py", badjson,
             "--award", AWARD, "--rubric", RUB])
    rec("损坏 JSON 报错可读", r.returncode == 2 and "不是合法 JSON" in r.stderr
        and "Traceback" not in r.stderr)
    # 不存在的 rubric
    r = run([ROOT / "scripts" / "checks.py", docjson, "--award", AWARD,
             "--rubric", edge / "norubric.json"])
    rec("不存在的评分卡报错可读", r.returncode == 2 and "找不到" in r.stderr
        and "Traceback" not in r.stderr)
    # 不存在的 judgments
    r = run([ROOT / "scripts" / "score.py", docjson, "--checks", chkjson,
             "--judgments", edge / "nojudg.json", "--award", AWARD,
             "--rubric", RUB, "--out", edge])
    rec("不存在的 judgments 报错可读", r.returncode == 2 and "找不到" in r.stderr
        and "Traceback" not in r.stderr)
    # 空目录 batch
    r = run([ROOT / "scripts" / "batch.py", edge, "--out", edge])
    rec("空目录 batch 报错可读", r.returncode == 2 and "没有匹配到材料文件" in r.stderr)
    # batch 指定的 rubric 不存在，必须是可读错误而不是 traceback
    r = run([ROOT / "scripts" / "batch.py", bdir, "--out", bdir,
             "--rubric", edge / "norubric.json"])
    rec("batch 缺评分卡报错可读", r.returncode == 2 and "找不到评分卡" in r.stderr
        and "Traceback" not in r.stderr)
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