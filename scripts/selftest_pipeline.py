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
        ws = wb[wb.sheetnames[0]]
        # 关键列必须存在
        hdr = {str(ws.cell(4, c).value).strip() for c in range(1, ws.max_column + 1)}
        need = {"申报材料 / 修改建议", "总分", "奖项区间", "改完提升", "改完总分",
                "原文是什么", "怎么改", "修改之后的句子"}
        # 「置信度」「需关注」已按要求去掉，不该再出现
        banned = {"置信度", "需关注"}
        rec("Excel 只有一张表", len(sheets) == 1, str(sheets))
        rec("汇总表关键列存在", need <= hdr, "缺 %s" % (need - hdr) if not need <= hdr else "")
        rec("已去掉置信度/需关注列", not (banned & hdr), "仍有 %s" % (banned & hdr) if banned & hdr else "")
        # 总计行必须在汇总表里
        all_txt = " ".join(str(c.value) for row in ws.iter_rows() for c in row if c.value)
        rec("汇总表含总计行", "总计" in all_txt and "平均" in all_txt)
        # 建议列必须有内容（不能是空列）
        # 首要修改建议列的列名含"首要修改建议"，用包含匹配
        adv_col = next((c for c in range(1, ws.max_column + 1)
                        if "修改建议" in str(ws.cell(4, c).value or "")), None)
        filled = 0
        if adv_col:
            filled = sum(1 for rw in range(5, ws.max_row + 1)
                         if ws.cell(rw, adv_col).value
                         and str(ws.cell(rw, adv_col).value).strip() not in ("", "—"))
        # 「预计改完提升」与「改完后总分」两列
        gain_col = next((c for c in range(1, ws.max_column + 1)
                         if "改完提升" in str(ws.cell(4, c).value or "")), None)
        after_col = next((c for c in range(1, ws.max_column + 1)
                          if "改完总分" in str(ws.cell(4, c).value or "")), None)
        rec("有「预计改完提升」列", gain_col is not None)
        rec("有「改完后总分」列", after_col is not None)
        if gain_col and after_col:
            pairs = []
            for rw in range(5, ws.max_row + 1):
                g = ws.cell(rw, gain_col).value
                a = ws.cell(rw, after_col).value
                sc = ws.cell(rw, 2).value
                if isinstance(g, str) and g.startswith("+") and isinstance(a, (int, float)) \
                        and isinstance(sc, (int, float)):
                    pairs.append((sc, float(g[1:]), float(a)))
            rec("提升值为正且改完分不超过100",
                bool(pairs) and all(a <= 100 and g > 0 for s, g, a in pairs),
                "%d 行有提升值" % len(pairs))
            # 改完分 ≈ 当前分 + 提升（允许合计封顶）
            ok_math = all(abs(min(100.0, s + g) - a) < 0.15 for s, g, a in pairs)
            rec("改完分 = 当前分 + 提升", ok_math)

        rec("首要修改建议列有内容", bool(adv_col) and filled >= 1,
            "%s 行有建议" % filled if adv_col else "未找到该列")
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

    # ---------- 6b2. Excel 与 Word 的建议条目必须一一对应 ----------
    print("\n[6b2] 建议条目对应性")
    xl2 = TMP / "corr.xlsx"
    r = run([ROOT / "scripts" / "export_xlsx.py", out, "--out", xl2])
    wdir2 = TMP / "corrwd"
    r2 = run([ROOT / "scripts" / "export_docx.py", out, "--out", wdir2])
    ok = r.returncode == 0 and r2.returncode == 0 and xl2.is_file()
    rec("两份导出都成功", ok)
    if ok:
        from openpyxl import load_workbook as _lw
        from docx import Document as _Doc
        import re as _re
        wb2 = _lw(xl2)
        rec("Excel 只有一张表", len(wb2.sheetnames) == 1, str(wb2.sheetnames))
        ws2 = wb2[wb2.sheetnames[0]]
        # 单表里建议编号出现在第一列（可能带前导全角空格），扫全表提取
        xl_nos = []
        for row in ws2.iter_rows(min_row=5, values_only=True):
            for v in row:
                if v is None:
                    continue
                m2 = _re.search(r"([0-9]+-M[0-9][0-9])\u3000", str(v))
                if m2:
                    xl_nos.append(m2.group(1))
                    break
        wd_nos, lack = [], 0
        for f2 in wdir2.glob("*.docx"):
            if "汇总" in f2.name:
                continue
            txt = "\n".join(p.text for p in _Doc(str(f2)).paragraphs)
            for seg in ("原文是什么", "怎么改", "修改之后的句子"):
                if seg not in txt:
                    lack += 1
            for p in _Doc(str(f2)).paragraphs:
                m = _re.match(r"^([0-9]+-M[0-9][0-9])\u3000", p.text.strip())
                if m:
                    wd_nos.append(m.group(1))
        rec("Word 与 Excel 条目数一致", len(wd_nos) == len(xl_nos) and len(xl_nos) > 0,
            "Word %d / Excel %d" % (len(wd_nos), len(xl_nos)))
        rec("Word 与 Excel 编号集合相同", sorted(wd_nos) == sorted(xl_nos))
        rec("每条建议都是三段式", lack == 0, "缺 %d 处" % lack if lack else "")
        # Word 必须是白底黑字：不允许出现彩色文字
        colored = 0
        for f2 in wdir2.glob("*.docx"):
            for p in _Doc(str(f2)).paragraphs:
                for rr_ in p.runs:
                    c = rr_.font.color
                    if c is not None and c.rgb is not None and str(c.rgb) not in ("000000", "555555"):
                        colored += 1
        rec("Word 无彩色文字", colored == 0, "彩色 run %d 处" % colored if colored else "")
        # 四节结构齐全（含免责说明）
        first_txt = "\n".join(p.text for p in _Doc(str(next(wdir2.glob("*.docx")))).paragraphs)
        has4 = all(sec in first_txt for sec in ("一、", "二、", "三、", "四、免责说明"))
        rec("Word 有四节（含免责说明）", has4)

        # 内容质量：结构齐全不等于内容可用。逐条检查三段都不是空话。
        short_o, fb_no_need, x_tpl, checked_items = 0, 0, 0, 0
        for f2 in wdir2.glob("*.docx"):
            if "汇总" in f2.name:
                continue
            paras2 = [p.text.strip() for p in _Doc(str(f2)).paragraphs]
            idx2 = [i for i, x in enumerate(paras2)
                    if _re.match(r"^[0-9]+-M[0-9][0-9]\u3000", x)]
            segs2 = ["原文是什么", "怎么改", "修改之后的句子（可直接复制）"]
            for j2, st2 in enumerate(idx2):
                en2 = idx2[j2 + 1] if j2 + 1 < len(idx2) else len(paras2)
                blk = paras2[st2:en2]
                vel = {}
                for s2 in segs2:
                    if s2 in blk:
                        i2 = blk.index(s2)
                        vel[s2] = blk[i2 + 1] if i2 + 1 < len(blk) and blk[i2 + 1] not in segs2 else ""
                checked_items += 1
                if len(vel.get(segs2[0], "")) < 12:
                    short_o += 1
                if "证据等级" in vel.get(segs2[0], "")[:6] and "需补充" not in vel.get(segs2[0], ""):
                    fb_no_need += 1
                if "X 方案" in vel.get(segs2[2], "") and "填：" not in vel.get(segs2[2], ""):
                    x_tpl += 1
        rec("原文段不空泛（>=12 字）", short_o == 0,
            "过短 %d/%d" % (short_o, checked_items) if short_o else "%d 条均合格" % checked_items)
        rec("证据不足时写明需补什么", fb_no_need == 0,
            "未说明 %d 条" % fb_no_need if fb_no_need else "")
        rec("可抄句式不含裸 X 占位", x_tpl == 0,
            "裸占位 %d 条" % x_tpl if x_tpl else "")

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

    # ---------- 10d. 参考材料与赛道一致性 ----------
    print("\n[10d] 参考材料不参与排名 / 赛道不一致要报出")
    refd = TMP / "refmix"; refd.mkdir(exist_ok=True)
    import shutil as _sh3
    _sh3.copy(docx_path, refd / docx_path.name)
    # 造一份"过往版本"参考材料
    ref_docx = refd / (AWARD + "-测试材料-过往版本.docx")
    _sh3.copy(docx_path, ref_docx)
    _sh3.copy(jp, refd / (AWARD + "-测试材料-过往版本.judgments.json"))
    _sh3.copy(jp, refd / jp.name)
    r = run([ROOT / "scripts" / "batch.py", refd, "--out", refd, "--rubric", RUB])
    lb = refd / "leaderboard.md"
    ok = r.returncode == 0 and lb.is_file()
    rec("含参考材料的批次正常完成", ok)
    if ok:
        txt = lb.read_text(encoding="utf-8")
        rec("参考材料单列且注明不计入排名",
            "参考材料（不计入排名）" in txt and "过往版本" in txt)
        mf2 = json.loads((refd / "manifest.json").read_text(encoding="utf-8"))
        rec("manifest 区分参赛与参考",
            mf2.get("references") == 1 and mf2.get("submissions", 0) == mf2.get("materials", 0),
            "submissions=%s references=%s materials=%s"
            % (mf2.get("submissions"), mf2.get("references"), mf2.get("materials")))

    # 赛道不一致：文件名写 A，正文写 B，必须被报出。
    # 注意：必须用当前评分卡里真实存在的奖项名做文件名前缀，否则 batch 先因
    # "无法判断奖项"跳过，根本走不到赛道核对那一步。
    mis = TMP / "mismatch"; mis.mkdir(exist_ok=True)
    import docx as _docx
    d = _docx.Document()
    for line in ["参赛赛道：这是另一个赛道的名称", "改造前电单耗 0.1069 kWh/吨",
                 "累计已实现节约电费 202.745 万元"]:
        d.add_paragraph(line)
    d.save(str(mis / (AWARD + "-赛道不一致材料.docx")))
    _sh3.copy(jp, mis / (AWARD + "-赛道不一致材料.judgments.json"))
    r = run([ROOT / "scripts" / "batch.py", mis, "--out", mis, "--rubric", RUB])
    both = (r.stdout or "") + (r.stderr or "")
    rec("赛道不一致被报出", "正文声明" in both,
        [x for x in both.splitlines() if "正文声明" in x][:1])

    # ---------- 10e. 格式问题只提醒不扣分 ----------
    print("\n[10e] 格式问题只提醒、不扣分")
    fmt_dir = TMP / "fmtonly"; fmt_dir.mkdir(exist_ok=True)
    import docx as _dx
    d0 = _dx.Document()
    for line in ["改造前电单耗 0.1069 kWh/吨，改造后 0.0784 kWh/吨",
                 "累计已实现节约电费 202.745 万元",
                 "已在东部大区 3 个区域 13 家水厂推广应用",
                 "填报单位：____    填报日期：XX年XX月XX日"]:
        d0.add_paragraph(line)
    d0.save(str(fmt_dir / (AWARD + "-含占位符材料.docx")))
    jf2 = json.loads(jp.read_text(encoding="utf-8"))
    (fmt_dir / (AWARD + "-含占位符材料.judgments.json")).write_text(
        json.dumps(jf2, ensure_ascii=False), encoding="utf-8")
    r = run([ROOT / "scripts" / "batch.py", fmt_dir, "--out", fmt_dir, "--rubric", RUB])
    sj = fmt_dir / (AWARD + "-含占位符材料.score.json")
    ok = r.returncode == 0 and sj.is_file()
    rec("含占位符材料能正常评分", ok)
    if ok:
        sd = json.loads(sj.read_text(encoding="utf-8"))
        cids = {x.get("check") for x in (sd.get("deductions") or [])}
        rec("占位符 C09 不产生扣分", "C09" not in cids, "扣分项: %s" % sorted(c for c in cids if c))
        cids17 = {x.get("check") for x in (sd.get("deductions") or [])}
        rec("截图 C17 不产生扣分", "C17" not in cids17)
        # 但必须出现在争议点里提醒
        pts = " ".join(str(x.get("point")) for x in (sd.get("disputes") or []))
        has_c09_notice = ("占位" in pts) and ("不扣分" in pts)
        rec("占位符问题仍作为提醒出现", has_c09_notice)
        # 规则本身也要守住：score.py 里 C09/C17 不许再出现 deduct(
        src_sc = (ROOT / "scripts" / "score.py").read_text(encoding="utf-8")
        bad_lines = [ln.strip() for ln in src_sc.splitlines()
                     if ("C09" in ln or "C17" in ln)
                     and "deduct(" in ln and not ln.strip().startswith("#")]
        rec("score.py 里 C09/C17 不得 deduct", not bad_lines,
            bad_lines[0][:60] if bad_lines else "")
        # 文档必须写明这条规则
        sk = (ROOT / "SKILL.md").read_text(encoding="utf-8")
        rec("SKILL.md 写明格式问题不扣分",
            ("C09" in sk and "C17" in sk)
            and ("不扣分" in sk or "not deducted" in sk or "no deduction" in sk))

    # ---------- 10f. 格式检查器本身必须能抓问题 ----------
    print("\n[10f] 格式检查器")
    vf = ROOT / "scripts" / "verify_format.py"
    rec("verify_format.py 存在", vf.is_file())
    if vf.is_file():
        # 正常产物应全过（自测批次不出 PDF，故跳过 PDF 检查）
        r = run([vf, bdir, "--no-pdf"])
        rec("正常产物格式验收通过", r.returncode == 0,
            (r.stdout.strip().splitlines() or [""])[-1][:50])
        # 删掉一个 pdf 必须报错（若有 pdf）
        pdfs = list(bdir.glob("*.pdf"))
        if pdfs:
            bak = pdfs[0].with_suffix(".pdf.bak")
            pdfs[0].rename(bak)
            r = run([vf, bdir])
            rec("verify_format 能按预期判断 PDF", r.returncode != 0)
            bak.rename(pdfs[0])
        else:
            # 本批没产出 PDF（未装转换器或未要 --with-pdf），这条无从验证，跳过不计失败
            print("        [SKIP] verify_format 能抓到缺 PDF（本批无 PDF）")

    # ---------- 10g. 证据等级封顶必须严格按 rubric ----------
    print("\n[10g] 证据等级封顶只认 rubric")
    src_sc2 = (ROOT / "scripts" / "score.py").read_text(encoding="utf-8")
    # 只看代码行，不看注释（注释里会提到历史常量名）
    code_lines = [ln for ln in src_sc2.splitlines()
                  if ln.strip() and not ln.strip().startswith("#")]
    rec("score.py 不再有 GRADE_FACTOR 常量",
        not any("GRADE_FACTOR = {" in ln for ln in code_lines),
        "发现第二套系数表")
    rec("score.py 有 grade_cap_value()", "def grade_cap_value(" in src_sc2)

    # 造全假引文，逼出降级，验证分数不超 rubric cap
    dg = TMP / "downgrade"; dg.mkdir(exist_ok=True)
    jd = json.loads(jp.read_text(encoding="utf-8"))
    for k in (jd.get("judgments") or {}):
        jd["judgments"][k]["grade"] = "A"                 # 自报最高级
        jd["judgments"][k]["quotes"] = ["这句话原文里绝对不存在假引文测试"]  # 全部编造
    (dg / (AWARD + "-降级测试.judgments.json")).write_text(
        json.dumps(jd, ensure_ascii=False), encoding="utf-8")
    import shutil as _sh4
    _sh4.copy(docx_path, dg / (AWARD + "-降级测试.docx"))
    r = run([ROOT / "scripts" / "batch.py", dg, "--out", dg, "--rubric", RUB])
    sfile = dg / (AWARD + "-降级测试.score.json")
    if sfile.is_file():
        sd = json.loads(sfile.read_text(encoding="utf-8"))
        rub_d = json.loads(RUB.read_text(encoding="utf-8"))
        caps = {k: v["cap"] for k, v in (rub_d.get("evidence_grades") or {}).items()}
        bad = []
        for x in sd["dimensions"]:
            g = x.get("verified_grade")
            if g in caps and x["score"] > x["max"] * caps[g] + 0.05:
                bad.append("%s %s级 %.1f>%.1f" % (x["name"], g, x["score"],
                                                  x["max"] * caps[g]))
        rec("降级后分数不超 rubric cap", not bad, "; ".join(bad[:2]))
        rec("编造引文确实被拒", all(
            not x.get("quotes_verified") for x in sd["dimensions"]))
    else:
        rec("降级后分数不超 rubric cap", False, "没产出 score.json")

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