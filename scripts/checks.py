#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""机检：对 extract_report.py 产出的 doc.json 做可复算的规则核查。

每个检查项输出 status(pass/warn/fail/na) + evidence(指到 表N/行M 或 原文片段)，
并把结果映射到受影响维度（rubric 里的 dimension id），供 score.py 封顶/扣分。
"""
from __future__ import annotations
import argparse, json, re, sys
from pathlib import Path

NUM = re.compile(r"-?\d+(?:,\d{3})*(?:\.\d+)?")
PCT = re.compile(r"(-?\d+(?:\.\d+)?)\s*%")
PLACEHOLDER = re.compile(r"(_{3,}|XXXX|××××|（文本）|\(文本\)|TODO|待填写)")

def num(s):
    if s is None: return None
    m = NUM.search(str(s).replace(" ", ""))
    if not m: return None
    try: return float(m.group(0).replace(",", ""))
    except ValueError: return None

def cell(row, i):
    return row[i] if i < len(row) else ""

def find_table(doc, need_header=(), min_rows=1):
    out = []
    for b in doc["blocks"]:
        if b["type"] != "table": continue
        head = " ".join(b.get("header") or [])
        if all(k in head for k in need_header) and b.get("n_rows", 0) >= min_rows:
            out.append(b)
    return out

def col_index(tbl, *keys):
    head = tbl.get("header") or []
    for i, h in enumerate(head):
        if all(k in h for k in keys): return i
    return None

def snippet(s, n=48):
    s = re.sub(r"\s+", " ", str(s or "")).strip()
    return s[:n] + ("..." if len(s) > n else "")

def R(cid, name, status, detail, evidence=None, dims=None, rule=None):
    return {"id": cid, "name": name, "status": status, "detail": detail,
            "evidence": evidence or [], "dims": dims or [], "rule": rule}

def profile(award, rub):
    """每个奖项在 rubric 里声明自己需要哪些专属检查（profile），检查按声明启用而不是按奖项名硬编码。"""
    try:
        return rub["awards"][award].get("profile") or {}
    except KeyError:
        return {}

def c01_required_tables(doc, award, rub):
    if not profile(award, rub).get("generic_tables", True):
        return R("C01", "结构表完整性", "na", "该奖项不使用通用三表模板，见本奖项的专属检查", dims=[])
    need = [("basic_fields", ["基础字段"]), ("baseline_effect", ["效益维度", "基准值"]), ("verify_items", ["核验项"])]
    missing = [k for k, h in need if not find_table(doc, h)]
    if missing:
        return R("C01", "结构表完整性", "fail", "缺少必备表格：" + ",".join(missing), dims=["value", "replication"])
    return R("C01", "结构表完整性", "pass", "基础字段/效益维度/核验项三表齐全")

def c02_actual_vs_plan(doc, award, rub):
    tbls = find_table(doc, ["效益维度", "基准值"])
    if not tbls:
        return R("C02", "已实现/在研分离", "na", "未找到效益维度表", dims=["value"])
    tb = tbls[0]
    ia = col_index(tb, "实际值"); ip = col_index(tb, "远期目标")
    if ia is None or ip is None:
        return R("C02", "已实现/在研分离", "fail", "效益表缺「实际值【已实现】」或「远期目标【规划/在研】」列",
                 evidence=["表%d 表头：%s" % (tb["index"], " | ".join(tb["header"]))], dims=["value"], rule="G2")
    filled = [(i, r) for i, r in enumerate(tb["rows"]) if num(cell(r, ia)) is not None]
    ev = ["表%d 行%d：%s -> 实际值=%s" % (tb["index"], i + 1, snippet(cell(r, 0), 12), snippet(cell(r, ia), 30)) for i, r in filled][:5]
    if not filled:
        return R("C02", "已实现/在研分离", "fail", "「实际值【已实现】」全空——所有效益均无已实现数据", evidence=ev, dims=["value"], rule="G2")
    if len(filled) <= 1:
        return R("C02", "已实现/在研分离", "warn", "仅 %d 项效益填写了已实现数值，其余为规划" % len(filled), evidence=ev, dims=["value"], rule="G2")
    return R("C02", "已实现/在研分离", "pass", "%d 项效益有已实现数值，且与远期目标分列" % len(filled), evidence=ev)

def c03_baseline(doc):
    tbls = find_table(doc, ["效益维度", "基准值"])
    if not tbls: return R("C03", "基线可算性", "na", "无效益维度表", dims=["value"])
    tb = tbls[0]
    ib = col_index(tb, "基准值"); ia = col_index(tb, "实际值")
    bad, ok = [], 0
    for i, r in enumerate(tb["rows"]):
        a, b = num(cell(r, ia)), num(cell(r, ib))
        if a is None: continue
        if b is None: bad.append(i)
        else: ok += 1
    if bad:
        return R("C03", "基线可算性", "warn", "%d 项有实际值但无基准值，变化率无法复算" % len(bad),
                 evidence=["表%d 行%d" % (tb["index"], i + 1) for i in bad[:5]], dims=["value"], rule="G4/D3")
    if not ok:
        return R("C03", "基线可算性", "na", "无同时具备基准值与实际值的行", dims=["value"])
    return R("C03", "基线可算性", "pass", "%d 项基准值/实际值成对，可复算" % ok)

def c04_evidence_column(doc):
    tbls = find_table(doc, ["效益维度", "基准值"])
    if not tbls: return R("C04", "效益佐证覆盖率", "na", "无效益维度表", dims=["value"])
    tb = tbls[0]
    ia, ie = col_index(tb, "实际值"), col_index(tb, "佐证")
    if ie is None: return R("C04", "效益佐证覆盖率", "fail", "效益表无「佐证材料」列", dims=["value"])
    used = [(i, r) for i, r in enumerate(tb["rows"]) if num(cell(r, ia)) is not None]
    if not used: return R("C04", "效益佐证覆盖率", "na", "无已实现数值行", dims=["value"])
    def has_ev(r):
        v = cell(r, ie).strip()
        return len(r) > ie and v not in ("", "—", "-", "/")
    with_ev = [(i, r) for i, r in used if has_ev(r)]
    ratio = len(with_ev) / len(used)
    st = "pass" if ratio >= 0.8 else ("warn" if ratio >= 0.4 else "fail")
    missing = [i for i, r in used if not has_ev(r)]
    return R("C04", "效益佐证覆盖率", st, "%d/%d 项已实现效益附带佐证材料（%.0f%%）" % (len(with_ev), len(used), ratio * 100),
             evidence=["表%d 行%d 缺佐证" % (tb["index"], i + 1) for i in missing][:5],
             dims=["value"], rule="G1")

def c05_recompute(doc):
    tbls = find_table(doc, ["效益维度", "基准值"])
    if not tbls: return R("C05", "数字复算一致性", "na", "无效益维度表", dims=["value"])
    tb = tbls[0]
    ib, ia = col_index(tb, "基准值"), col_index(tb, "实际值")
    checked, bad = 0, []
    clean = re.compile(r"^\s*-?\d+(?:\.\d+)?\s*(?:kWh|mg/L|元|万元|吨|m3|%|/)?", re.I)
    for i, r in enumerate(tb["rows"]):
        base_cell, act_cell = cell(r, ib), cell(r, ia)
        if not clean.match(str(base_cell)) or not clean.match(str(act_cell)): continue   # 非纯数字行跳过，避免口径污染
        b, a = num(base_cell), num(act_cell)
        if not b or a is None: continue
        stated = PCT.search(str(act_cell))                                              # 只认与数值同格的百分比
        if not stated: continue
        calc = (a - b) / b * 100
        checked += 1
        if abs(abs(calc) - abs(float(stated.group(1)))) > 5:
            bad.append("表%d 行%d：自述 %s%%，复算 %.1f%%" % (tb["index"], i + 1, stated.group(1), calc))
    if not checked: return R("C05", "数字复算一致性", "na", "无可复算的变化率对", dims=["value"])
    if bad: return R("C05", "数字复算一致性", "warn", "%d/%d 处变化率与复算不符（偏差>5pp）" % (len(bad), checked), evidence=bad[:5], dims=["value"], rule="D2")
    return R("C05", "数字复算一致性", "pass", "%d 处变化率复算一致" % checked)

def c06_promotion_list(doc, award, rub):
    if not profile(award, rub).get("promotion"):
        return R("C06", "推广清单", "na", "该奖项未启用推广清单核查", dims=["promotion"])
    cands = [b for b in doc["blocks"] if b["type"] == "table" and b.get("n_cols", 0) >= 6]
    plants, regions, src = set(), set(), None
    if cands:
        best = max(cands, key=lambda b: b["n_rows"])
        flat = [c for r in best["rows"] for c in r if c.strip()]          # 只数数据行，不数表头
        plants = {c for c in flat if re.search(r"(厂|公司|污水|净水)", c) and len(c) <= 14}
        regions = {c for c in flat if re.search(r"大区", c)}
        src = "表%d" % best["index"]
    text = doc["plain"]
    # 「13家应用项目」「17个项目推广」都要能识别。
    # 旧正则要求「家/个」后紧跟「项目/厂站/厂」，中间差一个修饰词就漏——
    # 实例：水厂设备卫士原文「东部大区13家应用项目」，C06 判 fail 并称"正文也无表述"。
    n_plants_txt = [int(x) for x in re.findall(
        r"(\d{1,3})\s*(?:家|个)(?:[\u4e00-\u9fa5]{0,4})?(?:项目|厂站|厂|项目公司|水厂|网点)", text)]
    n_region_txt = [int(x) for x in re.findall(r"(\d{1,2})\s*个\s*(?:大区|区域)", text)]
    guess = max([len(plants)] + n_plants_txt) if (plants or n_plants_txt) else 0
    n_reg = max([len(regions)] + n_region_txt) if (regions or n_region_txt) else 0
    if guess == 0:
        return R("C06", "推广清单", "fail", "未找到推广厂站清单，正文也无「N 家项目/N 个厂站」表述", dims=["promotion"], rule="金火炬-推广力度")
    lvl = "深度（10 个以上）" if guess > 10 else ("适度（6-10 个）" if guess >= 6 else "基础（1-5 个）")
    st = "pass" if guess > 10 else ("warn" if guess >= 6 else "fail")
    basis = ("清单" + src) if plants else "正文表述"
    return R("C06", "推广清单", st, "按%s识别推广项目 %d 个、覆盖大区或区域 %d 个 -> 官方数量档位：%s" % (basis, guess, n_reg, lvl),
             evidence=["正文出现：%s" % snippet(",".join(map(str, sorted(set(n_plants_txt))[:6])) + " 家/个 项目表述", 50)] if n_plants_txt else [],
             dims=["promotion"])

def c07_calc_steps(doc, award, rub):
    if not profile(award, rub).get("promotion"):
        return R("C07", "测算过程", "na", "该奖项未启用测算步骤核查", dims=[])
    # 表头可能写「公式」「计算步骤」「计算过程」「测算过程」「计算方法」任何一种。
    # 只认「公式」太窄：实例——水厂设备卫士表头写「计算过程」、公式级披露是全批最完整的之一，
    # 却被判 fail 并扣分。
    CALC_WORDS = ("公式", "计算步骤", "计算过程", "测算过程", "测算步骤", "计算方法", "计算口径")
    steps = [b for b in doc["blocks"] if b["type"] == "table"
             and any(any(w in str(h) for w in CALC_WORDS) for h in (b.get("header") or []))]
    if not steps:
        # 兜底：正文里出现「计算过程：」「测算过程：」这类小标题也算有测算
        body_hit = re.search(r"(计算步骤|计算过程|测算过程|测算步骤|计算方法)[：:]", doc.get("plain", ""))
        if not body_hit:
            return R("C07", "测算过程", "fail",
                     "未找到「计算步骤/公式」测算表，经济效益无法复算",
                     dims=["economics"], rule="金火炬-经济效益")
        return R("C07", "测算过程", "warn",
                 "正文有「%s」小标题，但未见结构化测算表，建议改成表格便于复算" % body_hit.group(1),
                 dims=["economics"], rule="金火炬-经济效益")
    total = sum(b["n_rows"] for b in steps)
    return R("C07", "测算过程", "pass", "%d 张测算表、共 %d 个计算步骤" % (len(steps), total),
             evidence=["表%d：%s" % (b["index"], snippet(" | ".join(b["header"]), 50)) for b in steps[:3]], dims=["economics"])

def c08_duplicate_benefit(doc):
    money = {}
    for b in doc["blocks"]:
        if b["type"] != "table": continue
        for i, r in enumerate(b["rows"]):
            joined = " ".join(r)
            if not re.search(r"(万元|元/|节约|节降|收益|效益)", joined): continue
            # 金额必须是完整数字，不能从中间截。
            # 旧正则 d{2,6} 至少两位，会把 "0.50万元" 从 "50" 开始匹配读成 50.0 万元，
            # 从而把 0.50 / 4.50 这类小额误判成重复计入（实例：Be-PIPE 被多扣 3 分）。
            loc = "表%d 行%d" % (b["index"], i + 1)
            for m in re.finditer(r"(?<![\d.])(\d+(?:\.\d+)?)\s*万元", joined):
                v = float(m.group(1))
                if v < 0.01:
                    continue
                # 同一行里同一金额出现多次不算"多处出现"——
                # 一个单元格写"0.3万元（含税、0.3万元以下）"是正常表述，
                # 不是重复计入。只有跨行/跨表复现才算。
                # 用原值做键，不 round——"741.0万元"在原文里根本不存在（原文是 741.3 或类似），
                # 四舍五入后的证据会让人以为机检读错了，反而降低可信度。
                bucket = money.setdefault(v, [])
                if loc not in bucket:
                    bucket.append(loc)
    dup = {k: v for k, v in money.items() if len(v) > 1}
    if len(dup) >= 3:
        return R("C08", "重复计入检查", "warn", "%d 个金额在多处出现，需确认是否重复计入" % len(dup),
                 evidence=["%s万元：%s" % (k, "、".join(v[:3])) for k, v in list(dup.items())[:4]], dims=["value", "economics"], rule="D1")
    return R("C08", "重复计入检查", "pass", "未发现明显的同一金额多处复用（启发式）")

def c09_placeholders(doc):
    hits = []
    for b in doc["blocks"]:
        texts = [b.get("text", "")] + ([c for r in b.get("rows", []) for c in r] if b["type"] == "table" else [])
        for t in texts:
            if PLACEHOLDER.search(str(t)):
                hits.append(snippet(t))
    if hits:
        return R("C09", "占位符残留", "warn", "%d 处未填写的占位（下划线/XXXX/（文本））" % len(hits), evidence=hits[:6], dims=["value", "replication"])
    return R("C09", "占位符残留", "pass", "未发现占位符残留")

def c10_length(doc, award, rub):
    m = rub["awards"][award]["min_chars"]
    n = doc["stats"]["chars"]
    st = "pass" if n >= m else ("warn" if n >= m * 0.6 else "fail")
    return R("C10", "材料篇幅", st, "正文 %d 字（该奖项参考下限 %d 字）" % (n, m), dims=[])

def c11_model_lab(doc, award, rub):
    if not profile(award, rub).get("model_lab"):
        return R("C11", "模型实验室/生产验证证据", "na", "该奖项未启用模型实验室核查", dims=["intelligence"])
    text = doc["plain"]
    keys = {"实验室部署": r"(模型实验室|线上部署)", "试点厂调用": r"(试点厂|生产调用|全流程生产)", "运行天数": r"(\d{2,3})\s*天"}
    found = {k: bool(re.search(v, text)) for k, v in keys.items()}
    days = [int(x) for x in re.findall(r"(\d{2,3})\s*天", text)]
    long_run = [d for d in days if d >= 90]
    st = "pass" if all(found.values()) and long_run else ("warn" if sum(found.values()) >= 2 else "fail")
    return R("C11", "模型实验室/生产验证证据", st,
             "实验室部署=%s、试点厂调用=%s、>=90天运行证据=%s" % (found["实验室部署"], found["试点厂调用"], bool(long_run)),
             evidence=["文中天数：%s" % sorted(set(days))[:8]] if days else [], dims=["intelligence"], rule="G3")

def c12_product_readiness(doc, award, rub):
    if not profile(award, rub).get("product"):
        return R("C12", "产品就绪度证据", "na", "该奖项未启用产品就绪度核查", dims=[])
    text = doc["plain"]
    lvl = None
    if re.search(r"(产品化|产品转化|已形成产品|销售收入|销售合同)", text): lvl = "非常符合（已实现产品转化）"
    elif re.search(r"(原型|[Pp]roto)", text): lvl = "符合（形成产品原型）"
    elif re.search(r"(方案阶段|尚未产品化|仍在方案)", text): lvl = "一般（未产品化）"
    elif re.search(r"产品就绪度", text): lvl = "一般（提到产品就绪度但未说明转化程度）"
    st = "pass" if lvl and lvl.startswith("非常") else ("warn" if lvl else "fail")
    return R("C12", "产品就绪度证据", st, "识别到：" + (lvl or "无明确表述"), dims=["readiness"], rule="G5")

def c13_iteration(doc, award, rub):
    """成果迭代性：必须点明"相比上一版本/上次参赛"的实质变化。

    判定需要两个要素同时成立（且相距不远，避免全篇分散命中）：
      ① 对比锚点：指向过去的参赛版本（上一届/往届/原版本/第N届(N<本届)…）
      ② 变化内容：迭代/升级/新增/改进/重构… 等实质变更表述
    只出现"迭代/升级"而没有对比对象的，只能算自述，降 warn。
    """
    if not profile(award, rub).get("product"):
        return R("C13", "成果迭代性（否决项）", "na", "该奖项未启用成果迭代性核查", dims=[])
    text = doc["plain"]

    # ① 对比锚点：只认"过去"，本届（第八届）不算
    ANCHOR = (r"(?:上一届|上届|上次参赛|往届|历届|过去获奖版本|此前获奖|"
              r"原版本|前版|原有版本|此前版本|上一版本|既有版本|原方案|"
              r"第[一二三四五六七]届)")
    # ② 变化内容
    # 注意：不把"拓展/扩展"算作变更词——它们在申报材料里多用于"推广范围拓展至…"，
    # 与"相比上一版本的迭代"无关，放进去会误判。
    # "优化/改进"在申报材料里也常用于"效率优化""工作安排优化"，需要带技术宾语才算版本迭代
    CHANGE = (r"(?:迭代|升级|重构|改造|替换|新增|加入|引入|上线|"
              r"[\u4e00-\u9fa5]{0,4}(?:功能|算法|模型|架构|方案|程序|逻辑|系统|设备|工艺|结构|性能)"
              r"[\u4e00-\u9fa5]{0,3}(?:优化|改进|增强|完善|提升))")

    anchors = [m for m in re.finditer(ANCHOR, text)]
    changes = [m for m in re.finditer(CHANGE, text)]

    if not changes:
        return R("C13", "成果迭代性（否决项）", "fail",
                 "全文未见任何迭代/升级/新增/改进类表述——无法支撑成果迭代性",
                 dims=["iteration"], rule="金火炬-否决项")

    # 锚点与变化相邻（±100 字）才算说清了"相比谁、变了什么"
    if anchors:
        for a in anchors:
            near = [c for c in changes if abs(c.start() - a.start()) <= 100]
            if near:
                ctx = text[max(0, min(a.start(), near[0].start()) - 20):
                           max(a.start(), near[0].start()) + 90].replace("\n", " ")
                return R("C13", "成果迭代性（否决项）", "pass",
                         "点明了「%s」与「%s」的对应关系" % (a.group(0), near[0].group(0)),
                         evidence=["…" + ctx.strip()[:110] + "…"], dims=["iteration"])

    # 有变化表述但没有对比对象 → 自述，交人工确认
    ctx = text[max(0, changes[0].start() - 40):changes[0].start() + 80].replace("\n", " ")
    return R("C13", "成果迭代性（否决项）", "warn",
             "出现「%s」但未点明相较哪一版本，属于自述，需人工确认" % changes[0].group(0),
             evidence=["…" + ctx.strip()[:110] + "…"], dims=["iteration"], rule="金火炬-否决项")

def c14_novelty_compare(doc, award, rub):
    text = doc["plain"]
    has_cmp = bool(re.search(r"(对比|比较|优于|显著低于|显著高于|优于原|原有办法|传统方式|相比)", text))
    has_data = bool(re.search(r"(\d+(?:\.\d+)?\s*%|\d+(?:\.\d+)?\s*倍|下降\s*\d|提升\s*\d)", text))
    if has_cmp and has_data:
        return R("C14", "创新性对照证据", "pass", "存在对比表述且带量化数据", dims=["novelty"])
    if has_cmp:
        return R("C14", "创新性对照证据", "warn", "有对比表述但缺量化对照", dims=["novelty"])
    return R("C14", "创新性对照证据", "fail", "全文未见与原有办法的对照", dims=["novelty"], rule="官方创新性维度")

def c15_strategy_fit(doc, award, rub):
    if not profile(award, rub).get("strategy"):
        return R("C15", "战略对齐与卡点", "na", "该奖项未启用战略对齐核查", dims=[])
    text = doc["plain"]
    strat = bool(re.search(r"(集团战略|十五五|战略方向|经营目标)", text))
    pain = bool(re.search(r"(卡点|痛点|堵点|瓶颈)", text))
    st = "pass" if strat and pain else ("warn" if strat or pain else "fail")
    return R("C15", "战略对齐与卡点", st, "战略对齐=%s、卡点描述=%s" % (strat, pain), dims=["usability"])

def c16_realism(doc):
    text = doc["plain"]
    soft = len(re.findall(r"(预计|有望|预期|拟)", text))
    total_rows = sum(b.get("n_rows", 0) for b in doc["blocks"] if b["type"] == "table")
    if total_rows == 0 and soft >= 3:
        return R("C16", "真实性软信号", "fail", "全文 %d 处「预计/有望」，且无结构化表格支撑" % soft, dims=["value"], rule="G1")
    if soft > total_rows * 2 and soft >= 8:
        return R("C16", "真实性软信号", "warn", "%d 处「预计/有望」相对 %d 行表格偏多，需人工判断是否包装" % (soft, total_rows), dims=["value"], rule="G1")
    return R("C16", "真实性软信号", "pass", "预计性表述与结构化数据比例正常")

def c17_image_evidence(doc):
    n = doc["stats"].get("images", 0)
    text = doc["plain"]
    shots = len(re.findall(r"(截图|截屏|系统界面)", text))
    st = "pass" if n >= 3 or shots >= 2 else ("warn" if n >= 1 or shots >= 1 else "fail")
    return R("C17", "图片/截图证据密度", st, "内嵌图片 %d 张、文中提及截图 %d 次" % (n, shots), dims=["value"], rule="D4")

def c18_torch_subscores(doc, award, rub):
    """该奖项三大子项各自的证据是否齐（案例佐证 / 测算与证明材料 / 交付形态）。"""
    if not profile(award, rub).get("promotion"):
        return R("C18", "子项证据完备性", "na", "该奖项未启用子项证据核查", dims=[])
    text = doc["plain"]
    ev_case = bool(re.search(r"(佐证|案例|证明材料|附件)", text))
    ev_money = bool(re.search(r"(测算|核算|成本效益|销售收入)", text))
    ev_prod = bool(re.search(r"(产品|交付|形态|服务)", text))
    n = sum([ev_case, ev_money, ev_prod])
    st = "pass" if n == 3 else ("warn" if n == 2 else "fail")
    return R("C18", "金火炬子项证据", st,
             "推广案例佐证=%s、经济测算=%s、产品形态=%s" % (ev_case, ev_money, ev_prod),
             dims=["promotion", "economics", "readiness"])

def c19_rongchuang_sections(doc, award, rub):
    """专属模板核查：背景 / 方案 / 价值主张 / ROI 核算 / 量化数据等章节是否齐全。"""
    if not profile(award, rub).get("template_sections"):
        return R("C19", "专属模板完整性", "na", "该奖项未启用模板完整性核查", dims=[])
    text = doc["plain"]
    need = {
        "实施背景": r"实施背景",
        "模型应用": r"模型应用|AI工具介绍",
        "创新性": r"创新性",
        "价值性": r"价值性",
        "复用性": r"复用性",
        "ROI核算": r"ROI|投入产出",
    }
    miss = [k for k, v in need.items() if not re.search(v, text)]
    st = "pass" if not miss else ("warn" if len(miss) <= 2 else "fail")
    return R("C19", "AI融创模板完整性", st,
             "缺章节：%s" % ("、".join(miss) if miss else "无（六段齐全）"),
             dims=["value", "replication", "novelty", "intelligence"],
             rule="AI融创官方模板（第3节「业务价值」在决赛中作为打分依据）")

def c20_roi(doc, award, rub):
    """ROI 是否有可复算的投入/收益两段。"""
    if not profile(award, rub).get("roi"):
        return R("C20", "ROI 可算性", "na", "该奖项未启用 ROI 核查", dims=["value"])
    text = doc["plain"]
    has_in = bool(re.search(r"(投入|成本|Inv|一次性投入|月度运行成本)", text))
    has_out = bool(re.search(r"(收益|节降|节约|ROI|回收期)", text))
    has_num = len(re.findall(r"\d+(?:\.\d+)?\s*(?:万元|元|%)", text)) >= 3
    st = "pass" if (has_in and has_out and has_num) else ("warn" if (has_in and has_out) else "fail")
    return R("C20", "ROI 可算性", st,
             "投入=%s、收益=%s、量化数字≥3=%s" % (has_in, has_out, has_num), dims=["value"])

def _default_rubric():
    """默认评分卡：优先同目录的 rubric.json，没有就用 rubric.example.json。

    两个文件名都认，是为了让「本地完整评分卡」和「公开脱敏示例卡」共用同一份脚本。
    """
    assets = Path(__file__).resolve().parents[1] / "assets"
    for name in ("rubric.json", "rubric.example.json"):
        p = assets / name
        if p.is_file():
            return p
    return assets / "rubric.json"


CHECKS = [c01_required_tables, c02_actual_vs_plan, c03_baseline, c04_evidence_column, c05_recompute,
          c06_promotion_list, c07_calc_steps, c08_duplicate_benefit, c09_placeholders, c10_length,
          c11_model_lab, c12_product_readiness, c13_iteration, c14_novelty_compare, c15_strategy_fit,
          c16_realism, c17_image_evidence, c18_torch_subscores, c19_rongchuang_sections, c20_roi]

def run(doc, award, rub):
    import inspect
    out = []
    for fn in CHECKS:
        try:
            n = len(inspect.signature(fn).parameters)
            out.append(fn(*[doc, award, rub][:n]))
        except Exception as e:
            out.append({"id": fn.__name__, "name": fn.__name__, "status": "error",
                        "detail": "%s: %s" % (type(e).__name__, e), "evidence": [], "dims": []})
    return out

def summarize(checks):
    from collections import Counter
    c = Counter(x["status"] for x in checks)
    return {"pass": c.get("pass", 0), "warn": c.get("warn", 0), "fail": c.get("fail", 0),
            "na": c.get("na", 0), "total": len(checks)}

def _load_json(path: str, what: str):
    """读 JSON 文件，缺失或损坏都给可读错误，不抛 traceback。"""
    p = Path(path)
    if not p.is_file():
        print("错误：找不到%s文件：%s" % (what, p), file=sys.stderr)
        raise SystemExit(2)
    try:
        return json.loads(p.read_text(encoding="utf-8"))
    except json.JSONDecodeError as e:
        print("错误：%s文件不是合法 JSON（%s）：\n  %s" % (what, e.msg, p), file=sys.stderr)
        raise SystemExit(2)


def main() -> int:
    ap = argparse.ArgumentParser(description="机检 doc.json -> checks.json")
    ap.add_argument("doc_json")
    ap.add_argument("--award", required=True, help="奖项名，必须与 rubric 里 awards 的某个键一致")
    ap.add_argument("--rubric", default=str(_default_rubric()))
    ap.add_argument("--out")
    a = ap.parse_args()
    doc = _load_json(a.doc_json, "材料 doc.json")
    rub = _load_json(a.rubric, "评分卡")
    checks = run(doc, a.award, rub)
    res = {"source": doc["source"], "name": doc["name"], "award": a.award,
           "summary": summarize(checks), "checks": checks}
    payload = json.dumps(res, ensure_ascii=False, indent=1)
    if a.out:
        Path(a.out).write_text(payload, encoding="utf-8")
    print(json.dumps({"name": res["name"], "award": a.award, "summary": res["summary"]}, ensure_ascii=False))
    for c in checks:
        mark = {"pass": "OK  ", "warn": "WARN", "fail": "FAIL", "na": "--  ", "error": "ERR "}.get(c["status"], "?")
        print(" %s %-4s %-18s %s" % (mark, c["id"], c["name"], c["detail"]))
    return 0

if __name__ == "__main__":
    raise SystemExit(main())