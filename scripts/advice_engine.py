#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""修改建议的唯一来源。

Excel 汇总表和 Word 意见书都从这里取条目，因此两边**天然一一对应**：
同样的编号、同样的顺序、同样的条数。任何一边想多写或少写，都做不到。

每条建议是三段式，面向申报人，能直接复制：
  original  原文是什么（定位到表几行几 / 引出原句）
  how       怎么改（做什么动作）
  revised   改之后的句子（可直接粘贴进材料）

用法：
  from advice_engine import build_items
  items = build_items(score_row)   # 返回 list[dict]
"""
from __future__ import annotations
import re

# ---------------------------------------------------------------------------
# 机检编号 -> (归属维度, 问题简述, 怎么改, 改后句子模板)
#
# 改后句子用 {orig} 引用原文片段；没有原文可引用时给通用的可抄句式。
# ---------------------------------------------------------------------------
# 机检项的完整定义已迁到 check_catalog.py。
# 下面三张表是它的**视图**，保留旧名字供现有调用点使用；
# 新增机检项只需在 check_catalog 里注册一次。
from check_catalog import (CATALOG as _CATALOG, get as _cat_get,
                           dims_of as _cat_dims, hints_of as _cat_hints,
                           points_of as _cat_points)


RULES_LEGACY: dict[str, dict] = {
    "C02": {
        "dims": {"value"},
        "title": "效益里分不清哪些已实现、哪些还在规划",
        "how": "把效益表拆成「实际值【已实现】」与「远期目标【规划/在研】」两栏，" 
               "已兑现的填实际值并附台账截图，未兑现的只填远期目标。",
        "revised": "本项目已实现效益 XX 万元（附《XX 台账》及运行记录截图），"
                   "预计后续可再实现效益 YY 万元（规划中，尚未发生）。",
    },
    "C03": {
        "dims": {"value"},
        "title": "有实际值但没有基准值，变化率算不出来",
        "how": "每项效益补上改造前的基准值、对比周期与计算口径，使变化率可复算。",
        "revised": "改造前电单耗 0.1069 kWh/吨（2024 年 1—12 月计量数据），"
                   "改造后 0.0784 kWh/吨（2025 年同期），降幅 26.7%。",
    },
    "C04": {
        "dims": {"value"},
        "title": "已实现的效益没有佐证材料",
        "how": "每个已实现效益补佐证材料：台账截图 / 运行日志 / 核算表，填进「佐证材料」列。",
        "revised": "上述效益已附 XX 系统导出报表（见附件 3，第 X 页）与 XX 部门确认的核算表。",
    },
    "C05": {
        "dims": {"value", "economics"},
        "title": "正文百分比与表格数字对不上",
        "how": "正文中的百分比一律由表格数字反算，偏差控制在 5 个百分点以内；"
               "或在正文注明计算方式。",
        "revised": "由表 4 数据计算：节约率 =（改造前 − 改造后）÷ 改造前 = 26.7%，与表内一致。",
    },
    "C06": {
        "dims": {"replication", "promotion"},
        "title": "没有可点数的推广清单",
        "how": "补一份推广清单：大区 / 区域 / 水厂逐个列名，正文写清覆盖几个大区、"
               "几个区域、几个项目。",
        "revised": "截至 2026 年 X 月，本成果已在 X 个大区、X 个区域、X 家项目公司推广应用"
                   "（明细见附表《推广清单》）。",
    },
    "C07": {
        "dims": {"economics"},
        "title": "没有计算步骤表，金额无法复算",
        "how": "补「计算步骤」表：每一步写清公式、代入值、结果、单位、含义。",
        "revised": "计算步骤：① 年处理水量 = 日均水量 × 365 = X 万 m³；"
                   "② 吨水电耗差 = 0.1069 − 0.0784 = 0.0285 kWh/m³；"
                   "③ 年节电 = ①×② = X 万 kWh；④ 年节约电费 = ③×电价 X 元 = X 万元。",
    },
    "C08": {
        "dims": {"value", "economics"},
        "title": "同一金额在多处出现，疑似重复计入",
        "how": "核对每个金额的出处，确认没有重复计入（例如吨水单耗与总费用各算一次）；"
               "如确属重复，删除其中之一并同步修正合计。",
        "revised": "本文所列 XX 万元为全周期合计，与表 3 的年度金额为包含关系，非重复计入。",
    },
    "C09": {
        "dims": {"_global"},
        "title": "还有未填的占位符",
        "how": "把下划线、XXXX、（文本）之类的占位全部补成实际内容；确无内容的写「无」。",
        "revised": "将「___0.384__ 万元」补为「0.384 万元」。",
    },
    "C11": {
        "dims": {"intelligence"},
        "title": "模型实验室 / 试点厂 / 90 天验证的证据不全",
        "how": "补三件套：模型实验室线上部署记录、试点厂生产调用记录、不少于 90 天的运行验证数据。",
        "revised": "模型于 2025 年 X 月在模型实验室完成部署（附部署记录），"
                   "X 月起在 X 厂投入生产调用（附调用日志），"
                   "已连续稳定运行 X 天，准确率 X%。",
    },
    "C12": {
        "dims": {"readiness"},
        "title": "没写清产品转化到什么程度",
        "how": "明确写清从方案到产品的转化程度：已产品化（有合同/订单/收入）→ 非常符合；"
               "仅有原型 → 符合；仍在方案 → 一般。",
        "revised": "本成果已形成标准化产品，2025 年签订 X 份供货合同、合同额 X 万元"
                   "（附合同及发票），具备对外交付能力。",
    },
    "C13": {
        "dims": {"iteration"},
        "title": "没有「相较上次参赛版本」的迭代说明",
        "how": "补一段迭代说明：相较上一届/上一版本的版本号、改了什么、指标提升多少。",
        "revised": "相较第七届参赛版本，本届新增「一表多点」采样与分区自决策两项能力，"
                   "仪表数量减少 X%、控制精度提升 X 个百分点。",
    },
    "C14": {
        "dims": {"novelty"},
        "title": "没有与原有办法的量化对照",
        "how": "补与原有做法的量化对照：同一场景下新旧方案的效果差值（成本 / 精度 / 时效），"
               "并说明差异来自哪个技术点。",
        "revised": "与传统人工控制相比，本方案吨水电耗降低 X%、碳源投加减少 X%，"
                   "差异主要来自「一表多点」采样带来的分区控制能力。",
    },
    "C15": {
        "dims": {"usability"},
        "title": "战略对齐或卡点描述缺一项",
        "how": "补两段：与集团战略方向的对应关系；当前卡点的具体表现（现状数据 + 卡在哪 + "
               "本方案如何解决）。",
        "revised": "本成果对应集团「十五五」数字化运营方向；当前卡点为 X 项指标依赖人工判断，"
                   "导致 X；本方案通过 X 实现 X。",
    },
    "C16": {
        "dims": {"value", "usability"},
        "title": "预计性表述偏多，实际数据不足",
        "how": "把「预计 / 有望 / 预期」尽量换成已发生的量；尚未发生的明确标注为规划值。",
        "revised": "已实现 X 万元（截至 2026 年 X 月）；预计全周期可达 Y 万元（规划值）。",
    },
    "C17": {
        "dims": {"_global"},
        "title": "缺少运行截图等图形证据",
        "how": "插入关键运行界面、台账、报告截图，并在图下加一句文字说明。",
        "revised": "图 X 为 XX 系统实时运行界面，可见 X 项指标处于 X 区间（截图时间 2026-XX-XX）。",
    },
    "C18": {
        "dims": {"replication", "promotion"},
        "title": "推广 / 测算 / 产品形态证据不齐",
        "how": "补齐推广案例佐证、经济测算过程、产品形态说明三块。",
        "revised": "推广案例见附表 1，测算过程见表 4，产品形态见附件 2 产品手册。",
    },
    "C19": {
        "dims": {"_global"},
        "title": "专属模板段落不全",
        "how": "补齐专属模板六段：实施背景、模型应用、业务价值、ROI 核算、下一步计划、量化数据。",
        "revised": "按模板补齐：①实施背景 …②模型应用 …③业务价值 …④ROI 核算 …"
                   "⑤下一步计划 …⑥量化数据 …",
    },
    "C20": {
        "dims": {"economics"},
        "title": "ROI 缺可复算的投入与收益",
        "how": "ROI 分投入与收益两段写，并配足可复算的量化数字。",
        "revised": "投入合计 X 万元（软件 X + 硬件 X + 实施 X）；年收益 X 万元；"
                   "投资回收期 = X ÷ X = X 年。",
    },
}

# 维度兜底：没有具体机检命中时，按维度给三段式建议
# RULES 以 catalog 为准：catalog 里有的全部覆盖进来，
# 这样一来 C01/C10 不再因为没有文案而被静默跳过。
RULES: dict[str, dict] = dict(RULES_LEGACY)
for _cid, _spec in _CATALOG.items():
    RULES[_cid] = {"dims": _spec["dims"], "title": _spec["title"],
                   "how": _spec["how"], "revised": _spec["revised"]}


DIM_FALLBACK: dict[str, dict] = {
    "value": {
        "title": "价值性论证不足，缺可复算的效益数字",
        "how": "把效益拆成「已实现」与「规划」两栏分别给数，每组数字都要能用公式复算。",
        "revised": "已实现效益：吨水电耗由 0.1069 降至 0.0784 kWh，年节约电费 X 万元"
                   "（计算过程见附表）；规划效益：X 万元。",
    },
    "replication": {
        "title": "推广性论证不足，缺可点数的推广事实",
        "how": "给出可点数的推广事实：已落地几个大区 / 区域 / 厂站，逐个列名，"
               "并说明可复制的原因。",
        "revised": "已在 X 个大区、X 个区域、X 家项目公司落地（明细见附表《推广清单》），"
                   "可复制性来自标准化部署包，单厂部署周期 X 天。",
    },
    "novelty": {
        "title": "创新性论证不足，缺与原有做法的量化差异",
        "how": "点明与原有做法的差别：同一场景下新旧方案的指标差，并说明差异来自哪个技术点。"
               "把下句括号里的内容换成你自己的数据即可。",
        "revised": "相较原有的（填：原做法名称）方案，本方案使（填：指标名，如吨水电耗）"
                   "由（填：改造前数值）降至（填：改造后数值），降幅（填：百分比），"
                   "差异主要来自（填：技术点，如「一表多点」采样带来的分区控制能力）。",
    },
    "intelligence": {
        "title": "智能化水平证据不足",
        "how": "补模型实验室部署记录、试点厂生产调用日志、以及不少于 90 天的运行表现。",
        "revised": "模型已在模型实验室完成部署（2025-XX），并在 X 厂生产调用，"
                   "连续稳定运行 X 天，准确率 X%。",
    },
    "promotion": {
        "title": "推广力度证据不足",
        "how": "补推广台账：按大区—区域—厂站三级列清单，标明部署时间与规模。",
        "revised": "推广台账：X 大区 X 区域 X 家，2025 年 X 月完成部署，覆盖 X 台设备。",
    },
    "economics": {
        "title": "经济效益缺计算过程",
        "how": "补一张计算步骤表：输入参数、公式、中间结果、最终金额、单位与口径各占一列。",
        "revised": "年节约费用 = 年节电量 X 万 kWh × 电价 X 元/kWh = X 万元（计算过程见表 4）。",
    },
    "readiness": {
        "title": "产品就绪度说明不足",
        "how": "写清交付形态与转化程度：是否已有合同/订单/收入，是否有标准产品文档与运维手册。",
        "revised": "已形成标准产品：X 份合同、合同额 X 万元，配套产品手册与运维手册各 1 套。",
    },
    "usability": {
        "title": "实用性论证不足",
        "how": "补战略对应关系与卡点具体表现（现状数据 + 卡在哪 + 怎么解决）。",
        "revised": "对应集团「十五五」X 方向；当前卡点为 X，本方案通过 X 解决。",
    },
}

# 纯格式 / 形式类检查：不写进修改意见。
# 理由：下划线没填、缺截图这类问题肉眼可见，写进意见是噪音，
# 会把真正需要论证补强的建议挤下去。这些由机检报告自己提示即可。
#   C09 占位符未填（下划线、XXXX）
#   C17 缺运行截图
FORMAT_ONLY = {"C09", "C17"}

# 补充材料能把该维度失分追回多少比例。
# 依据：证据等级反映论证强度，补数据能提高档次但通常回不到满分。
#   A 已有充分证据，只差业务结果 → 补齐后基本能拿满
#   B 论证方向对但不充分 → 补量化数据可追回大半
#   C 偏定性 → 补数字与佐证，追回一半左右
#   D 几乎没有证据 → 相当于重写，短期只能追回三成
RECOVER_RATE = {"A": 0.9, "B": 0.6, "C": 0.5, "D": 0.3}

NOISE = ("参赛赛道", "申报单位", "项目名称", "填报", "目录")
HEADING = re.compile(r"^\s*(\[H\d\]\s*)?(\d+(\.\d+)*\s*)?[^。；]{0,30}$")

# 各机检问题对应的原文线索词
HINTS = {
    "C02": ["效益", "实际值", "规划", "已实现"],
    "C03": ["基准", "改造前", "对比"],
    "C04": ["佐证", "台账", "截图"],
    "C06": ["推广", "大区", "区域", "厂"],
    "C07": ["测算", "计算", "公式", "节约", "万元"],
    "C08": ["万元", "节约", "费用"],
    "C11": ["模型实验室", "试点", "上线", "运行"],
    "C12": ["产品", "交付", "合同", "原型"],
    "C13": ["迭代", "升级", "版本"],
    "C14": ["对比", "优于", "原有", "传统"],
    "C15": ["战略", "卡点", "痛点"],
    "C17": ["截图", "图"],
    "C20": ["投入", "收益", "回收期"],
}


def _candidates(plain: str):
    """原文里可引用的句子，已滤掉标题行与封面字段。"""
    for s in re.split(r"[。；\n]", plain or ""):
        s = s.strip()
        if not (12 <= len(s) <= 240):
            continue
        if any(n in s for n in NOISE):
            continue
        if HEADING.match(s) and "。" not in s and len(s) < 40:
            continue
        yield s


def find_original(plain: str, keywords: list[str], max_len: int = 76,
                  fallback_terms: list[str] | None = None) -> str:
    """从原文里找一句能说明问题的句子，作为「原文是什么」。

    三级匹配，宁可少一句也不要给一句废话：
      1. 命中 >= 2 个线索词        —— 最能说明问题
      2. 命中 >= 1 个线索词 或 维度名 —— 次之
      3. 命中任意 fallback 词      —— 兜底

    跳过标题行、封面字段，否则会挑出「参赛赛道：成果应用赛道 金火炬」这种
    对问题毫无说明力的句子。
    """
    if not plain:
        return ""
    bands = [
        lambda s: sum(1 for k in keywords if k in s) >= 2,
        lambda s: sum(1 for k in keywords if k in s) >= 1,
        lambda s: bool(fallback_terms) and any(t in s for t in (fallback_terms or [])),
    ]
    for hit in bands:
        for s in _candidates(plain):
            if hit(s):
                return s if len(s) <= max_len else s[: max_len - 1] + "…"
    # 完全没有匹配：给该材料最长的一段（通常是最有内容的正文），并注明是概括
    best = ""
    for s in _candidates(plain):
        if len(s) > len(best):
            best = s
    if best:
        cut = best if len(best) <= max_len else best[: max_len - 1] + "…"
        return "（未找到直接对应的原句，以下为材料中的相关表述）" + cut
    return ""


def build_items(row: dict) -> list[dict]:
    """生成该材料的全部修改建议条目（三段式）。

    Excel 与 Word 都调用本函数，所以两边条目、顺序、编号完全一致。

    编号规则：M01、M02…… 按「可追回分数」从大到小排列。
    返回字段：
        no / dim / dim_id / check / title / original / how / revised / loss
    """
    checks = row.get("_checks") or {}
    plain = row.get("_plain") or ""
    caps = row.get("caps_and_gates") or []
    deds = row.get("deductions") or []

    # 维度 -> 封顶这条是哪个机检造成的
    cap_by_dim = {}
    for g in caps:
        if g.get("type") == "cap" and g.get("check"):
            cap_by_dim.setdefault(g.get("dim"), g["check"])

    items: list[dict] = []
    for dim in row.get("dimensions") or []:
        did = dim["id"]
        loss = round(dim["max"] - dim["score"], 1)

        # 收集该维度命中的检查编号（按优先级：封顶 > 本维度机检 > 归属扣分）
        cids: list[str] = []
        if cap_by_dim.get(did):
            cids.append(cap_by_dim[did])
        for cid, c in checks.items():
            if c.get("status") not in ("warn", "fail") or cid in FORMAT_ONLY:
                continue
            rule = RULES.get(cid)
            if not rule:
                continue
            if "_global" in rule["dims"]:
                continue
            if did not in rule["dims"]:
                continue
            cdims = c.get("dims") or []
            if cdims and did not in cdims:
                continue
            if cid not in cids:
                cids.append(cid)
        for d in deds:
            cid = d.get("check")
            if cid and cid in RULES and did in RULES[cid]["dims"] and cid not in cids:
                cids.append(cid)

        # 该维度确实失分，但没有任何具体机检 -> 用维度兜底
        if not cids and loss > 0.01 and did in DIM_FALLBACK:
            fb = DIM_FALLBACK[did]
            items.append({
                "dim": dim["name"], "dim_id": did, "check": "",
                "title": fb["title"],
                "original": _cap_or_grade(row, dim, plain, did),
                "how": fb["how"], "revised": fb["revised"], "loss": loss,
            })
            continue

        for cid in cids:
            rule = RULES[cid]
            ev = (checks.get(cid) or {}).get("evidence") or []
            orig = ""
            if ev:
                # 机检给的定位可能是「表3 行3」这种简短坐标——单独看说明不了问题，
                # 补一句该坐标的上下文，或与被引原句一起给出。
                pos = [str(x) for x in ev if re.match(r"^表\s*\d+", str(x))]
                quote = [str(x) for x in ev if not re.match(r"^表\s*\d+", str(x))]
                parts = []
                if quote:
                    parts.append(quote[0])
                if pos:
                    parts.append("位置：" + "、".join(pos[:3]))
                orig = "；".join(parts) if parts else "；".join(str(x) for x in ev[:2])
            else:
                orig = find_original(plain, HINTS.get(cid, []))
            if not orig:
                # 把维度名与检查名当线索再找一次，仍找不到才退回证据等级说明
                kw = HINTS.get(cid, []) + [dim["name"]]
                orig = find_original(plain, kw)
            if len(orig) < 12:
                # 定位太短（如「表3 行3」）时，用机检 detail 补充说明问题在哪
                det = (checks.get(cid) or {}).get("detail") or ""
                if det:
                    orig = (orig + "　—　" + det) if orig else det
            if not orig:
                orig = _cap_or_grade(row, dim, plain, did)
            items.append({
                "dim": dim["name"], "dim_id": did, "check": cid,
                "title": rule["title"],
                "original": orig,
                "how": rule["how"], "revised": rule["revised"], "loss": loss,
            })

    # 全局项（截图、模板等）单独成条，挂到"材料整体"；格式类跳过
    for cid, rule in RULES.items():
        if "_global" not in rule["dims"] or cid in FORMAT_ONLY:
            continue
        c = checks.get(cid)
        if not c or c.get("status") not in ("warn", "fail"):
            continue
        ev = c.get("evidence") or []
        items.append({
            "dim": "材料整体", "dim_id": "_global", "check": cid,
            "title": rule["title"],
            "original": "；".join(str(x) for x in ev[:2]) or (c.get("detail") or ""),
            "how": rule["how"], "revised": rule["revised"],
            "loss": _global_points(cid),
        })

    # 按可追回分数降序；同分保持维度顺序（稳定排序）
    items.sort(key=lambda x: -x["loss"])
    for i, it in enumerate(items, 1):
        it["no"] = "M%02d" % i

    # ---- 分派「预计改完提升」 ----
    # 两个必须处理的现实：
    #   1. 同一维度的失分只能计一次——经济效益 20 分同时被 C07、C08 命中时，
    #      两条各写 20 分相加就是 40 分，虚报一倍。记在首条，其余注明合并计入。
    #   2. 补齐材料拿不回该维度的全部分数——证据等级 B/C 意味着论证强度不足，
    #      补数据能让它往上走一档，不等于直接满分。按等级给定回收系数。
    dim_grade = {d["id"]: (d.get("verified_grade") or "") for d in (row.get("dimensions") or [])}
    seen: set[str] = set()
    for it in items:
        d = it["dim_id"]
        if d == "_global":
            it["gain"] = 0.0
            it["gain_note"] = "材料整体项，不计入维度分数"
            continue
        if d in seen:
            it["gain"] = 0.0
            it["gain_note"] = "与同维度首条合并计入"
            continue
        seen.add(d)
        rate = RECOVER_RATE.get(dim_grade.get(d, ""), 0.5)
        it["gain"] = round(it["loss"] * rate, 1)
        it["gain_note"] = "可追回该维度失分的 %d%%（证据等级 %s）" % (
            int(rate * 100), dim_grade.get(d) or "未知")
    return items


def total_gain(items: list[dict], ai_score: float) -> float:
    """把列出的建议全改完，总分能提升多少（不超过 100 分）。"""
    g = sum(it.get("gain") or 0.0 for it in items)
    return round(min(g, max(0.0, 100.0 - float(ai_score or 0))), 1)


# 每个维度"该有什么证据"，用于证据不足时给出可查的缺口
DIM_EVIDENCE = {
    "value": "效益金额、基准值、计算口径、佐证材料",
    "replication": "推广清单（大区/区域/厂站）、部署时间、可复制原因",
    "novelty": "与原有做法的同场景指标对照、技术点差异说明",
    "intelligence": "模型实验室部署记录、试点厂调用日志、≥90 天运行数据",
    "promotion": "推广台账、覆盖规模、单厂部署周期",
    "economics": "计算步骤表（公式/代入值/中间结果）、投入与收益明细",
    "readiness": "产品形态说明、合同/订单/收入、产品与运维手册",
    "usability": "战略对应关系、卡点现状数据、解决方案匹配",
    "iteration": "版本号、相较上届改了什么、指标提升多少",
}


def _cap_or_grade(row, dim, plain, did) -> str:
    """没有具体原文时，说明「缺什么证据」——比空泛说"论证不足"更有指向。"""
    if dim.get("cap_reason"):
        return dim["cap_reason"]
    g = dim.get("verified_grade") or ""
    need = DIM_EVIDENCE.get(did, "量化数据与佐证材料")
    return {
        "A": "该维度业务条件未完全达成，仍失分；需补充：%s" % need,
        "B": "论证方向对但证据强度不足（等级 B），缺可核验的量化数据；需补充：%s" % need,
        "C": "论证偏定性（等级 C），缺具体数字与佐证；需补充：%s" % need,
        "D": "几乎没有可引用的证据（等级 D）；需补充：%s" % need,
    }.get(g, "该维度证据不足；需补充：%s" % need)


def _global_points(cid: str) -> float:
    """全局项折算成"可追回分"，仅用于排序，不代表真实分值。"""
    return {"C09": 2.0, "C17": 2.0, "C01": 3.0, "C19": 1.5}.get(cid, 1.0)