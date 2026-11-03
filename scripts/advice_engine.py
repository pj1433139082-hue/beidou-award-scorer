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
RULES: dict[str, dict] = {
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
        "how": "点明与原有做法的差别：同一场景下新旧方案的指标差，并说明差异来自哪个技术点。",
        "revised": "相较原有 X 方案，本方案在 X 指标上提升 X%，差异来自 X 技术点。",
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


def find_original(plain: str, keywords: list[str], max_len: int = 76) -> str:
    """从原文里找一句能说明问题的句子，作为「原文是什么」。

    要求命中 >= 2 个线索词，跳过标题行、封面字段，否则会挑出
    「参赛赛道：成果应用赛道 金火炬」这种对问题毫无说明力的句子。
    """
    if not plain:
        return ""
    for s in re.split(r"[。；\n]", plain):
        s = s.strip()
        if not (12 <= len(s) <= 200):
            continue
        if any(n in s for n in NOISE):
            continue
        if HEADING.match(s) and "。" not in s and len(s) < 40:
            continue
        if sum(1 for k in keywords if k in s) >= 2:
            return s if len(s) <= max_len else s[: max_len - 1] + "…"
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
            if c.get("status") not in ("warn", "fail"):
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
                # 机检给了定位（表几行几 / 原句），直接用
                orig = "；".join(str(x) for x in ev[:2])
            else:
                orig = find_original(plain, HINTS.get(cid, []))
            if not orig:
                orig = _cap_or_grade(row, dim, plain, did)
            items.append({
                "dim": dim["name"], "dim_id": did, "check": cid,
                "title": rule["title"],
                "original": orig,
                "how": rule["how"], "revised": rule["revised"], "loss": loss,
            })

    # 全局项（占位符、截图等）单独成条，挂到"材料整体"
    for cid, rule in RULES.items():
        if "_global" not in rule["dims"]:
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
    return items


def _cap_or_grade(row, dim, plain, did) -> str:
    """没有具体原文时，用封顶原因或证据等级说明「原文哪里不好」。"""
    if dim.get("cap_reason"):
        return dim["cap_reason"]
    g = dim.get("verified_grade") or ""
    return {
        "A": "证据等级 A，但业务条件未完全达成，该项仍失分",
        "B": "证据等级 B：有论证但不够充分，缺少可核验的量化数据",
        "C": "证据等级 C：论证薄弱，多为定性描述，缺具体数字与佐证",
        "D": "证据等级 D：几乎没有可引用的证据",
    }.get(g, "该维度论证不足，未定位到具体原文问题")


def _global_points(cid: str) -> float:
    """全局项折算成"可追回分"，仅用于排序，不代表真实分值。"""
    return {"C09": 2.0, "C17": 2.0, "C01": 3.0, "C19": 1.5}.get(cid, 1.0)
