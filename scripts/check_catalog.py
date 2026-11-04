#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""CheckCatalog —— 每个机检项的完整定义，只写一遍。

## 为什么有这个模块

advice_engine 里有五张按 check id 索引的平行表：

    RULES          18 条：dims / title / how / revised
    HINTS          13 条：check -> 定位线索词
    DIM_FALLBACK    8 个维度
    DIM_EVIDENCE    9 个维度
    _global_points  4 条：check -> 可追回分数

而 checks.py 有 20 个检查。三张按同一个 id 索引的表长度互不相等，
代码用 .get() 兜住差异——兜住的意思是**静默失效**。

## 这个模块发现了一个真 bug

对齐检查发现 C01 与 C10 在 CHECKS 里有，在 RULES 里没有：

    19 份材料里 C01 warn/fail 2 份、C10 warn/fail 1 份
    这 3 份**完全丢失了建议**
    而 _global_points 给 C01 打了 3.0（全表最高权重）

原因是 build_items 里的 `if not rule: continue`——没有文案就跳过，
不报错、不提示。三张表按同一个 key 索引却必须手工保持同步，
而没有任何测试断言它们对齐。

## 设计

一张表取代五张。加一个机检项只改一处；缺字段当场报错而不是 .get() 兜住。
"""
from __future__ import annotations

# 每个机检项的完整定义。
#   dims      归属维度（空集合表示不影响维度，只提示）
#   title     问题的一句话描述
#   how       怎么改
#   revised   改完的示范句
#   hints     回到原文找定位的线索词
#   points    若该检查命中，最多可追回多少分
#
# 新增机检项时，这里的条目是必须的——tests 会检查每个 CHECKS 里的 id
# 都能在这张表里找到。
CATALOG: dict[str, dict] = {}


def register(cid: str, *, dims, title: str, how: str, revised: str,
             hints=(), points: float = 1.0) -> None:
    """登记一个机检项的完整定义。"""
    CATALOG[cid] = {
        "dims": set(dims),
        "title": title,
        "how": how,
        "revised": revised,
        "hints": list(hints),
        "points": points,
    }


def get(cid: str):
    """取一条定义；不存在返回 None（调用方应显式处理，不要静默跳过）。"""
    return CATALOG.get(cid)


def require(cid: str) -> dict:
    """取一条定义；不存在则报错。

    新代码应该用这个而不是 get()——静默跳过正是 C01/C10 丢掉建议的原因。
    """
    spec = CATALOG.get(cid)
    if spec is None:
        raise KeyError(
            "机检项 %s 在 CHECKS 里有，但 CheckCatalog 里没有定义。\n"
            "  没有 title/how/revised 就无法生成建议——会被静默丢弃。\n"
            "  请在 scripts/check_catalog.py 里补上它的定义。" % cid)
    return spec


def dims_of(cid: str) -> set:
    spec = CATALOG.get(cid)
    return spec["dims"] if spec else set()


def hints_of(cid: str) -> list:
    spec = CATALOG.get(cid)
    return spec["hints"] if spec else []


def points_of(cid: str) -> float:
    spec = CATALOG.get(cid)
    return spec["points"] if spec else 1.0


register(
    'C02',
    dims=['value'],
    title='效益里分不清哪些已实现、哪些还在规划',
    how='把效益表拆成「实际值【已实现】」与「远期目标【规划/在研】」两栏，已兑现的填实际值并附台账截图，未兑现的只填远期目标。',
    revised='本项目已实现效益 XX 万元（附《XX 台账》及运行记录截图），预计后续可再实现效益 YY 万元（规划中，尚未发生）。',
    hints=['效益', '实际值', '规划', '已实现'],
    points=1.0,
)
register(
    'C03',
    dims=['value'],
    title='有实际值但没有基准值，变化率算不出来',
    how='每项效益补上改造前的基准值、对比周期与计算口径，使变化率可复算。',
    revised='改造前电单耗 0.1069 kWh/吨（2024 年 1—12 月计量数据），改造后 0.0784 kWh/吨（2025 年同期），降幅 26.7%。',
    hints=['基准', '改造前', '对比'],
    points=1.0,
)
register(
    'C04',
    dims=['value'],
    title='已实现的效益没有佐证材料',
    how='每个已实现效益补佐证材料：台账截图 / 运行日志 / 核算表，填进「佐证材料」列。',
    revised='上述效益已附 XX 系统导出报表（见附件 3，第 X 页）与 XX 部门确认的核算表。',
    hints=['佐证', '台账', '截图'],
    points=1.0,
)
register(
    'C05',
    dims=['economics', 'value'],
    title='正文百分比与表格数字对不上',
    how='正文中的百分比一律由表格数字反算，偏差控制在 5 个百分点以内；或在正文注明计算方式。',
    revised='由表 4 数据计算：节约率 =（改造前 − 改造后）÷ 改造前 = 26.7%，与表内一致。',
    hints=[],
    points=1.0,
)
register(
    'C06',
    dims=['promotion', 'replication'],
    title='没有可点数的推广清单',
    how='补一份推广清单：大区 / 区域 / 水厂逐个列名，正文写清覆盖几个大区、几个区域、几个项目。',
    revised='截至 2026 年 X 月，本成果已在 X 个大区、X 个区域、X 家项目公司推广应用（明细见附表《推广清单》）。',
    hints=['推广', '大区', '区域', '厂'],
    points=1.0,
)
register(
    'C07',
    dims=['economics'],
    title='没有计算步骤表，金额无法复算',
    how='补「计算步骤」表：每一步写清公式、代入值、结果、单位、含义。',
    revised='计算步骤：① 年处理水量 = 日均水量 × 365 = X 万 m³；② 吨水电耗差 = 0.1069 − 0.0784 = 0.0285 kWh/m³；③ 年节电 = ①×② = X 万 kWh；④ 年节约电费 = ③×电价 X 元 = X 万元。',
    hints=['测算', '计算', '公式', '节约', '万元'],
    points=1.0,
)
register(
    'C08',
    dims=['economics', 'value'],
    title='同一金额在多处出现，疑似重复计入',
    how='核对每个金额的出处，确认没有重复计入（例如吨水单耗与总费用各算一次）；如确属重复，删除其中之一并同步修正合计。',
    revised='本文所列 XX 万元为全周期合计，与表 3 的年度金额为包含关系，非重复计入。',
    hints=['万元', '节约', '费用'],
    points=1.0,
)
register(
    'C09',
    dims=['_global'],
    title='还有未填的占位符',
    how='把下划线、XXXX、（文本）之类的占位全部补成实际内容；确无内容的写「无」。',
    revised='将「___0.384__ 万元」补为「0.384 万元」。',
    hints=[],
    points=2.0,
)
register(
    'C11',
    dims=['intelligence'],
    title='模型实验室 / 试点厂 / 90 天验证的证据不全',
    how='补三件套：模型实验室线上部署记录、试点厂生产调用记录、不少于 90 天的运行验证数据。',
    revised='模型于 2025 年 X 月在模型实验室完成部署（附部署记录），X 月起在 X 厂投入生产调用（附调用日志），已连续稳定运行 X 天，准确率 X%。',
    hints=['模型实验室', '试点', '上线', '运行'],
    points=1.0,
)
register(
    'C12',
    dims=['readiness'],
    title='没写清产品转化到什么程度',
    how='明确写清从方案到产品的转化程度：已产品化（有合同/订单/收入）→ 非常符合；仅有原型 → 符合；仍在方案 → 一般。',
    revised='本成果已形成标准化产品，2025 年签订 X 份供货合同、合同额 X 万元（附合同及发票），具备对外交付能力。',
    hints=['产品', '交付', '合同', '原型'],
    points=1.0,
)
register(
    'C13',
    dims=['iteration'],
    title='没有「相较上次参赛版本」的迭代说明',
    how='补一段迭代说明：相较上一届/上一版本的版本号、改了什么、指标提升多少。',
    revised='相较第七届参赛版本，本届新增「一表多点」采样与分区自决策两项能力，仪表数量减少 X%、控制精度提升 X 个百分点。',
    hints=['迭代', '升级', '版本'],
    points=1.0,
)
register(
    'C14',
    dims=['novelty'],
    title='没有与原有办法的量化对照',
    how='补与原有做法的量化对照：同一场景下新旧方案的效果差值（成本 / 精度 / 时效），并说明差异来自哪个技术点。',
    revised='与传统人工控制相比，本方案吨水电耗降低 X%、碳源投加减少 X%，差异主要来自「一表多点」采样带来的分区控制能力。',
    hints=['对比', '优于', '原有', '传统'],
    points=1.0,
)
register(
    'C15',
    dims=['usability'],
    title='战略对齐或卡点描述缺一项',
    how='补两段：与集团战略方向的对应关系；当前卡点的具体表现（现状数据 + 卡在哪 + 本方案如何解决）。',
    revised='本成果对应集团「十五五」数字化运营方向；当前卡点为 X 项指标依赖人工判断，导致 X；本方案通过 X 实现 X。',
    hints=['战略', '卡点', '痛点'],
    points=1.0,
)
register(
    'C16',
    dims=['usability', 'value'],
    title='预计性表述偏多，实际数据不足',
    how='把「预计 / 有望 / 预期」尽量换成已发生的量；尚未发生的明确标注为规划值。',
    revised='已实现 X 万元（截至 2026 年 X 月）；预计全周期可达 Y 万元（规划值）。',
    hints=[],
    points=1.0,
)
register(
    'C17',
    dims=['_global'],
    title='缺少运行截图等图形证据',
    how='插入关键运行界面、台账、报告截图，并在图下加一句文字说明。',
    revised='图 X 为 XX 系统实时运行界面，可见 X 项指标处于 X 区间（截图时间 2026-XX-XX）。',
    hints=['截图', '图'],
    points=2.0,
)
register(
    'C18',
    dims=['promotion', 'replication'],
    title='推广 / 测算 / 产品形态证据不齐',
    how='补齐推广案例佐证、经济测算过程、产品形态说明三块。',
    revised='推广案例见附表 1，测算过程见表 4，产品形态见附件 2 产品手册。',
    hints=[],
    points=1.0,
)
register(
    'C19',
    dims=['_global'],
    title='专属模板段落不全',
    how='补齐专属模板六段：实施背景、模型应用、业务价值、ROI 核算、下一步计划、量化数据。',
    revised='按模板补齐：①实施背景 …②模型应用 …③业务价值 …④ROI 核算 …⑤下一步计划 …⑥量化数据 …',
    hints=[],
    points=1.5,
)
register(
    'C20',
    dims=['economics'],
    title='ROI 缺可复算的投入与收益',
    how='ROI 分投入与收益两段写，并配足可复算的量化数字。',
    revised='投入合计 X 万元（软件 X + 硬件 X + 实施 X）；年收益 X 万元；投资回收期 = X ÷ X = X 年。',
    hints=['投入', '收益', '回收期'],
    points=1.0,
)
# --- 以下两条原本只在 CHECKS 里，RULES 没有文案 -> 命中时建议被静默丢弃 ---
# C01 是「必备表格」，_global_points 原本给它 3.0（全表最高），
# 但因为没有 title/how/revised，那 3.0 分从未真正用过。
register(
    "C01",
    dims={"value", "replication", "economics", "readiness"},
    title="缺必备表格，评审无法核验该维度的关键信息",
    how="按申报模板补齐必备表格（效益测算表、推广台账、对比数据等），"
        "每张表都要有表头、单位、数据来源；没有的数据留空并注明「暂无」，不要删表。",
    revised="表X-X <表名>（数据来源：<系统/台账名称>，统计周期：<起止>）",
    hints=["详见", "见附表", "如表"],
    points=3.0,
)

register(
    "C10",
    dims={"value", "novelty", "replication", "intelligence",
          "usability", "economics", "readiness", "promotion"},
    title="材料篇幅不足，关键论证展开不清",
    how="按主报告框架补足每节内容：实施背景写清痛点与量化现状，"
        "技术方案写清原理与关键参数，效益写清测算口径与佐证。",
    revised="（在各章节标题下补足 2—4 段实质内容，避免只有标题和一句话结论。）",
    hints=["实施背景", "技术方案", "效益"],
    points=1.0,
)
def missing_from(check_ids) -> list:
    """CHECKS 里有、CATALOG 里没有的 id。供自测与启动自检使用。"""
    return sorted(set(check_ids) - set(CATALOG))


def orphaned_in(check_ids) -> list:
    """CATALOG 里有、CHECKS 里没有的 id（定义指向不存在的机检）。"""
    return sorted(set(CATALOG) - set(check_ids))