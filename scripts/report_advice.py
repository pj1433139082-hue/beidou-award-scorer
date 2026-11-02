#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""为每份申报材料生成详细改进意见（Markdown）。

输入：含 *.score.json 与 *.checks.json 的目录
输出：每份一个 .意见.md，外加一份总览

用法：
  python scripts/report_advice.py out/ --out 意见/
"""
from __future__ import annotations
import argparse, json, re
from pathlib import Path

# ---- 机检编号 → 具体改法（面向申报人，说清补什么、补了能加多少） ----
FIX = {
    "C01": ("结构表完整性", "补齐必备表格：基本信息表、基线与效果对照表、验证清单。评委按这三张表定位证据，缺表就等于没有证据。"),
    "C02": ("已实现/在研分离", "效益表里「实际值【已实现】」与「远期目标【规划/在研】」必须分开填。已兑现的填实际值并附台账截图；没兑现的只填远期目标。两者混填，价值性最高只能拿 40%。"),
    "C03": ("基线可算性", "每项效益补上改造前的基准值和对比周期。没有基线，变化率无法复算，等于没有说服力。"),
    "C04": ("效益佐证覆盖率", "每个已实现效益都要配佐证材料（台账截图 / 运行日志 / 核算表），填进「佐证材料（附截图）」列。覆盖率低于 80% 价值性封顶 70%。"),
    "C05": ("数字自洽性", "正文里的百分比要和表格数字对得上（偏差不超过 5 个百分点），否则会被判为数字不自洽。"),
    "C06": ("推广清单", "补一份可点数的推广清单：大区 / 区域 / 水厂逐个列出，正文写清覆盖几个大区、几个区域、几个项目。只写「已广泛推广」不算证据。"),
    "C07": ("测算过程", "补「计算步骤」表：每步写清公式、结果、单位、含义。只给一个总金额无法复算，经济效益拿不到高分。"),
    "C08": ("重复计入", "同一金额若在多处出现，确认没有重复计入（例如吨水单耗与总费用各算一次）。重复计入会被直接扣分。"),
    "C09": ("占位符残留", "把没填的占位符（下划线、XXXX、（文本））全部补实，未填项会被判为材料未完成并扣分。"),
    "C11": ("模型实验室/生产验证", "补齐三件套：模型实验室线上部署记录、试点厂生产调用记录、不少于 90 天的运行验证数据。三者齐备才能拿满智能化水平分。"),
    "C12": ("产品就绪度", "写清从方案到产品的转化程度：已产品化（有合同/订单/收入）→ 非常符合；仅有原型 → 符合；仍在方案阶段 → 一般。附交付形态说明。"),
    "C13": ("成果迭代性", "补一段「相较上次参赛版本」的迭代说明：改了什么、指标提升多少。缺这段会被判为否决项待裁定。"),
    "C14": ("创新性对照", "补与原有办法的量化对照：同一场景下新旧方案的效果差值（成本、精度、时效），并说明差异来自哪个创新点。"),
    "C15": ("实用性", "补集团战略对齐说明（对应十五五 / 集团战略方向）与当前卡点的具体描述，说明方案如何直接解决该卡点。"),
    "C16": ("包装风险", "把「预计 / 有望 / 预期」换成实际发生的量：量化成效要能落到具体金额或比例。"),
    "C17": ("截图证据", "插入关键运行界面、台账、报告截图并加文字说明；只放图不说明不构成证据。"),
}

LEVEL = {"A": "证据充分", "B": "证据部分成立", "C": "证据薄弱", "D": "基本无证据"}

# 每条机检真正属于哪个维度。跨维度串建议会让意见显得错乱
# （例如把"金额重复计入"挂在"推广力度"上）。
BELONGS = {
    "C01": {"_global"},                 # 结构表：全局，不挂具体维度
    "C02": {"value"}, "C03": {"value"}, "C04": {"value"}, "C05": {"value", "economics"},
    "C06": {"replication", "promotion"},
    "C07": {"economics", "promotion"},
    "C08": {"value", "economics"},
    "C09": {"_global"},
    "C11": {"intelligence"},
    "C12": {"readiness"},
    "C13": {"iteration"},
    "C14": {"novelty"},
    "C15": {"usability"},
    "C16": {"value", "usability"},
    "C17": {"_global"},
}


def belongs(cid: str, dim_id: str) -> bool:
    """该机检是否应该挂在这个维度下。

    全局项（结构表 / 占位符 / 截图）不挂任何维度——它们是材料完成度问题，
    挂上去会让每个维度的建议都变成同一句话。
    """
    s = BELONGS.get(cid)
    if s is None:
        return True
    if "_global" in s:
        return False
    return dim_id in s


# 维度专属兜底建议：机检没点名时，按维度语义给方向
DIM_FALLBACK = {
    "value": "把效益拆成「已实现」与「规划」两栏分别给数：已实现的配台账或运行数据截图，规划中的注明预计时点。每组数字都要能用一个公式复算出来。",
    "replication": "给出可点数的推广事实：已落地几个大区/区域/厂站，逐个列名；再补一句可复制的原因（标准化程度、部署周期、依赖条件）。",
    "novelty": "点明与原有做法的差别在哪：同一场景下新旧方案的指标差（成本/精度/时效），并说明差异来自哪个技术点。",
    "intelligence": "补模型侧的证据链：模型实验室的线上部署与评估记录、试点厂的生产调用日志、以及不少于 90 天的运行表现数据。",
    "promotion": "补推广台账：按大区—区域—厂站三级列出落地清单，标明部署时间与规模，让覆盖范围可点数、可核对。",
    "economics": "补一张计算步骤表：输入参数、公式、中间结果、最终金额、单位与口径各占一列。只给总额无法复算，这一项就拿不到高分。",
    "readiness": "写清交付形态与转化程度：是否已有合同/订单/收入，是否有标准产品文档与运维手册，从试点到规模交付还差什么。",
    "usability": "补两段：一是与集团/十五五战略的对应关系，二是当前业务卡点的具体表现（现状数据 + 卡在哪里 + 本方案如何解决）。",
}


def short(t, n=120):
    t = re.sub(r"\s+", " ", str(t or "")).strip()
    return t if len(t) <= n else t[: n - 1] + "…"


def load(d: Path):
    out = []
    for sf in sorted(d.glob("*.score.json")):
        s = json.loads(sf.read_text(encoding="utf-8"))
        stem = sf.name.replace(".score.json", "")
        cf = d / (stem + ".checks.json")
        c = json.loads(cf.read_text(encoding="utf-8")) if cf.is_file() else {"checks": []}
        s["_checks"] = {x["id"]: x for x in c.get("checks", [])}
        s["_stem"] = stem
        out.append(s)
    return out


def advice_for(dim, score_row):
    """该维度的具体改法：优先用造成封顶/扣分的机检编号。"""
    cids = []
    # 封顶原因优先（封顶扣的分远多于固定扣分）
    for g in score_row.get("caps_and_gates") or []:
        if g.get("dim") == dim["id"] and g.get("check"):
            cids.append(g["check"])
    # 其次：该维度自身命中的机检
    for cid, c in score_row["_checks"].items():
        if (c.get("status") in ("warn", "fail")
                and dim["id"] in (c.get("dims") or [])
                and belongs(cid, dim["id"])):
            cids.append(cid)
    # 最后：固定扣分（只挂到它真正归属的维度）
    for dd in score_row.get("deductions") or []:
        cid = dd.get("check")
        if cid and belongs(cid, dim["id"]):
            cids.append(cid)
    seen, out = set(), []
    for cid in cids:
        if cid in seen or cid not in FIX:
            continue
        seen.add(cid)
        out.append((cid,) + FIX[cid])
    return out


def render_one(s) -> str:
    L = []
    A = L.append
    dims = s["dimensions"]
    total = s["ai_score"]

    A("# %s" % s["name"])
    A("")
    A("| | |")
    A("|---|---|")
    A("| 奖项 | %s |" % s["award"])
    A("| AI 评审分 | **%s / 100** |" % total)
    A("| 奖项区间 | %s |" % s["band"])
    A("| 置信度 | %s |" % s["confidence"])
    if s.get("veto_pending"):
        A("| 否决项 | **待裁定**（%s） |" % short(s.get("veto_note"), 60))
    A("")

    # 一句话结论
    by_ratio = sorted(dims, key=lambda d: d["score"] / d["max"] if d["max"] else 1)
    worst = by_ratio[:2]
    best = by_ratio[-1]
    A("**一句话**：最强项是%s（%s/%s）；最需要补的是%s。"
      % (best["name"], best["score"], best["max"],
         "、".join("%s（%s/%s）" % (d["name"], d["score"], d["max"]) for d in worst)))
    A("")

    # 维度逐项
    A("## 各维度得分与意见")
    A("")
    for d in dims:
        ratio = d["score"] / d["max"] if d["max"] else 0
        A("### %s　%s / %s（%d%%）" % (d["name"], d["score"], d["max"], round(ratio * 100)))
        A("")
        A("- 证据等级：**%s**（%s）——这一项说的是「引文能否在原文中逐字核验」，"
          "不等于业务达成度" % (d.get("verified_grade", "—"),
                          LEVEL.get(d.get("verified_grade"), "未知")))
        if d.get("claimed_grade") and d.get("claimed_grade") != d.get("verified_grade"):
            A("- ⚠ 等级被下调：%s → %s（引文未全部在原文中核验到）"
              % (d.get("claimed_grade"), d.get("verified_grade")))
        if d.get("cap_reason"):
            cap_pct = round((d.get("cap") or 1.0) * 100)
            if d.get("verified_grade") in ("A", "B") and cap_pct < 100:
                A("- **封顶原因**：%s" % d["cap_reason"])
                A("  ⚠ 注意：等级%s只说明引文找得到，但该项业务条件没达成，"
                  "所以分数被压到 %d%%。" % (d.get("verified_grade"), cap_pct))
            else:
                A("- **封顶原因**：%s（封顶至 %d%%）" % (d["cap_reason"], cap_pct))
        q = d.get("quotes_verified") or []
        if q:
            A("- 已核验原文引文：")
            for x in q[:3]:
                A("  > %s" % short(x, 100))
        rj = d.get("quotes_rejected") or []
        if rj:
            A("- ⚠ 未通过核验的引文（原文中找不到，不计分）：")
            for x in rj[:3]:
                txt = x.get("quote") if isinstance(x, dict) else str(x)
                A("  > %s" % short(txt, 90))
        fixes = advice_for(d, s)
        if fixes:
            A("")
            A("**怎么改**：")
            for cid, name, text in fixes[:2]:
                A("- （%s %s）%s" % (cid, name, text))
        else:
            fallback = DIM_FALLBACK.get(d["id"])
            if ratio < 0.8 and fallback:
                A("")
                A("**怎么改**：%s" % fallback)
            elif ratio < 0.8:
                A("- 建议再补一组可复算的数字与对照，把「做了什么—产生什么—怎么验证」串起来。")
            else:
                A("- 该维度已达标。")
        A("")

    # 争议点
    ds = s.get("disputes") or []
    if ds:
        A("## 争议点（答辩会被追问）")
        A("")
        for i, x in enumerate(ds, 1):
            A("%d. **%s**：%s" % (i, x.get("source", "—"), short(x.get("point"), 110)))
        A("")

    # 扣分
    if s.get("deductions"):
        A("## 固定扣分")
        A("")
        for x in s["deductions"]:
            A("- −%s 分：%s" % (x["points"], x["why"]))
        A("")

    # 优先改进顺序
    A("## 优先补什么（按性价比排序）")
    A("")
    rank = []
    for d in dims:
        loss = d["max"] - d["score"]
        if loss <= 0.01:
            continue
        fixes = advice_for(d, s)
        rank.append((loss, d, fixes))
    rank.sort(key=lambda x: -x[0])
    if not rank:
        A("各维度均已达标，无需补强。")
    for i, (loss, d, fixes) in enumerate(rank[:3], 1):
        if fixes:
            txt = fixes[0][2]
        else:
            txt = DIM_FALLBACK.get(d["id"]) or "补具体数据与佐证。"
        A("%d. **%s**（可追回 %s 分）：%s" % (i, d["name"], round(loss, 1), txt))
    A("")
    A("---")
    A("")
    A("*分数为自动机检 + 证据定级结果，需人工复核 judgments.json 后方可作为评审结论。*")
    return "\n".join(L) + "\n"


def main() -> int:
    ap = argparse.ArgumentParser(description="为每份材料生成详细改进意见")
    ap.add_argument("src", help="含 *.score.json 的目录")
    ap.add_argument("--out", default="意见")
    a = ap.parse_args()
    src, out = Path(a.src), Path(a.out)
    rows = load(src)
    if not rows:
        print("没找到 score.json", file=__import__("sys").stderr)
        return 1
    out.mkdir(parents=True, exist_ok=True)
    for s in rows:
        p = out / (s["_stem"] + ".意见.md")
        p.write_text(render_one(s), encoding="utf-8")
    print(json.dumps({"out": str(out), "count": len(rows)}, ensure_ascii=False))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())