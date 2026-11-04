#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""核对「上次给的建议，这一版落实了没有」。

## 解决什么

现在的循环断在最后一环：

    材料 v1 → 给建议 → 用户改 → 材料 v2 → 重新评
                                          ↑
                                但没人回答一个更基本的问题：
                                「上次让你补的，到底补了没有？」

重新评会给一个新分数，但看不出分数变化是因为**真的改了**，
还是因为评审换了口径。这两件事必须分开。

## 怎么判断落实

对每条建议，取它当初针对的**原文片段**在 v1 里的位置，
再到 v2 里找对应段落，比三个信号：

    1. 那段原文还在不在？（还在 = 没动过）
    2. 相关内容有没有变长/变多？（补了内容）
    3. 有没有出现建议要求的新要素？（基准值、厂名清单、日期…）

结论分四类：

    已落实    新要素出现了
    部分落实  内容变多但要素不全
    未改动    原文一字未变
    无法判断  找不到对应段落（结构大改或章节挪位）

## 为什么不用 LLM 判断

这一层的价值是**可复核**：给定两份材料，任何人跑一遍应得到同样结论。
让 LLM 判「这条算不算落实」会引入不可复现性，正好破坏它的用途。
需要语义判断时，脚本输出证据，由 agent 定夺。

用法：
    python scripts/check_compliance.py <新版材料目录> --baseline <旧版输出> --out <输出>
"""
from __future__ import annotations
import argparse, json, re, sys
from pathlib import Path


def load_doc(d: Path, name: str):
    """按材料名找 doc.json，容忍名字不完全一致。"""
    for f in Path(d).glob("*.doc.json"):
        try:
            x = json.loads(f.read_text(encoding="utf-8"))
        except Exception:
            continue
        n = x.get("name") or f.stem
        if n == name or (len(name) > 8 and (name[:12] in n or n[:12] in name)):
            return x
    return None


def load_items(d: Path) -> dict:
    """从汇总表里读上次的建议条目。布局与解析都在 advice_sheet 里定义。"""
    from advice_sheet import read as read_advice
    return read_advice(d)


# 每类建议要求出现的「新要素」信号词
# 用词表而不是 LLM：可复核、可解释、跑两次结果一样
SIGNALS = {
    "价值性": [
        (r"基准|改造前|原来|此前|v1|上一年", "基准值"),
        (r"已实现|实际值|台账|核算表", "已实现值/佐证"),
        (r"计算过程|测算过程|公式", "计算过程"),
    ],
    "推广性": [
        (r"\d+\s*家|\d+\s*个\s*(?:项目|厂|水厂)", "厂站数量"),
        (r"\d{4}\s*年\s*\d{1,2}\s*月|部署时间|上线时间", "部署时间"),
        (r"大区|区域", "区域归属"),
    ],
    "创新性": [
        (r"对比|相比|较之|原有|传统|与.*相比", "对照表述"),
        (r"\d+(\.\d+)?\s*%|降低|提升|减少|缩短", "量化差异"),
    ],
    "智能化水平": [
        (r"实验室|中试|试点", "验证场景"),
        (r"\d+\s*天|连续运行|稳定运行", "运行时长"),
        (r"调用|接口|上线|部署", "生产调用证据"),
    ],
    "实用性": [
        (r"卡点|痛点|难点", "问题陈述"),
        (r"战略|规划|方向", "战略对齐"),
    ],
    "经济效益": [
        (r"计算过程|测算|公式", "测算过程"),
        (r"基准|改造前", "基准值"),
    ],
    "产品就绪度": [
        (r"产品|架构|组件|模块|接口", "产品形态"),
        (r"转化|落地|上线|交付", "转化程度"),
    ],
    "推广力度": [
        (r"\d+\s*家|\d+\s*个", "数量"),
        (r"台账|清单|明细", "推广证据"),
    ],
}


def find_source_span(old_plain: str, item: dict):
    """从建议里反推它当初针对的原文片段。

    「原文是什么」列是机检的**结论描述**，不是材料原句，
    所以不能直接在 v2 里搜它。改为按维度信号词在 v1 里定位相关段落。
    """
    if not old_plain:
        return ""
    # 用建议的「修改之后的句子」里的关键词回到原文找上下文
    keys = [w for w in re.findall(r"[\u4e00-\u9fa5]{3,6}", item.get("revised") or "")][:6]
    for k in keys:
        i = old_plain.find(k)
        if i >= 0:
            return old_plain[max(0, i - 200):i + 260]
    # 退而用维度信号词定位
    sigs = SIGNALS.get(item.get("dim") or "", [])
    for pat, _ in sigs:
        m = re.search(pat, old_plain)
        if m:
            return old_plain[max(0, m.start() - 200):m.end() + 260]
    return ""


def judge(old_plain: str, new_plain: str, item: dict) -> dict:
    """判断一条建议在 v2 里的落实状态。"""
    sigs = SIGNALS.get(item.get("dim") or "", [])
    if not sigs:
        return {"status": "无法判断", "why": "该维度没有信号词表"}

    # 不能用「有没有命中」判断——旧版里往往已经命中过（比如本来就有"4个大区"），
    # 那样永远判不出新增。改为看**出现次数**是否明显增加：
    # 真正的补充会让证据密度上升，而不只是换个说法。
    def count(pat, s):
        return len(re.findall(pat, s or ""))

    gained, densified = [], []
    for pat, name in sigs:
        o, n = count(pat, old_plain), count(pat, new_plain)
        if o == 0 and n > 0:
            gained.append(name)
        elif n - o >= 2:
            # 同类证据至少多出 2 处，就算补充。
            # 不能用"翻倍"当门槛：原文本来 18 处、补到 21 处只是小幅增加，
            # 但补的那 3 处可能正是"厂站清单"这种关键证据。
            # 业务上"多写了具体厂名"和"多写了三遍套话"价值完全不同，
            # 脚本只负责指出密度变了，具体值不值得由人判断。
            densified.append("%s(%d→%d)" % (name, o, n))

    if gained:
        return {"status": "已落实", "why": "新增要素：" + "、".join(gained),
                "gained": gained}
    if densified:
        return {"status": "已落实", "why": "证据密度上升：" + "、".join(densified),
                "densified": densified}

    # 纯粹的长度增长也可能是有改动，但要素没增加，算部分落实
    old_len, new_len = len(old_plain or ""), len(new_plain or "")
    if new_len - old_len >= 20 or new_len > old_len * 1.005:
        missing = [name for pat, name in sigs if count(pat, new_plain) == 0]
        if missing:
            return {"status": "部分落实",
                    "why": "材料变长 %d 字，但以下要素仍缺：%s" % (
                        new_len - old_len, "、".join(missing))}

    old_hit = [name for pat, name in sigs if re.search(pat, old_plain or "")]
    new_hit = [name for pat, name in sigs if re.search(pat, new_plain or "")]
    if old_hit and set(new_hit) == set(old_hit):
        # 要素没变，看目标段落是不是一字未动
        span = find_source_span(old_plain, item)
        if span and span[:120] in (new_plain or ""):
            return {"status": "未改动", "why": "针对的段落与上一版一字未变"}
        if span:
            return {"status": "部分落实",
                    "why": "段落有改动，但要求的要素（%s）仍未见" % "、".join(
                        [n for pat, n in sigs if n not in new_hit])}
        return {"status": "无法判断", "why": "定位不到当初针对的段落"}
    if not old_hit and not new_hit:
        return {"status": "无法判断", "why": "两版都没有该维度的可识别信号"}
    return {"status": "无法判断", "why": "信号不匹配，需人工看"}


def main() -> int:
    ap = argparse.ArgumentParser(description="核对建议是否落实")
    ap.add_argument("new_dir", help="新版材料目录（含 *.doc.json）")
    ap.add_argument("--baseline", required=True, help="旧版输出目录（含 xlsx）")
    ap.add_argument("--out", default=".")
    a = ap.parse_args()

    new_dir = Path(a.new_dir)
    base = Path(a.baseline)
    if not new_dir.is_dir() or not base.is_dir():
        print("目录不存在", file=sys.stderr)
        return 2

    # 旧版原文：从 baseline 目录或它旁边的材料目录找
    items = load_items(base)
    if not items:
        print("baseline 里没有 北斗奖评分汇总.xlsx，无法拿到上次建议", file=sys.stderr)
        return 2

    rows = []
    stat = {"已落实": 0, "部分落实": 0, "未改动": 0, "无法判断": 0}
    for name, its in items.items():
        old_doc = load_doc(base, name)
        new_doc = load_doc(new_dir, name)
        if not new_doc:
            continue
        old_plain = (old_doc or {}).get("plain", "")
        new_plain = new_doc.get("plain", "")
        per = []
        for it in its:
            r = judge(old_plain, new_plain, it)
            r.update({"no": it["no"], "dim": it["dim"], "title": it["title"]})
            stat[r["status"]] = stat.get(r["status"], 0) + 1
            per.append(r)
        if per:
            rows.append({"name": name, "items": per})

    lines = ["# 建议落实情况核对", "",
             "新版：%s" % new_dir, "旧版：%s" % base, "",
             "## 汇总", "",
             "| 状态 | 条数 |", "|---|---|"]
    for k, v in stat.items():
        lines.append("| %s | %d |" % (k, v))
    lines.append("")

    for r in rows:
        lines.append("## %s" % r["name"][:44])
        lines.append("")
        lines.append("| 编号 | 维度 | 状态 | 说明 |")
        lines.append("|---|---|---|---|")
        for it in r["items"]:
            lines.append("| %s | %s | %s | %s |" % (
                it["no"], it["dim"], it["status"], (it.get("why") or "")[:50]))
        lines.append("")

    out = Path(a.out)
    out.mkdir(parents=True, exist_ok=True)
    p = out / "建议落实核对.md"
    p.write_text("\n".join(lines), encoding="utf-8")

    print(json.dumps({"out": str(p), "materials": len(rows), "stats": stat,
                      "note": "未改动=建议没被采纳；无法判断=需人工看原文"},
                     ensure_ascii=False, indent=1))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())