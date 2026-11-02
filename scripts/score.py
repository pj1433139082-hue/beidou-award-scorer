#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""按官方评审标准给单份申报材料打分。

输入：doc.json（extract_report.py）+ checks.json（checks.py）+ judgments.json（评审员的判断，含逐字引文）
输出：score.json + report.md

三条硬规则（对应官方 AI 评审迭代的三项要求）：
1. 评分导向校准  —— AI 不是加分项；技术属性只做分类（rubric.scoring_rules.ai_not_bonus）
2. 落地真实性核验 —— 每条判断必须附**逐字引文**，引文在原文里查不到即降证据等级
3. 争议点反馈    —— 自动汇总冲突项成《争议点清单》，供答辩追问
"""
from __future__ import annotations
import argparse, json, re, sys
from pathlib import Path

BAND_DEFAULT = [("一等奖候选", 85), ("二等奖候选", 75), ("三等奖候选", 65), ("入围候选", 55), ("待提升", 0)]
UNSCORED = -1.0


class NoOnsiteError(Exception):
    """给不设现场答辩的奖项传了 --onsite。"""


class UnknownAwardError(Exception):
    """奖项名不在评分卡里。"""

def _default_rubric():
    """默认评分卡：优先 rubric.json，没有就用 rubric.example.json。"""
    assets = Path(__file__).resolve().parents[1] / "assets"
    for name in ("rubric.json", "rubric.example.json"):
        p = assets / name
        if p.is_file():
            return p
    return assets / "rubric.json"

def validate_rubric_bands(rubric):
    """启动自检：证据等级分档必须互不重叠（左闭右开），否则同一分值会落进两档。

    借鉴 deepeval `g_eval/utils.py:212 validate_and_sort_rubrics` 的"逐对检查区间重叠即报错"。
    """
    grades = rubric.get("evidence_grades") or {}
    if not grades:
        return True
    bands = sorted(((k, float(v["band"][0]), float(v["band"][1])) for k, v in grades.items()), key=lambda x: x[1])
    problems = []
    for i in range(len(bands) - 1):
        a, b = bands[i], bands[i + 1]
        if b[1] < a[2] - 1e-9:      # 边界相接合法，真正交叠才报
            problems.append("%s[%.2f,%.2f] 与 %s[%.2f,%.2f] 重叠" % (a[0], a[1], a[2], b[0], b[1], b[2]))
    if problems:
        raise SystemExit("rubric 证据分档非法（会命中多个档）：" + "；".join(problems))
    return True
GRADE_FACTOR = {"A": 1.0, "B": 0.85, "C": 0.70, "D": 0.50}
ORDER = ["D", "C", "B", "A"]

def norm(s):
    return re.sub(r"\s+", "", str(s or ""))

def verify_quotes(quotes, plain_norm):
    ok, bad = [], []
    for q in quotes or []:
        n = norm(q)
        if len(n) < 6:
            bad.append({"quote": q, "why": "过短（<6 字），不足以定位"})
        elif n in plain_norm:
            ok.append(q)
        else:
            bad.append({"quote": q, "why": "原文中找不到该表述"})
    return ok, bad

def grade_cap(verified, claimed):
    """引文核验不足时给证据等级封顶：0 条 -> D，1 条 -> 最高 C，>=2 条 -> 按其自评。"""
    claimed = (claimed or "C").upper()
    if claimed not in ORDER: claimed = "C"
    if verified <= 0: return "D"
    if verified == 1: return ORDER[min(ORDER.index(claimed), 1)]
    return claimed

def apply_machine(checks, awards_caps):
    caps, deductions, gates = {}, [], []
    by_id = {c["id"]: c for c in checks}
    def cap(dim, ratio, why, cid=None):
        caps[dim] = min(caps.get(dim, 1.0), ratio)
        gates.append({"dim": dim, "type": "cap", "to": ratio, "why": why, "check": cid})
    def deduct(points, why, rule, cid=None):
        deductions.append({"points": points, "why": why, "rule": rule, "check": cid})
    c = lambda i: by_id.get(i, {"status": "na", "detail": ""})
    st = lambda i: c(i)["status"]

    if st("C02") == "fail": cap("value", 0.40, "「实际值【已实现】」全空：效益全为规划，价值性最高 40%", "C02")
    if st("C02") == "warn": cap("value", 0.75, "仅个别效益填了已实现值：价值性封顶 75%", "C02")
    if st("C03") == "warn": cap("value", 0.85, "存在有实际值无基准值的行，变化率不可复算", "C03")
    if st("C04") == "fail": cap("value", 0.70, "已实现效益普遍缺佐证材料", "C04")
    if st("C04") == "warn": cap("value", 0.85, "部分已实现效益缺佐证材料", "C04")
    if st("C05") == "warn": deduct(2, "变化率自述与复算不符（偏差>5pp）", "D2 正文数字与测算不一致", "C05")
    if st("C08") == "warn": deduct(3, "同一金额多处出现，存在重复计入风险", "D1 重复计入", "C08")
    if st("C09") == "warn": deduct(2, "存在未填写占位（下划线/XXXX/（文本））", "材料完成度", "C09")
    if st("C11") == "fail": cap("intelligence", 0.40, "未见模型实验室部署或试点厂生产调用证据", "C11")
    if st("C11") == "warn": cap("intelligence", 0.70, "实验室/试点厂证据不全，或缺 >=90 天运行验证", "C11")
    if st("C12") == "warn": cap("readiness", 0.65, "产品就绪度未达「已实现产品转化」", "C12")
    if st("C12") == "fail": cap("readiness", 0.45, "未能识别产品形态与转化程度", "C12")
    if st("C14") == "fail": cap("novelty", 0.60, "未见与原有办法的量化对照，创新性封顶 60%", "C14")
    if st("C14") == "warn": cap("novelty", 0.80, "有对比表述但缺量化对照", "C14")
    if st("C15") == "warn": cap("usability", 0.80, "战略对齐或卡点描述缺一项", "C15")
    if st("C16") == "warn": cap("value", 0.85, "预计性表述偏多，包装风险", "C16")
    if st("C17") == "fail": deduct(2, "全文无图片/截图证据", "D4 截图缺失", "C17")
    if awards_caps and st("C13") == "fail":
        # 「成果迭代性」是**否决项，不是加分项**（用户 2026-11-02 确认）——
        # 所以它不进入 dimensions、不占分值、不加分。
        # 但 P11 把最终裁定权给了评委综合审议（人工与 AI 差异 >10% 即转审议），
        # P12 明确 AI 的职责是"生成争议点清单，辅助答辩追问"，而非单方面定案。
        # 因此：分值不受影响，只置顶争议点 + 标注待裁定，由人和综合审议裁定。
        # 不额外扣分：官方未给否决项的扣分幅度，任何数字都是我们编的。
        gates.append({"dim": "iteration", "type": "veto_pending", "check": "C13",
                      "to": 0, "why": "成果迭代性未见说明：否决项，需人工/综合审议确认"})
    return caps, deductions, gates

def score_one(doc, checks, judgments, rubric, award, onsite=None):
    awards = rubric.get("awards") or {}
    if award not in awards:
        raise UnknownAwardError(
            "评分卡里没有「%s」这个奖项。\n"
            "  当前评分卡可用奖项：%s\n"
            "  若你在用公开版的示例卡（rubric.example.json），它只含示例奖项；\n"
            "  请用 --rubric 指向你自己的 rubric.json。" % (
                award, "、".join(awards.keys()) or "（空）"))
    spec = awards[award]
    if onsite is not None and spec.get("no_onsite"):
        raise NoOnsiteError(
            "%s 不设现场评审（规则原文：成果推广赛道不设答辩环节，采用 AI+人工辅助方式评定）。\n"
            "  不接受 --onsite；该奖项不参与现场 75%% + AI 25%% 合成。\n"
            "  去掉 --onsite 重跑。" % award)
    plain_norm = norm(doc["plain"])
    caps, deductions, gates = apply_machine(checks["checks"], [award])

    validate_rubric_bands(rubric)
    rows, unscored = [], []
    for d in spec["dimensions"]:
        j = ((judgments.get("judgments") or {}).get(d["id"])) or {}
        raw = j.get("score", None)
        if raw is None or float(raw) < 0:            # 未评分：材料不可读/抽取失败/评审员没给判断
            unscored.append({"id": d["id"], "name": d["name"], "max": d["max"],
                             "why": j.get("unscored_reason") or "评审员未给出该维度判断"})
            continue
        raw = float(raw)
        claimed = (j.get("grade") or "C").upper()
        if claimed not in ORDER: claimed = "C"
        ok, bad = verify_quotes(j.get("quotes") or [], plain_norm)
        grade = grade_cap(len(ok), claimed)
        cap = caps.get(d["id"], 1.0)
        after_cap = min(raw, d["max"] * cap)
        # 证据等级只在"降级"时生效：降一级按该等级系数收缩，避免"无证据也能拿高分"
        final = after_cap if grade == claimed else min(after_cap, d["max"] * GRADE_FACTOR[grade])
        rows.append({
            "id": d["id"], "name": d["name"], "max": d["max"], "official": d["official"],
            "llm_raw": round(raw, 1), "claimed_grade": claimed, "verified_grade": grade,
            "cap": cap,
            "cap_reason": next((g["why"] for g in gates if g["dim"] == d["id"] and g["type"] == "cap"), None),
            "score": round(final, 1), "rationale": j.get("rationale", ""),
            "quotes_verified": ok, "quotes_rejected": bad,
        })

    subtotal = round(sum(r["score"] for r in rows), 1)
    total_deduct = sum(x["points"] for x in deductions)
    ai_score = round(max(0.0, subtotal - total_deduct), 1)
    # 未评分维度排除出分母（借鉴 ossf/scorecard check_result.go:40 + GetAggregateScore 的 continue）
    max_total = sum(d["max"] for d in spec["dimensions"] if d["id"] not in {u["id"] for u in unscored})
    if unscored and max_total:
        ai_score = round(ai_score * 100.0 / max_total, 1)   # 折算回百分制，便于跨材料比较
        max_total = 100
    veto = [g for g in gates if g["type"] == "veto"]
    veto_pending = [g for g in gates if g["type"] == "veto_pending"]
    if veto: ai_score = 0.0
    band = next((n for n, th in BAND_DEFAULT if ai_score >= th), "待提升")
    if veto: band = "否决（不入围）"
    elif veto_pending: band = "暂列「%s」，否决项待综合审议确认" % band

    disputes = []
    for chk in checks["checks"]:
        if chk["status"] in ("warn", "fail"):
            disputes.append({"source": "机检 " + chk["id"],
                             "point": "%s：%s" % (chk["name"], chk["detail"]),
                             "evidence": chk.get("evidence", [])[:3]})
    for r in rows:
        for b in r["quotes_rejected"]:
            disputes.append({"source": "引文核验 " + r["name"],
                             "point": "引文在原文中找不到：%s" % str(b["quote"])[:60],
                             "evidence": [b["why"]]})
    for q in judgments.get("disputes") or []:
        disputes.append({"source": "评审员标记", "point": str(q)[:200], "evidence": []})
    for g in veto_pending:
        # 置顶：这一条决定项目生死，必须排在最前
        disputes.insert(0, {"source": "否决项待裁定", "point": g["why"],
                            "evidence": ["规则原文将成果迭代性标为否决项；P11 最终裁定权在评委综合审议，"
                                         "P12 要求 AI 输出争议点清单辅助答辩追问，故不直接归零。"]})

    warn_fail = sum(1 for c in checks["checks"] if c["status"] in ("warn", "fail"))
    confidence = "高"
    if warn_fail >= 6 or any(r["verified_grade"] in ("C", "D") for r in rows): confidence = "中"
    if warn_fail >= 9: confidence = "低"

    out = {"award": award, "track": spec["track"], "name": doc["name"], "source": doc["source"],
           "max_total": max_total, "ai_score": ai_score, "band": band, "confidence": confidence,
           "dimensions": rows, "unscored": unscored, "caps_and_gates": gates, "deductions": deductions,
           "disputes": disputes, "veto": bool(veto), "veto_pending": bool(veto_pending),
           "veto_note": ("；".join(g["why"] for g in veto_pending) or None),
           "rubric_edition": rubric.get("edition")}
    if onsite is not None:
        try:
            on = float(onsite)
            out["onsite_score"] = on
            out["final_weighted"] = round(0.75 * on + 0.25 * ai_score, 1)
            out["conflict_review"] = abs(on - ai_score) / max(on, 1e-9) > 0.10
        except ValueError:
            pass
    return out

def render_md(res):
    L = []
    A = L.append
    A("# %s｜%s 评分报告" % (res["name"], res["award"]))
    A("")
    A("> 依据：2026 年第八届北斗奖评审标准；赛道：%s；评分卡：%s" % (res["track"], res["rubric_edition"]))
    A("> 机检 + 推理评审混合；每条得分都带逐字引文，引文在原文查不到即降证据等级。")
    A("")
    A("## 一、结论")
    A("")
    A("| 项目 | 结果 |")
    A("|---|---|")
    A("| AI 评审总分 | **%s / %s** |" % (res["ai_score"], res["max_total"]))
    A("| 奖项区间 | %s |" % res["band"])
    A("| 结论置信度 | %s |" % res["confidence"])
    if res["veto"]: A("| 否决项 | **命中（见下）** |")
    if res.get("veto_pending"): A("| 否决项 | **待裁定**：%s |" % (res.get("veto_note") or "见争议点清单"))
    if "final_weighted" in res:
        A("| 现场评审分 | %s |" % res["onsite_score"])
        A("| 综合得分（现场 75%% + AI 25%%） | **%s** |" % res["final_weighted"])
        A("| 触发综合审议（差异>10%%） | %s |" % ("是" if res["conflict_review"] else "否"))
    A("")
    A("## 二、维度得分")
    A("")
    A("| 维度 | 官方分值 | 自评 | 证据等级 | 封顶 | 最终 |")
    A("|---|---|---|---|---|---|")
    for r in res["dimensions"]:
        cap_txt = ("x%.2f %s" % (r["cap"], (r["cap_reason"] or "").replace("%", "%%"))) if r["cap"] < 1 else "—"
        A("| %s | %s | %s | %s→%s | %s | **%s** |" % (r["name"], r["max"], r["llm_raw"],
                                                    r["claimed_grade"], r["verified_grade"], cap_txt, r["score"]))
    if res.get("unscored"):
        A("")
        A("> **未评分维度（已排除出分母，未按 0 分计）**：")
        for u in res["unscored"]:
            A("> - %s（满分 %s）：%s" % (u["name"], u["max"], u["why"]))
    A("")
    for r in res["dimensions"]:
        A("### %s（%s/%s）" % (r["name"], r["score"], r["max"]))
        A("")
        A("- 官方口径：%s" % r["official"])
        if r["rationale"]: A("- 判据：%s" % r["rationale"])
        if r["quotes_verified"]:
            A("- 已核验引文：")
            for q in r["quotes_verified"][:5]: A("  - 「%s」" % q)
        if r["quotes_rejected"]:
            A("- **未通过核验**：")
            for b in r["quotes_rejected"][:5]: A("  - 「%s」— %s" % (str(b["quote"])[:60], b["why"]))
        if r["cap_reason"]: A("- 封顶：%s" % r["cap_reason"])
        A("")
    if res["deductions"]:
        A("## 三、扣分明细")
        A("")
        A("| 扣分 | 原因 | 依据 |")
        A("|---|---|---|")
        for d in res["deductions"]: A("| -%s | %s | %s |" % (d["points"], d["why"], d["rule"]))
        A("")
    A("## 四、争议点清单（答辩追问用）")
    A("")
    if not res["disputes"]: A("无。")
    for i, d in enumerate(res["disputes"], 1):
        A("%d. **[%s]** %s" % (i, d["source"], d["point"]))
        for e in d["evidence"]: A("   - %s" % str(e)[:120])
    A("")
    A("## 五、改进建议")
    A("")
    todo = []
    for r in res["dimensions"]:
        if r["score"] < r["max"] * 0.7:
            todo.append("- **%s** 失分 %s 分：%s" % (r["name"], round(r["max"] - r["score"], 1),
                                                  r["cap_reason"] or "论证不足或证据强度不够"))
    for d in res["deductions"]:
        todo.append("- **固定扣分项**：%s（可修复，+%s 分）" % (d["why"], d["points"]))
    A("\n".join(todo) if todo else "- 无明显可快速修复的失分点。")
    A("")
    return "\n".join(L)

def _load_json(path: str, what: str):
    """读 JSON，缺失或损坏给可读错误。"""
    p = Path(path)
    if not p.is_file():
        print("错误：找不到%s文件：%s" % (what, p), file=sys.stderr)
        raise SystemExit(2)
    try:
        return json.loads(p.read_text(encoding="utf-8"))
    except json.JSONDecodeError as e:
        print("错误：%s文件不是合法 JSON（%s）：\n  %s" % (what, e.msg, p), file=sys.stderr)
        raise SystemExit(2)


def main():
    ap = argparse.ArgumentParser(description="北斗奖申报材料评分")
    ap.add_argument("doc_json")
    ap.add_argument("--checks", required=True)
    ap.add_argument("--judgments", required=True, help="评审员判断 json（含逐字引文）")
    ap.add_argument("--award", required=True, help="奖项名，必须与 rubric 里 awards 的某个键一致")
    ap.add_argument("--onsite", help="现场评审得分（可选，用于 75/25 合成与 10%% 争议判定）")
    ap.add_argument("--rubric", default=str(_default_rubric()))
    ap.add_argument("--out", default=".")
    a = ap.parse_args()
    doc = _load_json(a.doc_json, "材料 doc.json")
    checks = _load_json(a.checks, "机检结果 checks.json")
    judg = _load_json(a.judgments, "评审判断 judgments.json")
    rub = _load_json(a.rubric, "评分卡")
    try:
        res = score_one(doc, checks, judg, rub, a.award, a.onsite)
    except NoOnsiteError as exc:
        print("错误：" + str(exc), file=sys.stderr)
        return 2
    except UnknownAwardError as exc:
        print("错误：" + str(exc), file=sys.stderr)
        return 2
    out = Path(a.out); out.mkdir(parents=True, exist_ok=True)
    stem = Path(a.doc_json).name.replace(".doc.json", "")
    (out / (stem + ".score.json")).write_text(json.dumps(res, ensure_ascii=False, indent=1), encoding="utf-8")
    (out / (stem + ".report.md")).write_text(render_md(res), encoding="utf-8")
    print(json.dumps({"name": res["name"], "award": a.award, "ai_score": res["ai_score"],
                      "band": res["band"], "confidence": res["confidence"],
                      "disputes": len(res["disputes"])}, ensure_ascii=False))
    return 0

if __name__ == "__main__":
    raise SystemExit(main())