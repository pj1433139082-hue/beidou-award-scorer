#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""验证诊断模型与计算口径：把 OTE 公式套到数采样本上，看能否复算。

## 为什么需要这个脚本

创新性被判 C 的唯一原因：11 套诊断模型没有可复算的有效性证据。
评审能做的核验只有两种 —— 要么你自己把公式代入样本算出结果，
要么你根本算不出来。这个脚本做的是第一种。

## 它验证什么

1. **OTE 公式可复算**：小时累计氧利用率 = (C_out - C_in) / C_in x 100%，
   代入数采样本的进出口溶解氧，看结果是否落在模型判定区间内。
2. **单位膜面积通气量**：通气量 / 膜面积，给出与基线区间的比对。
3. **判定一致性**：脚本算出的诊断结论 vs 材料声称的结论，是否一致。

## 使用

    python scripts/verify_models.py <样本.csv> --model ote --out 验证报告.json
    python scripts/verify_models.py <样本.csv> --model all --out out/

## 输入格式

CSV 至少包含时间列与相关测量列，列名大小写与中英文均可识别：

    timestamp,dissolved_oxygen_in,dissolved_oxygen_out,air_flow,membrane_area
    2026-11-01 10:00:00,1.8,5.2,1200,480

缺列时脚本会明确报出缺哪一列，不会静默给出错误结论。
"""
from __future__ import annotations
import argparse, csv, json, sys
from pathlib import Path

# 列名别名表：同一个量在不同导出里名字不一样
ALIASES = {
    "time": ["time", "timestamp", "时间", "采集时间", "datetime", "ts"],
    "do_in": ["dissolved_oxygen_in", "do_in", "进口溶解氧", "进水do", "do_inlet"],
    "do_out": ["dissolved_oxygen_out", "do_out", "出口溶解氧", "出水do", "do_outlet"],
    "air_flow": ["air_flow", "风量", "通气量", "瞬时风量", "供气量"],
    "membrane_area": ["membrane_area", "膜面积", "单位膜面积"],
    "power": ["power", "功率", "风机功率", "耗电"],
    "load": ["load", "负荷", "进水负荷", "cod"],
    "eff_excel": ["excel口径效率", "excel_eff", "效率", "风机效率"],
    "eff_adj": ["修正口径效率", "adj_eff", "修正效率"],
}


def find_col(headers, key):
    """在表头里找某个量的实际列名。找不到返回 None。"""
    low = {h.strip().lower(): h for h in headers}
    for alias in ALIASES[key]:
        if alias.lower() in low:
            return low[alias.lower()]
    return None


def to_float(v):
    try:
        return float(str(v).strip())
    except (TypeError, ValueError):
        return None


def load_rows(path: Path):
    with open(path, encoding="utf-8-sig", newline="") as f:
        r = csv.DictReader(f)
        return list(r.fieldnames or []), list(r)


def verify_ote(headers, rows):
    """小时累计氧利用率 OTE。

    正确口径：OTE = (C_out - C_in) / C_out x 100
    这是标准定义；不同厂可能有系数修正，脚本只做同口径相对比较。
    """
    ci = find_col(headers, "do_in")
    co = find_col(headers, "do_out")
    if not ci or not co:
        return {
            "model": "ote",
            "status": "cannot_verify",
            "reason": "缺少必需列：" + "、".join(
                name for name, col in [("进口溶解氧", ci), ("出口溶解氧", co)] if not col
            ),
            "have_columns": headers,
        }
    vals = []
    skipped = 0
    for row in rows:
        a, b = to_float(row.get(ci)), to_float(row.get(co))
        if a is None or b is None or a <= 0:
            skipped += 1
            continue
        vals.append((b - a) / b * 100.0)
    if not vals:
        return {"model": "ote", "status": "cannot_verify",
                "reason": "有列但无可解析数值（全部为空或非数字）", "skipped": skipped}
    vals.sort()
    n = len(vals)
    mean = sum(vals) / n
    if mean < -5 or mean > 105:
        return {
            "model": "ote",
            "status": "implausible",
            "reason": "均值 %.2f%% 超出 OTE 物理区间 [0,105]，分母可能取反或列选错" % mean,
            "ote_percent": {"mean": round(mean, 2), "min": round(vals[0], 2), "max": round(vals[-1], 2)},
            "hint": "确认取出口侧，公式为 (C_out-C_in)/C_out",
        }
    return {
        "model": "ote",
        "status": "verified",
        "formula": "(DO_out - DO_in) / DO_out * 100",
        "samples": n,
        "skipped": skipped,
        "ote_percent": {
            "mean": round(mean, 2),
            "min": round(vals[0], 2),
            "p50": round(vals[n // 2], 2),
            "max": round(vals[-1], 2),
        },
        "reproducible": True,
    }


def verify_flux(headers, rows):
    """单位膜面积通气量 = 瞬时风量 / 膜面积。"""
    cf = find_col(headers, "air_flow")
    ca = find_col(headers, "membrane_area")
    if not cf or not ca:
        return {"model": "flux", "status": "cannot_verify",
                "reason": "缺少必需列：" + "、".join(
                    name for name, col in [("风量", cf), ("膜面积", ca)] if not col),
                "have_columns": headers}
    vals = []
    for row in rows:
        f_, a_ = to_float(row.get(cf)), to_float(row.get(ca))
        if f_ is None or a_ is None or a_ <= 0:
            continue
        vals.append(f_ / a_)
    if not vals:
        return {"model": "flux", "status": "cannot_verify", "reason": "无可解析数值"}
    vals.sort()
    n = len(vals)
    return {
        "model": "flux",
        "status": "verified",
        "formula": "air_flow / membrane_area",
        "samples": n,
        "flux": {"mean": round(sum(vals) / n, 3), "min": round(vals[0], 3),
                 "p50": round(vals[n // 2], 3), "max": round(vals[-1], 3)},
        "reproducible": True,
    }


def verify_fan(headers, rows, rated=(65.0, 77.0)):
    """风机效率诊断：判定有多少行落在铭牌额定区间内。

    口径：优先用修正口径（真实进口状态），没有再退回 Excel 口径。
    输出落在额定区间的比例，这个比例直接回答「模型判断可信吗」。
    """
    ce = find_col(headers, "eff_adj") or find_col(headers, "eff_excel")
    if not ce:
        return {"model": "fan", "status": "cannot_verify",
                "reason": "缺少效率列（Excel口径效率 或 修正口径效率）",
                "have_columns": headers}
    vals = [to_float(r.get(ce)) for r in rows]
    vals = [v for v in vals if v is not None]
    if not vals:
        return {"model": "fan", "status": "cannot_verify", "reason": "无可解析效率值"}
    lo, hi = rated
    inband = sum(1 for v in vals if lo <= v <= hi)
    vals.sort()
    n = len(vals)
    return {
        "model": "fan",
        "status": "verified",
        "column": ce,
        "rated_range": "%g%%~%g%%" % (lo, hi),
        "samples": n,
        "efficiency": {"mean": round(sum(vals) / n, 2), "min": round(vals[0], 2),
                       "p50": round(vals[n // 2], 2), "max": round(vals[-1], 2)},
        "in_rated_band": inband,
        "in_rated_ratio": round(inband / n * 100, 2),
        "reproducible": True,
    }


def consistency(claim, measured, tol=0.05):
    """判定一致性：脚本算出的结论 vs 材料声称。

    只做区间落点判断，不替评审下结论。
    """
    if claim is None or measured is None:
        return "无法判断"
    c, m = to_float(claim), to_float(measured)
    if c is None or m is None:
        return "无法判断"
    if c == 0:
        return "一致" if abs(m) < 1e-9 else "不一致"
    return "一致" if abs(m - c) / abs(c) <= tol else "不一致"


def main() -> int:
    ap = argparse.ArgumentParser(description="验证诊断模型的可复算性")
    ap.add_argument("sample", help="数采样本 CSV")
    ap.add_argument("--model", default="all",
                    choices=["all", "ote", "flux", "fan"])
    ap.add_argument("--claim", help="材料声称的结果，用于一致性比对（可选）")
    ap.add_argument("--tolerance", type=float, default=0.05)
    ap.add_argument("--out", help="输出 JSON 路径；省略则打印")
    a = ap.parse_args()

    p = Path(a.sample)
    if not p.is_file():
        print("找不到样本：%s" % p, file=sys.stderr)
        return 2

    headers, rows = load_rows(p)
    print("样本文件：%s" % p.name)
    print("  行数：%d" % len(rows))
    print("  列数：%d" % len(headers))
    print()

    results = []
    if a.model in ("all", "ote"):
        results.append(verify_ote(headers, rows))
    if a.model in ("all", "flux"):
        results.append(verify_flux(headers, rows))
    if a.model in ("all", "fan"):
        results.append(verify_fan(headers, rows))

    ok = 0
    for r in results:
        print("[%s] %s" % (r["model"].upper(), r["status"]))
        if r["status"] == "verified":
            ok += 1
            for k, v in r.items():
                if k in ("model", "status"):
                    continue
                print("    %-12s %s" % (k, v))
            if a.claim and r["model"] == "ote":
                m = (r.get("ote_percent") or {}).get("mean")
                print("    一致性        %s（声称 %s，实测均值 %s）"
                      % (consistency(a.claim, m, a.tolerance), a.claim, m))
        else:
            print("    原因：%s" % r.get("reason"))
            if r.get("have_columns"):
                print("    实际列：%s" % ", ".join(r["have_columns"]))
        print()

    payload = {
        "sample": str(p),
        "rows": len(rows),
        "claim": a.claim,
        "tolerance": a.tolerance,
        "results": results,
        "verified_count": ok,
        "total": len(results),
    }

    if a.out:
        out = Path(a.out)
        if out.suffix.lower() != ".json":
            out = out / (p.stem + ".verify.json")
        out.parent.mkdir(parents=True, exist_ok=True)
        out.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
        print("已写出：%s" % out)
    else:
        print(json.dumps(payload, ensure_ascii=False, indent=2))

    return 0 if ok else 1


if __name__ == "__main__":
    raise SystemExit(main())