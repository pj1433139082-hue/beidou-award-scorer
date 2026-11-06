#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""盲区模式：只给子智能体一份 docx + 一张最小规则卡。

## 和 blind_prepare.py 的区别

blind_prepare.py 复制整个 skill/（scripts + assets + references），
blind_minimal.py 只放：

    材料/申报主报告.docx      唯一业务输入
    规则.md                  维度表 + 等级定义 + 打分纪律（不含任何分数、示例、历史）
    输出/                   子智能体写 judgments.json

判断依据全部来自材料本身。子智能体不需要跑机检、不需要评分卡、
不需要看 SKILL.md —— 它只需要知道「评什么、怎么定级、引文要逐字」。

## 为什么更干净

  · 输入文件从 36 个降到 2 个
  · 没有 doc.json / checks.json，就没有 name/source 元数据泄露
  · 没有 rubric.example.json，就没有示例锚定
  · 没有 skill/scripts，子智能体无法顺手读到别的东西

## 代价（必须说清）

  · 引文由子智能体从 docx 直接取，可能与 extract_report.py 的 plain 有差异
    （实测 16/16 命中，但表格内文字要小心）。
  · 主会话拿到 judgments 后，仍应用 score.py 校验每条引文；
    查不到的引文会被拒收并自动降级，不会污染分数。
  · 机检仍由主会话在盲区外跑，子智能体不接触机检结论。

## 用法

    python scripts/blind_minimal.py <盲区> --material <材料.docx> --award 金智囊
    python scripts/blind_minimal.py <盲区> --verify
"""
from __future__ import annotations
import argparse, json, os, re, shutil, sys
from pathlib import Path

DEFAULT_NAME = "申报主报告"
LEAK_PATTERNS = (
    re.compile(r"\b\d{2,3}\.\d\s*(?:分\s*)?(?:终版|版本|迭代版)"),
    re.compile(r"[（(]\s*\d{2,3}\.\d\s*[^)）]*?(?:终版|版|优化)[^)）]*[)）]"),
    re.compile(r"(?:终版|优化版|补全版|迭代版)[（(]?\s*\d{2,3}\.\d"),
)

# 从 judgments-template.md 与 rubric 里抽出的「纯规则」，不含任何分数/示例
DIMENSIONS = {
    "金点子": [("value","价值性",35),("replication","推广性",35),("novelty","创新性",30)],
    "金刚钻": [("value","价值性",35),("replication","推广性",35),("novelty","创新性",30)],
    "金引擎": [("value","价值性",25),("replication","推广性",25),("novelty","创新性",30),("intelligence","智能化水平",20)],
    "金智囊": [("value","价值性",30),("replication","推广性",25),("novelty","创新性",20),("usability","实用性",25)],
    "金火炬": [("promotion","推广力度",40),("economics","经济效益",40),("readiness","产品就绪度",20)],
    "AI融创": [("value","价值性",25),("replication","推广性",25),("novelty","创新性",30),("intelligence","智能化水平",20)],
}

RULES = """# 评审规则（本文件不含任何分数、示例或历史判断）

你是北斗奖评审员。只读 {name}.docx，独立给出判断。

## 一、维度（奖项：{award}）

{dim_table}

## 二、证据等级（grade）

| 等级 | cap | 判定依据 |
|---|---|---|
| A | 1.00 | 每个关键数字都有改造前/基线值配对；至少一项效益配了佐证材料（台账/日志/截图/第三方报告）；能按材料给的公式或口径独立算出与文中一致的数。三条全满足才是 A。 |
| B | 0.75 | 有具体数字（不是「显著提升」这类表述），但缺基线值、佐证只覆盖部分时段/项目，或数字无法用材料自身信息复算。 |
| C | 0.50 | 只有文字描述做法或效果；或全文没有该维度的具体数字；或数字只在「预计/规划/目标」语境；或数据来源是自评表/PPT 结论。 |
| D | 0.20 | 该维度只有目标、口号或规划表述。 |

score 不得超过 满分 × 该等级 cap。四条档位左闭右开，互不重叠。

## 三、打分纪律

1. AI、数字化、传统技改只是分类，不是加分项。
2. 先认证据，再给分。不得为了让总分好看而调 grade。
3. 已实现 ≠ 计划中。效益若未区分「实际值【已实现】」与「远期目标【规划/在研】」，按在研计。
4. quotes 必须是材料正文里的逐字片段（>=6 字），不能改写、拼接、翻译。
5. 给不出逐字引文的判断，把 grade 降到 C 并说明缺什么材料。
6. 某维度确实无法判断，写 {{"score": -1, "unscored_reason": "..."}}，不要打 0。
7. 人工评审与 AI 差异 >10% 写进 disputes，不自行调分掩盖分歧。

## 四、输出格式

写到 输出/judgments.json：

```json
{{
  "award": "{award}",
  "judgments": {{
    "<维度id>": {{"score": 0, "grade": "A|B|C|D", "rationale": "一句话判据", "quotes": ["逐字片段"]}}
  }},
  "disputes": ["含糊、需答辩追问的点"]
}}
```

## 五、铁律

- 不看任何已有 score.json / judgments.json，工作区里没有，也不许去别处找。
- 不问目标分数，不猜别人给过多少分，只按证据给分。
- 文件名里的数字是版本号，不是分数，忽略。
- 用 python 校验每条 quote 确实出现在材料文本里；找不到的删掉或换成真实存在的。
"""


def build_rules(award, name):
    dims = DIMENSIONS.get(award)
    if not dims:
        raise SystemExit("未知奖项：%s（可用：%s）" % (award, "/".join(DIMENSIONS)))
    table = "\n".join("| `%s` | %s | %d |" % (i, n, m) for i, n, m in dims)
    header = "| 维度 id | 名称 | 满分 |\n|---|---|---|"
    return RULES.format(name=name, award=award, dim_table=header + "\n" + table)


def leaks_in(text):
    return [p.pattern for p in LEAK_PATTERNS if p.search(text)]


def prepare(ws, material, award, name):
    ws = Path(ws)
    if ws.exists():
        print("目录已存在，先删除或换名：%s" % ws, file=sys.stderr)
        return 2
    (ws / "材料").mkdir(parents=True)
    (ws / "输出").mkdir()

    src = Path(material)
    if not src.is_file():
        print("找不到材料：%s" % src, file=sys.stderr)
        return 2
    dst = ws / "材料" / (name + src.suffix.lower())
    shutil.copy2(src, dst)

    (ws / "规则.md").write_text(build_rules(award, name), encoding="utf-8")

    problems = []
    for root, dirs, files in os.walk(ws):
        for fn in files:
            p = Path(root) / fn
            rel = str(p.relative_to(ws))
            if leaks_in(fn):
                problems.append("文件名: " + rel)
            if p.suffix.lower() in (".md", ".json", ".txt", ".yaml"):
                try:
                    c = p.read_text(encoding="utf-8")
                except Exception:
                    continue
                if leaks_in(c):
                    problems.append("内容: " + rel)
    for pat in ("*.judgments.json", "*.score.json"):
        if list((ws / "输出").glob(pat)):
            problems.append("输出区残留 " + pat)

    n_files = sum(1 for _, _, fs in os.walk(ws) for _ in fs)
    result = {
        "workspace": str(ws),
        "award": award,
        "input_files": n_files,
        "contents": sorted(str(p.relative_to(ws)) for p in ws.rglob("*") if p.is_file()),
        "blind": not problems,
        "leaks": problems[:20],
        "note": "blind=true 才可派子智能体。",
    }
    print(json.dumps(result, ensure_ascii=False, indent=2))
    return 0 if not problems else 1


def verify(ws):
    ws = Path(ws)
    problems = []
    for root, dirs, files in os.walk(ws):
        for fn in files:
            p = Path(root) / fn
            if leaks_in(fn):
                problems.append("文件名: " + str(p.relative_to(ws)))
            if p.suffix.lower() in (".md", ".json", ".txt", ".yaml"):
                try:
                    c = p.read_text(encoding="utf-8")
                except Exception:
                    continue
                if leaks_in(c):
                    problems.append("内容: " + str(p.relative_to(ws)))
    n_out = len(list((ws / "输出").glob("*.judgments.json")))
    out = {"blind": not problems, "leaks": problems[:20],
           "judgments_written": n_out,
           "verdict": "盲区干净" if not problems else "存在泄露"}
    print(json.dumps(out, ensure_ascii=False, indent=2))
    return 0 if not problems else 1


def main() -> int:
    ap = argparse.ArgumentParser(description="最小盲区（只给 docx + 规则卡）")
    ap.add_argument("workspace")
    ap.add_argument("--material")
    ap.add_argument("--award", default="金智囊")
    ap.add_argument("--name", default=DEFAULT_NAME)
    ap.add_argument("--verify", action="store_true")
    a = ap.parse_args()
    if a.verify:
        return verify(Path(a.workspace))
    if not a.material:
        print("prepare 需要 --material", file=sys.stderr)
        return 2
    return prepare(Path(a.workspace), a.material, a.award, a.name)


if __name__ == "__main__":
    raise SystemExit(main())