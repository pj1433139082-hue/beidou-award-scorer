# 扩展评分卡

## 换赛事 / 换年份
只改 `assets/rubric.json`：

1. `edition` 改成新届次，`source` 写清新依据文件与页码。
2. `awards.<奖项>.dimensions` 里的 `max` 之和必须等于该奖项总分（当前都是 100）。
3. 维度 `id` 保持稳定（`value`/`replication`/`novelty`/`intelligence`/`usability`/`promotion`/`economics`/`readiness`），否则 `judgments.json` 会失效。
4. `gates`/`deductions` 里的 `effect` 只是文档说明，真正生效的是 `scripts/score.py: apply_machine()` 的映射——改闸门要同步改代码。

## 加一个新奖项（如"AI 融创奖"）

在 `awards` 下加一条：

```json
"AI融创": {
  "track": "特别奖",
  "full_name": "AI融创奖",
  "min_chars": 2500,
  "thesis": "业务逻辑成熟、AI 成熟度欠佳、应用潜力突出的高潜个人及团队",
  "dimensions": [
    {"id": "value", "name": "业务价值", "max": 60, "official": "实际ROI 40 + 预计年度ROI 20", "anchors": {}, "evidence": []},
    {"id": "replication", "name": "产出复用性", "max": 40, "official": "公众投榜 20 + 专业评审 20", "anchors": {}, "evidence": []}
  ]
}
```

再在 `scripts/checks.py` 加该奖项专属检查（参考 `c11_model_lab` 的写法），在 `scripts/batch.py` 的 `AWARDS` 列表里补上名字。

## 加一项机检

在 `scripts/checks.py` 写一个 `def cNN_xxx(doc, award): ... return R(...)`，然后加进 `CHECKS` 列表。
约束：
- 必须返回 `R(...)`，`dims` 填受影响的维度 id；
- 拿不准就返回 `na`，不要为了"有输出"硬判；
- `evidence` 必须能指回原文（表N/行M 或原文片段）。

## 换成别的材料格式
`scripts/extract_report.py` 目前支持 `.docx`/`.pdf`。加 `.md`/`.txt` 只需在 `build()` 里加一个分支，产出同样的 `blocks` 结构即可，下游不需要改。
