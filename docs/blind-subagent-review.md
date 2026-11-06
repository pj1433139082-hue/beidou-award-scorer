# 隔离评审：用无上下文子智能体写 judgments

> 状态：已实现（`scripts/blind_prepare.py`、`scripts/blind_minimal.py`），已实测。
> 适用范围：需要**独立复核**的评分，或对同一次评分有「是否自己抄自己」疑虑时。

## 为什么需要这一层

`isolate.py` 隔离的是**文件系统**：把历史 judgments / score 挪走，让后跑的
agent 在磁盘上找不到先跑的答案。2026-11-05 实测它封存了 51 个目录，
`check` 返回 `clean: true`。

但它挡不住另一条通道 —— **写 judgments 的那个 agent 自己的上下文**。
若主会话刚改完稿、刚给过某个分数，再由它来写独立判断，文件系统藏得再干净
也没用：它记得自己给过什么。这不是「不诚实」，是上下文本来就带着记忆。

实测差距：同一份《东部智脑》材料

| 评审者 | 分数 | 区间 |
|---|---|---|
| 主会话（带改稿记忆） | 89.3 | 一等奖候选 |
| 无上下文子智能体（盲评） | 57.0 | 入围候选 |

差 32.3 分、三个档位。盲评还抓到了主会话因为「作者视角」忽略的
**材料内部矛盾**：第（七）节称跨区域推广为「规划在建」，附录（十一）（十六）
却称「已在四个大区同时运行」。这类矛盾只有没参与改稿的评审者才看得见。

## 三层隔离，缺一不可

| 层 | 挡住什么 | 工具 |
|---|---|---|
| 1. 文件系统 | 磁盘上找得到历史答案 | `isolate.py prepare / check / end` |
| 2. 上下文 | agent 记得自己给过什么分 | 派 `subagent`（全新上下文），主会话不写 judgments |
| 3. 输入面 | 文件名、元数据、文档里的历史分数 | `blind_prepare.py` / `blind_minimal.py` |

**第 3 层最容易被忽略。** 实测踩过的坑：

- 材料复制进盲区，文件名仍叫「（89.3终版）.docx」；
- `checks.json` / `doc.json` 的 `name`、`source` 字段也带着该名字，
  子智能体读机检文件第一眼就看到分数；
- `skill/docs/judgments-rules.md` 里写着「同一批材料在这台电脑打 59.4 分，
  另一台打 41.5 分」，复制整个 skill 时历史分数会顺着文档泄露；
- `rubric.example.json` 有 13 处示例锚定。

所以：文件名要中性化、元数据要洗、`docs/` 不复制。

## 两种盲区，按需要选

| | `blind_prepare.py` | `blind_minimal.py` |
|---|---|---|
| 输入文件 | 36 个（含机检、技能副本） | **2 个**（docx + 规则.md） |
| 给子智能体 | 机检结论 + 评分卡 + 规则 | 只有材料和纯规则 |
| 引文来源 | `doc.json` 的 plain | 直接从 docx 抽取 |
| 机检 | 子智能体能看到 | 主会话在盲区外跑，子智能体不接触 |
| 适合 | 想让评审者参考机检 warn/fail | 追求最大隔离、纯人工判断 |

**默认推荐 `blind_minimal.py`。** 输入面越小，泄露面越小。
实测 16 条引文在 docx 直抽文本里 16/16 命中（对比 `doc.json` 的 plain 也是
16/16），但表格内文字要小心：两者长度不同（20314 vs 20718），`plain` 含表格内容。
主会话拿到 judgments 后仍用 `score.py` 校验引文，查不到的会被拒收并自动降级。

## 标准流程

```bash
# 1. 主会话：准备最小盲区（材料名中性化，输入只有 2 个文件）
python scripts/blind_minimal.py <盲区> --material <材料.docx> --award 金智囊
python scripts/blind_minimal.py <盲区> --verify      # 必须 blind: true

# 2. 主会话：派一个 subagent（全新上下文），prompt 里只给盲区路径
#    prompt 铁律：不看任何历史 score/judgments、不问目标分数、不猜别人给过分
#    子智能体在 输出/judgments.json 写判断，逐条校验 (q in 材料文本)

# 3. 子智能体完成后，主会话（不修改它的判断）：
python scripts/score.py <doc.json> --checks <checks.json> \
    --judgments <盲区>/输出/judgments.json --award 金智囊 \
    --rubric <rubric.json> --out <输出> --require-bound

#    --require-bound 前先 bind_judgments.py 绑定材料指纹

# 4. 复核盲区未被污染、对比两次结果
python scripts/blind_minimal.py <盲区> --verify
python scripts/isolate.py audit <结果A> <结果B>
```

## 子智能体 prompt 模板

```text
你是北斗奖评审员，独立完成一次评分判断。你没有任何上下文，也不要去找。

你的全部输入（只有两个文件）：
  材料：<盲区>/材料/申报主报告.docx
  规则：<盲区>/规则.md

先完整读规则，再读 docx 提取正文与表格文字。按规则为各维度写判断，
输出到 <盲区>/输出/judgments.json。

严格核对每条 A 档 all_of；quotes 必须逐字存在，用 python 做 (q in text) 校验；
score 不得超过 满分 x 该 grade 的 cap；不确定的写进 disputes。

铁律：不看任何已有 score.json / judgments.json；不问目标分数；
不猜别人给过多少分；不修改规则与材料；不因为材料写得长就抬分。
```

## 判读两次结果的差异

差异**不是错误**，是信息。用 `isolate.py audit` 看一致率：

- ≥95% 相同 → 可疑，几乎不可能是独立判断（多半共享了输入）；
- 70~95% → 可能共享了部分输入；
- <70% → 独立。

出现大幅差异时，逐维度比对 `rationale`：通常是**证据等级判定**不同，
而不是算术不同。把两份判据并排给评审委员会，比取平均分更有价值。

## 不要做的事

- 不要让主会话「参考」盲评结果后重写 judgments —— 那又变成同一个上下文了。
- 不要在盲区放 `docs/`、示例评分卡、历史报告。
- 不要给子智能体看机检结论后又称「纯独立」（可用，但要如实标注口径）。
- 不要用 `--from-scores` 复用旧分数冒充新一轮判断。

## 已验证

```
blind_minimal.py prepare     blind: true，输入 2 个文件
blind_minimal.py --verify    注入「（89.3终版）说明.txt」后立即 blind: false
blind_prepare.py             抓到 docs/ 历史分数、doc.json 元数据泄露
引文交叉验证                16/16 在 docx 与 doc.json 双向命中
盲评 vs 主会话              57.0 vs 89.3，差 32.3
```