# ADR-0003 · 材料的身份是正文声明的 name，不是文件名

**状态**：已采纳（2026-11-03） · 相关：GLOSSARY（材料）

## 背景

五处读取评分结果的模块，对「用什么标识一份材料」给不出统一答案：

```
export_xlsx.py:80   stem = sf.name.replace(".score.json","")     # 文件名
export_docx.py:246  stem = sf.name.replace(".score.json","")     # 文件名
modes.py:44         out[x.get("name") or f.stem] = x             # 正文 name
consistency.py:35   out[x.get("name") or f.stem] = x             # 正文 name
isolate.py:185      out[x["name"]] = x["ai_score"]               # 正文 name，无 fallback
```

「两种键恰好相等」是靠命名约定维持的，不是契约。证据：`modes.py:139-148` 已经为此打了模糊匹配补丁（`k in name or name in k or k[:12] == name[:12]`）——**补丁的存在就是键不统一的物证**。

还有一个具体的坑：`batch.py:158` 在 `--from-json` 模式下必须用两套算法算 stem，因为 `Path("x.doc.json").stem` 得到的是 `"x.doc"`。

## 决定

材料的唯一身份是评分结果里的 `name` 字段（源于原始材料名去扩展名）。

文件名只用于**定位**文件，不用于**标识**材料。

## 理由

文件名是运输层，内容是事实层。用文件名当身份，会在「改个名」和「换种序列化格式」时全部失效。`--from-json` 那个坑正是这个后果。

选择正文 name 之后，`modes.py` 的模糊匹配补丁可以删掉——那补丁在修补这个病，而不是材料真的重名。

## 后果

- 缺少 `name` 字段的评分结果被视为损坏，报错而非静默兜底（`isolate.py` 已如此）
- 文件名与 name 不一致时，以 name 为准，并提示一次
