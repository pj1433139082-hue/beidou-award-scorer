#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""AdviceSheet —— 汇总表里建议区的布局，只在这里定义一次。

## 为什么有这个模块

写入端 export_xlsx.py 按维度数动态算列位置，读取端却有**两处各自猜**：

    modes.py             row[9] / row[10] / row[11]     直接索引
    check_compliance.py  _col(row, 9) / _col(row, 10)   有越界保护

两处的正则逐字相同，取列方式却不一样——drift 已经发生。而写入端：

    COL_ORIG = COL_AFTER + 1        # 跟着维度数走

也就是说 9/10/11 这三个硬编码索引，只在「维度数恰好 5」时成立。
**写入端算列、读取端猜列，中间没有任何约定。**

同一组表头还被三种不同语义匹配：写入端写死、verify_format 子串匹配、
selftest 集合相等。

## 布局

    列 1        材料名 / 建议编号标题
    列 2        总分
    列 3        奖项区间
    列 4..4+n-1 各维度得分（n = 该批次最大维度数）
    列 +0       改完提升
    列 +1       改完总分
    列 +2       原文是什么
    列 +3       怎么改
    列 +4       修改之后的句子

前三列与建议区三列的位置固定，中间维度区宽度随批次变化——这就是为什么
读取端不能硬编码列号。

## 编号格式

建议条目在表里写作「　1-M01　价值性　标题」（全角空格分隔）。这个串是
Excel 与 Word 对齐的唯一主键，之前被 6 份正则、4 种变体重复解析。
现在格式与解析都只在这里。
"""
from __future__ import annotations
import re
from pathlib import Path

# 固定列（不随维度数变化）
COL_NAME = 1
COL_SCORE = 2
COL_BAND = 3
COL_DIM0 = 4          # 维度区起点

HEAD_NAME = "申报材料 / 修改建议"
HEAD_SCORE = "总分"
HEAD_BAND = "奖项区间"
HEAD_DIM = "维度得分（满分）"
HEAD_GAIN = "改完提升"
HEAD_AFTER = "改完总分"
HEAD_ORIG = "原文是什么"
HEAD_HOW = "怎么改"
HEAD_FIX = "修改之后的句子"

DATA_START_ROW = 5    # 表头在第 4 行，数据从第 5 行起
HEAD_ROW = 4

# 建议条目标题行的样子：「　1-M01　价值性　标题」
# 前导全角空格是 merge 单元格的视觉缩进，实际写进去的串就带着它。
ITEM_PAT = re.compile(
    r"^[\u3000\s]*(\d+-M\d\d)[\u3000\s]+([^\u3000\s]+)[\u3000\s]+(.+)")


def columns(max_dim: int) -> dict:
    """给定该批次最大维度数，算出全部列号。

    写入与读取都调它，所以两边不可能算出不同的列。
    """
    gain = COL_DIM0 + max_dim
    after = gain + 1
    orig = after + 1
    how = orig + 1
    fix = how + 1
    return {"name": COL_NAME, "score": COL_SCORE, "band": COL_BAND,
            "dim0": COL_DIM0, "max_dim": max_dim, "gain": gain,
            "after": after, "orig": orig, "how": how, "fix": fix,
            "last": fix}


def item_label(material_index: int, no: str, dim: str, title: str) -> str:
    """建议条目的标题串。格式唯一定义处。"""
    return "\u3000%d-%s\u3000%s\u3000%s" % (material_index, no, dim, title)


def parse_item_label(s: str):
    """解析标题串，返回 (编号, 维度, 标题)；不是建议行返回 None。"""
    m = ITEM_PAT.match(str(s or ""))
    if not m:
        return None
    return m.group(1), m.group(2), m.group(3)


def _col(row, i: int) -> str:
    """安全取列：索引越界返回空串。"""
    try:
        v = row[i]
        return "" if v is None else str(v)
    except (IndexError, TypeError):
        return ""


def read(path) -> dict:
    """读汇总表，返回 {材料短名: [建议条目]}。

    材料短名是表里第一列的原样（已去掉奖项前缀），调用方若要匹配材料身份，
    用 material_record.same_material() 还原。

    列号由表头反查，不硬编码：维度数不同的批次列会右移。
    """
    try:
        from openpyxl import load_workbook
    except ImportError:
        return {}
    p = Path(path)
    if p.is_dir():
        p = p / "北斗奖评分汇总.xlsx"
    if not p.is_file():
        return {}
    wb = load_workbook(p)
    ws = wb[wb.sheetnames[0]]

    # 从表头反查三列建议区的位置，而不是假定 9/10/11
    pos = {}
    for c in range(1, ws.max_column + 1):
        h = str(ws.cell(HEAD_ROW, c).value or "").strip()
        if h == HEAD_ORIG:
            pos["orig"] = c
        elif h == HEAD_HOW:
            pos["how"] = c
        elif h == HEAD_FIX:
            pos["fix"] = c
    if not pos:
        return {}

    items, cur = {}, None
    for row in ws.iter_rows(min_row=DATA_START_ROW, values_only=True):
        first = str(row[0] or "") if row else ""
        parsed = parse_item_label(first)
        if parsed:
            no, dim, title = parsed
            if cur:
                items.setdefault(cur, []).append({
                    "no": no, "dim": dim, "title": title,
                    "original": _col(row, pos.get("orig", 0) - 1),
                    "how": _col(row, pos.get("how", 0) - 1),
                    "revised": _col(row, pos.get("fix", 0) - 1),
                })
        else:
            t = first.strip().replace("\u3000", " ").strip()
            # 主行是材料短名：跳过表头、奖项分组行、次级标题
            if t and "-M" not in t and not re.match(r"^[第\(（]", t):
                cur = t
    return items


def header_names() -> list:
    """汇总表必须出现的表头。verify_format 与 selftest 都从这里取。"""
    return [HEAD_NAME, HEAD_SCORE, HEAD_BAND, HEAD_GAIN, HEAD_AFTER,
            HEAD_ORIG, HEAD_HOW, HEAD_FIX]


def banned_headers() -> list:
    """明确不该出现的表头（曾经有、已被删掉的列）。"""
    return ["置信度", "需关注"]