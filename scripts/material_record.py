#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""MaterialRecord —— 一份材料的评审结果，连同它的机检结论与原文。

## 为什么有这个模块

在它之前，有五个模块各自扫描 *.score.json 并各自决定「用什么标识一份材料」：

    export_xlsx.py      键 = 文件名 stem
    export_docx.py      键 = 文件名 stem
    modes.py            键 = 正文 name，或 stem
    consistency.py      键 = 正文 name，或 stem
    isolate.py          键 = 正文 name，缺了就崩

两种键「恰好相等」靠命名约定维持，不是契约。证据：modes.py 曾经为此打了模糊匹配
补丁（k in name or name in k or k[:12]==name[:12]）——补丁的存在就是键不统一的物证。

## 身份规则（ADR-0003）

材料的身份是评分结果里的 name 字段。文件名只用于**定位**文件，不用于**标识**材料。

缺 name 的评分结果视为损坏：报错，不猜。这是 ADR-0003 的后果条款——静默兜底会让
一份坏数据冒充成另一份材料，比直接失败更难查。

## 为什么还保留名字匹配

xlsx 汇总表第一列存的是 strip_award 之后的短名（给人看），而 name 带奖项前缀。
读表时要把它还原回 name。这里用**精确还原**（套用写入端同一个前缀规则），
不用模糊匹配——模糊匹配会把名字相近的两份材料搞混，而那种错误不会报错。
"""
from __future__ import annotations
import json, re
from pathlib import Path

AWARD_PREFIX = re.compile(
    r"^(金点子|金刚钻|金引擎|金智囊|金火炬|特别奖-[^\s\-－—]*)[-－—\s]*")


class CorruptRecord(Exception):
    """评分结果缺少必要字段，无法确定它属于哪份材料。"""


def strip_award(name: str) -> str:
    """去掉奖项前缀，得到给人看的短名。"""
    s = AWARD_PREFIX.sub("", str(name or "").strip())
    return re.sub(r"^[-－—\s]+", "", s).strip() or str(name or "")


def same_material(full_name: str, short_name: str) -> bool:
    """短名是否精确对应这份材料。

    精确还原而非模糊匹配：把 full_name 按写入端同一个规则去掉前缀，再比。
    模糊匹配（子串/前缀）会让「名字相近的两份材料」互相顶替而不报错。
    """
    if not full_name or not short_name:
        return False
    return full_name == short_name or strip_award(full_name) == short_name


class MaterialRecord:
    """一份材料的完整上下文。

    属性：
        name        材料身份（ADR-0003 指定用它）
        total       总分
        band        奖项区间
        award       奖项
        dimensions  维度列表（每个含 id/name/score/max/verified_grade/rationale）
        checks      机检结论 {check_id: check_dict}，可能为空
        plain       材料原文，可能为空
        advice      修改建议条目，可能为空
        stem        文件名 stem（仅用于定位文件，不作身份）
    """

    __slots__ = ("raw", "name", "award", "stem", "checks", "plain", "advice")

    def __init__(self, raw: dict, *, stem: str = "", checks=None, plain: str = ""):
        if not isinstance(raw, dict):
            raise CorruptRecord("评分结果不是 JSON 对象")
        name = raw.get("name")
        if not name:
            raise CorruptRecord(
                "评分结果缺少 name 字段，无法确定它属于哪份材料。\n"
                "  ADR-0003：材料身份用正文 name，不用文件名。缺 name 视为损坏，不猜。\n"
                "  文件：%s" % (stem or "（未知）"))
        self.raw = raw
        self.name = str(name)
        self.award = str(raw.get("award") or "")
        self.stem = stem
        self.checks = checks or {}
        self.plain = plain or ""
        self.advice = []

    # --- 常用属性，调用方不必知道 raw 的键名 ---
    @property
    def total(self) -> float:
        return float(self.raw.get("ai_score") or 0)

    @property
    def band(self) -> str:
        return str(self.raw.get("band") or "")

    @property
    def confidence(self) -> str:
        return str(self.raw.get("confidence") or "")

    @property
    def dimensions(self) -> list:
        return self.raw.get("dimensions") or []

    def dim(self, dim_id: str) -> dict:
        """按 id 取维度；取不到返回空 dict，调用方不必自己判空。"""
        for d in self.dimensions:
            if d.get("id") == dim_id:
                return d
        return {}

    def as_row(self) -> dict:
        """拼成导出器需要的 dict。

        导出器大量按 dict 取字段，迁移期保留这条适配路径。下划线字段的**名字**
        在这里定义一次——之前它们是调用方手工挂的，两边各写一遍。
        注意是 _items 不是 _advice：export_xlsx.py 全程按 _items 取。
        """
        d = dict(self.raw)
        d["_checks"] = self.checks
        d["_plain"] = self.plain
        d["_stem"] = self.stem
        d["_items"] = self.advice
        return d

    def __repr__(self) -> str:
        return "MaterialRecord(%s, %.1f)" % (self.name[:24], self.total)


def load_dir(src, *, with_context: bool = True, with_advice: bool = False) -> list:
    """读一个目录里所有评分结果。

    with_context  同时读 checks.json 与 doc.json（导出器需要）
    with_advice   同时算修改建议（只有 Excel/Word 导出需要）
    """
    src = Path(src)
    out = []
    for sf in sorted(src.glob("*.score.json")):
        try:
            raw = json.loads(sf.read_text(encoding="utf-8"))
        except Exception:
            continue
        if "dimensions" not in raw:
            continue
        stem = sf.name[: -len(".score.json")]
        checks, plain = {}, ""
        if with_context:
            cf = src / (stem + ".checks.json")
            if cf.is_file():
                try:
                    checks = {c["id"]: c for c in
                              (json.loads(cf.read_text(encoding="utf-8")).get("checks") or [])}
                except Exception:
                    checks = {}
            df = src / (stem + ".doc.json")
            if df.is_file():
                try:
                    plain = json.loads(df.read_text(encoding="utf-8")).get("plain", "")
                except Exception:
                    plain = ""
        rec = MaterialRecord(raw, stem=stem, checks=checks, plain=plain)
        if with_advice:
            try:
                from advice_engine import build_items
                rec.advice = build_items(rec.as_row())
            except Exception:
                rec.advice = []
        out.append(rec)
    return out


def load_index(src, *, recursive: bool = False) -> dict:
    """读一个目录，返回 {name: MaterialRecord}。

    用于只需要分数、不需要上下文的场景（一致性对比、隔离审计）。
    """
    src = Path(src)
    files = src.rglob("*.score.json") if recursive else src.glob("*.score.json")
    out = {}
    for f in sorted(files):
        try:
            raw = json.loads(f.read_text(encoding="utf-8"))
        except Exception:
            continue
        if "dimensions" not in raw:
            continue
        stem = f.name[: -len(".score.json")]
        try:
            rec = MaterialRecord(raw, stem=stem)
        except CorruptRecord:
            continue
        out[rec.name] = rec
    return out


def sort_by_total(recs: list) -> list:
    """按总分降序。排序规则唯一来源。"""
    return sorted(recs, key=lambda r: r.total, reverse=True)