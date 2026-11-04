"""MaterialRecord 的单元测试。

只测已确认的接缝 S1 / S2 / S3；文件读取属于集成测试，见 selftest_pipeline.py。
"""
import sys
from pathlib import Path

import pytest

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "scripts"))

from material_record import (  # noqa: E402
    CorruptRecord,
    MaterialRecord,
    same_material,
)


def _raw(**kw):
    """一份最小可用的评分结果。"""
    base = {
        "name": "金火炬-水厂设备卫士——污水厂关键设备监测",
        "award": "金火炬",
        "ai_score": 65.2,
        "band": "三等奖候选",
        "confidence": "中",
        "dimensions": [
            {"id": "promotion", "name": "推广力度", "score": 30.0, "max": 40,
             "verified_grade": "C"},
            {"id": "economics", "name": "经济效益", "score": 20.0, "max": 40,
             "verified_grade": "B"},
        ],
    }
    base.update(kw)
    return base


# ---------- S1：缺 name 必须拒绝 ----------

def test_缺_name_时拒绝构造():
    """ADR-0003：材料身份用正文 name。缺了就不能猜，否则一份结果会冒充另一份。"""
    raw = _raw()
    del raw["name"]
    with pytest.raises(CorruptRecord):
        MaterialRecord(raw)


def test_空_name_也拒绝():
    """空串和缺失等价——不能因为字段在就放过。"""
    with pytest.raises(CorruptRecord):
        MaterialRecord(_raw(name=""))


def test_有_name_时构造成功并读出属性():
    rec = MaterialRecord(_raw())
    assert rec.name == "金火炬-水厂设备卫士——污水厂关键设备监测"
    assert rec.award == "金火炬"
    assert rec.total == 65.2
    assert rec.band == "三等奖候选"
    assert len(rec.dimensions) == 2


def test_缺_ai_score_时总分为零而非报错():
    """总分缺失是「未评分」，不是「损坏」——损坏的判据是缺身份。"""
    raw = _raw()
    del raw["ai_score"]
    assert MaterialRecord(raw).total == 0


# ---------- S2：dim 取不到返回空 dict ----------

def test_dim_取到指定维度():
    rec = MaterialRecord(_raw())
    d = rec.dim("promotion")
    assert d["name"] == "推广力度"
    assert d["score"] == 30.0


def test_dim_取不到时返回空_dict_而不是_None():
    """调用方会直接 d.get(...)；返回 None 会让它崩，空 dict 才是安全默认值。"""
    rec = MaterialRecord(_raw())
    missing = rec.dim("no_such_dimension")
    assert missing == {}
    assert missing.get("score") is None  # 空 dict 上 .get 不抛异常


# ---------- S3：短名精确还原 ----------

def test_去掉奖项前缀后能对上短名():
    assert same_material("金火炬-水厂设备卫士", "水厂设备卫士")


def test_全名等于短名时也算同一份():
    assert same_material("水厂设备卫士", "水厂设备卫士")


def test_不把前缀截断当成另一份材料():
    """这是当初模糊匹配补丁的真正危险：'水厂设备' 是 '水厂设备卫士' 的前缀，
    模糊匹配会认为它们是同一份，而实际是两份不同材料。"""
    assert not same_material("金火炬-水厂设备卫士", "水厂设备")


def test_不同材料不会互相顶替():
    assert not same_material("金火炬-水厂设备卫士", "金刚钻-从感觉还行到数据说话")


def test_空值不匹配():
    assert not same_material("", "水厂设备卫士")
    assert not same_material("金火炬-水厂设备卫士", "")