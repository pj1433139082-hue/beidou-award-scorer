"""CheckCatalog 的单元测试。

只测 S6 与 register 的契约；与 checks.py 的对齐属于集成测试。
"""
import sys
from pathlib import Path

import pytest

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "scripts"))

import check_catalog as CC  # noqa: E402


# ---------- S6：未注册必须报错，不能静默 ----------

def test_require_未注册时抛_KeyError():
    """这是本模块存在的核心理由。

    `.get()` 返回 None 让「忘了写定义」和「本来就没有」无法区分——
    C01/C10 就是这样静默丢了建议。require() 必须抛。
    """
    with pytest.raises(KeyError):
        CC.require("ZZZ_NOT_A_REAL_CHECK")


def test_require_报错信息点出该改哪个文件():
    """报错要能直接指导修复，而不是只说找不到。"""
    with pytest.raises(KeyError) as ei:
        CC.require("ZZZ_NOT_A_REAL_CHECK")
    msg = str(ei.value)
    assert "ZZZ_NOT_A_REAL_CHECK" in msg
    assert "check_catalog" in msg


def test_get_未注册时返回_None():
    """get() 仍允许探测性地问「有没有」，与 require() 分工不同。"""
    assert CC.get("ZZZ_NOT_A_REAL_CHECK") is None


# ---------- 已注册项的行为 ----------

def test_require_已注册时返回完整定义():
    spec = CC.require("C03")
    assert set(spec) == {"dims", "title", "how", "revised", "hints", "points"}
    assert spec["title"]
    assert spec["how"]
    assert spec["revised"]


def test_访问器与定义一致():
    spec = CC.CATALOG["C03"]
    assert CC.dims_of("C03") == spec["dims"]
    assert CC.hints_of("C03") == spec["hints"]
    assert CC.points_of("C03") == spec["points"]


def test_未注册项的访问器给安全默认值():
    """访问器用于「有就取、没有就走默认」的场景，不该抛。"""
    assert CC.dims_of("ZZZ") == set()
    assert CC.hints_of("ZZZ") == []
    assert CC.points_of("ZZZ") == 1.0


# ---------- 那两个曾丢建议的检查 ----------

def test_C01_与_C10_必须有定义():
    """这两条曾经在 CHECKS 里有、CATALOG 里没有，命中时建议被静默丢弃。

    19 份材料里 C01 命中 2 份、C10 命中 1 份，全部丢失建议。
    """
    for cid in ("C01", "C10"):
        spec = CC.require(cid)
        assert spec["title"], "%s 缺 title" % cid
        assert spec["how"], "%s 缺 how" % cid
        assert spec["revised"], "%s 缺 revised" % cid


def test_每条定义都有非空文案():
    """任何一条缺文案，命中时就会静默丢建议。"""
    empty = [cid for cid, s in CC.CATALOG.items()
             if not (s["title"] and s["how"] and s["revised"])]
    assert empty == [], "以下检查缺文案：%s" % empty


def test_每条定义都至少归属一个维度():
    """不归属任何维度的检查不会产出建议（dims 为空的条目需显式确认）。"""
    orphan = [cid for cid, s in CC.CATALOG.items() if not s["dims"]]
    assert orphan == [], "以下检查没有归属维度：%s" % orphan