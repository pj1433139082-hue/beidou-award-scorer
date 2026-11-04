"""AdviceSheet 的单元测试。

只测 S4 / S5；读写真实 xlsx 属于集成测试。
"""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "scripts"))

from advice_sheet import (  # noqa: E402
    COL_NAME, COL_SCORE, COL_BAND, COL_DIM0,
    columns, item_label, parse_item_label, header_names, banned_headers,
)


# ---------- S4：列号随维度数右移 ----------

def test_列号随维度数右移():
    """写入端按维度数算列；硬编码 9/10/11 只在维度数=5 时成立。"""
    c3, c4, c5 = columns(3), columns(4), columns(5)
    assert c3["orig"] == 9
    assert c4["orig"] == 10
    assert c5["orig"] == 11
    # 每多一个维度，建议区整体右移一列
    assert c4["orig"] - c3["orig"] == 1
    assert c5["orig"] - c4["orig"] == 1


def test_三列建议区连续():
    """原文 / 怎么改 / 改后句子 必须紧挨着，中间不能跳列。"""
    for n in (2, 3, 5, 7):
        c = columns(n)
        assert c["how"] == c["orig"] + 1
        assert c["fix"] == c["how"] + 1
        assert c["last"] == c["fix"]


def test_固定前四列不随维度数变化():
    for n in (2, 5, 9):
        c = columns(n)
        assert c["name"] == COL_NAME == 1
        assert c["score"] == COL_SCORE == 2
        assert c["band"] == COL_BAND == 3
        assert c["dim0"] == COL_DIM0 == 4


def test_维度区宽度等于最大维度数():
    assert columns(4)["gain"] == COL_DIM0 + 4
    assert columns(7)["gain"] == COL_DIM0 + 7


# ---------- S5：编号格式往返 ----------

def test_编号格式往返一致():
    lab = item_label(1, "M01", "价值性", "缺基准值")
    assert parse_item_label(lab) == ("1-M01", "价值性", "缺基准值")


def test_编号含两位材料序号也能往返():
    lab = item_label(19, "M12", "推广力度", "推广证据不足")
    assert parse_item_label(lab) == ("19-M12", "推广力度", "推广证据不足")


def test_标题里含空格仍能解析():
    """标题是最后一段，里面的空格不该被当成分隔符。"""
    lab = item_label(3, "M02", "创新性", "第 2 代 与 第 3 代 对比缺失")
    got = parse_item_label(lab)
    assert got is not None
    assert got[0] == "3-M02" and got[1] == "创新性"
    assert got[2] == "第 2 代 与 第 3 代 对比缺失"


def test_非建议行返回_None():
    """材料名行、奖项分组行、空行都不能被当成建议条目。"""
    assert parse_item_label("水厂设备卫士——污水厂关键设备监测") is None
    assert parse_item_label("金火炬　·　2 份") is None
    assert parse_item_label("") is None
    assert parse_item_label("按分数降序") is None


def test_没有M编号的行不算建议():
    assert parse_item_label("　1-01　价值性　标题") is None


# ---------- 表头契约 ----------

def test_表头清单含全部必需列():
    need = header_names()
    assert "原文是什么" in need
    assert "怎么改" in need
    assert "修改之后的句子" in need
    assert "总分" in need


def test_已删的列在禁止清单里():
    """置信度与需关注是被有意删掉的列，不该再出现。"""
    banned = banned_headers()
    assert "置信度" in banned
    assert "需关注" in banned
    # 强制清单里不能又出现它们
    assert not set(banned) & set(header_names())