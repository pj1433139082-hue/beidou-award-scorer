"""install.py 的安全行为测试。

覆盖两条修复：
  1. 下载内容完整性校验（sha256 / SHA256SUMS）
  2. 执行下载来的代码前必须确认（confirm_trust）
"""
import hashlib
import importlib.util
import sys
from pathlib import Path

import pytest

ROOT = Path(__file__).resolve().parents[1]


def _load_install():
    spec = importlib.util.spec_from_file_location("install_mod", ROOT / "install.py")
    mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)
    return mod


@pytest.fixture(scope="module")
def inst():
    return _load_install()


# ---------- 完整性校验 ----------

def test_sha256_of_与标准库一致(inst):
    """不能自己实现摘要——算错了会让校验永远通过或永远失败。"""
    for data in (b"", b"hello", b"\x00\xff" * 100):
        assert inst.sha256_of(data) == hashlib.sha256(data).hexdigest()


def test_取不到清单时返回_None_而不是空串(inst, monkeypatch):
    """返回空串会让调用方误以为「期望值是空的」，从而永远匹配失败；
    None 才表示「没有可用清单」。"""
    import urllib.request

    def boom(*a, **k):
        raise OSError("network disabled")

    monkeypatch.setattr(urllib.request, "urlopen", boom)
    assert inst.fetch_expected_sha256("x/y", "main") is None


def test_能解析_SHA256SUMS_格式(inst, monkeypatch):
    """标准 sha256sum 输出是「哈希  文件名」。"""
    import io
    import urllib.request

    digest = "a" * 64
    payload = ("%s  beidou-award-scorer.zip\n" % digest).encode()

    class FakeResp:
        def __enter__(self):
            return io.BytesIO(payload)

        def __exit__(self, *a):
            return False

    monkeypatch.setattr(urllib.request, "urlopen", lambda *a, **k: FakeResp())
    assert inst.fetch_expected_sha256("x/y", "main") == digest


def test_能解析只有哈希一行的格式(inst, monkeypatch):
    import io
    import urllib.request

    digest = "b" * 64

    class FakeResp:
        def __enter__(self):
            return io.BytesIO((digest + "\n").encode())

        def __exit__(self, *a):
            return False

    monkeypatch.setattr(urllib.request, "urlopen", lambda *a, **k: FakeResp())
    assert inst.fetch_expected_sha256("x/y", "main") == digest


# ---------- 执行前确认 ----------

def test_本地目录来源不用确认(inst, monkeypatch):
    """dir 是用户自己的东西，问一遍只是噪音。

    必须把 isatty 设成 True —— 否则在 pytest 里 stdin 本来就不是 tty，
    这条断言会走「非交互直接放行」那条路，永远通过，测不到 dir 分支。
    （这个盲区是被故意破坏验证抓出来的：把 dir 判断删掉，测试居然还是绿的。）
    """
    monkeypatch.setattr(sys.stdin, "isatty", lambda: True, raising=False)
    monkeypatch.setattr("builtins.input",
                        lambda *a, **k: pytest.fail("不该向 dir 来源提问"))
    assert inst.confirm_trust("dir", False) is True


def test_yes_跳过确认(inst):
    """CI / 自动化要能显式跳过。"""
    assert inst.confirm_trust("github", True) is True
    assert inst.confirm_trust("zip", True) is True


def test_非交互环境放行并不阻塞(inst, monkeypatch):
    """CI 里 stdin 不是 tty，不能卡在 input() 上。"""
    monkeypatch.setattr(sys.stdin, "isatty", lambda: False, raising=False)
    assert inst.confirm_trust("github", False) is True


def test_交互环境用户拒绝时返回_False(inst, monkeypatch, capsys):
    """用户说不，就必须真的不执行——这是最后一道可控关口。"""
    monkeypatch.setattr(sys.stdin, "isatty", lambda: True, raising=False)
    monkeypatch.setattr("builtins.input", lambda *a, **k: "n")
    assert inst.confirm_trust("github", False) is False


def test_交互环境用户同意时返回_True(inst, monkeypatch):
    monkeypatch.setattr(sys.stdin, "isatty", lambda: True, raising=False)
    for ans in ("y", "yes", "Y", "是"):
        monkeypatch.setattr("builtins.input", lambda *a, **k: ans)
        assert inst.confirm_trust("github", False) is True


def test_输入中断时按取消处理(inst, monkeypatch):
    """Ctrl-C / EOF 不能被当成同意。"""
    monkeypatch.setattr(sys.stdin, "isatty", lambda: True, raising=False)

    def interrupted(*a, **k):
        raise EOFError()

    monkeypatch.setattr("builtins.input", interrupted)
    assert inst.confirm_trust("github", False) is False


# ---------- 确认信息必须说清风险 ----------

def test_确认信息点明来源与将要执行的脚本(inst, monkeypatch, capsys):
    """用户要能看见「这是什么、要跑什么」，才谈得上同意。"""
    monkeypatch.setattr(sys.stdin, "isatty", lambda: True, raising=False)
    monkeypatch.setattr("builtins.input", lambda *a, **k: "n")
    inst.confirm_trust("github", False)
    out = capsys.readouterr().out
    assert "setup.py" in out
    assert "GitHub" in out or "下载" in out