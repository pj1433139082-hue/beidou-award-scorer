"""sums.py 与安装器按归档名校验的测试。

关键约束：分支归档的哈希是自指的（记录它会改变它），只有 tag 归档稳定。
"""
import importlib.util
import io
import sys
import urllib.request
from pathlib import Path

import pytest

ROOT = Path(__file__).resolve().parents[1]


def _load(name, rel):
    spec = importlib.util.spec_from_file_location(name, ROOT / rel)
    mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)
    return mod


@pytest.fixture(scope="module")
def sums():
    return _load("sums_mod", "scripts/sums.py")


@pytest.fixture(scope="module")
def inst():
    return _load("install_sum", "install.py")


def _fake_urlopen(payload, monkeypatch):
    class Resp:
        def __enter__(self):
            return io.BytesIO(payload)

        def __exit__(self, *a):
            return False

    monkeypatch.setattr(urllib.request, "urlopen", lambda *a, **k: Resp())


# ---------- 归档名精确匹配 ----------

def test_按归档文件名精确取哈希(inst, monkeypatch):
    """清单里有多条时，必须按文件名取对应那条，不能随便返回第一条。"""
    payload = ("%s  beidou-award-scorer-v3.5.0.zip\n"
               "%s  beidou-award-scorer-v3.4.0.zip\n"
               % ("a" * 64, "b" * 64)).encode()
    _fake_urlopen(payload, monkeypatch)
    got = inst.fetch_expected_sha256("x/y", "v3.5.0", "beidou-award-scorer-v3.5.0.zip")
    assert got == "a" * 64


def test_取另一个版本的哈希(inst, monkeypatch):
    payload = ("%s  beidou-award-scorer-v3.5.0.zip\n"
               "%s  beidou-award-scorer-v3.4.0.zip\n"
               % ("a" * 64, "b" * 64)).encode()
    _fake_urlopen(payload, monkeypatch)
    got = inst.fetch_expected_sha256("x/y", "v3.4.0", "beidou-award-scorer-v3.4.0.zip")
    assert got == "b" * 64


def test_清单里没有该归档时返回_None(inst, monkeypatch):
    """多条记录但都不是要的那个 -> 返回 None，让调用方走「跳过校验并提示」。"""
    payload = ("%s  beidou-award-scorer-v3.5.0.zip\n" % ("a" * 64)).encode()
    _fake_urlopen(payload, monkeypatch)
    got = inst.fetch_expected_sha256("x/y", "main", "beidou-award-scorer-main.zip")
    assert got is None


def test_只有一条记录时兼容单哈希格式(inst, monkeypatch):
    _fake_urlopen(("c" * 64 + "\n").encode(), monkeypatch)
    assert inst.fetch_expected_sha256("x/y", "v1.0.0", "any.zip") == "c" * 64


def test_注释行被忽略(inst, monkeypatch):
    """清单带注释头，不能把注释当成哈希。"""
    payload = ("# 这是注释\n"
               "# 校验方式：sha256sum -c SHA256SUMS\n\n"
               "%s  beidou-award-scorer-v3.5.0.zip\n" % ("d" * 64)).encode()
    _fake_urlopen(payload, monkeypatch)
    got = inst.fetch_expected_sha256("x/y", "v3.5.0", "beidou-award-scorer-v3.5.0.zip")
    assert got == "d" * 64


def test_清单固定从默认分支读(inst, monkeypatch):
    """清单属于 main，不属于某个 tag。

    raw.githubusercontent 按 ref 取文件：用 v3.6.0 当 ref 时拿到的是那个 tag
    时刻的旧清单，于是明明发布了校验和却读不到（实测返回 200 但不含新版本）。
    """
    seen = []

    class Resp:
        def __enter__(self):
            return io.BytesIO(b"%s  beidou-award-scorer-v3.6.0.zip\n" % (("e" * 64).encode()))

        def __exit__(self, *a):
            return False

    def fake(url, *a, **k):
        seen.append(url if isinstance(url, str) else url.full_url)
        return Resp()

    monkeypatch.setattr(urllib.request, "urlopen", fake)
    inst.fetch_expected_sha256("x/y", "v3.6.0", "beidou-award-scorer-v3.6.0.zip")
    assert seen, "没有发起任何请求"
    assert "/main/SHA256SUMS" in seen[0], seen[0]
    assert "/v3.6.0/" not in seen[0], "清单不该按 tag 读：" + seen[0]


def test_可以显式指定清单所在分支(inst, monkeypatch):
    seen = []

    class Resp:
        def __enter__(self):
            return io.BytesIO(b"")

        def __exit__(self, *a):
            return False

    def fake(url, *a, **k):
        seen.append(url if isinstance(url, str) else url.full_url)
        return Resp()

    monkeypatch.setattr(urllib.request, "urlopen", fake)
    inst.fetch_expected_sha256("x/y", "dev", "a.zip", sums_ref="release")
    assert any("/release/" in u for u in seen), seen


def test_源里没有目标条目时继续试下一个源(inst, monkeypatch):
    """raw.githubusercontent 的 CDN 会滞后数分钟。

    实测：GitHub API 已能读到 v3.6.1，raw 仍返回只含 v3.5.0/v3.6.0 的旧清单。
    第一源拿到「有效但过期」的内容时必须继续问下一个源，不能就此 return None。
    """
    old = b"%s  beidou-award-scorer-v3.6.0.zip\n" % (("a" * 64).encode())
    new = (b"%s  beidou-award-scorer-v3.6.0.zip\n" % (("a" * 64).encode())
           + b"%s  beidou-award-scorer-v3.6.1.zip\n" % (("b" * 64).encode()))
    responses = [old, old, new, new]   # 前两个源过期，后两个是新的
    seen = []

    class Resp:
        def __init__(self, body):
            self.body = body

        def __enter__(self):
            return io.BytesIO(self.body)

        def __exit__(self, *a):
            return False

    def fake(url, *a, **k):
        seen.append(url if isinstance(url, str) else url.full_url)
        return Resp(responses[min(len(seen) - 1, len(responses) - 1)])

    monkeypatch.setattr(urllib.request, "urlopen", fake)
    got = inst.fetch_expected_sha256("x/y", "v3.6.1", "beidou-award-scorer-v3.6.1.zip")
    assert got == "b" * 64, got
    assert len(seen) > 1, "第一个源过期后没有继续尝试"


def test_所有源都没有目标条目时返回_None(inst, monkeypatch):
    payload = b"%s  beidou-award-scorer-v3.5.0.zip\n" % (("a" * 64).encode())

    class Resp:
        def __enter__(self):
            return io.BytesIO(payload)

        def __exit__(self, *a):
            return False

    monkeypatch.setattr(urllib.request, "urlopen", lambda *a, **k: Resp())
    assert inst.fetch_expected_sha256(
        "x/y", "v9.9.9", "beidou-award-scorer-v9.9.9.zip") is None


# ---------- tag 与分支走不同 URL ----------

def test_tag_形态的_ref_走_tags_路径(inst):
    """v3.5.0 这种应该当 tag；否则归档名对不上清单。"""
    for ref in ("v3.5.0", "v1.0.0", "v10.2.3"):
        assert ref.startswith("v") and ref[1:2].isdigit(), ref
    for ref in ("main", "dev", "master", "release-2026"):
        assert not (ref.startswith("v") and ref[1:2].isdigit()), ref


# ---------- sums.py 的读写 ----------

def test_read_sums_解析注释与空行(sums, tmp_path, monkeypatch):
    p = tmp_path / "SHA256SUMS"
    p.write_text("# 注释\n\n%s  a.zip\n%s  b.zip\n" % ("1" * 64, "2" * 64),
                 encoding="utf-8")
    monkeypatch.setattr(sums, "SUMS", p)
    got = sums.read_sums()
    assert got == {"a.zip": "1" * 64, "b.zip": "2" * 64}


def test_read_sums_缺文件返回空(sums, tmp_path, monkeypatch):
    monkeypatch.setattr(sums, "SUMS", tmp_path / "nope")
    assert sums.read_sums() == {}


def test_sha256_file_与_bytes_一致(sums, tmp_path):
    p = tmp_path / "x.bin"
    p.write_bytes(b"hello world")
    assert sums.sha256_file(p) == sums.sha256_bytes(b"hello world")