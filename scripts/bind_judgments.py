#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""给 judgments.json 绑定材料指纹，让它只能用在当初判断的那份材料上。

## 解决什么

judgments.json 里装的是答案（score / grade / rationale / quotes）。
如果它和材料放一起，任何进这个目录的人或 agent 一条 ls 就全看见了，
直接抄就是同样的分数——这正是之前两个 agent 结果 94.7% 相同的成因。

绑定之后：

    材料被改过      → 指纹变 → score.py 拒收
    换一台机器跑    → 重跑 batch.py 得到的抽取结果指纹相同（确定性），但
                      若 material 不同则拒收
    从别处抄来的    → 指纹对不上 → 拒收

## 用法

    # 绑定整个目录（每个 xxx.judgments.json 找同名的 xxx.doc.json）
    python scripts/bind_judgments.py <目录>

    # 只绑定一份
    python scripts/bind_judgments.py <目录> --only 水厂设备卫士

    # 在写 judgments 时顺手绑定（供 agent 调用）
    python scripts/judgments_stamp.py <材料.doc.json> <judgments.json>

绑定信息包含：
    material_sha256   材料指纹（硬校验，不匹配就拒收）
    material_name     材料名（信息性）
    judged_at         绑定时间
    tool              工具版本标记
"""
from __future__ import annotations
import argparse, datetime, json, re, sys
from pathlib import Path

TOOL = "beidou-award-scorer/bind_judgments"


def fingerprint(doc: dict) -> str:
    """与 score.py 的 material_fingerprint 必须完全一致。"""
    import hashlib
    norm = re.sub(r"\s+", "", str((doc or {}).get("plain") or ""))
    return hashlib.sha256(norm.encode("utf-8")).hexdigest()


def stamp(doc_path: Path, judg_path: Path) -> dict:
    doc = json.loads(doc_path.read_text(encoding="utf-8"))
    j = json.loads(judg_path.read_text(encoding="utf-8"))
    j["material_sha256"] = fingerprint(doc)
    j["material_name"] = doc.get("name") or doc_path.stem
    j["judged_at"] = datetime.datetime.now().astimezone().isoformat(timespec="seconds")
    j["tool"] = TOOL
    judg_path.write_text(json.dumps(j, ensure_ascii=False, indent=1), encoding="utf-8")
    return {"material": j["material_name"],
            "fingerprint": j["material_sha256"][:16],
            "judgments": str(judg_path)}


def main() -> int:
    ap = argparse.ArgumentParser(description="给 judgments 绑定材料指纹")
    ap.add_argument("src", help="含 *.judgments.json 与 *.doc.json 的目录")
    ap.add_argument("--only", help="只处理文件名含该关键词的")
    a = ap.parse_args()

    src = Path(a.src)
    if not src.is_dir():
        print("目录不存在：%s" % src, file=sys.stderr)
        return 2

    judg_files = sorted(src.glob("*.judgments.json"))
    if a.only:
        judg_files = [f for f in judg_files if a.only in f.name]
    if not judg_files:
        print("没有 *.judgments.json", file=sys.stderr)
        return 2

    done, skipped = [], []
    for jf in judg_files:
        stem = jf.name.replace(".judgments.json", "")
        df = src / (stem + ".doc.json")
        if not df.is_file():
            skipped.append((jf.name, "找不到同名 .doc.json"))
            continue
        try:
            done.append(stamp(df, jf))
        except Exception as e:
            skipped.append((jf.name, str(e)[:60]))

    print(json.dumps({"bound": len(done), "skipped": len(skipped),
                      "items": done[:20]}, ensure_ascii=False, indent=1))
    for n, why in skipped[:10]:
        print("  跳过 %s：%s" % (n[:50], why), file=sys.stderr)
    return 0 if done else 1


if __name__ == "__main__":
    raise SystemExit(main())