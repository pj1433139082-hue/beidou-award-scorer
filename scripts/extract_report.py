#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""把申报材料（docx / pdf）抽成结构化 doc.json：按正文顺序保留 段落 + 标题 + 表格。"""
from __future__ import annotations
import argparse, json, os, re, sys
from pathlib import Path

def _clean(t: str) -> str:
    if t is None: return ""
    t = t.replace("\u3000", " ").replace("\xa0", " ").replace("\u200b", "")
    t = re.sub(r"[ \t]+", " ", t)
    return t.strip()

def _style_level(style_name: str) -> int:
    s = (style_name or "").lower()
    m = re.search(r"(?:heading|标题)\s*([1-9])", s)
    if m: return int(m.group(1))
    return 0

NUM_PREFIX = re.compile(r"^\s*(?:[一二三四五六七八九十]+[、.．]|[（(][一二三四五六七八九十0-9]+[)）]|[0-9]{1,2}(?:\.[0-9]{1,2}){0,2}\s*[、.．]?)\s*")

def _heading_from_text(t: str) -> int:
    """无样式时靠正文形态判标题：短、无句号、带编号或后缀冒号。"""
    t = t.strip()
    if not t or len(t) > 40: return 0
    if t.endswith(("。", "；", ";")): return 0
    if re.match(r"^(?:附件|附表|附录)", t): return 1
    if re.match(r"^[0-9]{1,2}\s*[、.．]\s*\S", t) and not re.search(r"[0-9]{2,}", t): return 2
    if re.match(r"^[（(][一二三四五六七八九十0-9]{1,3}[)）]", t): return 3
    if re.match(r"^(?:项目|创新|实施|考核|计划|推广|经济)[^\s]{0,10}(?:背景|介绍|分析|建议|内容|说明)$", t): return 1
    return 0

def docx_blocks(path: Path):
    import docx
    from docx.table import Table
    from docx.text.paragraph import Paragraph
    d = docx.Document(str(path))
    body = d.element.body
    blocks, tidx, pidx = [], 0, 0
    for child in body.iterchildren():
        tag = child.tag.split("}")[-1]
        if tag == "p":
            p = Paragraph(child, d)
            raw = p.text or ""
            txt = _clean(raw)
            if not txt: continue
            pidx += 1
            lvl = _style_level(p.style.name if p.style is not None else "")
            if lvl == 0: lvl = _heading_from_text(txt)
            if lvl:
                blocks.append({"type": "heading", "level": lvl, "text": txt, "n": pidx})
            else:
                blocks.append({"type": "para", "text": txt, "n": pidx})
        elif tag == "tbl":
            t = Table(child, d)
            rows, seen = [], set()
            for r in t.rows:
                cells = []
                for c in r.cells:
                    v = _clean((c.text or "").replace("\n", " / "))
                    if not cells or cells[-1] != v: cells.append(v)
                key = "|".join(cells)
                if key in seen: continue
                seen.add(key); rows.append(cells)
            if not rows: continue
            tidx += 1
            blocks.append({"type": "table", "index": tidx, "header": rows[0], "rows": rows[1:],
                           "n_rows": len(rows), "n_cols": max(len(r) for r in rows)})
    n_img = sum(1 for _ in d.inline_shapes)
    return blocks, {"images": n_img, "tables": tidx}

def pdf_blocks(path: Path):
    import fitz
    doc = fitz.open(str(path))
    blocks, pidx, tidx = [], 0, 0
    for page in doc:
        for line in (page.get_text() or "").splitlines():
            txt = _clean(line)
            if not txt: continue
            pidx += 1
            lvl = _heading_from_text(txt)
            blocks.append({"type": "heading" if lvl else "para", "level": lvl, "text": txt, "n": pidx} if lvl
                          else {"type": "para", "text": txt, "n": pidx})
    return blocks, {"images": 0, "tables": tidx, "pages": doc.page_count}

def build(src: Path) -> dict:
    ext = src.suffix.lower()
    if ext == ".docx": blocks, stats = docx_blocks(src)
    elif ext in (".pdf",): blocks, stats = pdf_blocks(src)
    else: raise SystemExit(f"不支持的格式：{ext}（只吃 .docx / .pdf）")
    # 章节路径
    stack, path_now = [], ""
    for b in blocks:
        if b["type"] == "heading":
            while stack and stack[-1][0] >= b["level"]: stack.pop()
            stack.append((b["level"], b["text"]))
            path_now = " > ".join(t for _, t in stack)
        b["section"] = path_now
    def _render(b):
        if b["type"] == "heading": return "[H%d] %s" % (b["level"], b["text"])
        if b["type"] == "table":
            head = " | ".join(b["header"])
            body = "\n".join(" | ".join(r) for r in b["rows"][:40])
            return "[表%d %d行]\n%s\n%s" % (b["index"], b["n_rows"], head, body)
        return b["text"]
    plain = "\n".join(_render(b) for b in blocks)
    chars = len(re.sub(r"\s", "", plain))
    return {"source": str(src), "name": src.stem, "kind": ext.lstrip("."),
            "stats": {**stats, "chars": chars, "blocks": len(blocks)}, "blocks": blocks, "plain": plain}

def main() -> int:
    ap = argparse.ArgumentParser(description="申报材料 → doc.json")
    ap.add_argument("inputs", nargs="+")
    ap.add_argument("--out", default=".")
    a = ap.parse_args()
    out = Path(a.out); out.mkdir(parents=True, exist_ok=True)
    for i in a.inputs:
        src = Path(i)
        d = build(src)
        dst = out / (src.stem + ".doc.json")
        dst.write_text(json.dumps(d, ensure_ascii=False, indent=1), encoding="utf-8")
        print(f"[ok] {src.name} -> {dst.name}  chars={d['stats']['chars']} tables={d['stats'].get('tables',0)} blocks={d['stats']['blocks']}")
    return 0

if __name__ == "__main__":
    raise SystemExit(main())
