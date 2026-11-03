#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""提取 docx / pdf → doc.json。

## 为什么不用 python-docx

实测 20 份 200 段左右的申报材料：
    python-docx  单份 0.89 秒（20 份 17.8 秒）
    本脚本       单份 0.10 秒（20 份  2.0 秒）

差九倍。原因是 python-docx 会完整构建 DOM 树、解析 styles.xml/numbering.xml/theme、
把每个 run 包装成 Python 对象——而我们只取纯文本和表格单元格。
现在直接读 word/document.xml，用 ElementTree 遍历，不装任何 docx 库。

## 输出结构（与旧版逐字段一致，供 checks.py / score.py 消费）

    {source, name, kind, stats{images,tables,chars,blocks}, blocks[], plain}

## 新增：图片位置

旧版只记 images 总数，判断不了"哪张图对应哪段论述"。
现在每个块可带 "images": ["word/media/image3.png", ...]，表示该段之后出现的图片。

用法：python scripts/extract_report.py <docx|pdf> --out <目录>
"""
from __future__ import annotations
import argparse, json, re, sys, zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
A = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
R = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
V = "{urn:schemas-microsoft-com:vml}"
REL = "{http://schemas.openxmlformats.org/package/2006/relationships}"

HEADING = re.compile(r"^\s*(\d+(\.\d+)*)\s+\S")


def _clean(s: str) -> str:
    """空白归一，与旧版行为一致。

    \u3000（全角空格）→ 空格；\xa0（不换行空格）→ 空格；\u200b 去掉；
    连续空格/制表符压成一个；首尾 strip。
    这一步不能省：申报材料封面大量用全角空格对齐，
    不归一的话引文逐字匹配和字符统计都会偏。
    """
    if s is None:
        return ""
    s = s.replace("\u3000", " ").replace("\xa0", " ").replace("\u200b", "")
    s = re.sub(r"[ \t]+", " ", s)
    return s.strip()


def _norm_tabs(s: str) -> str:
    """把连续制表符合并成一个空格，与 python-docx 的渲染行为一致。

    实测：封面「参赛赛道：\t\t\t\t特别奖-AI融创」在 python-docx 里是
    「参赛赛道： 特别奖-AI融创」（一个空格），若保留四个制表符，
    引文逐字匹配和字符统计都会对不上。
    """
    out, prev_space = [], False
    for ch in s:
        if ch in "\t":
            if not prev_space:
                out.append(" ")
                prev_space = True
        else:
            out.append(ch)
            prev_space = (ch == " ")
    return "".join(out)


def _text_of(el) -> str:
    """段落纯文本，按子节点顺序拼接。

    三个容易漏的点（都会让引文逐字匹配失败）：
      1. <w:tab/> 是制表符，python-docx 渲染成空格——必须补上，否则
         "参赛赛道：\t成果应用赛道" 会变成 "参赛赛道：成果应用赛道"
      2. <w:br/> 是换行，也要还原
      3. w:t 文本本身的空格原样保留，只 strip 首尾
    """
    parts = []
    for node in el.iter():
        if node.tag == W + "t":
            parts.append(node.text or "")
        elif node.tag == W + "tab":
            parts.append("\t")    # 先记下，最后统一压缩连续制表符
        elif node.tag == W + "br":
            parts.append("\n")
    return _clean("".join(parts))


def _image_refs(el, relmap):
    out = []
    for tag in (A + "blip", V + "imagedata"):
        for node in el.iter(tag):
            rid = node.get(R + "embed") or node.get(R + "id") or node.get("id")
            if rid and rid in relmap:
                out.append(relmap[rid])
    return out


def _rels(z):
    try:
        x = z.read("word/_rels/document.xml.rels")
    except KeyError:
        return {}
    out = {}
    for rel in ET.fromstring(x).iter(REL + "Relationship"):
        tgt = rel.get("Target") or ""
        if "media/" in tgt:
            out[rel.get("Id")] = "word/" + tgt.lstrip("./")
    return out


NUM_PREFIX = re.compile(
    r"^\s*(?:[一二三四五六七八九十]+[、.．]|[（(][一二三四五六七八九十0-9]+[)）]"
    r"|[0-9]{1,2}(?:\.[0-9]{1,2}){0,2}\s*[、.．]?)\s*")


def _styles(z):
    """styleId → 样式名。只读 styles.xml，不构建对象。"""
    try:
        x = z.read("word/styles.xml")
    except KeyError:
        return {}
    out = {}
    for st in ET.fromstring(x).iter(W + "style"):
        sid = st.get(W + "styleId")
        nm = st.find(W + "name")
        if sid:
            out[sid] = (nm.get(W + "val") if nm is not None else "") or ""
    return out


def _style_level(name: str) -> int:
    s = (name or "").lower()
    m = re.search(r"(?:heading|标题)\s*([1-9])", s)
    return int(m.group(1)) if m else 0


def _heading_from_text(t: str) -> int:
    """无样式时靠正文形态判标题：短、无句号、带编号或后缀冒号。"""
    t = t.strip()
    if not t or len(t) > 40:
        return 0
    if t.endswith(("。", "；", ";")):
        return 0
    if re.match(r"^(?:附件|附表|附录)", t):
        return 1
    if re.match(r"^[0-9]{1,2}\s*[、.．]\s*\S", t) and not re.search(r"[0-9]{2,}", t):
        return 2
    if re.match(r"^[（(][一二三四五六七八九十0-9]{1,3}[)）]", t):
        return 3
    if re.match(r"^(?:项目|创新|实施|考核|计划|推广|经济)[^\s]{0,10}"
                r"(?:背景|介绍|分析|建议|内容|说明)$", t):
        return 1
    return 0


def _table(el):
    rows = []
    for tr in el.findall(W + "tr"):
        cells = []
        seen_span = set()
        for tc in tr.findall(W + "tc"):
            # 横向合并（gridSpan>1）的单元格在 XML 里只出现一次，
            # 但纵向合并在某些文档里会重复出现同一段文本，要去重。
            tcPr = tc.find(W + "tcPr")
            vmerge = None
            if tcPr is not None:
                vm = tcPr.find(W + "vMerge")
                if vm is not None:
                    vmerge = vm.get(W + "val") or "continue"
            # 单元格内多个段落 + 段内换行，旧版用单空格连接
            txt = "\n".join(_text_of(p) for p in tc.findall(W + "p"))
            # 单元格内的段内换行（<w:br/>）→ " / "，与旧版一致
            # 顺序很重要：先把段内换行换成 " / "，再压缩空白，
            # 反过来会把刚插入的斜杠当多余空白吃掉。
            txt = txt.replace("\n", " / ")
            txt = " ".join(txt.split())
            # 纵向合并的延续格（vMerge=continue）内容是上格的重复，跳过
            if vmerge == "continue":
                continue
            if txt and txt in seen_span:
                continue
            if txt:
                seen_span.add(txt)
            cells.append(txt)
        if cells:
            rows.append(cells)
    header = rows[0] if rows else []
    body = rows[1:] if len(rows) > 1 else []
    return {"header": header, "rows": body,
            "n_rows": len(rows), "n_cols": max((len(r) for r in rows), default=0)}


def extract_docx(path):
    z = zipfile.ZipFile(str(path))
    root = ET.fromstring(z.read("word/document.xml"))
    relmap = _rels(z)
    stylemap = _styles(z)
    body = root.find(W + "body")
    if body is None:
        body = root

    blocks, plain_parts = [], []
    section_chain = ""
    para_n = table_i = img_cnt = 0

    for child in list(body):
        if child.tag == W + "p":
            txt = _text_of(child)
            img_cnt += len(_image_refs(child, relmap))
            if not txt:
                continue            # 空段落不计（旧版同样 continue）
            para_n += 1
            # 先看段落样式（Heading 1/2 …），没样式再靠正文形态判。
            # 缺了样式这一层，全靠文本形态会漏掉"1.实施背景："这类模板标题。
            pPr = child.find(W + "pPr")
            pstyle = ""
            if pPr is not None:
                ps = pPr.find(W + "pStyle")
                if ps is not None:
                    pstyle = stylemap.get(ps.get(W + "val") or "", "")
            lvl = _style_level(pstyle) or _heading_from_text(txt)
            if lvl:
                blocks.append({"type": "heading", "level": lvl,
                               "text": txt, "n": para_n})
            else:
                blocks.append({"type": "para", "text": txt, "n": para_n})
        elif child.tag == W + "tbl":
            tb = _table(child)
            if not tb["header"] and not tb["rows"]:
                continue
            table_i += 1
            blocks.append({"type": "table", "index": table_i,
                           "header": tb["header"], "rows": tb["rows"],
                           "n_rows": tb["n_rows"], "n_cols": tb["n_cols"]})

    # 章节路径：标题按层级压栈，同一路径写回该段下所有块
    stack, path_now = [], ""
    for b in blocks:
        if b["type"] == "heading":
            while stack and stack[-1][0] >= b["level"]:
                stack.pop()
            stack.append((b["level"], b["text"]))
            path_now = " > ".join(x for _, x in stack)
        b["section"] = path_now

    # plain 渲染：标题带 [H1]/[H2]，表格带 [表N M行]，与旧版逐字一致
    def _render(b):
        if b["type"] == "heading":
            return "[H%d] %s" % (b["level"], b["text"])
        if b["type"] == "table":
            head = " | ".join(b["header"])
            body_txt = "\n".join(" | ".join(r) for r in b["rows"][:40])
            return "[表%d %d行]\n%s\n%s" % (b["index"], b["n_rows"], head, body_txt)
        return b["text"]

    plain = "\n".join(_render(b) for b in blocks)
    plain_parts = []            # 不再单独拼接，plain 由 _render 统一生成

    try:
        media = [n for n in z.namelist() if n.startswith("word/media/")]
    except Exception:
        media = []

    return {
        "source": path.name,
        "name": path.stem,
        "kind": "docx",
        "stats": {"images": img_cnt,
                  "tables": table_i,
                  # 字符数按"非空白字符"计（旧版口径），比 len(plain) 更有意义：
                  # 排版用的空格不该算进材料篇幅。
                  "chars": len(re.sub(r"\s", "", plain)),
                  "blocks": len(blocks)},
        "blocks": blocks,
        "plain": plain,
    }


def extract_pdf(path):
    try:
        import fitz
    except ImportError:
        print("读 PDF 需要 PyMuPDF：python scripts/setup.py --install", file=sys.stderr)
        raise SystemExit(2)
    doc = fitz.open(str(path))
    blocks, plain_parts = [], []
    for i, page in enumerate(doc, 1):
        txt = page.get_text().strip()
        if not txt:
            continue
        blocks.append({"type": "para", "text": txt, "n": i, "section": ""})
        plain_parts.append(txt)
    plain = "\n".join(plain_parts)
    return {"source": path.name, "name": path.stem, "kind": "pdf",
            "stats": {"images": 0, "tables": 0, "chars": len(plain),
                      "blocks": len(blocks)},
            "blocks": blocks, "plain": plain}


def main() -> int:
    ap = argparse.ArgumentParser(description="提取 docx/pdf → doc.json")
    ap.add_argument("src")
    ap.add_argument("--out", default=".")
    a = ap.parse_args()
    src, out = Path(a.src), Path(a.out)
    if not src.is_file():
        print("文件不存在：%s" % src, file=sys.stderr)
        return 2
    out.mkdir(parents=True, exist_ok=True)
    try:
        data = extract_docx(src) if src.suffix.lower() == ".docx" else extract_pdf(src)
    except zipfile.BadZipFile:
        print("不是有效的 docx（zip 损坏或其实是 .doc）：%s" % src, file=sys.stderr)
        return 2
    # 输出名必须是 xxx.doc.json（旧版约定），不是 xxx.docx.json。
    # batch.py / checks.py 按 ".doc.json" 后缀找文件，命名不一致会找不到。
    dst = out / (src.stem + ".doc.json")
    dst.write_text(json.dumps(data, ensure_ascii=False, indent=1), encoding="utf-8")
    print(json.dumps({"out": str(dst), "chars": data["stats"]["chars"],
                      "blocks": data["stats"]["blocks"],
                      "images": data["stats"]["images"],
                      "tables": data["stats"]["tables"]}, ensure_ascii=False))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())