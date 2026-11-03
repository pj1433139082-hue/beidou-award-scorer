"""【历史参考，不作为现行流程使用】

这是 2026-11-03 千问办公跑评分时的后处理脚本，收录在此只作参考。

留着的理由：它展示了一种保守的改造方式——不改已装技能，在外层重排输出。

什么时候还有用：
  · 需要在不修改技能源码的前提下改输出格式（别人装了技能但不让改）；
  · 想看 Word COM 转 PDF 的健壮写法（处理了进程泄漏：_list_winword_pids /
    _start_word / _stop_word，Quit 无效时按 PID 强杀）。

什么时候不该用它：
  · 当前技能的 export_docx.py 已原生支持白底黑字 + 四、免责说明，
    不需要这层重排。直接用 python scripts/export_docx.py 即可。

注意：它把 SKILL_SCRIPTS 硬编码成 qwen 的安装路径，换机器要改。

--- 原始说明 ---

千问版后处理：把 batch.py 已经产出的 docx 重排成用户要求的白底黑字四段式，
补齐参考材料 docx，然后用 Word COM 把所有 docx 转成 pdf。

为什么当时要单独一份：技能里的 export_docx.py 当时用绿色主题
（ACCENT=0B6B57 / FIXED=0B4F3F），用户要求白底黑字 + 显式的「四、免责说明」，
而改已装技能需单独授权，所以在外部做最后一层重排。

用法：python qwen_render_reference.py <输出目录>
"""
from __future__ import annotations
import argparse, json, os, re, sys, time
from pathlib import Path

# 借用技能里的建议引擎：三段式内容、追回比例、编号规则都是唯一来源
SKILL_SCRIPTS = Path(r"C:\Users\14331\.qwenworkcn\skills\beidou-award-scorer\scripts")
sys.path.insert(0, str(SKILL_SCRIPTS))
from advice_engine import build_items  # noqa: E402
from material_role import is_reference  # noqa: E402

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor

BLACK = RGBColor(0, 0, 0)
FONT = "微软雅黑"


def strip_award(name: str) -> str:
    s = re.sub(r"^(金点子|金刚钻|金引擎|金智囊|金火炬|特别奖-[^\s\-－—]*)[-－—\s]*", "",
               str(name or "").strip())
    return re.sub(r"^[-－—\s]+", "", s).strip() or str(name or "")


def set_run(run, size=10.5, bold=False, italic=False, name=FONT):
    run.font.name = name
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.italic = italic
    run.font.color.rgb = BLACK
    rpr = run._element.get_or_add_rPr()
    rf = rpr.find(qn("w:rFonts"))
    if rf is None:
        rf = OxmlElement("w:rFonts"); rpr.append(rf)
    rf.set(qn("w:eastAsia"), name)


def para(doc, text="", size=9.5, bold=False, italic=False,
         before=0, after=1, indent=0, keep=False, align=None):
    p = doc.add_paragraph()
    if align is not None:
        p.alignment = align
    pf = p.paragraph_format
    pf.space_before = Pt(before)
    pf.space_after = Pt(after)
    pf.line_spacing = 1.0
    if indent:
        pf.left_indent = Cm(indent)
    if keep:
        pf.keep_with_next = True
    if text:
        set_run(p.add_run(text), size=size, bold=bold, italic=italic)
    return p


def h(doc, text, size=9.5, before=2):
    return para(doc, text, size=size, bold=True, before=before, after=0, keep=True)


def seg_label(doc, text):
    p = doc.add_paragraph()
    pf = p.paragraph_format
    pf.space_before = Pt(0)
    pf.space_after = Pt(0)
    pf.left_indent = Cm(0.4)
    pf.line_spacing = 1.0
    pf.keep_with_next = True
    set_run(p.add_run(text), size=8, bold=True)
    return p


def body_text(doc, text, indent=0.4, size=8.5):
    p = doc.add_paragraph()
    pf = p.paragraph_format
    pf.left_indent = Cm(indent)
    pf.space_before = Pt(0)
    pf.space_after = Pt(0)
    pf.line_spacing = 1.0
    set_run(p.add_run(text), size=size)
    return p


def new_doc():
    """A4，四周 1.5cm 边距（比默认 2.2/2.0 紧，让每份材料能塞进一页）。"""
    doc = Document()
    sec = doc.sections[0]
    sec.page_width, sec.page_height = Cm(21), Cm(29.7)
    sec.left_margin = sec.right_margin = Cm(1.5)
    sec.top_margin = sec.bottom_margin = Cm(1.3)
    st = doc.styles["Normal"]
    st.font.name = FONT
    st.font.size = Pt(9.5)
    st.element.rPr.rFonts.set(qn("w:eastAsia"), FONT)
    return doc


def write_title(doc, name):
    """长标题在「——」处断成主标题+副标题（用户明确要求）。"""
    title = strip_award(name)
    if len(title) > 28 and "——" in title:
        main, sub = title.split("——", 1)
        main = main.strip().rstrip("-－—").strip()
        sub = sub.strip()
        p = para(doc, "", size=11, bold=True, after=0)
        set_run(p.add_run(main), size=11, bold=True)
        p2 = para(doc, "", size=9, bold=True, after=1)
        set_run(p2.add_run("——" + sub), size=9, bold=True)
    else:
        para(doc, title, size=11, bold=True, after=1)


def build_scored(data: dict, out_path: Path, doc=None, mat_index: int = 0):
    """一份有 score.json 的材料的意见书：四段（判断/建议/追问/免责）。"""
    if doc is None:
        doc = new_doc()

    write_title(doc, data["name"])
    # meta 合并到一行，减少竖向占用
    meta = "奖项：%s　得分：%s/100　区间：%s　置信度：%s" % (
        data["award"], data["ai_score"], data["band"], data["confidence"])
    para(doc, meta, size=8.5, after=0)
    if data.get("veto_pending"):
        para(doc, "否决项：待裁定（%s）" % (data.get("veto_note") or ""),
             size=8.5, after=1)

    # 一、总体判断
    h(doc, "一、总体判断")
    dims = data["dimensions"]
    by_ratio = sorted(dims, key=lambda d: d["score"] / d["max"] if d["max"] else 1)
    best = by_ratio[-1]
    worst = by_ratio[:2]
    para(doc, "最强项是%s（%s/%s）。最需要补的是%s。" % (
        best["name"], best["score"], best["max"],
        "、".join("%s（%s/%s）" % (d["name"], d["score"], d["max"]) for d in worst)),
        size=9)

    # 二、逐条修改建议（三段式）
    items = build_items(data)
    h(doc, "二、逐条修改建议")
    if not items:
        para(doc, "各维度均已达标，无需补强。", size=9)
    for it in items:
        p = doc.add_paragraph()
        pf = p.paragraph_format
        pf.space_before = Pt(2)
        pf.space_after = Pt(0)
        pf.line_spacing = 1.0
        pf.keep_with_next = True
        no = ("%d-%s" % (mat_index, it["no"])) if mat_index else it["no"]
        set_run(p.add_run(no + "　"), size=9, bold=True)
        set_run(p.add_run("%s　%s" % (it["dim"], it["title"])), size=9, bold=True)
        g = it.get("gain") or 0.0
        if g > 0:
            set_run(p.add_run("　（预计提升 %.1f 分）" % g), size=7.5)
        elif it.get("gain_note"):
            set_run(p.add_run("　（%s）" % it["gain_note"]), size=7.5)

        seg_label(doc, "原文是什么")
        body_text(doc, it["original"] or "（该项未定位到具体原文，见问题描述）")
        seg_label(doc, "怎么改")
        body_text(doc, it["how"])
        seg_label(doc, "修改之后的句子（可直接复制）")
        body_text(doc, it["revised"])

    # 三、答辩可能被追问
    ds = data.get("disputes") or []
    h(doc, "三、答辩可能被追问")
    if not ds:
        para(doc, "本份材料机检未发现需要在答辩中追问的争议点。", size=8.5, indent=0.4)
    for i, x in enumerate(ds, 1):
        pt = (x.get("point") or "").strip()
        src_ = (x.get("source") or "").strip()
        tail = ("　（来源：%s）" % src_) if src_ else ""
        para(doc, "Q%02d　%s%s" % (i, pt, tail), size=8.5, indent=0.4, after=0)

    # 四、免责说明
    h(doc, "四、免责说明")
    para(doc, "分数为机检+证据定级自动产出，引用需回原文核实；最终得分以评委综合审议为准。",
         size=8, indent=0.4, after=0)

    if out_path is not None:
        doc.save(str(out_path))
    return out_path


def build_reference(stem: str, out_path: Path, doc=None):
    """参考材料（过往版本）：一页说明，不参与评分。"""
    if doc is None:
        doc = new_doc()
    write_title(doc, stem)
    para(doc, "奖项：（不参与评分）　　角色：参考材料（过往版本）",
         size=9.5, after=2)

    h(doc, "一、总体判断")
    para(doc, "本份材料文件名含「过往版本」标记，被识别为参考材料，"
              "用途是供当届申报书做「成果迭代性」对照，不作为独立参赛项目参与排名。"
              "机检仍然运行（用于确认迭代证据），但不生成 score.json / report.md，"
              "也不进入排行榜与「改完提升」统计。", size=10.5)

    h(doc, "二、逐条修改建议")
    para(doc, "参考材料不出具修改建议。若要把「过往版本 → 当届版本」的迭代证据补进当届申报书，"
              "请在当届材料的「成果迭代性」章节写明：版本号、相较上届改了什么、"
              "指标提升多少（对应机检项 C13）。", size=10.5)

    h(doc, "三、答辩可能被追问")
    para(doc, "本份为参考材料，不进入答辩追问清单。相关追问由当届版本材料承担。",
         size=10, indent=0.5)

    h(doc, "四、免责说明")
    para(doc, "参考材料的识别仅依据文件名标记；若判定有误，请在文件名中移除"
              "「过往版本 / 上一版 / 旧版」等标记后重跑。", size=9.5, indent=0.5)

    doc.save(str(out_path))
    return out_path


def build_merged(rows, out_path: Path):
    """合并册：封面 + 目录 + 每份材料一页。"""
    doc = new_doc()

    # 封面
    for _ in range(8):
        para(doc, "", after=0)
    p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    set_run(p.add_run("北斗奖申报材料评分意见"), size=22, bold=True)
    p2 = doc.add_paragraph(); p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
    set_run(p2.add_run("共 %d 份" % len(rows)), size=11)

    # 目录
    doc.add_page_break()
    h(doc, "目录", before=0)
    order = {}
    for r in rows:
        order.setdefault(r["award"], []).append(r)
    for award, items in order.items():
        para(doc, award, size=10.5, bold=True, before=8, after=2)
        for r in sorted(items, key=lambda x: -x["ai_score"]):
            nm = strip_award(r["name"])
            if len(nm) > 32:
                nm = nm[:31] + "…"
            idx = rows.index(r) + 1
            para(doc, "    %2d. %5.1f 分　%s" % (idx, r["ai_score"], nm),
                 size=9.5, after=1)

    # 每份材料一页
    for mi, r in enumerate(rows, 1):
        doc.add_page_break()
        build_scored(r, None, doc=doc, mat_index=mi)

    doc.save(str(out_path))
    return out_path


def load_rows(out: Path):
    """从 out 目录里加载 score.json，附带 checks + doc.json 的原文供建议引擎用。"""
    rows = []
    for sf in sorted(out.glob("*.score.json")):
        d = json.loads(sf.read_text(encoding="utf-8"))
        stem = sf.name.replace(".score.json", "")
        cf = out / (stem + ".checks.json")
        d["_checks"] = {x["id"]: x for x in
                        (json.loads(cf.read_text(encoding="utf-8")).get("checks", [])
                         if cf.is_file() else [])}
        df = out / (stem + ".doc.json")
        d["_plain"] = (json.loads(df.read_text(encoding="utf-8")).get("plain", "")
                       if df.is_file() else "")
        d["_stem"] = stem
        rows.append(d)
    # 排序必须与 export_xlsx.py 完全一致（ai_score 降序），
    # 否则材料序号对不上，Word 与 Excel 的编号就会错位。
    rows.sort(key=lambda r: r.get("ai_score", 0), reverse=True)
    return rows


def find_reference_stems(out: Path):
    """有 doc.json 但没有 score.json 的材料 = 参考材料。"""
    scored = {sf.name.replace(".score.json", "") for sf in out.glob("*.score.json")}
    refs = []
    for df in sorted(out.glob("*.doc.json")):
        stem = df.name.replace(".doc.json", "")
        if stem in scored:
            continue
        if is_reference(stem):
            refs.append(stem)
    return refs


# ---------------------------------------------------------------------------
# PDF：本机没装 LibreOffice，用 Word COM（pywin32）走 Word 自己转，
# 版面最贴近 docx，中文字体也不会掉。
#
# 坑：Word COM 在批量转换时会不定期崩掉（RPC 服务器不可用 / 远程过程调用失败），
# 一次崩溃后同实例上的后续调用全部失败。所以每份文件失败要重启 Word 再试；
# 每 N 份主动重启一次做保守防御。
#
# 另一个坑：word.Quit() 有时不真退（内部错误 36915），进程留在后台锁住 docx，
# 让下一轮脚本无法覆写。用户 2026-11-03 授权：本脚本起的 Word 实例如果 Quit 后
# 仍在，可以 taskkill /F 强杀（只杀本脚本起的，不动用户自己开的 Word）。
# ---------------------------------------------------------------------------
def _list_winword_pids() -> set[int]:
    """当前所有 WINWORD.EXE 的 PID（用 tasklist，无额外依赖）。"""
    import subprocess
    try:
        r = subprocess.run(
            ["tasklist", "/FI", "IMAGENAME eq WINWORD.EXE", "/FO", "CSV", "/NH"],
            capture_output=True, text=True, timeout=10)
    except Exception:
        return set()
    pids = set()
    for line in (r.stdout or "").splitlines():
        parts = [x.strip('"') for x in line.split('","')]
        if len(parts) >= 2 and parts[0].upper() == "WINWORD.EXE":
            try: pids.add(int(parts[1]))
            except ValueError: pass
    return pids


def _start_word():
    """起一个新的 Word 实例，返回 (word, pids)。pids 用于 Quit 后兜底强杀。"""
    import win32com.client
    before = _list_winword_pids()
    word = win32com.client.DispatchEx("Word.Application")
    word.Visible = False
    word.DisplayAlerts = 0  # wdAlertsNone
    # 关掉自动保存 / 拼写检查，减少转换过程中的干扰
    try:
        word.Options.CheckGrammarAsYouType = False
        word.Options.CheckSpellingAsYouType = False
        word.Options.SaveInterval = 0
    except Exception:
        pass
    after = _list_winword_pids()
    return word, (after - before)


def _stop_word(word, pids=None):
    """Quit + 兜底 taskkill：只杀本次 _start_word 起的那些 PID。"""
    if word is None:
        return
    try: word.Quit()
    except Exception: pass
    pids = set(pids or set())
    if not pids:
        return
    # 等 3 秒让 Quit 生效
    for _ in range(6):
        time.sleep(0.5)
        alive = _list_winword_pids() & pids
        if not alive:
            return
    # 仍在，强杀（用户 2026-11-03 授权，只杀本脚本起的）
    import subprocess
    for pid in sorted(pids):
        try:
            subprocess.run(["taskkill", "/PID", str(pid), "/F"],
                           capture_output=True, text=True, timeout=10)
            print(f"  [qwen] Word PID {pid} Quit 未生效，已 taskkill /F")
        except Exception as e:
            print(f"  [qwen] taskkill PID {pid} 失败: {e}", file=sys.stderr)


def _convert_one(word, src: Path, dst: Path) -> None:
    """成功不返回；失败抛异常。"""
    doc = None
    try:
        # ConfirmConversions=False, ReadOnly=True, AddToRecentFiles=False
        doc = word.Documents.Open(str(src.resolve()), False, True, False)
        # 17 = wdFormatPDF
        doc.SaveAs2(str(dst.resolve()), FileFormat=17)
    finally:
        if doc is not None:
            try: doc.Close(False)
            except Exception: pass


def _convert_via_temp(word, src: Path, dst: Path) -> None:
    """有些文件名 Word 打不开（长中文+全角引号+连字符的组合会报「未能引发事件」）。
    回退方案：复制到短 ASCII 临时名，转成 pdf 再改回目标名。"""
    import shutil, uuid
    tmp_dir = src.parent
    tag = uuid.uuid4().hex[:8]
    tmp_docx = tmp_dir / ("_qwen_tmp_%s.docx" % tag)
    tmp_pdf = tmp_dir / ("_qwen_tmp_%s.pdf" % tag)
    try:
        shutil.copy2(src, tmp_docx)
        _convert_one(word, tmp_docx, tmp_pdf)
        if not tmp_pdf.is_file():
            raise RuntimeError("临时 pdf 未生成")
        shutil.move(str(tmp_pdf), str(dst))
    finally:
        for p in (tmp_docx, tmp_pdf):
            try:
                if p.exists(): p.unlink()
            except Exception:
                pass


def convert_pdf(docx_paths: list[Path], out_dir: Path,
                restart_every: int = 5, max_retry: int = 3,
                force: bool = False) -> tuple[list[Path], list[str]]:
    """把 docx 批量转成 pdf。返回 (成功的 pdf 路径, 失败的错误消息)。

    force=False 时，已存在且 >1KB 的 pdf 会被跳过（便于中断续跑）；
    force=True 时全部重转（改了 docx 版面后用）。
    """
    import pythoncom

    made, failed = [], []
    pythoncom.CoInitialize()
    word = None
    word_pids: set = set()
    since_restart = 0
    try:
        for i, src in enumerate(docx_paths, 1):
            dst = out_dir / (src.stem + ".pdf")
            if not force and dst.is_file() and dst.stat().st_size > 1024:
                made.append(dst)
                print("  [%2d/%d] 已存在，跳过 %s" % (i, len(docx_paths), src.name[:40]))
                continue
            ok = False
            last_err = None
            for attempt in range(max_retry):
                if word is None or since_restart >= restart_every:
                    _stop_word(word, word_pids)
                    word, word_pids = _start_word()
                    since_restart = 0
                # 第一次直接开；后面几次改用短 ASCII 临时名，
                # 绕开某些长中文文件名触发的 "Word 未能引发事件"。
                fn = _convert_one if attempt == 0 else _convert_via_temp
                try:
                    fn(word, src, dst)
                    if dst.is_file() and dst.stat().st_size > 1024:
                        ok = True
                        since_restart += 1
                        break
                    last_err = "PDF 未生成或过小"
                except Exception as e:
                    last_err = str(e)
                    # 崩溃后必须重启 Word，同实例上的后续调用都会失败
                    _stop_word(word, word_pids)
                    word, word_pids = None, set()
                    time.sleep(1.0 + attempt)
            if ok:
                made.append(dst)
                print("  [%2d/%d] OK  %s" % (i, len(docx_paths), src.name[:60]))
            else:
                failed.append("%s: %s" % (src.name, last_err))
                print("  [%2d/%d] FAIL %s: %s"
                      % (i, len(docx_paths), src.name[:40], last_err), file=sys.stderr)
    finally:
        _stop_word(word, word_pids)
        pythoncom.CoUninitialize()
    return made, failed


def main() -> int:
    ap = argparse.ArgumentParser(description="千问版后处理：docx 重排 + 参考材料补齐 + Word COM 转 PDF")
    ap.add_argument("out", help="batch.py 的输出目录（会就地覆盖 docx / 新增 pdf）")
    ap.add_argument("--skip-pdf", action="store_true", help="只重排 docx，不转 pdf")
    ap.add_argument("--force-pdf", action="store_true",
                    help="已存在的 pdf 也重转（改了 docx 版面后用）")
    a = ap.parse_args()

    out = Path(a.out)
    if not out.is_dir():
        print("输出目录不存在：%s" % out, file=sys.stderr); return 2

    rows = load_rows(out)
    if not rows:
        print("目录里没有 score.json，先跑 batch.py。", file=sys.stderr); return 2
    print("[qwen] 加载 %d 份评分材料" % len(rows))

    # 1. 逐份 docx（覆盖 batch.py 的绿字版）
    for mi, r in enumerate(rows, 1):
        build_scored(r, out / (r["_stem"] + ".docx"), mat_index=mi)
    print("[qwen] 重出 %d 份单份 docx（白底黑字四段）" % len(rows))

    # 2. 合并册
    build_merged(rows, out / "北斗奖评分意见汇总.docx")
    print("[qwen] 重出合并册 docx")

    # 3. 参考材料补一份 docx
    refs = find_reference_stems(out)
    for stem in refs:
        build_reference(stem, out / (stem + ".docx"))
    if refs:
        print("[qwen] 补齐 %d 份参考材料 docx：%s"
              % (len(refs), "、".join(s[:24] for s in refs)))

    # 4. PDF：把 out 目录下所有 docx 转成 pdf
    if a.skip_pdf:
        print("[qwen] 跳过 PDF（--skip-pdf）")
        return 0

    docxs = sorted(out.glob("*.docx"))
    print("[qwen] Word COM 转 PDF：%d 份 docx%s"
          % (len(docxs), "（--force-pdf 全部重转）" if a.force_pdf else ""))
    t0 = time.time()
    made, failed = convert_pdf(docxs, out, force=a.force_pdf)
    dt = time.time() - t0
    print("[qwen] PDF 完成 %d 份，用时 %.1fs" % (len(made), dt))
    if failed:
        print("[qwen] 失败 %d 份：" % len(failed), file=sys.stderr)
        for f in failed[:5]:
            print("  - " + f, file=sys.stderr)
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())