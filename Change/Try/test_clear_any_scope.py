"""
二次安全审查（commit 84a773e 之后）—— 主题 2：clear_any 的识别边界。

**不改实现**：本文件先把 `clear_any()/_remove_in_part(clear_any=True)` 到底按什么条件
识别“水印”记录下来，并把「可能误删用户原有内容」的场景写成**复现测试**。
这些复现测试当前是失败态（xfail），它们刻画的是**已知风险**，不是回归缺陷；
一旦将来收紧了 clear_any 的判定，它们会变成 XPASS，提醒回来收口。

识别条件（engine_docx._remove_in_part + engine_docx.clear_watermark）：
  1. 扫描范围 = 每个 section 的**页眉 + 页脚**（含首页/偶数页）内的 `w:drawing`。
     正文（document body）与表格**不在**扫描范围内 —— 这一条是安全的。
  2. 命中即删的条件有两个，任一成立：
     - 图形上带本工具私有标记（wp:docPr/@descr 或 @name ∈ {WB_WATERMARK_TEXT,
       WB_WATERMARK_IMG}）；
     - **或**（仅当 clear_any=True）该图形是 `wp:anchor[behindDoc="1"]`
       的浮动图形 —— **不看归属**，只凭“衬于文字下方”这个结构特征。
  3. 删除手法 `_detach_drawing`：整条 `w:r`（run）一起摘掉，不只是图形。
  4. 只认 `wp:anchor`；`wp:inline`（行内图片）与未设 behindDoc 的浮动图形不参与。

结论一句话：**clear_any 把「衬于文字下方的浮动图形」当成了“疑似水印”**。
Word 原生水印确实长这样（所以按设计要清掉），但用户在页眉里放的、同样
“衬于文字下方”的背景图/Logo 也会被一起删掉 —— 这是可能的误删。
"""
from __future__ import annotations

import io
import os
import sys
import tempfile

import pytest
from docx import Document
from PIL import Image

from lxml import etree

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from watermark_tool import engine_docx


def _png(color=(5, 5, 5), size=(32, 32)):
    b = io.BytesIO()
    Image.new("RGB", size, color).save(b, "PNG")
    return b.getvalue()


def _make_doc(path):
    doc = Document()
    doc.add_paragraph("正文第一段")
    t = doc.add_table(rows=1, cols=1)
    t.cell(0, 0).text = "表格内容"
    doc.save(path)
    return path


def _user_graphic(hdr, name="用户页眉背景图", behind=True, docpr_id=900):
    """在页眉里放一个「用户自己的」浮动图形（无本工具私有标记，可选 behindDoc）。"""
    rId, _ = hdr.part.get_or_add_image(io.BytesIO(_png()))
    drawing = engine_docx._make_drawing(rId, 400000, 400000, 0, name, "", docpr_id, 0, 0)
    if not behind:                      # 造一个非 behindDoc 的浮动图形
        anchor = drawing.find(".//" + engine_docx.qn("wp:anchor"))
        anchor.set("behindDoc", "0")
    engine_docx._add_drawing_to_part(hdr, drawing)
    return drawing


def _native_watermark(hdr):
    """模拟 Word 原生水印：behindDoc=1、显示名像系统水印、无本工具标记。"""
    rId, _ = hdr.part.get_or_add_image(io.BytesIO(_png()))
    return engine_docx._make_drawing(rId, 400000, 400000, 0,
                                     "PowerPlusWaterMarkObject", "", 901, 0, 0)


def _behind_marked_drawings(part):
    """返回页眉里「behindDoc=1 且没有本工具标记」的图形数量。"""
    n = 0
    for dr in part._element.iter(engine_docx.qn("w:drawing")):
        anchor = dr.find(".//" + engine_docx.qn("wp:anchor"))
        if anchor is not None and anchor.get("behindDoc") == "1" \
                and engine_docx._mark_of(dr) is None:
            n += 1
    return n


@pytest.mark.xfail(reason="已知风险：clear_any 只凭 behindDoc=1 判定，会删掉用户在页眉里的"
                          "衬于文字下方图形（复现测试，未改实现）", strict=False)
def test_clear_any_deletes_user_own_behinddoc_header_picture():
    """用户在页眉放的背景图（behindDoc=1、无本工具标记）应被保留。"""
    tmp = tempfile.mkdtemp()
    p = os.path.join(tmp, "d.docx")
    _make_doc(p)
    doc = Document(p)
    hdr = doc.sections[0].header
    _user_graphic(hdr)
    doc.save(p)

    assert _behind_marked_drawings(hdr.part) == 1
    engine_docx.clear_watermark(p)      # kinds=None → 走 clear_any=True 分支
    doc2 = Document(p)
    assert _behind_marked_drawings(doc2.sections[0].header.part) == 1, (
        "用户在页眉里自己的 behindDoc 图形被 clear_any 删掉了（潜在误删）"
    )


def test_clear_any_leaves_inline_and_front_floating_graphics():
    """安全对照：行内图片、非 behindDoc 浮动图形都不在删除范围内。"""
    tmp = tempfile.mkdtemp()
    p = os.path.join(tmp, "d.docx")
    _make_doc(p)
    doc = Document(p)
    hdr = doc.sections[0].header
    pr = hdr.paragraphs[0]._p
    # 行内图片（Word 里最常见的页眉 Logo）
    hdr.paragraphs[0].add_run().add_picture(io.BytesIO(_png()))
    # 非 behindDoc 的浮动图形
    _user_graphic(hdr, name="前景浮动图形", behind=False)
    doc.save(p)

    engine_docx.clear_watermark(p)

    doc2 = Document(p)
    hdr2 = doc2.sections[0].header
    hdr_el = doc2.sections[0].header._element
    inline = sum(1 for _ in hdr_el.iter(engine_docx.qn("wp:inline")))
    floating = 0
    for dr in hdr2._element.iter(engine_docx.qn("w:drawing")):
        anchor = dr.find(".//" + engine_docx.qn("wp:anchor"))
        if anchor is not None and anchor.get("behindDoc") != "1" \
                and engine_docx._mark_of(dr) is None:
            floating += 1
    assert inline == 1, "页眉行内图片（Logo）不应被删除"
    assert floating == 1, "非 behindDoc 的浮动图形不应被删除"
    print("[review2] 安全对照：行内图片 / 非 behindDoc 浮动图形 均未被删除")


def test_clear_any_keeps_body_and_table_intact():
    """扫描范围不含正文与表格：clear_any 不应影响正文/表格/正文图片。"""
    tmp = tempfile.mkdtemp()
    p = os.path.join(tmp, "d.docx")
    _make_doc(p)
    doc = Document(p)
    doc.add_paragraph("正文第二段")
    doc.paragraphs[0].runs[0].text = "正文第一段（用户改过）"
    doc.save(p)
    before = ([x.text for x in Document(p).paragraphs],
              [[c.text for c in r.cells] for t in Document(p).tables for r in t.rows])

    engine_docx.clear_watermark(p)

    after = ([x.text for x in Document(p).paragraphs],
             [[c.text for c in r.cells] for t in Document(p).tables for r in t.rows])
    assert after == before, "clear_any 不应改动正文或表格"
    print("[review2] clear_any 扫描范围不含正文/表格，正文与表格未被改动")


def test_clear_any_still_removes_word_native_watermark_by_design():
    """按设计：Word 原生水印（behindDoc=1 且名字像系统水印）应当被清掉。"""
    tmp = tempfile.mkdtemp()
    p = os.path.join(tmp, "d.docx")
    _make_doc(p)
    doc = Document(p)
    hdr = doc.sections[0].header
    engine_docx._add_drawing_to_part(hdr, _native_watermark(hdr.part))
    doc.save(p)

    res = engine_docx.clear_watermark(p)
    doc2 = Document(p)
    left = _behind_marked_drawings(doc2.sections[0].header.part)
    assert left == 0 and res["removed"] >= 1, "clear_any 的目的就是清掉原生水印"
    print("[review2] 按设计：Word 原生水印被清除")


def test_clear_with_kinds_limited_is_safe():
    """限定 kinds 的路径（GUI「只去文字水印」）只删本工具标记，不碰用户图形。"""
    tmp = tempfile.mkdtemp()
    p = os.path.join(tmp, "d.docx")
    _make_doc(p)
    doc = Document(p)
    hdr = doc.sections[0].header
    _user_graphic(hdr, name="用户背景图", docpr_id=902)
    doc.save(p)

    engine_docx.clear_watermark(p, kinds=["text"])

    doc2 = Document(p)
    assert _behind_marked_drawings(doc2.sections[0].header.part) == 1, (
        "kinds 限定路径不应碰用户图形"
    )
    print("[review2] kinds 限定路径安全：只删本工具标记，用户图形保留")


def test_clear_any_deletes_whole_run_including_sibling_text():
    """放大效应：页眉 run 里“文字 + behindDoc 图形”并存时，文字**不许**被连带删掉。

    该用例原先是 xfail（复现 `_detach_drawing` 整条 run 一起摘的老问题）。
    修复后 `_detach_drawing` 只摘图形、run 非空就不动它，这里升级为正式回归断言。
    """
    tmp = tempfile.mkdtemp()
    p = os.path.join(tmp, "d.docx")
    _make_doc(p)
    doc = Document(p)
    hdr = doc.sections[0].header
    run = hdr.paragraphs[0].add_run()
    run.text = "页眉必须保留的说明文字"
    rId, _ = hdr.part.get_or_add_image(io.BytesIO(_png()))
    run._r.append(engine_docx._make_drawing(rId, 300000, 300000, 0,
                                            "页眉旧插图", "", 903, 0, 0))
    doc.save(p)

    engine_docx.clear_watermark(p)

    doc2 = Document(p)
    texts = [x.text for x in doc2.sections[0].header.paragraphs]
    assert "页眉必须保留的说明文字" in texts, (
        "同 run 内的文字被 clear_any 连带删除了"
    )
    print("[review2] 同 run 文字未被删除")


WPS_URI = "http://schemas.microsoft.com/office/word/2010/wordprocessingShape"


def _user_shape(hdr, behind=True, name="用户插入的形状"):
    """在页眉里放一个「用户自己画的形状」。

    Word 保存 2010+ 形状时用的正是 mc:AlternateContent + w:drawing + wps:wsp
    这套结构（`mc:Choice` 里是 DrawingML，`mc:Fallback` 里是 VML）。也就是说，
    一个形状在 OOXML 里同样会以 `w:drawing` 出现，只是 `a:graphicData/@uri`
    不是图片命名空间，而是 wordprocessingShape。
    """
    mc = "{http://schemas.openxmlformats.org/markup-compatibility/2006}"
    el = hdr._element
    ac = etree.SubElement(el, mc + "AlternateContent")
    choice = etree.SubElement(ac, mc + "Choice")
    choice.set(mc + "Requires", "wps")
    drawing = etree.SubElement(choice, engine_docx.qn("w:drawing"))
    anchor = etree.SubElement(drawing, engine_docx.qn("wp:anchor"))
    anchor.set("behindDoc", "1" if behind else "0")
    docPr = etree.SubElement(anchor, engine_docx.qn("wp:docPr"))
    docPr.set("id", "777")
    docPr.set("name", name)
    graphic = etree.SubElement(anchor, engine_docx.qn("a:graphic"))
    gd = etree.SubElement(graphic, engine_docx.qn("a:graphicData"))
    gd.set("uri", WPS_URI)
    etree.SubElement(gd, "{" + WPS_URI + "}wsp")
    return drawing


def _count_shapes(part):
    """页眉里非图片类的 w:drawing 数量（即用户形状）。"""
    n = 0
    for dr in part._element.iter(engine_docx.qn("w:drawing")):
        gd = dr.find(".//" + engine_docx.qn("a:graphicData"))
        if gd is None or gd.get("uri") != engine_docx.PIC:
            n += 1
    return n


def test_clear_any_keeps_user_drawing_shape():
    """用户自己在页眉画的形状（wps 命名空间，非图片）不应被 clear_any 删掉。"""
    tmp = tempfile.mkdtemp()
    p = os.path.join(tmp, "d.docx")
    _make_doc(p)
    doc = Document(p)
    hdr = doc.sections[0].header
    _user_shape(hdr, behind=True)          # 用户把它「置于底层」
    # 同时放一个 Word 原生水印做对照：它必须仍被清掉
    engine_docx._add_drawing_to_part(hdr, _native_watermark(hdr.part))
    doc.save(p)
    assert _count_shapes(hdr.part) == 1

    engine_docx.clear_watermark(p)         # kinds=None → clear_any=True

    hdr2 = Document(p).sections[0].header.part
    assert _count_shapes(hdr2) == 1, (
        "用户在页眉里自己画的形状被 clear_any 当成水印删掉了——"
        "它既不是本工具水印，也不是 Word 原生水印（graphicData uri 是 wps 而非 picture）。"
        f"(left={_count_shapes(hdr2)})"
    )
    print("[review3] 用户页眉形状被保留，Word 原生水印按设计被清除")


def test_clear_any_keeps_front_floating_shape():
    """形状即便没有 behindDoc 也不该被碰（安全对照，本就成立）。"""
    tmp = tempfile.mkdtemp()
    p = os.path.join(tmp, "d.docx")
    _make_doc(p)
    doc = Document(p)
    hdr = doc.sections[0].header
    _user_shape(hdr, behind=False)
    doc.save(p)

    engine_docx.clear_watermark(p)

    assert _count_shapes(Document(p).sections[0].header.part) == 1
    print("[review3] 前景浮动形状未被删除")


def test_clear_any_scope_summary():
    """把识别边界以可执行的形式固化下来（本次审查的结论，防止日后無意间扩大范围）。"""
    tmp = tempfile.mkdtemp()
    p = os.path.join(tmp, "d.docx")
    _make_doc(p)
    doc = Document(p)
    hdr = doc.sections[0].header
    # 三种图形：本工具水印 / Word 原生水印 / 用户自己的 behindDoc 图
    _wm_rid, _ = hdr.part.get_or_add_image(io.BytesIO(_png()))
    engine_docx._add_drawing_to_part(hdr, engine_docx._make_drawing(
        _wm_rid, 300000, 300000, 0,
        engine_docx.MARK_TEXT, engine_docx.MARK_TEXT, 910, 0, 0))
    engine_docx._add_drawing_to_part(hdr, _native_watermark(hdr.part))
    _user_graphic(hdr, name="用户背景图", docpr_id=911)
    doc.save(p)

    engine_docx.clear_watermark(p)

    doc2 = Document(p)
    hdr_el = doc2.sections[0].header._element
    ours_left = sum(1 for dr in hdr_el.iter(engine_docx.qn("w:drawing"))
                    if engine_docx._mark_of(dr) is not None)
    user_left = sum(1 for dr in hdr_el.iter(engine_docx.qn("w:drawing"))
                    if engine_docx._mark_of(dr) is None)
    assert ours_left == 0, "本工具水印应被清掉"
    # Word 原生水印按设计被清掉；用户自己的 behindDoc 图形是否保留 = 已知风险点
    assert user_left in (0, 1)
    print(f"[review2] 边界固化：本工具水印残留={ours_left}（应为 0），"
          f"残留无标记图形={user_left}（1 表示尚未收紧 → 已知风险）")
