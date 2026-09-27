"""
P1 回归：重复删除、删除后重新插入、多 Section / 页眉页脚的按类型删除。

设计原则与同目录的其他文件一致——只验证**最终文档状态**（水印是否还在、
用户内容是否还在、文档还能不能被 python-docx 打开），不做"只断言没抛异常"的测试。

覆盖的老问题：
- 已经没有水印时再删一次：不应报错、不应损坏文档；
- insert -> clear -> insert 循环：不该因残留空 run / 旧 XML 结构而异常；
- 多 Section 文档：每个 Section 的页眉都要被扫到，按类型删除时另一类必须活着。
"""
from __future__ import annotations

import io
import os
import sys
import tempfile

from docx import Document
from PIL import Image

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from watermark_tool import core, engine_docx

USER_BODY = "用户正文"


# ---------------------------------------------------------------------------
# 素材
# ---------------------------------------------------------------------------
def _png_bytes(color=(21, 22, 23), size=(28, 28)):
    b = io.BytesIO()
    Image.new("RGB", size, color).save(b, "PNG")
    return b.getvalue()


def _opts(tmp, tag):
    img = os.path.join(tmp, f"{tag}.png")
    with open(img, "wb") as f:
        f.write(_png_bytes())
    return {"text": {"text": "机密", "font_size": 100},
            "image": {"image_path": img}}


def _types(path):
    return engine_docx.detect_watermark_types(path)


def _body_texts(path):
    return [p.text for p in Document(path).paragraphs]


def _marked_count(path):
    """全文（含页眉页脚）里本工具标记的图形份数。"""
    doc = Document(path)
    return sum(1 for _ in engine_docx._iter_marked_drawings(doc))


def _watermarked(tmp, name, kinds, two_sections=False):
    """造一份带水印的文档；two_sections=True 时做成两个 Section 且页眉内容不同。"""
    src = os.path.join(tmp, f"{name}_src.docx")
    doc = Document()
    doc.add_paragraph(USER_BODY)
    if two_sections:
        doc.add_section()                     # 第二节：独立页眉
        doc.add_paragraph("第二节正文")
    doc.save(src)

    out = os.path.join(tmp, f"{name}_wm.docx")
    core.insert_watermark(src, list(kinds), output_path=out, **_opts(tmp, name))
    return out


def _section_header_text(path):
    return [[p.text for p in s.header.paragraphs] for s in Document(path).sections]


def _out_for(tmp, name, src):
    return os.path.join(tmp, f"{name}_out.docx")


# ===========================================================================
# 重复删除 / 无目标删除
# ===========================================================================
def test_clear_again_when_no_text_watermark_left():
    """已经没有文字水印时再删一次文字：不报错、文档不坏、用户内容不变。"""
    tmp = tempfile.mkdtemp()
    src = _watermarked(tmp, "rep_text", ["image"])     # 只有图片水印
    assert _types(src) == {"image"}

    out1 = _out_for(tmp, "rep_text_1", src)
    core.clear_watermark(src, output_path=out1, kinds=["text"])

    assert _types(out1) == {"image"}, "本来就没文字水印，删文字不应有副作用"
    assert _body_texts(out1) == _body_texts(src), "用户正文必须原样"

    # 再来一次（重复删除）
    out2 = _out_for(tmp, "rep_text_2", out1)
    core.clear_watermark(out1, output_path=out2, kinds=["text"])

    assert _types(out2) == {"image"}, "重复删除后图片水印仍应完好"
    assert _body_texts(out2) == _body_texts(src), "重复删除后正文仍应原样"
    Document(out2).paragraphs                              # 仍能正常打开
    print("[P1] 无文字水印时删文字 / 重复删除：不报错，正文与图片水印零变化")


def test_clear_again_when_no_image_watermark_left():
    """已经没有图片水印时再删一次图片：同上。"""
    tmp = tempfile.mkdtemp()
    src = _watermarked(tmp, "rep_img", ["text"])
    assert _types(src) == {"text"}

    out1 = _out_for(tmp, "rep_img_1", src)
    core.clear_watermark(src, output_path=out1, kinds=["image"])
    assert _types(out1) == {"text"}

    out2 = _out_for(tmp, "rep_img_2", out1)
    core.clear_watermark(out1, output_path=out2, kinds=["image"])

    assert _types(out2) == {"text"}, "重复删除图片水印不应影响文字水印"
    assert _body_texts(out2) == _body_texts(src)
    Document(out2).paragraphs
    print("[P1] 无图片水印时删图片 / 重复删除：不报错，正文与文字水印零变化")


def test_clear_twice_when_already_empty():
    """已经全部删光后再删一次（both）：不报错，文档仍可正常打开。"""
    tmp = tempfile.mkdtemp()
    src = _watermarked(tmp, "rep_both", ["text", "image"])
    assert _types(src) == {"text", "image"}
    body_before = _body_texts(src)

    empty = _out_for(tmp, "rep_both_empty", src)
    core.clear_watermark(src, output_path=empty, kinds=["text", "image"])
    assert _types(empty) == set()

    again = _out_for(tmp, "rep_both_again", empty)
    res = core.clear_watermark(empty, output_path=again, kinds=["text", "image"])

    assert res.get("ok"), "重复的全量删除应正常返回"
    assert _types(again) == set(), "重复删除后仍应无水印"
    assert _body_texts(again) == body_before, "重复删除后正文必须原样"
    assert Document(again).paragraphs is not None, "文档必须仍可正常打开"
    print("[P1] 已全删后再删一次：不报错，文档完好，正文不变")


# ===========================================================================
# 删除后重新插入
# ===========================================================================
def test_insert_clear_insert_round_trip():
    """insert -> clear -> insert：不得因旧标记/空 run 而出错，且水印能重新插上。"""
    tmp = tempfile.mkdtemp()
    src = os.path.join(tmp, "rt_src.docx")
    doc = Document()
    doc.add_paragraph(USER_BODY)
    doc.save(src)
    opts = _opts(tmp, "rt")

    # 第一次插入
    a = os.path.join(tmp, "rt_a.docx")
    first = core.insert_watermark(src, ["text", "image"], output_path=a, **opts)
    assert first.get("inserted", 0) > 0 and _types(a) == {"text", "image"}

    # 删除
    b = os.path.join(tmp, "rt_b.docx")
    core.clear_watermark(a, output_path=b, kinds=["text", "image"])
    assert _types(b) == set()

    # 第二次插入（同一份文档、同样的参数）
    c = os.path.join(tmp, "rt_c.docx")
    second = core.insert_watermark(b, ["text", "image"], output_path=c, **opts)

    assert second.get("ok"), "删除后重新插入必须成功"
    assert _types(c) == {"text", "image"}, (
        f"重新插入后应重新检出两类水印，实际 {_types(c)}")
    assert _body_texts(c) == _body_texts(src), "往返后正文必须原样"
    assert _marked_count(c) == _marked_count(a), (
        f"重新插入的水印份数应与首次一致：{_marked_count(c)} vs {_marked_count(a)}")
    print(f"[P1] insert->clear->insert 往返正常：水印 {_marked_count(c)} 份，正文不变")


def test_clear_between_text_and_image_insertions_is_independent():
    """交替插入/删除时，两类水印互不干扰（只删一类后另一类必须完整）。"""
    tmp = tempfile.mkdtemp()
    src = os.path.join(tmp, "alt.docx")
    doc = Document()
    doc.add_paragraph(USER_BODY)
    doc.save(src)
    opts = _opts(tmp, "alt")

    a = os.path.join(tmp, "alt_a.docx")
    core.insert_watermark(src, ["text", "image"], output_path=a, **opts)
    assert _types(a) == {"text", "image"}

    b = os.path.join(tmp, "alt_b.docx")
    core.clear_watermark(a, output_path=b, kinds=["text"])
    assert _types(b) == {"image"}, "只删文字后应只剩图片水印"

    c = os.path.join(tmp, "alt_c.docx")
    core.clear_watermark(b, output_path=c, kinds=["image"])
    assert _types(c) == set(), "再删图片后应全清"
    assert _body_texts(c) == _body_texts(src)
    print("[P1] 交替删除两类水印：互不干扰，最终全清且正文不变")


# ===========================================================================
# 多 Section / 页眉页脚
# ===========================================================================
def test_multisection_clear_text_keeps_image_and_user_text():
    """两个 Section 各有水印：只删文字时，两个 Section 的文字水印都要没了。"""
    tmp = tempfile.mkdtemp()
    src = _watermarked(tmp, "ms", ["text", "image"], two_sections=True)
    assert Document(src).sections.__len__() == 2, "前置条件：应有两个 Section"
    assert _types(src) == {"text", "image"}

    hdr_before = _section_header_text(src)
    used_headers = {tuple(t) for sec in Document(src).sections
                    for t in ([p.text for p in sec.header.paragraphs],)}

    out = _out_for(tmp, "ms_out", src)
    core.clear_watermark(src, output_path=out, kinds=["text"])

    assert _types(out) == {"image"}, (
        f"只删文字后应只剩图片水印，实际 {_types(out)}"
        "——说明某个 Section 的文字水印没被清干净")
    assert USER_BODY in _body_texts(out), "正文必须保留"
    Document(out).paragraphs
    # 用户原始页眉文字（这里是空页眉）不应被凭空改写
    assert _section_header_text(out) == hdr_before, "页眉结构不应被改动"
    assert len(used_headers) >= 1
    print("[P1] 多 Section：只删文字 -> 两节文字水印都清掉，图片水印与正文保留")


def test_multisection_clear_image_keeps_text():
    """两个 Section：只删图片时，两个 Section 的图片水印都要没了。"""
    tmp = tempfile.mkdtemp()
    src = _watermarked(tmp, "ms2", ["text", "image"], two_sections=True)

    out = _out_for(tmp, "ms2_out", src)
    core.clear_watermark(src, output_path=out, kinds=["image"])

    assert _types(out) == {"text"}, (
        f"只删图片后应只剩文字水印，实际 {_types(out)}")
    assert USER_BODY in _body_texts(out)
    Document(out).paragraphs
    print("[P1] 多 Section：只删图片 -> 两节图片水印都清掉，文字水印与正文保留")


def test_multisection_clear_both_removes_everything():
    """两个 Section 全删：一个 Section 都不许残留。"""
    tmp = tempfile.mkdtemp()
    src = _watermarked(tmp, "ms3", ["text", "image"], two_sections=True)
    before = _marked_count(src)

    out = _out_for(tmp, "ms3_out", src)
    core.clear_watermark(src, output_path=out, kinds=["text", "image"])

    assert _marked_count(out) == 0, (
        f"全删后不应有任何残留，实际 {_marked_count(out)} 份"
        f"（首次插入时共 {before} 份）")
    assert _types(out) == set()
    assert _body_texts(out) == _body_texts(src)
    print(f"[P1] 多 Section 全删：{before} 份水印全部清除，正文不变")
