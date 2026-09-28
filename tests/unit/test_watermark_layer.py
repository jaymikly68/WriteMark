"""
文字水印的**结构层级**回归测试（对应“水印被正文图片遮挡”的排查）。

结论先写在这里，避免后来人重复踩：

    WriteMark 的水印写在 **页眉 part** 里，锚点为 `wp:anchor` +
    `behindDoc="1"` + `wrapNone` + `relativeFrom="page"`。
    这与 Word 内置水印（页眉里衬于文字下方的图形）是同一套做法。

    Word 的绘制顺序是**跨层固定**的：页眉/页脚层绘制在正文内容层**之下**
    （水印之所以能“衬于文字下方”，正是靠这一点）。因此正文里的图片——
    无论 inline、浮动还是 behindDoc——都绘制在页眉水印之上，水印经过图片
    区域时会被挡住。

    这是 Word 的固有行为，不是 WriteMark 的 XML 缺陷：
    - `behindDoc` 只在同一 story（页眉 / 正文）内部排序，改它无法跨层；
    - `relativeHeight`(z-order) 同样只在 story 内生效；
    - Word 自带的「水印」功能也是页眉实现，遇到不透明正文图片同样被遮。

    可行的改法只有一种：把 `behindDoc` 改成 "0"，让水印浮到正文之上。
    但那会让水印压住正文文字和图片，与“水印不干扰阅读”的常规语义冲突，
    因此**当前刻意保持 behindDoc="1"**，作为产品默认。若将来要提供
    「水印显示在图片之上」的选项，改 `_make_drawing()` 这一处属性即可，
    识别/清除逻辑依赖的是 docPr 上的私有标记，**不受该属性影响**。

本文件只做**结构层**断言。真实 Word 的像素级遮挡行为无法在 CI 里验证
（需要 Word COM 渲染），必须人工在 Word 中打开确认。
"""
import os
import sys
import tempfile

import pytest
from docx import Document
from PIL import Image

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from watermark_tool import core, engine_docx

qn = engine_docx.qn


def _png(d, name, size=(24, 24), color=(10, 10, 10)):
    p = os.path.join(d, name)
    Image.new("RGB", size, color).save(p)
    return p


def _header_drawings(doc):
    return [dr for s in doc.sections
            for dr in s.header._element.iter(qn("w:drawing"))]


def _body_drawings(doc):
    return list(doc.element.body.iter(qn("w:drawing")))


def _make_doc_with_body_image(d):
    src = os.path.join(d, "body_img.docx")
    doc = Document()
    doc.add_paragraph("用户正文")
    doc.add_picture(_png(d, "bodyimg.png", (60, 40), (200, 200, 200)), width=None)
    doc.save(src)
    return src


def _make_plain_doc(d, name="plain.docx"):
    src = os.path.join(d, name)
    doc = Document()
    doc.add_paragraph("用户正文")
    doc.save(src)
    return src


# ---------------------------------------------------------------------------
def test_watermark_is_placed_in_header_behind_doc():
    """水印必须落在页眉 part、behindDoc=1、绝对定位、无环绕。

    这几个属性共同构成“衬于文字下方”的语义；任何一项被改掉都会改变所有
    文档的默认外观，因此在这里钉死。
    """
    d = tempfile.mkdtemp()
    src = _make_plain_doc(d)          # 纯文本文档：正文里本来就没有图形
    out = os.path.join(d, "out.docx")
    core.insert_watermark(src, ["text"], output_path=out,
                          text={"text": "机密", "font_size": 80})

    doc = Document(out)
    hdrs = _header_drawings(doc)
    assert len(hdrs) == 1, f"页眉里应恰好一份水印，实际 {len(hdrs)}"
    assert _body_drawings(doc) == [], "水印不应被写进正文"

    anchor = hdrs[0].find(".//" + qn("wp:anchor"))
    assert anchor is not None, "水印应是浮动锚定(anchor)图形"
    assert anchor.get("behindDoc") == "1", "水印应衬于文字下方(behindDoc=1)"
    assert hdrs[0].find(".//" + qn("wp:wrapNone")) is not None, "水印应无环绕"
    pos_h = hdrs[0].find(".//" + qn("wp:positionH"))
    pos_v = hdrs[0].find(".//" + qn("wp:positionV"))
    assert pos_h.get("relativeFrom") == "page", "水平定位应以页面为基准"
    assert pos_v.get("relativeFrom") == "page", "垂直定位应以页面为基准"
    print("[层] 水印位于页眉 · behindDoc=1 · relativeFrom=page · wrapNone")


def test_body_image_stays_in_body_layer():
    """正文图片留在正文层：与页眉水印分属不同 story。

    这不是“谁对谁错”，而是把「遮挡来自跨层绘制顺序」这一事实钉住：
    水印在页眉、图片在正文，Word 先画页眉后画正文，所以正文图片在上。
    """
    d = tempfile.mkdtemp()
    src = _make_doc_with_body_image(d)
    out = os.path.join(d, "out.docx")
    core.insert_watermark(src, ["text"], output_path=out,
                          text={"text": "机密", "font_size": 80})

    doc = Document(out)
    body_imgs = _body_drawings(doc)
    assert len(body_imgs) == 1, f"正文图片应留在正文层，实际 {len(body_imgs)}"
    assert len(_header_drawings(doc)) == 1, "页眉里只应有水印，不应混入正文图片"
    print("[层] 正文图片在 body / 水印在 header —— 分属两层，遮挡由 Word 绘制顺序决定")


def test_watermark_marker_independent_of_behind_doc():
    """识别与清除依赖 docPr 上的私有标记，不依赖 behindDoc。

    因此将来若把 behindDoc 改成 "0"（让水印显示在正文图片之上），
    检测/删除逻辑**不需要**跟着改——这个测试提前把这条不变式钉住。
    """
    d = tempfile.mkdtemp()
    src = _make_doc_with_body_image(d)
    out = os.path.join(d, "out.docx")
    core.insert_watermark(src, ["text"], output_path=out,
                          text={"text": "机密", "font_size": 80})

    assert engine_docx.detect_watermark_types(out) == {"text"}

    doc = Document(out)
    dr = _header_drawings(doc)[0]
    docPr = dr.find(".//" + qn("wp:docPr"))
    assert engine_docx._mark_of(dr) == engine_docx.MARK_TEXT
    # 人为改成浮于文字上方，识别与清除仍必须正常
    dr.find(".//" + qn("wp:anchor")).set("behindDoc", "0")
    doc.save(out)
    assert engine_docx.detect_watermark_types(out) == {"text"}, \
        "识别依赖标记而非 behindDoc：切换层级后仍应识别"
    cleared = os.path.join(d, "cleared.docx")
    core.clear_watermark(out, output_path=cleared, kinds=["text"])
    assert engine_docx.detect_watermark_types(cleared) == set(), \
        "清除依赖标记而非 behindDoc：切换层级后仍应删掉"
    print("[层] behindDoc 改为 0 后仍可识别/删除 —— 标记与层级解耦")


def test_insert_then_delete_roundtrip_with_body_image():
    """带正文图片的文档：插入→删除→重新插入，正文图片必须始终完好。"""
    d = tempfile.mkdtemp()
    src = _make_doc_with_body_image(d)
    wm = os.path.join(d, "wm.docx")
    core.insert_watermark(src, ["text"], output_path=wm,
                          text={"text": "机密", "font_size": 80})

    cleared = os.path.join(d, "cleared.docx")
    core.clear_watermark(wm, output_path=cleared, kinds=["text"])

    doc = Document(cleared)
    assert len(_body_drawings(doc)) == 1, "删除水印不得连带删掉正文图片"
    assert "用户正文" in [p.text for p in doc.paragraphs], "正文文字必须保留"

    again = os.path.join(d, "again.docx")
    core.insert_watermark(cleared, ["text"], output_path=again,
                          text={"text": "机密2", "font_size": 80})
    assert engine_docx.detect_watermark_types(again) == {"text"}
    assert len(_body_drawings(Document(again))) == 1, "重插后正文图片仍应完好"
    print("[层] 插入→删除→重插 往返：正文图片与文字始终完好")
