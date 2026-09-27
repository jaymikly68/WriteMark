"""
视频/Word 图像水印的“PDF取任意页”能力（用户诉求：多页 PDF 当图片水印时可选任意页）：

- GUI 新增“PDF页码”微调框，仅在当前图片为 PDF 时可用，并按实际页数设上限；
- “PDF取任意页”短语悬浮时，其下方蓝色小字说明显现；
- 引擎 prepare_image_png 支持 page 参数（1-based，越界自动夹取）；
- _gather_opts 把 pdf_page 带入 image 配置，预览与插入均按所选页渲染。
"""
from __future__ import annotations

import os
import sys
import tempfile

sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", ".."))

from PySide6.QtCore import QEvent

from watermark_tool import engine_docx
from watermark_tool.gui import App


def _make_pdf(path, widths=(300, 500, 700)):
    import pymupdf
    doc = pymupdf.open()
    for i, w_pt in enumerate(widths):
        page = doc.new_page(width=w_pt, height=400)
        page.insert_text((50, 200), f"PAGE{i+1}")
    doc.save(path)
    doc.close()


def test_pdf_page_spin_hidden_until_pdf(app, tmp_path):
    w = App()
    w.image_enabled = True  # 图像水印启用后，PDF 页码才可能被启用
    # 默认（无图片）应禁用
    assert not w.img_pdf_page_spin.isEnabled()
    # 非 PDF 图片仍禁用
    txt = tmp_path / "a.png"
    txt.write_bytes(b"")
    w.img_edit.setText(str(txt))
    w._refresh_pdf_page_state()
    assert not w.img_pdf_page_spin.isEnabled()
    # PDF 应启用，且上限 = 实际页数
    pdf = tmp_path / "m.pdf"
    _make_pdf(str(pdf))
    w.img_edit.setText(str(pdf))
    w._refresh_pdf_page_state()
    assert w.img_pdf_page_spin.isEnabled()
    assert w.img_pdf_page_spin.maximum() == 3


def test_pdf_page_picker_selects_page(app, tmp_path):
    pdf = tmp_path / "m.pdf"
    _make_pdf(str(pdf))
    # 第 1 页（窄）与第 2 页（宽）渲染后宽度应不同，验证 page 参数生效
    img1 = engine_docx.prepare_image_png(str(pdf), 255, page=1)
    img2 = engine_docx.prepare_image_png(str(pdf), 255, page=2)
    assert img2.width > img1.width, "page=2 应渲染更宽的第 2 页"
    # 越界页码应夹取到有效范围而不报错
    img_big = engine_docx.prepare_image_png(str(pdf), 255, page=999)
    assert img_big is not None and img_big.width > 0


def test_pdf_page_in_gather(app, tmp_path):
    w = App()
    w.image_enabled = True
    pdf = tmp_path / "m.pdf"
    _make_pdf(str(pdf))
    w.img_edit.setText(str(pdf))
    w._refresh_pdf_page_state()
    w.img_pdf_page_spin.setValue(2)
    opts = w._gather_opts()
    assert opts["image"]["pdf_page"] == 2


def test_pdf_hover_helper_toggles(app):
    w = App()
    w.show()  # 让控件树进入可见态，isVisible() 才能反映显隐
    assert not w.img_pdf_helper.isVisible()
    w.img_pdf_hint.enterEvent(QEvent(QEvent.Type.Enter))
    assert w.img_pdf_helper.isVisible(), "悬浮“PDF取任意页”应显示蓝色说明"
    w.img_pdf_hint.leaveEvent(QEvent(QEvent.Type.Leave))
    assert not w.img_pdf_helper.isVisible(), "移出后应隐藏说明"


if __name__ == "__main__":
    import pytest
    pytest.main([__file__, "-q"])
