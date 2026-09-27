"""
三处 UI 修复的回归测试（用户 2026-09-27 反馈）：

1. 悬停「PDF取任意页」时**只**出现蓝色小字说明，不再同时弹出系统 tooltip（红/黄底那行）——
   根因是 _make_hover_hint 里额外调了 phrase.setToolTip(helper_text)。
2. 「时刻(秒)」微调框宽度足够，上下箭头不会被挤到错位——
   根因是 setMaximumWidth(80) 太窄。
3. 「浏览...」「选择图片...」按钮统一为蓝底样式（BTN_BLUE），共 6 处。
"""
from __future__ import annotations

import os
import sys

sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", ".."))

from PySide6.QtCore import QEvent
from PySide6.QtWidgets import QPushButton

from watermark_tool import gui as gui_mod
from watermark_tool.gui import App


# ---------------------------------------------------------------- 1) hover 只留蓝字
def test_hover_hint_has_no_system_tooltip(app):
    w = App()
    assert w.img_pdf_hint.toolTip() == "", \
        "悬浮短语不应再带系统 tooltip（否则与蓝字说明重复出现两份）"


def test_hover_hint_blue_helper_toggles(app):
    w = App()
    w.show()
    assert not w.img_pdf_helper.isVisible()
    w.img_pdf_hint.enterEvent(QEvent(QEvent.Type.Enter))
    assert w.img_pdf_helper.isVisible(), "悬浮应显示蓝色说明"
    # 蓝色：样式里必须是蓝字（#1565c0），不是红字
    assert "#1565c0" in w.img_pdf_helper.styleSheet()
    w.img_pdf_hint.leaveEvent(QEvent(QEvent.Type.Leave))
    assert not w.img_pdf_helper.isVisible()


# ---------------------------------------------------------------- 2) 时刻(秒) 箭头
def test_time_spin_uses_explicit_up_down_buttons(app):
    """「时刻(秒)」不再用 Qt 自带的窄箭头（点击热区会错位，点上箭头却落进输入框），
    改成一对独立的 ▲/▼ 按钮：上 +0.5 秒、下 -0.5 秒。"""
    from PySide6.QtWidgets import QAbstractSpinBox
    w = App()
    assert w.v_t_spin.buttonSymbols() == QAbstractSpinBox.NoButtons
    w.v_t_spin.setValue(1.0)
    w.v_t_up_btn.click()
    assert abs(w.v_t_spin.value() - 1.5) < 1e-6, "上箭头应 +0.5 秒"
    w.v_t_dn_btn.click()
    w.v_t_dn_btn.click()
    assert abs(w.v_t_spin.value() - 0.5) < 1e-6, "下箭头应 -0.5 秒"


# ---------------------------------------------------------------- 3) 蓝底浏览按钮
def _find_buttons(widget):
    return widget.findChildren(QPushButton)


def test_browse_buttons_are_blue(app):
    w = App()
    btns = [b for b in _find_buttons(w)
            if b.text().startswith("浏览...") or b.text().startswith("选择图片...")]
    assert len(btns) >= 6, f"应至少 6 个浏览/选择图片按钮，实际 {len(btns)}"
    for b in btns:
        ss = b.styleSheet()
        assert "#d3e3fd" in ss and "#0b57d0" in ss, \
            f"「{b.text()}」应为蓝底样式（BTN_BLUE），实际：{ss!r}"


def test_btn_blue_style_definition():
    assert "#d3e3fd" in gui_mod.BTN_BLUE
    assert "#0b57d0" in gui_mod.BTN_BLUE
    # 与主操作实蓝底区分：底色明显更浅
    assert "#1a73e8" not in gui_mod.BTN_BLUE


if __name__ == "__main__":
    import pytest
    pytest.main([__file__, "-q"])
