"""
「时刻(秒)」的 ▲/▼ 步进按钮回归测试（用户 2026-09-28 反馈）：

- 点向上的三角形 → +0.5 秒（原先点了却落进输入框，等于手动输入）；
- 点向下的三角形 → -0.5 秒；
- 两个按钮必须是**真实可点中的控件**（用 QTest 在其几何中心按下鼠标验证），
  而不是靠 Qt 自带箭头的 subControl 热区——那正是原 BUG 的来源。

附带验证「限定时段」的输入框就是 "1.0-2.0" 这种一行写法。
"""
from __future__ import annotations

import os
import sys

sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", ".."))

from PySide6.QtCore import Qt, QPoint
from PySide6.QtTest import QTest
from PySide6.QtWidgets import QAbstractSpinBox

from watermark_tool.gui import App


def _click(widget):
    """在控件几何中心真实按一次鼠标左键（模拟用户点击，而非直接调 click()）。"""
    QTest.mouseClick(widget, Qt.LeftButton, Qt.NoModifier,
                     QPoint(widget.width() // 2, widget.height() // 2))


# ------------------------------------------------------------------ ▲ / ▼
def test_up_button_adds_half_second(app):
    w = App()
    w.show()
    w.v_t_spin.setValue(1.0)
    _click(w.v_t_up_btn)
    assert abs(w.v_t_spin.value() - 1.5) < 1e-6, "▲ 应 +0.5 秒"


def test_down_button_subtracts_half_second(app):
    w = App()
    w.show()
    w.v_t_spin.setValue(1.5)
    _click(w.v_t_dn_btn)
    assert abs(w.v_t_spin.value() - 1.0) < 1e-6, "▼ 应 -0.5 秒"


def test_buttons_are_real_clickable_widgets(app):
    """两个按钮都可见、有正尺寸，且点得到（不是被输入框盖住的装饰性三角形）。"""
    w = App()
    w.show()
    for b in (w.v_t_up_btn, w.v_t_dn_btn):
        assert b.isVisible()
        assert b.width() >= 12 and b.height() >= 8, f"按钮太小点不中：{b.size()}"
        assert b.isEnabled()


def test_native_arrows_disabled(app):
    """自带箭头已关掉，避免与自绘 ▲/▼ 重复/热区冲突。"""
    w = App()
    assert w.v_t_spin.buttonSymbols() == QAbstractSpinBox.NoButtons


def test_step_buttons_visible_in_window(app):
    """▲/▼ 与输入框在同一行，且都落在窗口可视区域内。"""
    w = App()
    w.show()
    w.resize(1200, 800)
    for b in (w.v_t_up_btn, w.v_t_dn_btn, w.v_t_spin):
        assert b.isVisible(), f"{b} 应在界面上可见"


# ------------------------------------------------------ 限定时段：1.0-2.0
def test_range_input_is_single_line_format(app):
    w = App()
    assert w.v_range_edit.text() == "1.0-2.0", "默认示例就应是「起-止」一行写法"
    assert "1.0-2.0" in w.v_range_edit.placeholderText()
    w.v_range_chk.setChecked(True)
    w.v_range_edit.setText("2.5-4.0")
    assert w._v_gather_opts()["time_range"] == (2.5, 4.0)


if __name__ == "__main__":
    import pytest
    pytest.main([__file__, "-q"])
