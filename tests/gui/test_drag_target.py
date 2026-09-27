"""
验证「文字 / 图片水印分别定位」（对应用户诉求：先定文字、再定图片，一次搞定，不必分两次）：

- 视频预览拖拽：选「图片」拖只改图片 X/Y，文字不动；选「文字」同理；
- Word 预览拖拽：选「图片」拖只改图像偏移，文字不动；选「文字」同理；
- _DragLabel 多标记：setMarkers 可同时画两个带标号的标记，且旧接口 setMarker 仍可用。
"""
from __future__ import annotations

import os
import sys

sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", ".."))

from PySide6.QtWidgets import QApplication

from watermark_tool.gui import App, _DragLabel, _combo_key


def test_video_drag_target_image_only(app, monkeypatch):
    w = App()
    w.v_text_chk.setChecked(True)
    w.v_img_chk.setChecked(True)
    w.v_drag_chk.setChecked(True)
    w.v_drag_target_combo.setCurrentIndex(w.v_drag_target_combo.findData("image"))
    t_x0, t_y0 = w.v_text_x_spin.value(), w.v_text_y_spin.value()
    w._on_video_frame_drag(0.2, 0.3)
    assert w.v_img_x_spin.value() == 20 and w.v_img_y_spin.value() == 30
    assert w.v_text_x_spin.value() == t_x0 and w.v_text_y_spin.value() == t_y0, \
        "选「图片」拖拽不应改动文字水印位置"
    assert _combo_key(w.v_img_pos_combo) == "自定义"


def test_video_drag_target_text_only(app, monkeypatch):
    w = App()
    w.v_text_chk.setChecked(True)
    w.v_img_chk.setChecked(True)
    w.v_drag_chk.setChecked(True)
    w.v_drag_target_combo.setCurrentIndex(w.v_drag_target_combo.findData("text"))
    i_x0, i_y0 = w.v_img_x_spin.value(), w.v_img_y_spin.value()
    w._on_video_frame_drag(0.7, 0.1)
    assert w.v_text_x_spin.value() == 70 and w.v_text_y_spin.value() == 10
    assert w.v_img_x_spin.value() == i_x0 and w.v_img_y_spin.value() == i_y0, \
        "选「文字」拖拽不应改动图片水印位置"


def test_word_drag_target_image_only(app, monkeypatch):
    w = App()
    w.text_enabled = True
    w.image_enabled = True
    w.preview_drag_chk.setChecked(True)
    w.preview_drag_target_combo.setCurrentIndex(
        w.preview_drag_target_combo.findData("image"))
    t_x0, t_y0 = w.text_offx_spin.value(), w.text_offy_spin.value()
    w._on_preview_drag(0.8, 0.2)   # 右上
    assert w.img_offx_spin.value() == 30 and w.img_offy_spin.value() == -30
    assert w.text_offx_spin.value() == t_x0 and w.text_offy_spin.value() == t_y0, \
        "选「图片」拖拽不应改动文字水印偏移"


def test_word_drag_target_text_only(app, monkeypatch):
    w = App()
    w.text_enabled = True
    w.image_enabled = True
    w.preview_drag_chk.setChecked(True)
    w.preview_drag_target_combo.setCurrentIndex(
        w.preview_drag_target_combo.findData("text"))
    i_x0, i_y0 = w.img_offx_spin.value(), w.img_offy_spin.value()
    w._on_preview_drag(0.2, 0.8)   # 左下
    assert w.text_offx_spin.value() == -30 and w.text_offy_spin.value() == 30
    assert w.img_offx_spin.value() == i_x0 and w.img_offy_spin.value() == i_y0, \
        "选「文字」拖拽不应改动图片水印偏移"


def test_drag_label_multi_markers(app):
    lbl = _DragLabel()
    lbl.setMarkers([
        {"fx": 0.2, "fy": 0.3, "color": (225, 6, 0), "label": "文", "key": "text"},
        {"fx": 0.7, "fy": 0.8, "color": (0, 120, 215), "label": "图", "key": "image"},
    ])
    assert len(lbl._markers) == 2
    # 兼容旧接口：setMarker 仍可用（单标记）
    lbl.setMarker(0.5, 0.5)
    assert len(lbl._markers) == 1 and lbl._markers[0]["key"] == "single"
    # 同时，clearMarkers / clearMarker 等价
    lbl.clearMarkers()
    assert lbl._markers == []
    lbl.setMarker(0.1, 0.9)
    lbl.clearMarker()
    assert lbl._markers == []


def test_drag_label_active_target(app):
    """setActiveMarkerKey 让拖拽只更新被选中的那一个标记。"""
    lbl = _DragLabel()
    lbl.setMarkers([
        {"fx": 0.2, "fy": 0.3, "color": (225, 6, 0), "label": "文", "key": "text"},
        {"fx": 0.7, "fy": 0.8, "color": (0, 120, 215), "label": "图", "key": "image"},
    ])
    lbl.setActiveMarkerKey("image")
    lbl._set_active(0.9, 0.9)   # 模拟拖到右上
    assert lbl._markers[1]["fx"] == 0.9 and lbl._markers[1]["fy"] == 0.9
    assert lbl._markers[0]["fx"] == 0.2 and lbl._markers[0]["fy"] == 0.3, \
        "拖拽只应移动 active 标记（图片），文字标记不动"


if __name__ == "__main__":
    app = QApplication([])
    import pytest
    pytest.main([__file__, "-q"])
    app.quit()
