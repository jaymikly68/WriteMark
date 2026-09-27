"""
视频文字水印的“中文字体 / 西文字体”自定义（对标 Word 文本水印的中西文分字体）。

关键事实：视频没有“导入文档”动作，故字体来源与 Word 一致——本机已安装字体
（Windows 字体注册表）。实现上：
- GUI 视频文字区新增两个下拉框（复用 _make_font_combo）；
- _v_gather_opts() 把所选字体写入 text_cfg，引擎 _build_text_layer 已支持分字体渲染；
- 启动时用 FontWorker（仅读注册表、无需授权）扩充这两个下拉框。
"""
from __future__ import annotations

import os
import sys

sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", ".."))

from watermark_tool.gui import App


def test_video_font_combos_exist(app):
    w = App()
    assert hasattr(w, "v_cn_font_combo") and hasattr(w, "v_latin_font_combo")
    # 默认可编辑 + 含常见中西文字体
    cn = [w.v_cn_font_combo.itemText(i) for i in range(w.v_cn_font_combo.count())]
    la = [w.v_latin_font_combo.itemText(i) for i in range(w.v_latin_font_combo.count())]
    assert "微软雅黑" in cn
    assert "Arial" in la
    assert w.v_cn_font_combo.isEditable() and w.v_latin_font_combo.isEditable()


def test_video_gather_opts_carries_fonts(app):
    w = App()
    w.v_text_chk.setChecked(True)
    w.v_cn_font_combo.setCurrentText("楷体")
    w.v_latin_font_combo.setCurrentText("Times New Roman")
    opts = w._v_gather_opts()
    tc = opts["text"]
    assert tc.get("cn_font_name") == "楷体"
    assert tc.get("latin_font_name") == "Times New Roman"


def test_video_fonts_expand_on_scan(app):
    """_on_video_fonts_loaded 把系统字体去重并入两个下拉框，并保持当前选择。"""
    w = App()
    w.v_cn_font_combo.setCurrentText("微软雅黑")
    w.v_latin_font_combo.setCurrentText("Arial")
    w._on_video_fonts_loaded(["新字体甲", "新字体乙", "微软雅黑"], "系统字体")
    cn = [w.v_cn_font_combo.itemText(i) for i in range(w.v_cn_font_combo.count())]
    la = [w.v_latin_font_combo.itemText(i) for i in range(w.v_latin_font_combo.count())]
    assert "新字体甲" in cn and "新字体乙" in cn
    assert "新字体甲" in la and "新字体乙" in la
    # 重复项（如 微软雅黑）不应被重复加入
    assert cn.count("微软雅黑") == 1
    # 当前选择保持
    assert w.v_cn_font_combo.currentText() == "微软雅黑"
    assert w.v_latin_font_combo.currentText() == "Arial"


def test_video_startup_scan_guard(app, monkeypatch):
    """_startup_load_video_fonts 只触发一次（守护标记），且确实启动了后台扫描。"""
    w = App()
    started = {}
    monkeypatch.setattr(type(w), "_on_video_fonts_loaded",
                        lambda self, fonts, source: None)
    # 用真实 FontWorker 会在后台扫注册表，这里只验证“启动”与“去重”语义，不依赖其结果
    w._startup_load_video_fonts()
    assert w._video_fonts_loaded is True
    assert getattr(w, "_video_font_worker", None) is not None
    # 第二次调用应为 no-op（守卫生效）
    first = w._video_font_worker
    w._startup_load_video_fonts()
    assert w._video_font_worker is first


if __name__ == "__main__":
    import pytest
    pytest.main([__file__, "-q"])
