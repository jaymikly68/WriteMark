"""
视频水印「生效时段」（用户诉求：3 秒视频只在第 2 秒加水印，用来遮住那几秒才出现的
敏感信息；第 1、3 秒画面保持干净）。

覆盖：
- 默认不勾选 = 全片生效（不下发 time_range）；
- 勾选且区间合法 → _v_gather_opts 带 time_range；区间非法（止<=起）→ 仍按全片；
- 引擎 _in_time_range 的半开区间 [起, 止) 语义；
- 预览 compose_on_frame 与导出 add_video_watermark **同一套判定**：
  时段外无水印、时段内有水印（端到端读回成品帧验证）。
"""
from __future__ import annotations

import os
import sys

import numpy as np
import imageio.v2 as iio
from PIL import Image

sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", ".."))

from watermark_tool import video
from watermark_tool.gui import App


def _make_video(path, W=160, H=120, N=72, fps=24):
    """3 秒灰底视频（72 帧 @24fps），灰色背景上任何红/蓝都只能来自水印。"""
    w = iio.get_writer(path, fps=fps, macro_block_size=1)
    for _ in range(N):
        w.append_data(np.full((H, W, 3), 100, np.uint8))
    w.close()


def _red_png(path):
    Image.new("RGBA", (40, 40), (255, 0, 0, 255)).save(path)


def _opts(img_path, time_range=None):
    o = {
        "kinds": ["image"],
        "motion": "fixed",
        "image": {"image_path": img_path, "alpha": 255, "img_frac": 0.25,
                  "position": (0.1, 0.1)},
        "text": {"text": ""},
    }
    if time_range is not None:
        o["time_range"] = time_range
    return o


# ---------------------------------------------------------------------------
def test_in_time_range_semantics():
    assert video._in_time_range(0.0, None) is True          # 不限时段
    assert video._in_time_range(5.0, ()) is True
    assert video._in_time_range(1.5, (1.0, 2.0)) is True    # 区间内
    assert video._in_time_range(1.0, (1.0, 2.0)) is True    # 起点含
    assert video._in_time_range(2.0, (1.0, 2.0)) is False   # 终点不含（半开）
    assert video._in_time_range(0.5, (1.0, 2.0)) is False   # 区间前
    assert video._in_time_range(9.9, (1.0, 2.0)) is False   # 区间后
    assert video._in_time_range(1.5, (2.0, 1.0)) is True    # 非法区间 → 全片


def test_gather_default_full_video(app):
    w = App()
    assert "time_range" not in w._v_gather_opts(), "默认未勾选应全片生效"


def test_gather_time_range(app):
    """用户填法就是一行「1.0-2.0」，勾选后下发 (1.0, 2.0)。"""
    w = App()
    w.v_range_chk.setChecked(True)
    w.v_range_edit.setText("1.0-2.0")
    assert w._v_gather_opts()["time_range"] == (1.0, 2.0)
    # 区间非法（止 <= 起）→ 不下发，引擎按全片处理
    w.v_range_edit.setText("3.0-1.0")
    assert "time_range" not in w._v_gather_opts()
    # 填不出来（只有一个数字 / 全是字）→ 同样按全片
    w.v_range_edit.setText("2.5")
    assert "time_range" not in w._v_gather_opts()


def test_parse_time_range_accepts_common_separators():
    """分隔符怎么顺手怎么填：- ~ – 到 逗号 空格 都认。"""
    from watermark_tool.gui import parse_time_range
    assert parse_time_range("1.0-2.0") == (1.0, 2.0)
    assert parse_time_range("1.0~2.0") == (1.0, 2.0)
    assert parse_time_range("1.0 到 2.0") == (1.0, 2.0)
    assert parse_time_range("1.0，2.0") == (1.0, 2.0)
    assert parse_time_range("第1秒到第3秒") == (1.0, 3.0)
    assert parse_time_range("1-2") == (1.0, 2.0)
    assert parse_time_range("") is None
    assert parse_time_range("abc") is None


def test_range_check_toggles_input_enabled(app):
    w = App()
    w.v_range_chk.setChecked(False)
    assert not w.v_range_edit.isEnabled()
    w.v_range_chk.setChecked(True)
    assert w.v_range_edit.isEnabled()


def test_check_range_snaps_preview_into_range(app):
    """勾选时若当前预览时刻在时段外，自动跳到起点，便于立刻看到“这段有水印”。"""
    w = App()
    w.v_t_spin.setValue(5.0)          # 时段外
    w.v_range_edit.setText("2.0-3.0")
    w.v_range_chk.setChecked(True)
    assert w.v_t_spin.value() == 2.0


def test_range_hint_text(app):
    """填对了给蓝字说明，填错了给红字提示——用户一眼知道自己被认成了几秒到几秒。"""
    w = App()
    w.v_range_chk.setChecked(True)
    w.v_range_edit.setText("1.0-2.0")
    assert "1.0" in w.v_range_hint.text() and "2.0" in w.v_range_hint.text()
    assert "#1565c0" in w.v_range_hint.styleSheet(), "合法应为蓝字"
    w.v_range_edit.setText("abc")
    assert "#d93025" in w.v_range_hint.styleSheet(), "非法应为红字提示"


def test_compose_preview_gating(app):
    """预览与导出同一套判定：时段外的帧不应出现水印。"""
    raw = Image.new("RGB", (320, 180), (100, 100, 100))
    opts = _opts(None, time_range=(1.0, 2.0))
    opts["kinds"] = ["text"]
    opts["text"] = {"text": "机密", "color": (255, 0, 0), "alpha": 255,
                    "size_frac": 0.3, "position": (0.1, 0.1),
                    "cn_font_name": "微软雅黑", "latin_font_name": "Arial"}
    inside = video.compose_on_frame(raw, opts, t_sec=1.5)
    outside = video.compose_on_frame(raw, opts, t_sec=0.5)
    raw_px = list(raw.getdata())
    assert list(inside.getdata()) != raw_px, "时段内应叠上水印"
    assert list(outside.getdata()) == raw_px, "时段外画面应保持原样"


def test_export_time_range_end_to_end(app, tmp_path):
    """端到端：3 秒视频 + 时段 1.0~2.0 → 第 1、3 秒干净，第 2 秒含水印。"""
    src = str(tmp_path / "src.mp4")
    out = str(tmp_path / "out.mp4")
    img = str(tmp_path / "wm.png")
    _make_video(src)          # 72 帧 @24fps = 3 秒
    _red_png(img)

    video.add_video_watermark(src, out, _opts(img, time_range=(1.0, 2.0)))
    assert os.path.exists(out)

    r = iio.get_reader(out, "ffmpeg")
    frames = [f for f in r]
    r.close()
    assert len(frames) >= 60, f"成品帧数异常：{len(frames)}"

    def redness(f):
        a = np.asarray(f).astype(int)
        return int((a[:, :, 0] - a[:, :, 1]).max())   # 红 - 绿，灰底≈0

    assert redness(frames[12]) < 40, "t=0.5s（时段前）不应有水印"
    assert redness(frames[36]) > 100, "t=1.5s（时段内）应有水印"
    assert redness(frames[60]) < 40, "t=2.5s（时段后）不应有水印"


if __name__ == "__main__":
    import pytest
    pytest.main([__file__, "-q"])
