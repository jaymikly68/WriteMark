"""
video.py（风险 j 那一次修改）的收尾语义验证。

84a773e 把「writer.close() 前置到 _mux_audio 之前」和「finally 清理临时 mp4」
加进了 add_video_watermark。本文件逐条验证这条维护约定是否真的成立：

1. 正常完成时 writer 是否已正确关闭（缓冲刷盘）；
2. `_mux_audio()` 是否一定发生在 writer.close() 之后（否则 mux 读到的是半成品）；
3. 正常成功（含 mux 失败降级）后临时文件是否删除；
4. 中途异常后临时文件是否删除；
5. finally 里第二次 `writer.close()` 是否安全（重复 close 不应掀翻流程）。

手法：用假 mkstemp 抓住 video.py 自己创建的那条临时文件路径，直接断言它最终
不存在；用假 writer 记录事件顺序来证明时序；另有一条走真实 ffmpeg 的用例。
"""
from __future__ import annotations

import os
import sys
import tempfile

import numpy as np
import imageio.v2 as iio
import pytest

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from watermark_tool import video  # noqa: E402


def _make_video(path, W=96, H=72, N=6):
    w = iio.get_writer(path, fps=24, macro_block_size=1)
    try:
        for _ in range(N):
            w.append_data(np.full((H, W, 3), 100, np.uint8))
    finally:
        w.close()


def _text_opts():
    return {"kinds": ["text"], "motion": "fixed", "scroll_speed": 0,
            "text": {"text": "X", "color": (255, 0, 0), "alpha": 200,
                     "size_frac": 0.3, "position": (0.2, 0.2)}}


def _grab_tmp_path(monkeypatch):
    """让 video.py 的 mkstemp 返回一条我们可控的 (fd, path)（真实建出文件）。"""
    path = os.path.join(tempfile.gettempdir(),
                        "wm_review_tmp_%d.mp4" % os.getpid())
    fd = os.open(path, os.O_CREAT | os.O_WRONLY, 0o600)
    monkeypatch.setattr(video.tempfile, "mkstemp", lambda *a, **k: (fd, path))
    return path


# ---------------------------------------------------------------------------
def test_video_temp_removed_after_success(tmp_path, monkeypatch):
    """正常成功（真实 writer + 真实 mux）：成品有内容，临时无声视频被删掉。"""
    d = str(tmp_path)
    src = os.path.join(d, "src.mp4")
    out = os.path.join(d, "out.mp4")
    _make_video(src)

    tmp = _grab_tmp_path(monkeypatch)          # 抓住 video.py 用的临时文件

    res = video.add_video_watermark(src, out, _text_opts())
    assert res.get("ok"), res
    assert os.path.exists(out) and os.path.getsize(out) > 0, "成品应真实存在且有内容"
    assert not os.path.exists(tmp), f"成功路径残留临时文件: {tmp}"
    print("[review3] video 成功路径：成品有内容，临时无声视频已删除")


def test_video_temp_removed_when_mux_falls_back(tmp_path, monkeypatch):
    """mux 失败降级（无音轨 / ffmpeg 异常）：os.replace 兜底后临时文件同样不存在。"""
    d = str(tmp_path)
    src = os.path.join(d, "src.mp4")
    out = os.path.join(d, "out.mp4")
    _make_video(src)

    tmp =     _grab_tmp_path(monkeypatch)
    monkeypatch.setattr(video, "_mux_audio", lambda *a, **k: False)

    res = video.add_video_watermark(src, out, _text_opts())
    assert res.get("ok"), res
    assert os.path.exists(out), "mux 失败时应退化为无声视频（把临时文件改名成成品）"
    assert not os.path.exists(tmp), f"降级路径残留临时文件: {tmp}"
    print("[review3] video mux 失败降级：临时文件被 rename 成成品，无残留")


def test_video_temp_removed_on_midway_exception(tmp_path, monkeypatch):
    """中途异常（写帧失败）：异常照常抛出，临时文件仍被 finally 清掉。"""
    d = str(tmp_path)
    src = os.path.join(d, "src.mp4")
    out = os.path.join(d, "out.mp4")
    _make_video(src)

    tmp = _grab_tmp_path(monkeypatch)
    monkeypatch.setattr(video, "_mux_audio", lambda *a, **k: True)

    class _BoomWriter:
        def __init__(self, *a, **k):
            pass

        def append_data(self, data):
            raise RuntimeError("模拟写入失败（磁盘满）")

        def close(self):
            pass

    monkeypatch.setattr(video.iio, "get_writer", lambda *a, **k: _BoomWriter())

    with pytest.raises(RuntimeError, match="模拟写入失败"):
        video.add_video_watermark(src, out, _text_opts())

    assert not os.path.exists(tmp), \
        f"中途异常后残留临时文件（会堆在 %TMP%）: {tmp}"
    print("[review3] video 中途异常：异常向上抛出，临时文件已清理")


def test_writer_closed_before_mux_and_twice_is_safe(tmp_path, monkeypatch):
    """时序与幂等：writer 先关闭（刷盘）→ 再 mux；finally 里的第二次 close 不报错。"""
    d = str(tmp_path)
    src = os.path.join(d, "src.mp4")
    out = os.path.join(d, "out.mp4")
    _make_video(src)

    events = []
    monkeypatch.setattr(video, "_mux_audio",
                        lambda *a, **k: events.append("mux") or True)

    class _SpyWriter:
        def __init__(self, *a, **k):
            pass

        def append_data(self, data):
            events.append("frame")

        def close(self):
            events.append("close")

    monkeypatch.setattr(video.iio, "get_writer", lambda *a, **k: _SpyWriter())

    res = video.add_video_watermark(src, out, _text_opts())
    assert res.get("ok"), res

    assert events.index("close") < events.index("mux"), (
        f"_mux_audio 必须在 writer.close() 之后执行，实际事件顺序: {events}"
    )
    assert events.count("close") >= 2, (
        f"finally 里应再 close 一次以兜底，实际事件顺序: {events}"
    )
    assert events[-1] == "close", f"事件序列应以 finally 的 close 收尾: {events}"
    print(f"[review3] video 时序：{events} —— mux 在 writer 刷盘之后，重复 close 安全")
