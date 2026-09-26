"""
video.py「失败路径」专项验证（本轮新增，针对两个静默失败的缺陷）。

1. 静默假成功：mux 失败后要用 os.replace 把无声临时文件改名成成品，
   若这一步也失败（输出路径不可写 / 成品被占用 / 磁盘满），函数**绝不能**
   返回 {"ok": True}——否则用户看到“成功”提示，output 却根本不存在。
   这条测试按你要求模拟 _mux_audio 返回 False 且 os.replace 抛异常。

2. 0 字节孤儿临时文件：mkstemp 建出的空临时 mp4 与 iio.get_writer 的构造
   原来落在 try/finally 保护范围之外，一旦 get_writer 自己抛异常
   （自带的 ffmpeg 缺失/损坏），%TMP% 下就会留下一个 0 字节 .mp4。

两条都用假 mkstemp 抓住 video.py 自己创建的那条临时路径来断言「最终不存在」。
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
    """让 video.py 的 mkstemp 返回一条我们可控的 (fd, path)，并真实建出这个空文件。

    返回 (path, fd)：调用方需自己 os.close(fd)（video.py 本就会这么做）。
    """
    path = os.path.join(tempfile.gettempdir(),
                        "wm_review_fail_%d.mp4" % os.getpid())
    fd = os.open(path, os.O_CREAT | os.O_WRONLY, 0o600)
    monkeypatch.setattr(video.tempfile, "mkstemp", lambda *a, **k: (fd, path))
    return path, fd


# ---------------------------------------------------------------------------
# 缺陷 1：mux 失败 + 改名也失败 → 静默假成功
# ---------------------------------------------------------------------------
def test_mux_fail_and_replace_fail_must_not_return_ok(tmp_path, monkeypatch):
    """mux 失败且 os.replace 抛异常：函数绝不能返回 ok: True。

    契约允许两种修法——抛 RuntimeError，或返回 {"ok": False, ...}；
    本测试锁的是「不得静默成功」这个结果，不锁具体实现。
    """
    d = str(tmp_path)
    src = os.path.join(d, "src.mp4")
    out = os.path.join(d, "out.mp4")
    _make_video(src)

    tmp, _fd = _grab_tmp_path(monkeypatch)
    monkeypatch.setattr(video, "_mux_audio", lambda *a, **k: False)

    def _boom(*a, **k):
        raise OSError(13, "Permission denied", a[0] if a else "")
    # video.py 用的是全局 os 模块，这里按模块属性打桩，monkeypatch 会还原
    monkeypatch.setattr(video.os, "replace", _boom)

    try:
        res = video.add_video_watermark(src, out, _text_opts())
    except RuntimeError as e:
        assert "输出文件未生成" in str(e), (
            f"改名失败应明确报错，实际：{e}")
    except Exception as e:
        raise AssertionError(
            f"失败路径抛出了非预期异常（应抛 RuntimeError 或返回 ok:False）：{e!r}")
    else:
        assert not res.get("ok"), (
            f"os.replace 改名失败，output 并未生成，却仍返回成功：{res}")

    assert not os.path.exists(out), (
        f"改名失败后不应声称产出成品，但 {out} 存在")
    assert not os.path.exists(tmp), f"改名失败后残留临时文件: {tmp}"
    print("[review5] video 静默假成功已消除：mux+改名双失败 -> 明确失败，无临时文件残留")


def test_mux_fail_replace_fail_error_names_output(tmp_path, monkeypatch):
    """错误信息必须说清“输出文件未生成”，而不是留一句无上下文的 Permission denied。"""
    d = str(tmp_path)
    src = os.path.join(d, "src.mp4")
    out = os.path.join(d, "out.mp4")
    _make_video(src)

    tmp, _fd = _grab_tmp_path(monkeypatch)
    monkeypatch.setattr(video, "_mux_audio", lambda *a, **k: False)
    monkeypatch.setattr(video.os, "replace",
                        lambda *a, **k: (_ for _ in ()).throw(
                            OSError(28, "No space left on device")))

    with pytest.raises(RuntimeError) as ei:
        video.add_video_watermark(src, out, _text_opts())

    msg = str(ei.value)
    assert "输出文件未生成" in msg, f"错误信息应点明输出未生成，实际：{msg}"
    assert "No space left on device" in msg, (
        f"错误信息应带上底层原因方便排障，实际：{msg}")
    assert not os.path.exists(tmp), f"失败后残留临时文件: {tmp}"
    print(f"[review5] video 失败信息可读：{msg}")


# ---------------------------------------------------------------------------
# 缺陷 2：get_writer 构造失败 -> 0 字节孤儿 + 异常被 NameError 掩盖
# ---------------------------------------------------------------------------
def test_get_writer_failure_removes_empty_tmp_and_keeps_reason(tmp_path, monkeypatch):
    """iio.get_writer 抛异常：刚建的空临时 mp4 必须被清掉，且真实异常原因原样上抛。"""
    d = str(tmp_path)
    src = os.path.join(d, "src.mp4")
    out = os.path.join(d, "out.mp4")
    _make_video(src)

    # 注意：这个 fd 由 video.py 自己 os.close()，测试侧不能再关一次（double close
    # 会触发 EBADF，反而掩盖真正想验证的东西）。
    tmp, _fd = _grab_tmp_path(monkeypatch)
    assert os.path.exists(tmp) and os.path.getsize(tmp) == 0, (
        "前置条件：mkstemp 建出的是一个 0 字节 .mp4（这正是会留下的孤儿文件）")

    def _boom_get_writer(*a, **k):
        raise RuntimeError("模拟自带的 ffmpeg 二进制损坏")
    monkeypatch.setattr(video.iio, "get_writer", _boom_get_writer)

    with pytest.raises(RuntimeError) as ei:
        video.add_video_watermark(src, out, _text_opts())

    # 关键：异常原因必须是底层那个，而不是 finally 里 writer.close() 造成的 NameError
    assert "模拟自带的 ffmpeg 二进制损坏" in str(ei.value), (
        f"真实异常原因被掩盖了，实际抛出：{type(ei.value).__name__}: {ei.value}")
    assert not os.path.exists(tmp), (
        f"get_writer 构造失败后残留 0 字节孤儿临时文件（会堆在 %TMP%）: {tmp}")
    print("[review5] video get_writer 失败：异常原因保留，0 字节孤儿临时文件已清理")


def test_get_writer_failure_leaves_no_orphan_across_repeats(tmp_path, monkeypatch):
    """连续两次失败（模拟用户在 ffmpeg 坏掉后反复点导出）：%TMP% 下不留任何本函数的 .mp4。"""
    d = str(tmp_path)
    src = os.path.join(d, "src.mp4")
    out = os.path.join(d, "out.mp4")
    _make_video(src)

    before = set(f for f in os.listdir(tempfile.gettempdir()) if f.endswith(".mp4"))

    for _ in range(2):
        path = os.path.join(tempfile.gettempdir(),
                            "wm_review_rep_%d.mp4" % os.getpid())
        if os.path.exists(path):
            os.remove(path)
        fd = os.open(path, os.O_CREAT | os.O_WRONLY, 0o600)
        monkeypatch.setattr(video.tempfile, "mkstemp", lambda *a, **k: (fd, path))
        monkeypatch.setattr(video.iio, "get_writer",
                            lambda *a, **k: (_ for _ in ()).throw(
                                RuntimeError("模拟 ffmpeg 缺失")))
        with pytest.raises(RuntimeError):
            video.add_video_watermark(src, out, _text_opts())
        monkeypatch.undo()

    after = set(f for f in os.listdir(tempfile.gettempdir()) if f.endswith(".mp4"))
    leaked = (after - before)
    assert not leaked, f"失败后 %TMP% 下残留了 .mp4: {leaked}"
    print("[review5] video 连续失败：%TMP% 下未新增任何 .mp4 残留")
