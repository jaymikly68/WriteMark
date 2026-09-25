"""视频水印端到端 + Word 按类型清除 回归测试。

覆盖：
1) video.add_video_watermark 对合成视频逐帧加水印（固定 / 滚动），产出文件含水印像素、音轨 mux 不崩。
2) engine_docx.detect_watermark_types 与 clear_watermark(kinds=...) 按类型清除正确。
3) GUI _clear 在“文字+图片”共存时弹「想要去除水印？」并正确映射选项。
"""
import os
import sys
import tempfile
import numpy as np
import imageio.v2 as iio

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

from PySide6.QtWidgets import QApplication, QMessageBox
from docx import Document
from PIL import Image

from watermark_tool import engine_docx, core
from watermark_tool.gui import App


def _make_docx_with_both(path, img):
    Document().save(path)
    engine_docx.insert_watermark(path, ["text", "image"],
        text={"text": "机密", "font_size": 80, "color": (200, 0, 0), "alpha": 120},
        image={"image_path": img, "alpha": 120})


def _make_video(path, W=160, H=120, N=24):
    w = iio.get_writer(path, fps=24, macro_block_size=1)
    for i in range(N):
        f = np.full((H, W, 3), 100, np.uint8)   # 中性灰背景，红/蓝只能来自水印
        w.append_data(f)
    w.close()


# ---------------------------------------------------------------------------
def test_video_fixed_and_scroll():
    d = tempfile.mkdtemp()
    src = os.path.join(d, "src.mp4")
    img = os.path.join(d, "wm.png")
    Image.new("RGBA", (40, 40), (0, 0, 255, 200)).save(img)
    _make_video(src)

    for mode in ("fixed", "scroll"):
        out = os.path.join(d, f"{mode}.mp4")
        res = engine_run = None
        from watermark_tool import video
        res = video.add_video_watermark(src, out, {
            "kinds": ["text", "image"], "mode": mode,
            "text": {"text": "机密", "color": (255, 0, 0), "alpha": 200, "size_frac": 0.15},
            "image": {"image_path": img, "alpha": 200, "img_frac": 0.18},
            "position": (0.8, 0.8), "scroll_speed": 0.15,
        })
        assert res["ok"], f"{mode} 处理失败: {res}"
        assert os.path.exists(out) and os.path.getsize(out) > 0
        # 末帧（fixed 稳定；scroll 此时水印已滚入画面）应同时含文字(红)与图片(蓝)水印
        fr = iio.get_reader(out, "ffmpeg").get_data(23)
        red = int((fr[:, :, 0] > 150).sum())
        blue = int((fr[:, :, 2] > 150).sum())
        print(f"[{mode}] 末帧 红(文字)={red} 蓝(图片)={blue}")
        assert red > 0, f"{mode} 末帧应含文字水印"
        assert blue > 0, f"{mode} 末帧应含图片水印"
    print("video 固定/滚动 端到端 PASS")


def test_docx_detect_and_clear_by_kind():
    d = tempfile.mkdtemp()
    src = os.path.join(d, "doc.docx")
    img = os.path.join(d, "wm.png")
    Image.new("RGBA", (40, 40), (255, 0, 0, 128)).save(img)
    _make_docx_with_both(src, img)

    types = engine_docx.detect_watermark_types(src)
    assert types == {"text", "image"}, f"应同时检测到文字与图片，实际 {types}"

    r = engine_docx.clear_watermark(src, kinds=["text"])
    assert r["removed"] >= 1
    assert engine_docx.detect_watermark_types(src) == {"image"}, "只去文字后应剩图片"

    engine_docx.clear_watermark(src, kinds=["image"])
    assert engine_docx.detect_watermark_types(src) == set(), "去图片后应无本工具水印"
    print("docx 按类型检测与清除 PASS")


def test_gui_clear_choice_dialog(monkeypatch):
    """_clear 在“文字+图片”共存时弹窗，且三选项正确映射 kinds。

    关键：GUI 的 _clear 通过 _run(QThread) 异步执行 clear_watermark，
    这里把 _run 替换成“同步立即执行 job”，从而能同步断言真正传给
    core.clear_watermark 的 kinds 参数；_ask_clear_choice 也直接被替换
    成返回对应选项，避免依赖真实 QMessageBox 交互。
    """
    d = tempfile.mkdtemp()
    src = os.path.join(d, "doc.docx")
    img = os.path.join(d, "wm.png")
    Image.new("RGBA", (40, 40), (255, 0, 0, 128)).save(img)

    app = QApplication.instance() or QApplication([])
    w = App()
    w.file_path = src
    w.out_edit.setText(core.default_output_path(src))

    # 记录真正调用 core.clear_watermark 的 kinds
    captured = {}
    real_clear = core.clear_watermark
    def fake_clear(p, output_path=None, kinds=None):
        captured["kinds"] = kinds
        return real_clear(p, output_path=output_path, kinds=kinds)
    monkeypatch.setattr(core, "clear_watermark", fake_clear)

    # 把异步 _run 改成“同步立刻执行 job”
    def sync_run(fn):
        fn()
    monkeypatch.setattr(w, "_run", sync_run)

    for expect_label, expect_choice, expect_kinds in (
            ("去除文字水印", "text", ["text"]),
            ("去除图片水印", "image", ["image"]),
            ("文字和图片水印都去除", "both", ["text", "image"])):
        # 每次都重置“文字+图片”共存的源文档
        _make_docx_with_both(src, img)
        w.file_path = src
        captured.clear()
        w._ask_clear_choice = lambda: expect_choice
        w._clear()
        assert captured["kinds"] == expect_kinds, \
            f"{expect_label}：期望 {expect_kinds}，实际 {captured['kinds']}"
    print("GUI 清除弹窗选项映射 PASS")


if __name__ == "__main__":
    from pytest import MonkeyPatch
    test_video_fixed_and_scroll()
    test_docx_detect_and_clear_by_kind()
    test_gui_clear_choice_dialog(MonkeyPatch())
    print("\nALL PASS")
