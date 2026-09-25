"""视频水印端到端 + Word 按类型清除 回归测试。

覆盖：
1) video.add_video_watermark 对合成视频逐帧加水印（固定 / 滚动），产出文件含水印像素、音轨 mux 不崩。
2) 播放方式 x 水印类型 的搭配矩阵：固定 / 滚动 / 固定+滚动，以及逐类型覆盖
   （如“文字固定 + 图片滚动”）；文字与图片各有独立位置锚点。
3) engine_docx.detect_watermark_types 与 clear_watermark(kinds=...) 按类型清除正确。
4) GUI _clear 在“文字+图片”共存时弹「想要去除水印？」并正确映射选项。
5) GUI 视频分区的播放方式三选一与类型勾选能如实转成引擎参数。
"""
import os
import sys
import tempfile
import numpy as np
import imageio.v2 as iio

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

from PySide6.QtCore import QCoreApplication
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
        from watermark_tool import video
        res = video.add_video_watermark(src, out, {
            "kinds": ["text", "image"], "motion": mode,
            "text": {"text": "机密", "color": (255, 0, 0), "alpha": 200, "size_frac": 0.15,
                     "position": (0.8, 0.8)},
            "image": {"image_path": img, "alpha": 200, "img_frac": 0.18,
                      "position": (0.1, 0.1)},
            "scroll_speed": 0.15,
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


def test_video_motion_matrix():
    """播放方式 × 水印类型 的搭配矩阵必须全部可用且叠加正确。

    覆盖：固定 / 滚动 / 固定+滚动，以及“文字固定+图片滚动”这类逐类型覆盖。
    """
    d = tempfile.mkdtemp()
    src = os.path.join(d, "src.mp4")
    img = os.path.join(d, "wm.png")
    Image.new("RGBA", (60, 60), (0, 0, 255, 220)).save(img)
    _make_video(src, W=200, H=140, N=30)

    from watermark_tool import video

    cases = [
        # (名称, kinds, text位置, image位置, text motion, image motion, 期望 motion)
        ("固定+文字", ["text"], (0.8, 0.8), None, "fixed", None, ["fixed"]),
        ("滚动+图片", ["image"], None, (0.8, 0.8), None, "scroll", ["scroll"]),
        ("固定+滚动+文字", ["text"], (0.8, 0.8), None, "both", None, ["fixed", "scroll"]),
        ("固定+滚动+图片", ["image"], None, (0.8, 0.8), None, "both", ["fixed", "scroll"]),
        ("文字固定+图片滚动+双类型", ["text", "image"], (0.8, 0.8), (0.05, 0.05),
         "fixed", "scroll", ["fixed", "scroll"]),
        ("文字滚动+图片固定+双类型", ["text", "image"], (0.8, 0.8), (0.05, 0.05),
         "scroll", "fixed", ["fixed", "scroll"]),
        ("跟随全局双图层", ["text", "image"], (0.8, 0.8), (0.05, 0.05),
         "跟随", "跟随", ["fixed", "scroll"]),
    ]
    for name, kinds, tpos, ipos, tmotion, imotion, expect in cases:
        text_cfg = {"text": "机密", "color": (255, 0, 0), "alpha": 220, "size_frac": 0.12}
        image_cfg = {"image_path": img, "alpha": 220, "img_frac": 0.16}
        if tpos:
            text_cfg["position"] = tpos
        if ipos:
            image_cfg["position"] = ipos
        if tmotion:
            text_cfg["motion"] = tmotion
        if imotion:
            image_cfg["motion"] = imotion

        out = os.path.join(d, f"{name.replace('+', 'p').replace(' ', '')}.mp4")
        res = video.add_video_watermark(src, out, {
            "kinds": kinds, "motion": "both",
            "text": text_cfg, "image": image_cfg, "scroll_speed": 0.2,
        })
        assert res["ok"], f"{name} 处理失败: {res}"
        assert res.get("motion") == "+".join(expect), \
            f"{name}: 期望 motion={'+'.join(expect)}，实际 {res.get('motion')}"
        assert os.path.exists(out) and os.path.getsize(out) > 0
        fr = iio.get_reader(out, "ffmpeg").get_data(29)
        red = int((fr[:, :, 0] > 150).sum())
        blue = int((fr[:, :, 2] > 150).sum())
        if "text" in kinds:
            assert red > 0, f"{name} 末帧应含文字水印(红)，实际红={red}"
        if "image" in kinds:
            assert blue > 0, f"{name} 末帧应含图片水印(蓝)，实际蓝={blue}"
        print(f"[{name}] motion={res.get('motion')} 红={red} 蓝={blue}")

    # 归一化：motion 字段的所有写法都应收敛到合法值
    assert video._normalize_motion("fixed") == ["fixed"]
    assert video._normalize_motion("scroll") == ["scroll"]
    for alias in ("both", "all", "fixed+scroll", "固定+滚动"):
        assert video._normalize_motion(alias) == ["fixed", "scroll"], alias
    # 旧字段 mode 仍被兼容
    assert video._normalize_motion("scroll") == ["scroll"]
    print("video 播放方式 x 水印类型 搭配矩阵 PASS")


def test_video_export_resolution_and_fps():
    """导出分辨率 / 帧率必须按 opts 生效，且水印依然被叠加。

    历史问题：水印图层按“源分辨率”构建，放大导出时等于把小图二次拉伸，
    结果就是“图片水印又糊又小”。现在图层按输出分辨率构建。
    """
    from watermark_tool import video

    d = tempfile.mkdtemp()
    src = os.path.join(d, "src.mp4")
    _make_video(src, W=320, H=240, N=12)
    out = os.path.join(d, "out_720.mp4")

    res = video.add_video_watermark(src, out, {
        "kinds": ["text"], "motion": "fixed", "scroll_speed": 0,
        "out_size": (720, 480), "fps": 30, "crf": 18,
        "text": {"text": "机密", "color": (255, 0, 0), "alpha": 220, "size_frac": 0.15,
                 "position": (0.8, 0.8)},
    })
    assert res["ok"], res
    meta = iio.get_reader(out, "ffmpeg").get_meta_data()
    w, h = meta["size"]
    assert (w, h) == (720, 480), f"导出分辨率应为 720x480，实际 {w}x{h}"
    assert abs(float(meta["fps"]) - 30) < 1.5, f"导出帧率应约 30，实际 {meta['fps']}"

    fr = iio.get_reader(out, "ffmpeg").get_data(11)
    red = int((fr[:, :, 0] > 150).sum())
    assert red > 0, "放大到 720x480 后末帧仍应含文字水印"
    assert res["size"] == (720, 480) and abs(res["fps"] - 30) < 0.01 and res["crf"] == 18
    print("video 导出分辨率/帧率生效 PASS")


def test_video_crf_passthrough(monkeypatch):
    """清晰度参数：不能再用 imageio 默认的 crf≈25，必须把 crf 传进 ffmpeg。"""
    from watermark_tool import video

    captured = {}

    class _FakeWriter:
        def __init__(self, *a, **k):
            captured.update(k)

        def append_data(self, data):
            pass

        def close(self):
            pass

    d = tempfile.mkdtemp()
    src = os.path.join(d, "src.mp4")
    _make_video(src, W=64, H=48, N=6)
    # 注意：必须在生成测试视频之后再打这个补丁（iio 是同一个模块对象）
    monkeypatch.setattr(video.iio, "get_writer",
                        lambda *a, **k: _FakeWriter(*a, **k))
    video.add_video_watermark(src, os.path.join(d, "o.mp4"), {
        "kinds": ["text"], "motion": "fixed", "scroll_speed": 0,
        "crf": 14,
        "text": {"text": "X", "color": (255, 0, 0), "alpha": 200, "size_frac": 0.2},
    })
    assert captured.get("quality") is None, "必须关掉 imageio 的 quality 默认档"
    params = captured.get("output_params") or []
    assert "-crf" in params and params[params.index("-crf") + 1] == "14", \
        f"crf 未传进 ffmpeg：{params}"
    print("video crf 参数透传 PASS")


def test_image_layer_supersample_sharpness():
    """小图大幅放大时走超采样 + 轻锐化，锐度不能低于直接 LANCZOS。"""
    from watermark_tool import video
    from PIL import Image, ImageFilter

    d = tempfile.mkdtemp()
    p = os.path.join(d, "small.png")
    # 高频边缘图：能否保住边缘正是“糊”的判定点
    im = Image.new("RGBA", (24, 24), (0, 0, 0, 0))
    px = im.load()
    for y in range(24):
        for x in range(24):
            if (x // 2 + y // 2) % 2 == 0:
                px[x, y] = (255, 255, 255, 255)
    im.save(p)

    layer = video._build_image_layer(p, 1920, 1080, {"img_frac": 0.2, "alpha": 255})
    target_w = 1920 * 0.2
    assert abs(layer.size[0] - target_w) <= 2, f"图层宽度异常: {layer.size}"
    assert layer.size[0] > 24 * 1.5, "应触发放大分支"

    def _grad(img):
        g = img.convert("L").filter(ImageFilter.FIND_EDGES)
        a = np.asarray(g, dtype=float)
        return float(a.std())

    direct = im.resize((int(target_w), int(round(24 * target_w / 24))), Image.LANCZOS)
    sharp = video._resize_for_watermark(im, (int(target_w), int(round(24 * target_w / 24))))
    assert _grad(sharp) >= _grad(direct) * 0.98, \
        f"超采样结果应不比直接放大更糊: {_grad(sharp)} vs {_grad(direct)}"
    print("video 图片水印超采样锐度 PASS")


def test_gui_export_spec_defaults_and_limits():
    """导出规格：默认 1080P / 跟随屏幕刷新率，且一律不超过显示器上限。"""
    from watermark_tool import video
    from watermark_tool.gui import App, _screen_geometry, _screen_refresh_rate

    app = QApplication.instance() or QApplication([])
    w = App()
    disp_w, disp_h = _screen_geometry()
    rate = _screen_refresh_rate()

    assert w.v_res_items, "至少应提供一个可选分辨率"
    for lb, iw, ih in w.v_res_items:
        assert iw <= disp_w and ih <= disp_h, \
            f"选项 {lb} 超过显示器上限 {disp_w}x{disp_h}"
    idx1080 = [i for i, (lb, _a, _b) in enumerate(w.v_res_items) if lb.startswith("1080P")]
    if idx1080:
        assert w.v_res_combo.currentIndex() == idx1080[0], "默认导出分辨率应为 1080P"
    assert w._v_out_spec() == w.v_res_items[w.v_res_combo.currentIndex()][1:], \
        "解析出的分辨率与界面选项不一致"

    for f in w.v_fps_items:
        assert f <= rate, f"帧率选项 {f} 超过屏幕刷新率 {rate}"
    assert w._v_out_fps() == max(w.v_fps_items), "默认帧率应跟随屏幕刷新率"
    assert w._v_out_crf() == video.CRF_PRESETS["标准"], "默认画质应为标准"

    ow, oh = w._v_out_spec()
    assert isinstance(ow, int) and isinstance(oh, int) and ow > 0 and oh > 0, \
        f"导出分辨率应为正整数，实际 {w._v_out_spec()}"
    print("GUI 导出规格默认与上限 PASS")


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


def test_gui_video_run_button_clicked(monkeypatch):
    """回归：点击视频分区「开始加水印」必须真的跑起来，且按钮是蓝底黑字高亮。

    历史 bug：_vrun 里写成 self.v_gather_opts()（实际方法名是 _v_gather_opts），
    点击后抛 AttributeError 被 Qt 静默吞掉 —— 冻结 exe 无控制台，表现为
    “点了没反应”。这里用真实视频走一遍点击路径，确保产出文件。
    """
    d = tempfile.mkdtemp()
    src = os.path.join(d, "src.mp4")
    out = os.path.join(d, "out.mp4")
    _make_video(src)

    app = QApplication.instance() or QApplication([])
    w = App()

    # 弹窗不阻塞测试
    monkeypatch.setattr(QMessageBox, "information", lambda *a, **k: None)
    monkeypatch.setattr(QMessageBox, "warning", lambda *a, **k: None)
    monkeypatch.setattr(QMessageBox, "critical", lambda *a, **k: None)

    # 按钮高亮（蓝底黑字）
    ss = w.v_run_btn.styleSheet()
    assert "#1a73e8" in ss and "color:#000" in ss, f"「开始加水印」应为蓝底黑字，实际样式: {ss}"

    w.v_src_edit.setText(src)
    w.v_out_edit.setText(out)
    w.v_text_chk.setChecked(True)
    w.v_img_chk.setChecked(False)

    # 播放方式三选一必须齐备，且默认“固定”
    assert w.v_motion_fixed.text() == "固定" and w.v_motion_scroll.text() == "滚动" \
        and w.v_motion_both.text() == "固定+滚动"
    assert w.v_motion_fixed.isChecked(), "默认播放方式应为「固定」"

    import time
    w._v_run()          # 直接走槽函数入口（含异常兜底）
    # 必须手动泵事件循环：worker 的 result/progress/finished 都是排队连接，
    # 主线程只 sleep 的话信号不会投递，按钮就一直停在“处理中”。
    for _ in range(160):
        QCoreApplication.processEvents()
        worker = getattr(w, "_v_worker", None)
        if not (worker and worker.isRunning()):
            break
        time.sleep(0.25)
    QCoreApplication.processEvents()

    assert os.path.exists(out) and os.path.getsize(out) > 0, "点击后应产出带水印的视频"
    assert w.v_run_btn.isEnabled(), "处理结束后「开始加水印」应恢复可用"
    print("GUI 视频「开始加水印」点击生效 PASS")


def test_gui_motion_and_kind_matrix():
    """GUI 侧：播放方式三选一 + 文字/图片类型搭配，_v_gather_opts 必须如实反映。"""
    from watermark_tool.gui import App, _V_MOTION_MAP

    app = QApplication.instance() or QApplication([])
    w = App()

    def pick(label):
        for rb in (w.v_motion_fixed, w.v_motion_scroll, w.v_motion_both):
            if rb.text() == label:
                rb.setChecked(True)
                return
        raise AssertionError(f"找不到播放方式按钮：{label}")

    img = os.path.join(tempfile.mkdtemp(), "wm.png")
    Image.new("RGBA", (40, 40), (0, 0, 255, 200)).save(img)

    w.v_text_chk.setChecked(True)
    w.v_img_chk.setChecked(True)
    w.v_img_edit.setText(img)

    for label, expect in (("固定", "fixed"), ("滚动", "scroll"), ("固定+滚动", "both")):
        pick(label)
        opts = w._v_gather_opts()
        assert opts["motion"] == expect, f"全局{label} -> {opts['motion']}，期望 {expect}"
        # 两个下拉默认“跟随”，所以逐类型的 motion 也应是同一个值
        assert opts["text"]["motion"] == expect, label
        assert opts["image"]["motion"] == expect, label
    # 只有滚动生效时，滚动速度才被保留；纯固定时应归零（引擎侧会有兜底）
    pick("固定")
    assert w._v_gather_opts()["scroll_speed"] == 0, "纯固定模式下滚动速度应为 0"

    # 逐类型覆盖：文字固定 / 图片滚动（与全局“固定+滚动”并存也不冲突）
    pick("固定+滚动")
    w.v_text_motion_combo.setCurrentText("固定")
    w.v_img_motion_combo.setCurrentText("滚动")
    opts = w._v_gather_opts()
    assert opts["motion"] == "both", "全局仍应是 fixed+scroll"
    assert opts["text"]["motion"] == "fixed"
    assert opts["image"]["motion"] == "scroll"
    # 位置各自独立，避免同时固定时完全重叠
    assert opts["text"]["position"] != opts["image"]["position"], \
        "文字与图片应有各自的位置锚点"
    assert opts["text"]["position"] in {(0.82, 0.85), (0.82, 0.08), (0.08, 0.85),
                                        (0.08, 0.08), (0.35, 0.40)}
    assert opts["image"]["position"] in {(0.82, 0.85), (0.82, 0.08), (0.08, 0.85),
                                         (0.08, 0.08), (0.35, 0.40)}

    # 只勾选文字时不拼图片配置进 kinds
    w.v_img_chk.setChecked(False)
    assert w._v_gather_opts()["kinds"] == ["text"]
    # 一个都不勾选时应被 _v_run_inner 拦住
    w.v_text_chk.setChecked(False)
    assert w._v_gather_opts()["kinds"] == []
    assert _V_MOTION_MAP["固定+滚动"] == "both"
    print("GUI 播放方式 x 水印类型 搭配 PASS")


if __name__ == "__main__":
    from pytest import MonkeyPatch
    test_video_fixed_and_scroll()
    test_video_motion_matrix()
    test_docx_detect_and_clear_by_kind()
    test_gui_clear_choice_dialog(MonkeyPatch())
    test_gui_video_run_button_clicked(MonkeyPatch())
    test_gui_motion_and_kind_matrix()
    print("\nALL PASS")
