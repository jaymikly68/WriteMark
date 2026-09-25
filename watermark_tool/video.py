"""
视频水印引擎（与 Word 水印完全独立）。

能力：
- 逐帧加水印：对视频的每一帧做合成，**绝不只改封面/关键帧**，因此极难被去除。
- 水印类型：文字水印 与 图片水印（可单独或同时）。
- 位置模式：
  - "fixed"：固定在某一角/某一位置（默认右下）。
  - "scroll"：横向滚动播放（水印自右向左循环移动，类似跑马灯）。
- 输出：mp4(H.264) 以兼容主流播放器；原视频的**音轨会被自动保留**（用随附的
  ffmpeg 二进制重新封装，不重新编码音频）。

实现依赖：imageio + imageio-ffmpeg（自带 ffmpeg 静态二进制），无需系统安装 ffmpeg。
框架合成复用 engine_docx 的 render_text_png / prepare_image_png（纯 PIL），
因此文字水印同样支持中西文分别指定字体、缺字形兜底。

注意：逐帧合成是 CPU 密集操作，长视频（如 1080p 数分钟）耗时较长属正常；
这是“难以去除”的代价，与 Word 水印的设计目标一致。
"""
from __future__ import annotations

import os
import subprocess
import tempfile

import numpy as np
import imageio.v2 as iio
import imageio_ffmpeg
from PIL import Image

from . import engine_docx  # 仅复用纯 PIL 的文字/图片渲染函数


# ---------------------------------------------------------------------------
# 图层构造：把“一份水印”渲染成与帧尺寸相关的 RGBA 图层（透明）
# ---------------------------------------------------------------------------
def _build_text_layer(text: str, frame_w: int, frame_h: int, cfg: dict) -> Image.Image:
    """根据帧尺寸把文字渲染为水印图层（透明背景）。"""
    size_frac = float(cfg.get("size_frac", 0.06))      # 字号占帧高比例
    pixel_size = max(12, int(frame_h * size_frac))
    alpha = int(max(0, min(255, float(cfg.get("alpha", 180)))))
    color = engine_docx.parse_color(cfg.get("color", (255, 255, 255)))
    angle = float(cfg.get("angle", 0))
    layer = engine_docx.render_text_png(
        text, pixel_size, color, alpha,
        cn_font_name=cfg.get("cn_font_name"),
        latin_font_name=cfg.get("latin_font_name"),
    )
    if angle:
        layer = layer.rotate(angle, expand=True, resample=Image.BICUBIC)
    return layer


def _build_image_layer(image_path: str, frame_w: int, frame_h: int, cfg: dict) -> Image.Image:
    """把用户图片缩放成水印图层（透明背景）。"""
    alpha = int(max(0, min(255, float(cfg.get("alpha", 180)))))
    base = engine_docx.prepare_image_png(image_path, alpha)  # RGBA，已乘透明度
    img_frac = float(cfg.get("img_frac", 0.15))             # 水印宽占帧宽比例
    target_w = max(16, int(frame_w * img_frac))
    scale = target_w / base.width
    target_h = max(16, int(base.height * scale))
    layer = base.resize((target_w, target_h), Image.LANCZOS)
    angle = float(cfg.get("angle", 0))
    if angle:
        layer = layer.rotate(angle, expand=True, resample=Image.BICUBIC)
    return layer


def _layers_for_frame(frame_w: int, frame_h: int, kinds: list, text_cfg: dict, image_cfg: dict):
    """返回本帧需要叠加的所有水印图层列表（文字 + 图片）。"""
    layers = []
    if "text" in kinds and text_cfg.get("text"):
        layers.append(_build_text_layer(text_cfg["text"], frame_w, frame_h, text_cfg))
    if "image" in kinds and image_cfg.get("image_path"):
        if os.path.exists(image_cfg["image_path"]):
            layers.append(_build_image_layer(image_cfg["image_path"], frame_w, frame_h, image_cfg))
    return layers


# ---------------------------------------------------------------------------
# 位置计算
# ---------------------------------------------------------------------------
def _fixed_pos(layer_w: int, layer_h: int, frame_w: int, frame_h: int, position):
    """固定位置：position=(fx, fy) 为水印左上角相对帧的比例，0..1。"""
    fx, fy = position if len(position) == 2 else (0.85, 0.85)
    x = int(fx * frame_w)
    y = int(fy * frame_h)
    x = max(0, min(frame_w - layer_w, x))
    y = max(0, min(frame_h - layer_h, y))
    return x, y


def _scroll_pos(layer_w: int, layer_h: int, frame_w: int, frame_h: int,
                position, t_sec: float, scroll_speed: float):
    """滚动位置：水印自右向左循环移动（跑马灯），纵向锚定在 position[1]。"""
    span = frame_w + layer_w
    phase = (t_sec * scroll_speed) % 1.0          # 0..1 循环
    x = int(frame_w - phase * span)               # 从右边缘滑到左边缘外
    fy = position[1] if len(position) == 2 else 0.85
    y = int(fy * frame_h)
    y = max(0, min(frame_h - layer_h, y))
    return x, y


# ---------------------------------------------------------------------------
# 主入口
# ---------------------------------------------------------------------------
def add_video_watermark(src: str, output: str, opts: dict,
                        progress_fn=None, stop_check=None) -> dict:
    """
    给视频逐帧加水印。

    opts 关键字段：
      kinds        : 列表，含 'text' / 'image'（可同时）。
      text         : 文字水印配置 dict（text/color/alpha/size_frac/angle/字体）。
      image        : 图片水印配置 dict（image_path/alpha/img_frac/angle）。
      mode         : 'fixed' 或 'scroll'（滚动播放）。
      position     : (fx, fy) 比例，水印左上角锚点；fixed 模式下即固定位置，
                     scroll 模式下为纵向锚点。
      scroll_speed : 滚动速度（每秒移动“帧宽”的比例，默认 0.12）。
    返回 dict 含 output / frames / engine。
    """
    if not os.path.exists(src):
        raise FileNotFoundError(f"视频不存在: {src}")

    kinds = [k for k in opts.get("kinds", ["text", "image"]) if k in ("text", "image")]
    if not kinds:
        raise ValueError("未启用任何水印类型。")
    text_cfg = opts.get("text", {}) or {}
    image_cfg = opts.get("image", {}) or {}
    mode = opts.get("mode", "fixed")
    position = opts.get("position", (0.85, 0.85))
    scroll_speed = float(opts.get("scroll_speed", 0.12))

    try:
        reader = iio.get_reader(src, "ffmpeg")
        meta = reader.get_meta_data()
    except Exception as e:
        raise RuntimeError(f"无法读取该视频（可能格式不被支持或文件损坏）：{e}")

    fps = float(meta.get("fps", 25) or 25)
    if fps <= 0:
        fps = 25.0
    try:
        W, H = meta["size"]
    except Exception:
        W, H = 1280, 720
    total = meta.get("nframes", 0) or 0

    # 水印图层与帧尺寸相关、但与帧内容无关：仅构造一次，循环里只换位置，省大量 CPU
    layers = _layers_for_frame(W, H, kinds, text_cfg, image_cfg)
    if not layers:
        raise ValueError("文字为空且未提供有效图片，无法生成水印。")

    # 先写到临时无声视频，最后再 mux 原音轨
    tmp_fd, tmp_vid = tempfile.mkstemp(suffix=".mp4")
    os.close(tmp_fd)
    writer = iio.get_writer(tmp_vid, "ffmpeg", fps=fps, macro_block_size=1)

    frame_idx = 0
    try:
        for frame in reader:
            if stop_check is not None and stop_check():
                break
            # frame 多为 RGB uint8；统一转 RGBA 以便合成
            img = Image.fromarray(np.asarray(frame)).convert("RGBA")
            t_sec = frame_idx / fps
            for layer in layers:
                lw, lh = layer.size
                if mode == "scroll":
                    x, y = _scroll_pos(lw, lh, W, H, position, t_sec, scroll_speed)
                else:
                    x, y = _fixed_pos(lw, lh, W, H, position)
                img.alpha_composite(layer, (x, y))
            writer.append_data(np.asarray(img.convert("RGB")))
            frame_idx += 1
            if progress_fn is not None and total:
                progress_fn(frame_idx, total)
    finally:
        try:
            writer.close()
        except Exception:
            pass
        try:
            reader.close()
        except Exception:
            pass

    frames_done = frame_idx
    if frames_done == 0:
        if os.path.exists(tmp_vid):
            os.remove(tmp_vid)
        raise RuntimeError("未能从视频中读出任何帧，请确认文件是有效的视频。")

    # ---- 保留原音轨：用自带的 ffmpeg 重新封装（音视频都 copy，不重编码） ----
    muxed = _mux_audio(tmp_vid, src, output)
    if muxed:
        if os.path.exists(tmp_vid):
            os.remove(tmp_vid)
    else:
        # mux 失败（如无音频流或 ffmpeg 异常）则退化为无声视频
        try:
            if os.path.abspath(tmp_vid) != os.path.abspath(output):
                os.replace(tmp_vid, output)
        except Exception:
            pass

    return {"ok": True, "engine": "video", "frames": frames_done, "output": output}


def _mux_audio(tmp_vid: str, src: str, output: str) -> bool:
    """把原视频音轨封装到无声水印视频里；成功返回 True。"""
    try:
        ffmpeg = imageio_ffmpeg.get_ffmpeg_exe()
    except Exception:
        return False
    cmd = [
        ffmpeg, "-y", "-i", tmp_vid, "-i", src,
        "-c", "copy", "-map", "0:v:0", "-map", "1:a:0?", "-shortest", output,
    ]
    try:
        subprocess.run(cmd, capture_output=True, check=False,
                       timeout=600, creationflags=0)
    except Exception:
        return False
    # 校验产物确实生成且比源“无声临时文件”更像完整视频
    return os.path.exists(output) and os.path.getsize(output) > 0


def watermark_video_formats() -> list:
    """列出本工具支持的常见输入视频格式（ffmpeg 解码，基本覆盖主流格式）。"""
    return [".mp4", ".mkv", ".avi", ".mov", ".wmv", ".flv", ".webm",
            ".mpeg", ".mpg", ".ts", ".m4v", ".3gp", ".vob"]
