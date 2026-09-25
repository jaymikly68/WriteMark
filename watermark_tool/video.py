"""
视频水印引擎（与 Word 水印完全独立）。

能力：
- 逐帧加水印：对视频的每一帧做合成，**绝不只改封面/关键帧**，因此极难被去除。
- 水印类型：文字水印 与 图片水印（可单独或同时）。
- 播放方式（motion，可自由搭配）：
  - "fixed" 固定在某一角/某一位置。
  - "scroll" 横向滚动播放（水印自右向左循环移动，类似跑马灯）。
  - "both" 固定+滚动：同一份水印同时保留固定副本与滚动副本，双重覆盖。
- 文字与图片各自可单独选择播放方式与锚点位置，因此可以做出
  “文字滚动 + 图片固定”“文字固定 + 图片固定（不同角）”等任意搭配。
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
from PIL import Image, ImageFilter

from . import engine_docx  # 仅复用纯 PIL 的文字/图片渲染函数


# H.264 恒定质量因子：越小越清晰。imageio 默认 quality=5 会换算成 crf≈25，
# 那就是“导出视频发糊”的根源，这里默认提到 18（接近视觉无损），并允许用户再调。
DEFAULT_CRF = 18
CRF_PRESETS = {"省空间": 23, "标准": 18, "高": 14, "极高": 10}


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


def _resize_for_watermark(img: Image.Image, size: tuple[int, int]) -> Image.Image:
    """水印图的缩放策略——图片水印“糊”多半出在这里。

    小图直接放大（一步 LANCZOS）会得到很软的边缘：先超采样再收缩，
    最后补一道极轻微的锐化，能明显改善 logo/字标一类水印的锐度。
    """
    tw, th = size
    w, h = img.size
    if tw <= 0 or th <= 0:
        return img
    if tw == w and th == h:
        return img
    if tw < w:
        # 缩小：LANCZOS 内部自带预滤波，直接一步即可
        return img.resize((tw, th), Image.LANCZOS)
    if tw <= w * 1.5:
        # 小幅放大：普通 LANCZOS 就够，锐化反而会出毛刺
        return img.resize((tw, th), Image.LANCZOS)
    # 大幅放大：3 倍超采样 -> 收回到目标 -> 轻微锐化
    big = img.resize((w * 3, max(1, int(round(h * 3)))), Image.LANCZOS)
    work = big.resize((tw, th), Image.LANCZOS)
    try:
        work = work.filter(ImageFilter.UnsharpMask(radius=1.0, percent=55, threshold=2))
    except Exception:
        pass
    return work


def _build_image_layer(image_path: str, frame_w: int, frame_h: int, cfg: dict) -> Image.Image:
    """把用户图片缩放成水印图层（透明背景）。"""
    alpha = int(max(0, min(255, float(cfg.get("alpha", 180)))))
    base = engine_docx.prepare_image_png(image_path, alpha)  # RGBA，已乘透明度
    img_frac = float(cfg.get("img_frac", 0.15))             # 水印宽占帧宽比例
    target_w = max(16, int(frame_w * img_frac))
    scale = target_w / base.width
    target_h = max(16, int(base.height * scale))
    layer = _resize_for_watermark(base, (target_w, target_h))
    angle = float(cfg.get("angle", 0))
    if angle:
        layer = layer.rotate(angle, expand=True, resample=Image.BICUBIC)
    return layer


def _layer_specs(frame_w: int, frame_h: int, kinds: list, text_cfg: dict, image_cfg: dict):
    """返回本帧需要叠加的所有水印图层规格。

    每项为 dict：{'kind', 'img', 'position'}。
    position 取自该类型自己的配置（text.position / image.position），
    因此文字与图片即使同时固定，也不会叠在同一处。
    """
    specs = []
    if "text" in kinds and (text_cfg.get("text") or "").strip():
        specs.append({
            "kind": "text",
            "img": _build_text_layer(text_cfg["text"], frame_w, frame_h, text_cfg),
            "position": text_cfg.get("position"),
        })
    if "image" in kinds and image_cfg.get("image_path"):
        if os.path.exists(image_cfg["image_path"]):
            specs.append({
                "kind": "image",
                "img": _build_image_layer(image_cfg["image_path"], frame_w, frame_h, image_cfg),
                "position": image_cfg.get("position"),
            })
    return specs


def _normalize_motion(motion):
    """把用户选的播放方式解析成 ['fixed'] / ['scroll'] / ['fixed','scroll']。

    兼容旧字段 mode='fixed'|'scroll' 以及 motion='fixed+scroll'|'both'|'all'。
    传 '跟随'/'' /None 等“未指定”写法时返回 None，由调用方回落到全局播放方式。
    """
    raw = str(motion or "").strip().lower()
    if raw in ("", "none", "跟随", "auto", "default"):
        return None
    if raw in ("both", "all", "fixed+scroll", "fixed_and_scroll", "固定+滚动"):
        return ["fixed", "scroll"]
    if raw in ("scroll", "scrolling", "滚动"):
        return ["scroll"]
    return ["fixed"]


def _effective_motions(opts: dict) -> list:
    """汇总本次实际会出现的播放方式（含类型级覆盖），用于 GUI 判断滚动速度是否可用。"""
    kinds = [k for k in opts.get("kinds", []) if k in ("text", "image")]
    motions = set(_normalize_motion(opts.get("motion", opts.get("mode", "fixed"))) or [])
    for key in kinds:
        per_type = (opts.get(key) or {}).get("motion")
        resolved = _normalize_motion(per_type)
        if resolved:
            motions.update(resolved)
    return sorted(motions)


# ---------------------------------------------------------------------------
# 位置计算
# ---------------------------------------------------------------------------
def _fixed_pos(layer_w: int, layer_h: int, frame_w: int, frame_h: int,
               position=None):
    """固定位置：position=(fx, fy) 为水印左上角相对帧的比例，0..1。"""
    fx, fy = position if (position and len(position) == 2) else (0.85, 0.85)
    x = int(fx * frame_w)
    y = int(fy * frame_h)
    x = max(0, min(frame_w - layer_w, x))
    y = max(0, min(frame_h - layer_h, y))
    return x, y


def _scroll_pos(layer_w: int, layer_h: int, frame_w: int, frame_h: int,
                position=None, t_sec: float = 0.0, scroll_speed: float = 0.12):
    """滚动位置：水印自右向左循环移动（跑马灯），纵向锚定在 position[1]。"""
    span = frame_w + layer_w
    phase = (t_sec * scroll_speed) % 1.0          # 0..1 循环
    x = int(frame_w - phase * span)               # 从右边缘滑到左边缘外
    fy = position[1] if (position and len(position) == 2) else 0.85
    y = int(fy * frame_h)
    y = max(0, min(frame_h - layer_h, y))
    return x, y


# ---------------------------------------------------------------------------
# 主入口
# ---------------------------------------------------------------------------
def _resolve_out_size(opts: dict, src_w: int, src_h: int) -> tuple[int, int]:
    """解析 opts['out_size']；缺省或非法则沿用源分辨率。"""
    raw = opts.get("out_size")
    if not raw:
        return src_w, src_h
    try:
        w = int(raw[0])
        h = int(raw[1])
    except Exception:
        return src_w, src_h
    if w <= 0 or h <= 0:
        return src_w, src_h
    return w, h


def add_video_watermark(src: str, output: str, opts: dict,
                        progress_fn=None, stop_check=None) -> dict:
    """
    给视频逐帧加水印。

    opts 关键字段：
      kinds        : 列表，含 'text' / 'image'（可同时）。
      text         : 文字水印配置 dict（text/color/alpha/size_frac/angle/position/字体）。
      image        : 图片水印配置 dict（image_path/alpha/img_frac/angle/position）。
      motion       : 'fixed' 固定 / 'scroll' 滚动 / 'both' 固定+滚动（同一图层两份）。
                     兼容旧字段 mode。
      scroll_speed : 滚动速度（每秒移动“帧宽”的比例，默认 0.12）。
      out_size     : (w, h) 输出分辨率；缺省/非法则用源分辨率。水印图层按
                     **输出分辨率**构建，因此放大导出时水印依然锐利。
      fps          : 输出帧率，缺省跟随源视频；与滚动速度都以时间为基准，
                     改帧率不会影响滚动节奏感。
      crf          : H.264 恒定质量（10~30，越小越清晰，默认 18）。
    每个类型可用 position 单独指定锚点；缺省 (0.85, 0.85)。固定+滚动时该锚点
    既是固定副本的位置，也是滚动副本的纵向锚点。
    返回 dict 含 output / frames / engine / motion。
    """
    if not os.path.exists(src):
        raise FileNotFoundError(f"视频不存在: {src}")

    kinds = [k for k in opts.get("kinds", ["text", "image"]) if k in ("text", "image")]
    if not kinds:
        raise ValueError("未启用任何水印类型。")
    text_cfg = opts.get("text", {}) or {}
    image_cfg = opts.get("image", {}) or {}
    # 播放方式：优先 motion，兼容旧字段 mode
    motions = _normalize_motion(opts.get("motion", opts.get("mode", "fixed"))) or ["fixed"]
    scroll_speed = float(opts.get("scroll_speed", 0.12) or 0.12)

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
    # 部分容器（mp4/mov 等）imageio 无法预知总帧数，nframes 可能是 inf/0/极大值。
    # 直接把这种值透传给进度回调会让 Qt 信号溢出 -> 整个任务被判失败，故在此收敛为 0。
    total = 0
    try:
        _nf = meta.get("nframes", 0)
        _nf = float(_nf)
        if _nf == _nf and _nf > 0 and _nf < 10_000_000:  # 排除 nan / inf / 异常大值
            total = int(_nf)
    except Exception:
        total = 0

    # ---- 输出规格：分辨率 / 帧率 / 画质 ----
    W, H = _resolve_out_size(opts, W, H)      # 水印要按“导出分辨率”构建才不会糊
    out_fps = float(opts.get("fps", fps) or fps)
    if out_fps <= 0:
        out_fps = fps
    try:
        crf = int(opts.get("crf", DEFAULT_CRF))
    except Exception:
        crf = DEFAULT_CRF
    crf = max(10, min(30, crf))
    need_resize = True     # 后续统一按输出尺寸规整，简单且不会漏帧

    # 水印图层与帧尺寸相关、但与帧内容无关：仅构造一次，循环里只换位置，省大量 CPU
    layers = _layer_specs(W, H, kinds, text_cfg, image_cfg)
    if not layers:
        raise ValueError("文字为空且未提供有效图片，无法生成水印。")
    if scroll_speed <= 0:
        # GUI 在无滚动图层时会传 0，这里兜底回默认速度，避免“开了滚动却不滚动”
        scroll_speed = 0.12
    # 每个图层可自带 motion 覆盖全局（实现“文字滚动 + 图片固定”这类自由搭配）
    for spec in layers:
        cfg = text_cfg if spec["kind"] == "text" else image_cfg
        per_type = _normalize_motion(cfg.get("motion"))
        spec["motions"] = per_type or list(motions)

    # 先写到临时无声视频，最后再 mux 原音轨
    tmp_fd, tmp_vid = tempfile.mkstemp(suffix=".mp4")
    os.close(tmp_fd)
    # quality=None 关掉 imageio 的 crf≈25 默认档（那就是导出发糊的根源），
    # 改用显式 crf 控制清晰度；尺寸由每帧统一 resize 保证一致，无需再传 size。
    writer = iio.get_writer(
        tmp_vid, "ffmpeg", fps=out_fps, macro_block_size=1, quality=None,
        output_params=["-crf", str(crf), "-preset", "veryfast",
                       "-pix_fmt", "yuv420p"],
    )

    frame_idx = 0
    try:
        for frame in reader:
            if stop_check is not None and stop_check():
                break
            # frame 多为 RGB uint8；统一转 RGBA 并按导出分辨率放大/缩小
            img = Image.fromarray(np.asarray(frame)).convert("RGBA")
            if need_resize:
                img = img.resize((W, H), Image.LANCZOS)
            t_sec = frame_idx / out_fps
            painted = set()   # “固定+滚动”下避免同一图层在完全相同坐标被叠加两次
            for spec in layers:
                layer = spec["img"]
                lw, lh = layer.size
                pos = spec.get("position")
                for motion in spec.get("motions") or motions:
                    if motion == "scroll":
                        x, y = _scroll_pos(lw, lh, W, H, pos, t_sec, scroll_speed)
                    else:
                        x, y = _fixed_pos(lw, lh, W, H, pos)
                    if (x, y) in painted:
                        continue
                    painted.add((x, y))
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

    # 返回“实际生效”的播放方式合集：逐类型覆盖会让实际叠加方式多于全局选择
    used = sorted({m for s in layers for m in (s.get("motions") or motions)})
    return {"ok": True, "engine": "video", "frames": frames_done,
            "output": output, "motion": "+".join(used),
            "size": (W, H), "fps": round(out_fps, 3), "crf": crf}


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
