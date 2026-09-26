"""
水印预览模块：用 Pillow 把水印按与 Word 中一致的比例/角度/透明度/旋转，
渲染到一张 A4 白纸上，使“加上的水印长什么样”脱离 Word 也可直接查看。

支持同时叠加文本水印与图像水印（kinds 含 'text' 与/或 'image'）。

比例约定与 engine_docx.insert_watermark 保持一致：
- 文本水印覆盖约 0.85 倍页宽，图像水印约 0.6 倍页宽，再乘用户 scale。
- 旋转方向与 Word 一致：正角度为顺时针（Pillow 需取负）。
"""
from __future__ import annotations

from io import BytesIO

from PIL import Image

from .engine_docx import render_text_png, prepare_image_png, text_base, tile_layout

# A4 @ 96 DPI（像素），用于示意预览
A4_W_PX = 794
A4_H_PX = 1123


def _composite(page, layer, x, y):
    """把 layer 混合到 page 的 (x, y) 处；越界部分自动裁剪，避免负坐标报错。"""
    x = int(round(x))
    y = int(round(y))
    lx, ly = max(0, x), max(0, y)
    rx, ry = min(page.width, x + layer.width), min(page.height, y + layer.height)
    if rx <= lx or ry <= ly:
        return
    piece = layer.crop((lx - x, ly - y, rx - x, ry - y))
    page.alpha_composite(piece, (lx, ly))


def _draw_layer(page, wm, angle, scale, base, page_w, page_h, offset_x=0.0, offset_y=0.0,
                tile=False, tile_rows=1, tile_cols=1):
    """把单层水印缩放、旋转、按偏移定位后，以透明度混合到 page（RGBA）上。

    offset_x/offset_y 为占页宽/页高的百分比（0=居中，正=右/下，负=左/上）。
    tile=True 时按 rows×cols 网格平铺满页（与 engine_docx 的布局算法一致）。
    """
    if tile:
        # 平铺：每格一份，格内等比缩放；旋转后按格心对齐（与 Word 中观感一致）
        boxes = tile_layout(page_w, page_h, wm.width, wm.height,
                            tile_rows, tile_cols, scale, offset_x, offset_y)
        for (bx, by, bw, bh) in boxes:
            piece = wm.resize((max(1, int(bw)), max(1, int(bh))), Image.LANCZOS)
            # 旋转：Word 正角度=顺时针，Pillow 正角度=逆时针，故取负
            piece = piece.rotate(-angle, expand=True, resample=Image.BICUBIC)
            _composite(page, piece, bx + bw / 2.0 - piece.width / 2.0,
                       by + bh / 2.0 - piece.height / 2.0)
        return

    disp_w = int(page_w * base * scale)
    ratio = wm.height / wm.width if wm.width else 1.0
    disp_h = int(disp_w * ratio)
    wm = wm.resize((max(1, disp_w), max(1, disp_h)), Image.LANCZOS)
    # 旋转：Word 正角度=顺时针，Pillow 正角度=逆时针，故取负
    wm = wm.rotate(-angle, expand=True, resample=Image.BICUBIC)
    cx = (page_w - wm.width) // 2 + offset_x / 100.0 * page_w
    cy = (page_h - wm.height) // 2 + offset_y / 100.0 * page_h
    _composite(page, wm, cx, cy)


def render_preview(opts: dict, kinds, page_w=A4_W_PX, page_h=A4_H_PX) -> Image.Image:
    """返回一张 RGB 白底 A4 页面，中央叠加各启用层的水印（与 Word 中观感一致）。

    kinds: 包含 'text' 和/或 'image' 的可迭代对象。
    文本与图像各自的参数（角度/透明度/缩放）相互独立，分别从
    opts["text"] / opts["image"] 读取，互不影响。
    """
    kinds = [k for k in kinds if k in ("text", "image")]
    text_opts = opts.get("text", {}) or {}
    image_opts = opts.get("image", {}) or {}
    # 防去除加固：平铺满页（两类水印共用同一套网格参数）
    tile = bool(opts.get("tile", False))
    tile_rows = int(opts.get("tile_rows", 4) or 4)
    tile_cols = int(opts.get("tile_cols", 3) or 3)

    page = Image.new("RGBA", (page_w, page_h), (255, 255, 255, 255))

    if "text" in kinds:
        angle = float(text_opts.get("angle", 45))
        transparency = float(text_opts.get("transparency", 0.5))
        scale = float(text_opts.get("scale", 1.0))
        offset_x = float(text_opts.get("offset_x", 0.0))
        offset_y = float(text_opts.get("offset_y", 0.0))
        alpha = int(max(0, min(255, (1 - transparency) * 255)))
        text = text_opts.get("text", "水印")
        font_size = int(text_opts.get("font_size", 120))
        color = text_opts.get("color", (128, 128, 128))
        # 中西文分别指定字体：汉字用 cn_font_name、西文用 latin_font_name
        cn_font_name = text_opts.get("cn_font_name")
        latin_font_name = text_opts.get("latin_font_name")
        wm = render_text_png(text, font_size, color, alpha, cn_font_name, latin_font_name)
        # 与 engine_docx 一致：基准覆盖比例由字号决定，再乘 scale
        base = text_base(font_size)
        _draw_layer(page, wm, angle, scale, base, page_w, page_h, offset_x, offset_y,
                    tile=tile, tile_rows=tile_rows, tile_cols=tile_cols)

    if "image" in kinds:
        angle = float(image_opts.get("angle", 45))
        transparency = float(image_opts.get("transparency", 0.5))
        scale = float(image_opts.get("scale", 1.0))
        offset_x = float(image_opts.get("offset_x", 0.0))
        offset_y = float(image_opts.get("offset_y", 0.0))
        alpha = int(max(0, min(255, (1 - transparency) * 255)))
        image_path = image_opts.get("image_path")
        if not image_path:
            raise ValueError("未选择水印图片")
        wm = prepare_image_png(image_path, alpha)
        _draw_layer(page, wm, angle, scale, 0.6, page_w, page_h, offset_x, offset_y,
                    tile=tile, tile_rows=tile_rows, tile_cols=tile_cols)

    if not kinds:
        raise ValueError("未启用任何水印（文本与图像均未启用）。")

    return page.convert("RGB")


def preview_to_bytes(img: Image.Image, fmt: str = "PNG") -> bytes:
    bio = BytesIO()
    img.save(bio, fmt)
    return bio.getvalue()
