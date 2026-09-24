"""
Word COM 水印引擎（需要本机安装 Microsoft Word）。

用途：
- 处理 .doc（Word 97-2003）格式——纯 Python 无法处理，必须用 Word 自动化。
- 当本机有 Word 时，也用它来处理 .docx，水印是 Word 原生的“形状水印”，更可靠。

实现：在每一节页眉里插入 Shape（文本用 AddTextEffect，图片用 AddPicture），
设为衬于文字下方、半透明、居中、可旋转。文本与图像可同时添加：文本打 MARK_TEXT、
图像打 MARK_IMG，二者互不覆盖；清除时删除所有本工具标记的形状，不动正文。
"""
from __future__ import annotations

import os

# Word 常量
WD_HEADER_FOOTER_PRIMARY = 1
WD_HEADER_FOOTER_EVEN_PAGES = 2
WD_HEADER_FOOTER_FIRST_PAGE = 3
MSO_SEND_BEHIND_TEXT = 4          # ZOrder 参数：衬于文字下方
WD_WRAP_NONE = 3                  # 绕排方式：无（配合衬于下方）
MSO_TEXT_EFFECT_1 = 0             # AddTextEffect 的预设样式索引（0~39 均可）

MARK_TEXT = "WB_WATERMARK_TEXT"
MARK_IMG = "WB_WATERMARK_IMG"
MARK_NAMES = (MARK_TEXT, MARK_IMG)


def _rgb(rgb_tuple):
    """把 (r,g,b) 转成 Word 需要的整型颜色值（BGR 顺序）。"""
    r, g, b = rgb_tuple
    return (b << 16) | (g << 8) | r


def _word_app():
    import win32com.client
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False
    word.DisplayAlerts = False
    return word


def _header_indices(doc, sec):
    """返回该节需要处理的页眉类型列表（主/偶数页/首页）。"""
    idxs = [WD_HEADER_FOOTER_PRIMARY]
    try:
        if doc.PageSetup.OddAndEvenPagesHeaderFooter:
            idxs.append(WD_HEADER_FOOTER_EVEN_PAGES)
    except Exception:
        pass
    try:
        if sec.PageSetup.DifferentFirstPageHeaderFooter:
            idxs.append(WD_HEADER_FOOTER_FIRST_PAGE)
    except Exception:
        pass
    return idxs


def _add_text_watermark(header, text, cn_font_name, latin_font_name, font_size,
                        color, angle, transparency):
    # 西文字体作为 WordArt 主字体；若本机有对应字体则用西文，否则回退到中文字体
    main_font = latin_font_name or cn_font_name or "微软雅黑"
    shape = header.Shapes.AddTextEffect(
        MSO_TEXT_EFFECT_1, text, main_font, font_size, 0, 0, 0, 0
    )
    shape.Name = MARK_TEXT
    shape.Rotation = angle
    shape.Fill.ForeColor.RGB = _rgb(color)
    shape.Fill.Transparency = transparency
    shape.Line.Visible = False
    shape.WrapFormat.Type = WD_WRAP_NONE
    shape.ZOrder(MSO_SEND_BEHIND_TEXT)
    # 尽量让中文用“中文字体”、西文用“西文字体”（Word 文本范围的 FarEast 字体）
    if cn_font_name:
        try:
            tf = shape.TextFrame.TextRange.Font
            tf.NameFarEast = cn_font_name
            if latin_font_name:
                tf.Name = latin_font_name
        except Exception:
            pass
    return shape


def _add_image_watermark(header, image_path, angle, transparency):
    shape = header.Shapes.AddPicture(
        image_path, False, True, 0, 0, -1, -1
    )
    shape.Name = MARK_IMG
    shape.Rotation = angle
    try:
        shape.PictureFormat.Transparency = transparency
    except Exception:
        pass
    shape.Line.Visible = False
    shape.WrapFormat.Type = WD_WRAP_NONE
    shape.ZOrder(MSO_SEND_BEHIND_TEXT)
    return shape


def _center(shape, page_w, page_h, offset_x=0.0, offset_y=0.0):
    """居中放置形状，并按百分比偏移（offset 正=右/下，负=左/上）。

    page_w/page_h 与 shape.Left/Top 单位均为 Word 的“磅(point)”。
    """
    try:
        cx = (page_w - shape.Width) / 2 + offset_x / 100.0 * page_w
        cy = (page_h - shape.Height) / 2 + offset_y / 100.0 * page_h
        shape.Left = cx
        shape.Top = cy
    except Exception:
        pass


def _remove_in_header(header, clear_any=False):
    """移除页眉中的形状。

    clear_any=False：仅删本工具标记（MARK_NAMES）的形状（插入前“替换”用）。
    clear_any=True：额外删掉名称含 "watermark"（不区分大小写）的形状，
    以清掉 Word 原生水印或其它工具留下的水印。
    """
    removed = 0
    for s in list(header.Shapes):
        try:
            nm = s.Name
            is_ours = nm in MARK_NAMES
            is_native = (not is_ours) and clear_any and ("watermark" in nm.lower())
            if is_ours or is_native:
                s.Delete()
                removed += 1
        except Exception:
            pass
    return removed


def _has_in_header(header):
    for s in header.Shapes:
        try:
            if s.Name in MARK_NAMES:
                return True
        except Exception:
            pass
    return False


def _iter_doc_headers(doc):
    """生成 (header, page_w, page_h) 元组，覆盖所有节与页眉类型。"""
    for i in range(1, doc.Sections.Count + 1):
        sec = doc.Sections(i)
        page_w = float(sec.PageSetup.PageWidth)
        page_h = float(sec.PageSetup.PageHeight)
        for hf in _header_indices(doc, sec):
            try:
                header = sec.Headers(hf)
            except Exception:
                continue
            yield header, page_w, page_h


# ---------------------------------------------------------------------------
# 对外 API
# ---------------------------------------------------------------------------
def insert_watermark(path: str, kinds, output_path: str = None, **opts) -> dict:
    import win32com.client
    path = os.path.abspath(path)
    kinds = [k for k in kinds if k in ("text", "image")]
    text_opts = opts.get("text", {}) or {}
    image_opts = opts.get("image", {}) or {}

    text = text_opts.get("text", "水印")
    cn_font_name = text_opts.get("cn_font_name", "微软雅黑")
    latin_font_name = text_opts.get("latin_font_name", "微软雅黑")
    font_size = int(text_opts.get("font_size", 72))
    color = text_opts.get("color", (128, 128, 128))
    text_angle = float(text_opts.get("angle", 45))
    text_transparency = float(text_opts.get("transparency", 0.5))
    text_offset_x = float(text_opts.get("offset_x", 0.0))
    text_offset_y = float(text_opts.get("offset_y", 0.0))

    image_path = image_opts.get("image_path")
    img_angle = float(image_opts.get("angle", 45))
    img_transparency = float(image_opts.get("transparency", 0.5))
    img_offset_x = float(image_opts.get("offset_x", 0.0))
    img_offset_y = float(image_opts.get("offset_y", 0.0))

    if not kinds:
        raise ValueError("未启用任何水印（文本与图像均未启用），请至少启用一种。")

    word = _word_app()
    doc = None
    try:
        doc = word.Documents.Open(path, False, False)
        inserted = 0
        for header, page_w, page_h in _iter_doc_headers(doc):
            _remove_in_header(header)  # 先清旧，防叠加
            if "text" in kinds:
                shape = _add_text_watermark(header, text, cn_font_name, latin_font_name,
                                            font_size, color, text_angle, text_transparency)
                _center(shape, page_w, page_h, text_offset_x, text_offset_y)
                inserted += 1
            if "image" in kinds:
                if not image_path or not os.path.exists(image_path):
                    raise FileNotFoundError(f"图片不存在: {image_path}")
                shape = _add_image_watermark(header, image_path, img_angle, img_transparency)
                _center(shape, page_w, page_h, img_offset_x, img_offset_y)
                inserted += 1
        if output_path and os.path.abspath(output_path) != path:
            doc.SaveAs(os.path.abspath(output_path))  # 另存为，保留原文件
        else:
            doc.Save()
        return {"ok": True, "engine": "com", "inserted": inserted, "kinds": kinds}
    finally:
        if doc is not None:
            try:
                doc.Close(False)
            except Exception:
                pass
        try:
            word.Quit()
        except Exception:
            pass


def clear_watermark(path: str, output_path: str = None) -> dict:
    import win32com.client
    path = os.path.abspath(path)
    word = _word_app()
    doc = None
    try:
        doc = word.Documents.Open(path, False, False)
        removed = 0
        for header, _, _ in _iter_doc_headers(doc):
            removed += _remove_in_header(header, clear_any=True)
        if output_path and os.path.abspath(output_path) != path:
            doc.SaveAs(os.path.abspath(output_path))  # 另存为，保留原文件
        else:
            doc.Save()
        return {"ok": True, "engine": "com", "removed": removed}
    finally:
        if doc is not None:
            try:
                doc.Close(False)
            except Exception:
                pass
        try:
            word.Quit()
        except Exception:
            pass


def has_watermark(path: str) -> bool:
    import win32com.client
    path = os.path.abspath(path)
    word = _word_app()
    doc = None
    try:
        doc = word.Documents.Open(path, False, True)  # 只读
        for header, _, _ in _iter_doc_headers(doc):
            if _has_in_header(header):
                return True
        return False
    finally:
        if doc is not None:
            try:
                doc.Close(False)
            except Exception:
                pass
        try:
            word.Quit()
        except Exception:
            pass
