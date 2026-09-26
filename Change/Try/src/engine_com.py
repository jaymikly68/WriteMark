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

import logging
import os

log = logging.getLogger(__name__)

# Word 常量
from .engine_docx import decoy_name, tile_layout

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
    """启动一个不可见的 Word。返回 (word, 调用前的 WINWORD 进程快照)。

    快照用于收尾时区分“我们自己起的 Word”和用户自己打开的那些，
    避免为了清理残留而误杀用户的 Word。
    """
    import win32com.client
    from . import com_cleanup
    before = com_cleanup.winword_pids()
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False
    word.DisplayAlerts = False
    return word, before


def _close_word(word, before):
    """退出 Word 并确保进程真的消失（否则会留下看不见的 Word 拖慢系统）。

    处理较大文档时 Word 收尾会慢一些，所以给足 5 秒自行退出，仍不退才强制结束。
    """
    from . import com_cleanup
    return com_cleanup.quit_word(word, before_pids=before, grace=5.0)


def _header_indices(doc, sec):
    """返回该节需要处理的页眉类型列表（主/偶数页/首页）。"""
    idxs = [WD_HEADER_FOOTER_PRIMARY]
    try:
        if doc.PageSetup.OddAndEvenPagesHeaderFooter:
            idxs.append(WD_HEADER_FOOTER_EVEN_PAGES)
    except Exception as e:
        log.debug("读取奇偶页页眉设置失败（跳过）: %s", e)
    try:
        if sec.PageSetup.DifferentFirstPageHeaderFooter:
            idxs.append(WD_HEADER_FOOTER_FIRST_PAGE)
    except Exception as e:
        log.debug("读取首页不同页眉设置失败（跳过）: %s", e)
    return idxs


def _tag(shape, mark, disp_name=None):
    """给形状打上本工具的私有标记。

    加固时显示的 Name 用无害的伪装名（Word/WPS 的“删除水印”与多数去水印脚本
    都按名字/结构匹配水印），私有标记改存 AlternativeText（Word 会保留），
    因此伪装不影响本工具的识别与一键清除。
    """
    try:
        shape.Name = disp_name or mark
    except Exception as e:
        log.warning("设置水印形状 Name 失败（将影响识别与清除）: %s", e)
    try:
        shape.AlternativeText = mark
    except Exception as e:
        log.warning("设置水印形状 AlternativeText 标记失败（将影响识别与清除）: %s", e)


def _is_ours(shape):
    """判断形状是否本工具所加（兼容旧版把标记写在 Name 上的文件）。"""
    try:
        nm = shape.Name
        if nm in MARK_NAMES:
            return True
    except Exception as e:
        log.debug("读取形状 Name 失败: %s", e)
    try:
        return shape.AlternativeText in MARK_NAMES
    except Exception as e:
        log.debug("读取形状 AlternativeText 失败: %s", e)
        return False


def _add_text_watermark(header, text, cn_font_name, latin_font_name, font_size,
                        color, angle, transparency, disp_name=None):
    # 西文字体作为 WordArt 主字体；若本机有对应字体则用西文，否则回退到中文字体
    main_font = latin_font_name or cn_font_name or "微软雅黑"
    shape = header.Shapes.AddTextEffect(
        MSO_TEXT_EFFECT_1, text, main_font, font_size, 0, 0, 0, 0
    )
    _tag(shape, MARK_TEXT, disp_name)
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
        except Exception as e:
            log.debug("设置中/西文字体失败（使用 Word 默认字体）: %s", e)
    return shape


def _add_image_watermark(header, image_path, angle, transparency, disp_name=None):
    shape = header.Shapes.AddPicture(
        image_path, False, True, 0, 0, -1, -1
    )
    _tag(shape, MARK_IMG, disp_name)
    shape.Rotation = angle
    try:
        shape.PictureFormat.Transparency = transparency
    except Exception as e:
        log.warning("设置图片水印透明度失败（图片可能不透明）: %s", e)
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
    except Exception as e:
        log.warning("水印居中定位失败（位置可能偏移）: %s", e)


def _remove_in_header(header, clear_any=False):
    """移除页眉/页脚中的形状。

    clear_any=False：仅删本工具标记（MARK_NAMES）的形状（插入前“替换”用）。
    clear_any=True：额外删掉名称含 "watermark"（不区分大小写）的形状，
    以清掉 Word 原生水印或其它工具留下的水印。
    """
    removed = 0
    for s in list(header.Shapes):
        try:
            is_ours = _is_ours(s)
            nm = ""
            try:
                nm = (s.Name or "").lower()
            except Exception as e:
                log.debug("读取形状 Name 失败: %s", e)
            is_native = (not is_ours) and clear_any and ("watermark" in nm)
            if is_ours or is_native:
                s.Delete()
                removed += 1
        except Exception as e:
            log.warning("删除水印形状失败（可能残留）: %s", e)
    return removed


def _has_in_header(header):
    for s in header.Shapes:
        if _is_ours(s):
            return True
    return False


def _count_in_doc(doc) -> int:
    """统计文档（含页脚）中本工具水印形状的份数。"""
    n = 0
    for part, _, _ in _iter_doc_parts(doc, redundant=True):
        for s in part.Shapes:
            if _is_ours(s):
                n += 1
    return n


def _iter_doc_parts(doc, redundant=False):
    """生成 (part, page_w, page_h) 元组，覆盖所有节的页眉（+可选页脚）类型。"""
    for i in range(1, doc.Sections.Count + 1):
        sec = doc.Sections(i)
        page_w = float(sec.PageSetup.PageWidth)
        page_h = float(sec.PageSetup.PageHeight)
        for hf in _header_indices(doc, sec):
            try:
                yield sec.Headers(hf), page_w, page_h
            except Exception as e:
                log.debug("读取页眉失败（跳过该页眉）: %s", e)
                continue
            if redundant:
                try:
                    yield sec.Footers(hf), page_w, page_h
                except Exception as e:
                    log.debug("读取页脚失败（跳过该页脚）: %s", e)
                    continue


def _iter_doc_headers(doc):
    """兼容旧调用名：仅页眉。"""
    return _iter_doc_parts(doc, redundant=False)


def _place(shape, x, y, w=None, h=None):
    """把形状按“左上角坐标(磅) + 可选尺寸(磅)”绝对定位。"""
    try:
        if w and h:
            shape.Width = w
            shape.Height = h
        shape.Left = x
        shape.Top = y
    except Exception as e:
        log.warning("水印绝对定位失败（位置可能偏移）: %s", e)


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
    text_scale = float(text_opts.get("scale", 1.0))

    image_path = image_opts.get("image_path")
    img_angle = float(image_opts.get("angle", 45))
    img_transparency = float(image_opts.get("transparency", 0.5))
    img_offset_x = float(image_opts.get("offset_x", 0.0))
    img_offset_y = float(image_opts.get("offset_y", 0.0))
    img_scale = float(image_opts.get("scale", 1.0))

    # 防去除加固（与纯 Python 引擎同一套语义）
    tile = bool(opts.get("tile", False))
    tile_rows = int(opts.get("tile_rows", 5) or 5)
    tile_cols = int(opts.get("tile_cols", 2) or 2)
    redundant = bool(opts.get("redundant", False))

    if not kinds:
        raise ValueError("未启用任何水印（文本与图像均未启用），请至少启用一种。")

    idx = 0        # 伪装显示名序号
    inserted = 0
    word, _before = _word_app()
    doc = None
    try:
        doc = word.Documents.Open(path, False, False)
        for header, page_w, page_h in _iter_doc_parts(doc, redundant=redundant):
            _remove_in_header(header)  # 先清旧，防叠加
            if "text" in kinds:
                if tile:
                    shapes = []
                    for _ in range(tile_rows * tile_cols):
                        shapes.append(_add_text_watermark(
                            header, text, cn_font_name, latin_font_name, font_size,
                            color, text_angle, text_transparency, decoy_name(idx)))
                        idx += 1
                    ratio = (shapes[0].Height / shapes[0].Width) if shapes[0].Width else 0.2
                    boxes = tile_layout(page_w, page_h, 100.0, 100.0 * ratio,
                                        tile_rows, tile_cols, text_scale,
                                        text_offset_x, text_offset_y)
                    for s, (x, y, w, h) in zip(shapes, boxes):
                        _place(s, x, y, w, h)
                    inserted += len(shapes)
                else:
                    shape = _add_text_watermark(header, text, cn_font_name, latin_font_name,
                                                font_size, color, text_angle, text_transparency,
                                                decoy_name(idx) if redundant else None)
                    idx += 1
                    _center(shape, page_w, page_h, text_offset_x, text_offset_y)
                    inserted += 1
            if "image" in kinds:
                if not image_path or not os.path.exists(image_path):
                    raise FileNotFoundError(f"图片不存在: {image_path}")
                if tile:
                    shapes = []
                    for _ in range(tile_rows * tile_cols):
                        shapes.append(_add_image_watermark(
                            header, image_path, img_angle, img_transparency, decoy_name(idx)))
                        idx += 1
                    ratio = (shapes[0].Height / shapes[0].Width) if shapes[0].Width else 1.0
                    boxes = tile_layout(page_w, page_h, 100.0, 100.0 * ratio,
                                        tile_rows, tile_cols, img_scale,
                                        img_offset_x, img_offset_y)
                    for s, (x, y, w, h) in zip(shapes, boxes):
                        _place(s, x, y, w, h)
                    inserted += len(shapes)
                else:
                    shape = _add_image_watermark(header, image_path, img_angle,
                                                 img_transparency,
                                                 decoy_name(idx) if redundant else None)
                    idx += 1
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
        _close_word(word, _before)


def clear_watermark(path: str, output_path: str = None) -> dict:
    import win32com.client
    path = os.path.abspath(path)
    word, _before = _word_app()
    doc = None
    try:
        doc = word.Documents.Open(path, False, False)
        removed = 0
        # 始终扫描页脚：加固模式下水印也会写进页脚，否则会清不干净
        for header, _, _ in _iter_doc_parts(doc, redundant=True):
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
        _close_word(word, _before)


def count_watermarks(path: str) -> int:
    """返回本工具水印形状的份数（守护按“份数”判断是否被删过）。"""
    path = os.path.abspath(path)
    word, _before = _word_app()
    doc = None
    try:
        doc = word.Documents.Open(path, False, True)  # 只读
        return _count_in_doc(doc)
    finally:
        if doc is not None:
            try:
                doc.Close(False)
            except Exception:
                pass
        _close_word(word, _before)


def has_watermark(path: str) -> bool:
    import win32com.client
    path = os.path.abspath(path)
    word, _before = _word_app()
    doc = None
    try:
        doc = word.Documents.Open(path, False, True)  # 只读
        for header, _, _ in _iter_doc_parts(doc, redundant=True):
            if _has_in_header(header):
                return True
        return False
    finally:
        if doc is not None:
            try:
                doc.Close(False)
            except Exception:
                pass
        _close_word(word, _before)
