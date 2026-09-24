"""
docx 纯 Python 水印引擎（无需安装 Word 即可处理 .docx）。

水印实现思路：
- Word 的水印本质是“垫在页眉里、衬于文字下方(behindDoc)的半透明图形”。
- 文本水印：先用 Pillow 把文字渲染成透明 PNG（含颜色/透明度），再作为一个
  锚定(anchor)、旋转、衬于文字下方的图形插入每一节的页眉。
- 图像水印：把用户图片转成带透明度的 PNG，同样方式插入。
- 文本与图像水印可同时添加：文本用标记 MARK_TEXT、图像用标记 MARK_IMG，
  二者互不影响、可独立叠加；清除时只删带本工具标记的元素，绝不动正文/用户图片。

标记名称可改，但插入与清除必须一致，否则清除找不到。
"""
from __future__ import annotations

import os
from io import BytesIO

from PIL import Image, ImageDraw, ImageFont
from docx import Document
from docx.oxml.ns import qn
from lxml import etree

from . import word_fonts

# ---------------------------------------------------------------------------
# 命名空间与常量
# ---------------------------------------------------------------------------
W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
WP = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
A = "http://schemas.openxmlformats.org/drawingml/2006/main"
PIC = "http://schemas.openxmlformats.org/drawingml/2006/picture"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

EMU_PER_INCH = 914400
EMU_PER_PT = 12700

# 文本与图像分别打不同标记，便于“同时添加”时互不覆盖、也便于独立清除
MARK_TEXT = "WB_WATERMARK_TEXT"
MARK_IMG = "WB_WATERMARK_IMG"
MARK_NAMES = (MARK_TEXT, MARK_IMG)

# 文本水印尺寸模型：font_size 作为“字号/实际大小”的主控制项（与 Word COM 引擎一致），
# scale 作为微调倍数。默认 font_size=120、scale=1.0 时文本覆盖约 0.85 倍页宽，
# 与历史外观保持一致；字号越小水印越小（此前 .docx 路径字号只影响清晰度、不改大小，已修正）。
TEXT_COVER = 0.85
TEXT_REF_SIZE = 120.0


def text_base(font_size: float) -> float:
    """把字号换算成“占页宽比例”的基准值（再乘用户 scale 得到最终覆盖比例）。"""
    return TEXT_COVER * (float(font_size) / TEXT_REF_SIZE)

# 让 lxml 序列化时使用标准前缀（a:/pic:），而非 ns2:/ns3:，
# Word 能容忍任意前缀，但 WPS 等实现对此更敏感，标准前缀兼容性最好。
etree.register_namespace("a", A)
etree.register_namespace("pic", PIC)
etree.register_namespace("wp", WP)


# ---------------------------------------------------------------------------
# 字体与图片准备
# ---------------------------------------------------------------------------
def find_font() -> str | None:
    """在 Windows 常见位置找一个能渲染中文的字体，找不到返回 None（用默认点阵）。"""
    candidates = [
        "C:/Windows/Fonts/msyh.ttc",     # 微软雅黑
        "C:/Windows/Fonts/msyhbd.ttc",
        "C:/Windows/Fonts/simhei.ttf",   # 黑体
        "C:/Windows/Fonts/simsun.ttc",   # 宋体
        "C:/Windows/Fonts/arial.ttf",
    ]
    for c in candidates:
        if os.path.exists(c):
            return c
    return None


def is_cjk(ch: str) -> bool:
    """判断单个字符是否属于中日韩（CJK）文字脚本。

    命中即视为“中文字符”，应由“中文字体”渲染；否则（拉丁字母、数字、
    半角标点等）由“西文字体”渲染——这正是 Word 字体对话框里
    “中文字体 / 西文字体”分别设置的语义。
    """
    cp = ord(ch)
    return (0x3000 <= cp <= 0x303F or       # CJK 符号与标点
            0x3040 <= cp <= 0x30FF or       # 日文平假名/片假名
            0x3400 <= cp <= 0x4DBF or       # 扩展A
            0x4E00 <= cp <= 0x9FFF or       # 基本汉字
            0xF900 <= cp <= 0xFAFF or       # 兼容汉字
            0xAC00 <= cp <= 0xD7AF or       # 韩文音节
            0xFF00 <= cp <= 0xFFEF)         # 半/全角形式区


def render_text_png(text: str, font_size: int, color, alpha: int,
                    cn_font_name: str | None = None,
                    latin_font_name: str | None = None) -> Image.Image:
    """把文本渲染成透明背景 PNG。color=(r,g,b)，alpha=0~255（越小越透明）。

    支持 Word 式“中文 / 西文分别指定字体”：
    - CJK 字符（汉字、假名、韩文…）用 cn_font_name 渲染；
    - 其余字符（拉丁字母、数字、半角标点、空格…）用 latin_font_name 渲染。
    两者默认都回退到 find_font() 的中文字体，保证缺省外观与历史一致。
    若所选字体确无某字符字形（例如把中文字体设为纯西文 Arial），再二次回退
    到 find_font() 兜底，绝不把字符画成方框/豆腐块。
    """
    fallback_path = find_font()
    cn_path = word_fonts.resolve_font_path(cn_font_name) if cn_font_name else None
    if cn_path is None:
        cn_path = fallback_path
    latin_path = word_fonts.resolve_font_path(latin_font_name) if latin_font_name else None
    if latin_path is None:
        latin_path = fallback_path

    def _load(p):
        try:
            return ImageFont.truetype(p, font_size) if p else ImageFont.load_default()
        except Exception:
            return ImageFont.load_default()

    cn_font = _load(cn_path)
    latin_font = _load(latin_path)
    fallback_font = _load(fallback_path) if fallback_path else cn_font

    r, g, b = color[0], color[1], color[2]

    # ---- 字形存在性判定（栅格化后与“缺字形”参照比较），用于二次兜底 ----
    def _glyph_bytes(font, ch):
        gimg = Image.new("L", (font.size * 2, font.size * 2), 0)
        ImageDraw.Draw(gimg).text((0, 0), ch, font=font, fill=255)
        return gimg.tobytes()

    notdef_cn = _glyph_bytes(cn_font, "\uffff")     # 中文字体“缺字形”参照
    notdef_la = _glyph_bytes(latin_font, "\uffff")  # 西文字体“缺字形”参照

    def _has_glyph(font, ch, ref):
        if ch == "\uffff":
            return False
        m = _glyph_bytes(font, ch)
        if ref and m == ref:
            return False          # 与“缺字形”位图一致 → 视为无字形
        if not any(m):
            return False          # 全零（空）位图 → 视为无字形
        return True

    def _font_for(ch):
        if is_cjk(ch):
            f, ref = cn_font, notdef_cn
        else:
            f, ref = latin_font, notdef_la
        # 空白（空格/制表）本身没有可见字形，但仍应沿用该脚本字体的字宽，
        # 否则会被误判为“缺字形”而切到兜底字体，导致中西文混排时间距变形。
        if ch.isspace():
            return f
        # 二次兜底：所选字体无此字形时改用兜底字体，彻底杜绝方框
        if not _has_glyph(f, ch, ref):
            return fallback_font
        return f

    # 把连续“同一字体”的字符合并成 run，减少图层数量、保持排版自然
    runs = []
    for ch in text:
        f = _font_for(ch)
        if runs and runs[-1][1] is f:
            runs[-1] = (runs[-1][0] + ch, f)
        else:
            runs.append((ch, f))

    # ---- 混排渲染：所有 run 画在同一透明层上、共用一条基线 ----
    # 关键：中文字体与西文字体的 ascent/descent 不同，若按各自 bbox 垂直居中拼接，
    # “机密C”里的 C 会明显偏上/偏下。这里取各字体的最大 ascent 作为统一基线，
    # 与 Word “中文字体/西文字体分别设置”的实际排版效果一致。
    try:
        asc = max(f.getmetrics()[0] for _, f in runs)
        desc = max(f.getmetrics()[1] for _, f in runs)
    except Exception:
        asc, desc = int(font_size * 1.2), int(font_size * 0.4)

    advs = []
    for chunk, f in runs:
        try:
            advs.append(float(f.getlength(chunk)))
        except Exception:  # 极老 Pillow / 位图字体没有 getlength
            advs.append(float(font_size) * len(chunk) * 0.5)

    content_w = max(1, int(round(sum(advs))))
    content_h = max(1, asc + desc)
    layer = Image.new("RGBA", (content_w, content_h), (0, 0, 0, 0))
    d = ImageDraw.Draw(layer)
    baseline = asc
    x = 0.0
    for (chunk, f), adv in zip(runs, advs):
        try:
            # anchor="ls"：以基线为 y 基准排版，保证跨字体基线对齐
            d.text((x, baseline), chunk, font=f, fill=(r, g, b, alpha), anchor="ls")
        except Exception:
            # 兜底：不支持 anchor 的字体，用其自身 ascent 近似对齐到同一基线
            try:
                a_f = f.getmetrics()[0]
            except Exception:
                a_f = asc
            d.text((x, baseline - a_f), chunk, font=f, fill=(r, g, b, alpha))
        x += adv

    pad = max(20, int(font_size * 0.25))
    img = Image.new("RGBA", (content_w + 2 * pad, content_h + 2 * pad), (0, 0, 0, 0))
    img.alpha_composite(layer, (pad, pad))
    return img


def prepare_image_png(path: str, alpha: int) -> Image.Image:
    """打开用户图片并整体乘以目标透明度，返回 RGBA 图像。"""
    img = Image.open(path).convert("RGBA")
    a_channel = img.split()[3].point(lambda p: int(p * alpha / 255))
    img.putalpha(a_channel)
    return img


# ---------------------------------------------------------------------------
# DrawingML 构造
# ---------------------------------------------------------------------------
def _make_drawing(rId: str, cx: int, cy: int, rot_deg: float, name: str, docpr_id: int,
                 pos_x: int, pos_y: int) -> etree._Element:
    """构造一个衬于文字下方、可旋转、按绝对坐标定位的锚定图形 XML。

    pos_x / pos_y 为图形左上角相对页面原点（页面左上角）的 EMU 坐标，
    由调用方根据“居中基准 + 百分比偏移”算出，从而支持上下左右独立调整。
    """
    drawing = etree.Element(qn("w:drawing"))
    anchor = etree.SubElement(
        drawing, qn("wp:anchor"),
        {
            "distT": "0", "distB": "0", "distL": "0", "distR": "0",
            "simplePos": "0", "relativeHeight": "0", "behindDoc": "1",
            "locked": "0", "layoutInCell": "1", "allowOverlap": "1",
        },
    )
    etree.SubElement(anchor, qn("wp:simplePos"), {"x": "0", "y": "0"})

    posH = etree.SubElement(anchor, qn("wp:positionH"), {"relativeFrom": "page"})
    etree.SubElement(posH, qn("wp:posOffset")).text = str(int(round(pos_x)))
    posV = etree.SubElement(anchor, qn("wp:positionV"), {"relativeFrom": "page"})
    etree.SubElement(posV, qn("wp:posOffset")).text = str(int(round(pos_y)))

    etree.SubElement(anchor, qn("wp:extent"), {"cx": str(cx), "cy": str(cy)})
    etree.SubElement(anchor, qn("wp:effectExtent"), {"l": "0", "t": "0", "r": "0", "b": "0"})
    etree.SubElement(anchor, qn("wp:wrapNone"))

    docPr = etree.SubElement(anchor, qn("wp:docPr"), {"id": str(docpr_id), "name": name, "descr": name})
    etree.SubElement(anchor, qn("wp:cNvGraphicFramePr"))

    graphic = etree.SubElement(anchor, qn("a:graphic"))
    graphicData = etree.SubElement(graphic, qn("a:graphicData"), {"uri": PIC})
    pic = etree.SubElement(graphicData, qn("pic:pic"))

    nvPicPr = etree.SubElement(pic, qn("pic:nvPicPr"))
    etree.SubElement(nvPicPr, qn("pic:cNvPr"), {"id": "0", "name": name + "_img", "descr": name})
    etree.SubElement(nvPicPr, qn("pic:cNvPicPr"))

    blipFill = etree.SubElement(pic, qn("pic:blipFill"))
    etree.SubElement(blipFill, qn("a:blip"), {qn("r:embed"): rId})
    stretch = etree.SubElement(blipFill, qn("a:stretch"))
    etree.SubElement(stretch, qn("a:fillRect"))

    spPr = etree.SubElement(pic, qn("pic:spPr"))
    xfrm = etree.SubElement(spPr, qn("a:xfrm"))
    xfrm.set("rot", str(int(round(rot_deg * 60000))))  # 旋转单位为 1/60000 度
    etree.SubElement(xfrm, qn("a:off"), {"x": "0", "y": "0"})
    etree.SubElement(xfrm, qn("a:ext"), {"cx": str(cx), "cy": str(cy)})
    prstGeom = etree.SubElement(spPr, qn("a:prstGeom"), {"prst": "rect"})
    etree.SubElement(prstGeom, qn("a:avLst"))
    return drawing


# ---------------------------------------------------------------------------
# 页眉遍历 / 插入 / 清除
# ---------------------------------------------------------------------------
def _iter_headers(document, section):
    """返回该节所有“非链接”的页眉（主/首页/偶数页），避免重复插入。"""
    headers = [section.header]
    try:
        if section.different_first_page_header_footer:
            headers.append(section.first_page_header)
    except Exception:
        pass
    try:
        if document.part.settings.odd_and_even_pages_header_footer:
            headers.append(section.even_page_header)
    except Exception:
        pass
    # 去重：链接到前一节的页眉会共享同一个 part，只处理一次
    seen = set()
    result = []
    for h in headers:
        if h.part in seen:
            continue
        seen.add(h.part)
        result.append(h)
    return result


def _detach_drawing(drawing):
    """从文档中安全移除一个 drawing 元素（连同其外层 w:r 一并移除）。"""
    run = drawing.getparent()  # w:r
    if run is not None and run.tag == qn("w:r") and run.getparent() is not None:
        run.getparent().remove(run)
    elif drawing.getparent() is not None:
        drawing.getparent().remove(drawing)


def _remove_in_header(header, name=None, clear_any=False):
    """移除页眉中的水印图形。

    - name 给定时：仅删除该名称的图形（插入前“替换”用）。
    - name 为 None 且 clear_any=False：删除所有本工具标记的水印（插入前“替换”用）。
    - name 为 None 且 clear_any=True：删除本工具标记的水印 **以及** 任何“衬于文字下方”
      (behindDoc=1) 的图形——这样能清掉 Word 原生水印或其它工具留下的水印，
      满足“去掉原本就有水印的 Word”的需求。
    """
    root = header._element
    removed = 0
    for drawing in list(root.iter(qn("w:drawing"))):
        anchor = drawing.find(".//" + qn("wp:anchor"))
        behind = anchor is not None and anchor.get("behindDoc") == "1"
        docPr = drawing.find(".//" + qn("wp:docPr"))
        nm = docPr.get("name") if docPr is not None else None
        if name is not None:
            if nm == name:
                _detach_drawing(drawing)
                removed += 1
        elif nm in MARK_NAMES or (clear_any and behind):
            _detach_drawing(drawing)
            removed += 1
    return removed


def _add_drawing_to_header(header, drawing):
    """把 drawing 追加到页眉的一个新 run 中（必要时先为该节创建独立页眉）。"""
    # 触发 python-docx 为“链接到前一节”的页眉创建独立 part
    if not header.paragraphs:
        header.add_paragraph()
    p = header.paragraphs[0]
    r = p.add_run()
    r._r.append(drawing)


def _has_watermark_in_doc(document) -> bool:
    for section in document.sections:
        for header in _iter_headers(document, section):
            for drawing in header._element.iter(qn("w:drawing")):
                docPr = drawing.find(".//" + qn("wp:docPr"))
                if docPr is not None and docPr.get("name") in MARK_NAMES:
                    return True
    return False


# ---------------------------------------------------------------------------
# 对外 API
# ---------------------------------------------------------------------------
def insert_watermark(path: str, kinds, **opts) -> dict:
    """向 .docx 插入水印。kinds 为包含 'text' / 'image' 的列表，可同时含两者。

    文本与图像各自的参数（角度/透明度/缩放）相互独立、可分别调整：
    通过 opts["text"] 与 opts["image"] 两份独立字典传入，互不影响。
    当 kinds 同时含 text 与 image 时，文本与图像水印会同时叠加且各自使用自己的参数。
    """
    document = Document(path)
    kinds = [k for k in kinds if k in ("text", "image")]
    text_opts = opts.get("text", {}) or {}
    image_opts = opts.get("image", {}) or {}

    # 准备各层水印图像：文本层覆盖约 0.85 页宽，图像层约 0.6 页宽
    prepared = []
    if "text" in kinds:
        angle = float(text_opts.get("angle", 45))
        transparency = float(text_opts.get("transparency", 0.5))  # 0=不透明,1=全透明
        scale = float(text_opts.get("scale", 1.0))
        offset_x = float(text_opts.get("offset_x", 0.0))  # 水平偏移（占页宽百分比，0=居中）
        offset_y = float(text_opts.get("offset_y", 0.0))  # 垂直偏移（占页高百分比，0=居中）
        alpha = int(max(0, min(255, (1 - transparency) * 255)))
        text = text_opts.get("text", "水印")
        font_size = int(text_opts.get("font_size", 120))
        color = text_opts.get("color", (128, 128, 128))
        # 中西文分别指定字体：汉字用 cn_font_name、拉丁字母/数字用 latin_font_name
        cn_font_name = text_opts.get("cn_font_name")
        latin_font_name = text_opts.get("latin_font_name")
        img = render_text_png(text, font_size, color, alpha, cn_font_name, latin_font_name)
        # 基准覆盖比例由字号决定（再乘 scale），使“字号”真实控制水印大小
        base = text_base(font_size)
        prepared.append((MARK_TEXT, img, base, angle, scale, offset_x, offset_y))
    if "image" in kinds:
        angle = float(image_opts.get("angle", 45))
        transparency = float(image_opts.get("transparency", 0.5))
        scale = float(image_opts.get("scale", 1.0))
        offset_x = float(image_opts.get("offset_x", 0.0))
        offset_y = float(image_opts.get("offset_y", 0.0))
        alpha = int(max(0, min(255, (1 - transparency) * 255)))
        image_path = image_opts.get("image_path")
        if not image_path or not os.path.exists(image_path):
            raise FileNotFoundError(f"图片不存在: {image_path}")
        img = prepare_image_png(image_path, alpha)
        prepared.append((MARK_IMG, img, 0.6, angle, scale, offset_x, offset_y))

    if not prepared:
        raise ValueError("未启用任何水印（文本与图像均未启用），请至少启用一种。")

    # 预先把每层 PNG 字节准备好，循环内直接复用（避免重复渲染/序列化）
    layers = []
    for mark, img, base, angle, scale, offset_x, offset_y in prepared:
        bio = BytesIO()
        img.save(bio, "PNG")
        layers.append((mark, bio.getvalue(), img.width, img.height, base, angle, scale,
                       offset_x, offset_y))

    docpr_counter = 1
    inserted = 0
    processed_parts = set()  # 全局按 part 去重，避免链接页眉被重复插入
    for section in document.sections:
        page_w = int(section.page_width)
        page_h = int(section.page_height)
        for header in _iter_headers(document, section):
            if header.part in processed_parts:
                continue
            processed_parts.add(header.part)
            _remove_in_header(header)  # 先清本页眉所有本工具水印（replace 语义：仅保留本次所选类型）
            for (mark, png, nat_w, nat_h, base, angle, scale, offset_x, offset_y) in layers:
                bio = BytesIO(png)
                rId, img_part = header.part.get_or_add_image(bio)
                disp_w = int(page_w * base * scale)
                disp_h = int(disp_w * nat_h / nat_w) if nat_w else disp_w
                # 居中基准 + 百分比偏移；offset 正=右/下，负=左/上
                pos_x = page_w // 2 - disp_w // 2 + offset_x / 100.0 * page_w
                pos_y = page_h // 2 - disp_h // 2 + offset_y / 100.0 * page_h
                drawing = _make_drawing(rId, disp_w, disp_h, angle, mark, docpr_counter,
                                       pos_x, pos_y)
                docpr_counter += 1
                _add_drawing_to_header(header, drawing)
                inserted += 1

    document.save(path)
    return {"ok": True, "engine": "docx", "inserted": inserted, "kinds": kinds}


def clear_watermark(path: str) -> dict:
    """清除 .docx 中的水印：本工具添加的“以及”Word 原生的（衬于文字下方的图形）。

    既能清掉本工具留下的文本/图像水印，也能清掉原本就带水印的 Word 文档。
    """
    document = Document(path)
    removed = 0
    processed_parts = set()
    for section in document.sections:
        for header in _iter_headers(document, section):
            if header.part in processed_parts:
                continue
            processed_parts.add(header.part)
            removed += _remove_in_header(header, clear_any=True)  # 清掉所有水印类图形
    document.save(path)
    return {"ok": True, "engine": "docx", "removed": removed}


def has_watermark(path: str) -> bool:
    document = Document(path)
    return _has_watermark_in_doc(document)
