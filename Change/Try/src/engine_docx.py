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

from PIL import Image, ImageColor, ImageDraw, ImageFont
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

# “防去除加固”用的伪装显示名。
# Word/WPS 的“删除水印”按钮、以及多数去水印脚本，都是按图形名称或结构去匹配
# 水印图形的；这里改用普通插图那样的无害名字，让它们认不出来、删不干净。
# 本工具自己的识别标记改存到 wp:docPr/@descr（标准属性，Word 会保留），
# 因此“伪装”不影响本工具的识别与一键清除。
DECOY_NAMES = ("图片", "Image", "Shape", "Graphic", "Object", "Content")


def decoy_name(idx: int) -> str:
    """生成第 idx 个图形的伪装显示名（形如“图片 101”，与 Word 自动命名风格一致）。"""
    base = DECOY_NAMES[idx % len(DECOY_NAMES)]
    return f"{base} {100 + idx}"

# 文本水印尺寸模型：font_size 作为“字号/实际大小”的主控制项（与 Word COM 引擎一致），
# scale 作为微调倍数。默认 font_size=120、scale=1.0 时文本覆盖约 0.85 倍页宽，
# 与历史外观保持一致；字号越小水印越小（此前 .docx 路径字号只影响清晰度、不改大小，已修正）。
TEXT_COVER = 0.85
TEXT_REF_SIZE = 120.0


def text_base(font_size: float) -> float:
    """把字号换算成“占页宽比例”的基准值（再乘用户 scale 得到最终覆盖比例）。"""
    return TEXT_COVER * (float(font_size) / TEXT_REF_SIZE)


def tile_layout(page_w, page_h, nat_w, nat_h, rows, cols, scale=1.0,
                offset_x=0.0, offset_y=0.0):
    """平铺布局：把页面切成 rows×cols 个单元格，每个格内放一份水印。

    返回 [(x, y, w, h), ...]（与 page_w/page_h 同单位）。
    单个瓦片按单元格等比缩放（保持原始宽高比），再整体按百分比偏移。

    为什么要平铺：稀疏的单个居中水印，PS 的内容识别填充 / AI 修图去水印很容易
    抹掉且不留痕迹；覆盖整页的密集纹理要想去掉就得把整页重画，难度陡增。
    """
    rows = max(1, int(rows))
    cols = max(1, int(cols))
    cell_w = page_w / float(cols)
    cell_h = page_h / float(rows)
    # 瓦片尽量占满格宽（scale>1 时会与相邻瓦片重叠，密度更高、更难被修图抹掉）；
    # 纵向按格高的 90% 作为上限，避免文字水印被拉得过高变形。
    want_w = cell_w * scale
    want_h = cell_h * 0.9 * scale
    if nat_w and nat_h:
        disp_w = min(want_w, want_h * float(nat_w) / float(nat_h))
    else:
        disp_w = want_w
    disp_w = max(1.0, disp_w)
    disp_h = max(1.0, disp_w * (float(nat_h) / float(nat_w) if nat_w else 1.0))

    shift_x = offset_x / 100.0 * page_w
    shift_y = offset_y / 100.0 * page_h
    boxes = []
    for r in range(rows):
        for c in range(cols):
            x = c * cell_w + (cell_w - disp_w) / 2.0 + shift_x
            y = r * cell_h + (cell_h - disp_h) / 2.0 + shift_y
            boxes.append((x, y, disp_w, disp_h))
    return boxes

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


def parse_color(color) -> tuple[int, int, int]:
    """把颜色入参统一成 (r, g, b)——兼容 (r,g,b) / "#RRGGBB" / "#RGB" / 颜色名。

    必要性：以前直接写 r, g, b = color[0], color[1], color[2]，一旦有人按常见习惯传入
    十六进制字符串 "#FF0000"，取到的会是字符 '#' / 'F' / 'F'，直到 Pillow 绘图时才抛
    "TypeError: 'str' object cannot be interpreted as an integer"，非常难定位。
    现在统一在这里兜住，非法入参回退为中性灰而不是崩溃。
    """
    if isinstance(color, str):
        s = color.strip().lstrip("#")
        if len(s) == 3:
            s = "".join(ch * 2 for ch in s)
        if len(s) >= 6:
            try:
                return (int(s[0:2], 16), int(s[2:4], 16), int(s[4:6], 16))
            except ValueError:
                pass
        try:
            return tuple(int(v) for v in ImageColor.getrgb(color)[:3])
        except Exception:
            return (128, 128, 128)
    try:
        if len(color) >= 3:
            return tuple(int(v) for v in color[:3])
    except Exception:
        pass
    return (128, 128, 128)


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

    def _load(p, index=0):
        try:
            # index= 用于 TTC 多 face 字体（如“微软雅黑 Light”对应 msyh.ttc 第 1 面）；
            # 非 TTC 字体 index=0 即默认行为，与改动前一致（风险 k）。
            return ImageFont.truetype(p, font_size, index=index) if p else ImageFont.load_default()
        except Exception:
            return ImageFont.load_default()

    # 解析用户所选中/西文字体对应的 TTC face 索引，保证多 face 集合字体的正确字重被加载
    cn_index = word_fonts.font_index_of(cn_font_name) if cn_font_name else 0
    latin_index = word_fonts.font_index_of(latin_font_name) if latin_font_name else 0
    cn_font = _load(cn_path, cn_index)
    latin_font = _load(latin_path, latin_index)
    fallback_font = _load(fallback_path) if fallback_path else cn_font

    r, g, b = parse_color(color)

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
    """打开用户图片并整体乘以目标透明度，返回 RGBA 图像。

    支持 png / jpg / jpeg / bmp / gif / webp 等常见位图；
    若传入 .pdf，则用 PyMuPDF 渲染首页为位图（仅首页），再按透明度处理。
    """
    ext = os.path.splitext(path)[1].lower()
    if ext == ".pdf":
        import pymupdf  # PyMuPDF（已在 spec hiddenimports 中声明，保证打进包）
        doc = pymupdf.open(path)
        try:
            page = doc.load_page(0)                  # 只取首页
            # 以 200dpi 渲染，保证插入 Word 后的清晰度
            zoom = max(1.0, 200.0 / 72.0)
            pix = page.get_pixmap(matrix=pymupdf.Matrix(zoom, zoom), alpha=True)
            img = Image.frombytes("RGBA", (pix.width, pix.height), pix.samples)
        finally:
            doc.close()
    else:
        img = Image.open(path).convert("RGBA")
    a_channel = img.split()[3].point(lambda p: int(p * alpha / 255))
    img.putalpha(a_channel)
    return img


# ---------------------------------------------------------------------------
# DrawingML 构造
# ---------------------------------------------------------------------------
def _make_drawing(rId: str, cx: int, cy: int, rot_deg: float, name: str, descr: str,
                 docpr_id: int, pos_x: int, pos_y: int) -> etree._Element:
    """构造一个衬于文字下方、可旋转、按绝对坐标定位的锚定图形 XML。

    pos_x / pos_y 为图形左上角相对页面原点（页面左上角）的 EMU 坐标，
    由调用方根据“居中基准 + 百分比偏移”算出，从而支持上下左右独立调整。

    name 为对外显示名（加固时用伪装名，避免被“删除水印”按名字抓走），
    descr 存本工具的私有标记——Word 会保留该标准属性，故识别与清除不受影响。
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

    docPr = etree.SubElement(anchor, qn("wp:docPr"),
                             {"id": str(docpr_id), "name": name, "descr": descr})
    etree.SubElement(anchor, qn("wp:cNvGraphicFramePr"))

    graphic = etree.SubElement(anchor, qn("a:graphic"))
    graphicData = etree.SubElement(graphic, qn("a:graphicData"), {"uri": PIC})
    pic = etree.SubElement(graphicData, qn("pic:pic"))

    nvPicPr = etree.SubElement(pic, qn("pic:nvPicPr"))
    etree.SubElement(nvPicPr, qn("pic:cNvPr"),
                     {"id": str(docpr_id), "name": name + "_img", "descr": descr})
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
    return _dedup_parts(headers)


def _iter_footers(document, section):
    """返回该节所有“非链接”的页脚（主/首页/偶数页）。加固时把水印也放进页脚。"""
    footers = [section.footer]
    try:
        if section.different_first_page_header_footer:
            footers.append(section.first_page_footer)
    except Exception:
        pass
    try:
        if document.part.settings.odd_and_even_pages_header_footer:
            footers.append(section.even_page_footer)
    except Exception:
        pass
    return _dedup_parts(footers)


def _dedup_parts(parts):
    """按底层 part 去重：链接到前一节的页眉/页脚共享同一 part，只处理一次。"""
    seen = set()
    result = []
    for p in parts:
        if p.part in seen:
            continue
        seen.add(p.part)
        result.append(p)
    return result


def _iter_parts(document, section, include_footers=False):
    """返回该节需要写入水印的所有页眉（+可选页脚）。"""
    parts = _iter_headers(document, section)
    if include_footers:
        parts = parts + _iter_footers(document, section)
    return parts


def _detach_drawing(drawing):
    """从文档中安全移除一个 drawing 元素。

    安全约束：**只摘掉目标图形本身，绝不连带删掉同 run 内的其它内容**。

    一个 `w:r` 里可以同时放 `w:t`（用户自己的文字）和 `w:drawing`（水印），
    例如 Word 重新排版后合并 run、或旧版本工具留下的结构。若像以前那样
    “整条 w:r 一起摘掉”，用户那部分文字会跟着水印一起消失。

    因此这里分两步：
    1. 先把目标 `w:drawing` 从 `w:r` 上摘下来；
    2. 只有在该 run 剥掉图形后**已无任何子元素**时，才把空 run 一起移除
       ——水印 run 正是这种情况，所以正常路径的结果与改动前完全一致。
    """
    run = drawing.getparent()  # w:r
    if run is not None and run.tag == qn("w:r"):
        run.remove(drawing)
        # 还有内容（用户文字 / 另一张图 / w:br 等）就必须保留这条 run
        if len(run) == 0 and run.getparent() is not None:
            run.getparent().remove(run)
    elif drawing.getparent() is not None:
        drawing.getparent().remove(drawing)


def _mark_of(drawing):
    """取出图形上本工具的私有标记（descr 优先，兼容旧版写在 name 上的文件）。

    加固模式下显示名是伪装名，标记只存在 descr 里，因此必须两个都查。
    """
    docPr = drawing.find(".//" + qn("wp:docPr"))
    if docPr is None:
        return None
    ds = docPr.get("descr")
    if ds in MARK_NAMES:
        return ds
    nm = docPr.get("name")
    if nm in MARK_NAMES:
        return nm
    return None


def _is_picture(drawing) -> bool:
    """图形是否为「图片」类（a:graphicData/@uri == drawingml/2006/picture）。

    clear_any 的收窄条件：只有**图片**才允许凭“衬于文字下方 (behindDoc=1)”
    这条结构特征被推定为水印。Word 原生水印本身就是图片，所以照旧会被清掉；
    但用户在页眉里自己画的形状（Word 存成 mc:AlternateContent + wps:wsp，
    uri 是 wordprocessingShape 而非 picture）或图表并不是水印，不该被连带删掉。
    """
    gd = drawing.find(".//" + qn("a:graphicData"))
    return gd is not None and gd.get("uri") == PIC


def _remove_in_part(part, name=None, clear_any=False):
    """移除页眉/页脚中的水印图形。

    - name 给定时：仅删除该标记的图形（插入前“替换”用）。
    - name 为 None 且 clear_any=False：删除所有本工具标记的水印（插入前“替换”用）。
    - name 为 None 且 clear_any=True：删除本工具标记的水印 **以及** 任何“衬于文字下方”
      (behindDoc=1) 的**图片**图形——这样能清掉 Word 原生水印或其它工具留下的水印，
      满足“去掉原本就有水印的 Word”的需求。仅限图片是为了不去动用户自己画
      的形状（见 _is_picture）。
    """
    root = part._element
    removed = 0
    for drawing in list(root.iter(qn("w:drawing"))):
        anchor = drawing.find(".//" + qn("wp:anchor"))
        behind = anchor is not None and anchor.get("behindDoc") == "1"
        mark = _mark_of(drawing)
        if name is not None:
            if mark == name:
                _detach_drawing(drawing)
                removed += 1
        elif mark is not None or (clear_any and behind and _is_picture(drawing)):
            _detach_drawing(drawing)
            removed += 1
    return removed


# 兼容旧调用名
_remove_in_header = _remove_in_part


def _add_drawing_to_part(part, drawing):
    """把 drawing 追加到页眉/页脚的一个新 run 中（必要时先创建独立 part）。

    注意（私有 API 依赖，风险 g）：`r._r` 是 python-docx 的“私有”内部属性（CT_R 元素）。
    python-docx 没有公开 API 能把一个现成的 DrawingML anchor 注入到 run 里
    （`run.add_picture` 只能新建行内图片，而水印需要“衬于文字下方 behindDoc”的
    浮动型 anchor），因此这里只能用内部属性。一旦 python-docx 大版本改动 CT_R 结构，
    这里可能失效——下面针对 AttributeError 做了可读的错误提示，便于第一时间定位，
    而不是抛出晦涩的底层异常。成功路径与历史 XML 格式完全一致，行为不变。
    """
    # 触发 python-docx 为“链接到前一节”的页眉/页脚创建独立 part
    if not part.paragraphs:
        part.add_paragraph()
    p = part.paragraphs[0]
    r = p.add_run()
    try:
        r._r.append(drawing)
    except AttributeError as e:  # python-docx 内部结构变动导致 _r 不可用
        raise RuntimeError(
            "无法将水印图形写入页眉/页脚：python-docx 内部接口(_r)不可用，"
            "可能是 python-docx 版本不兼容，请升级本工具或固定 python-docx 版本。"
        ) from e


# 兼容旧调用名
_add_drawing_to_header = _add_drawing_to_part


def _iter_marked_drawings(document):
    """遍历文档（含页脚）里所有带本工具标记的图形。"""
    for section in document.sections:
        for part in _iter_parts(document, section, include_footers=True):
            for drawing in part._element.iter(qn("w:drawing")):
                if _mark_of(drawing) is not None:
                    yield drawing


def _has_watermark_in_doc(document) -> bool:
    for _ in _iter_marked_drawings(document):
        return True
    return False


def _count_marks_in_doc(document) -> int:
    """统计本工具添加的水印图形份数（守护用它判断“够不够份”）。"""
    return sum(1 for _ in _iter_marked_drawings(document))


# ---------------------------------------------------------------------------
# 对外 API
# ---------------------------------------------------------------------------
def insert_watermark(path: str, kinds, **opts) -> dict:
    """向 .docx 插入水印。kinds 为包含 'text' / 'image' 的列表，可同时含两者。

    文本与图像各自的参数（角度/透明度/缩放）相互独立、可分别调整：
    通过 opts["text"] 与 opts["image"] 两份独立字典传入，互不影响。
    当 kinds 同时含 text 与 image 时，文本与图像水印会同时叠加且各自使用自己的参数。

    防去除加固（均可选，默认关闭，不改变原有外观）：
    - opts["tile"]=True + tile_rows/tile_cols：改为平铺满页（对抗 PS/AI 修图去除）；
    - opts["redundant"]=True：页眉之外也写进页脚，并用伪装显示名（对抗
      Word/WPS“删除水印”按钮与按名字匹配的去水印脚本）。
    """
    document = Document(path)
    kinds = [k for k in kinds if k in ("text", "image")]
    text_opts = opts.get("text", {}) or {}
    image_opts = opts.get("image", {}) or {}
    tile = bool(opts.get("tile", False))
    tile_rows = int(opts.get("tile_rows", 4) or 4)
    tile_cols = int(opts.get("tile_cols", 3) or 3)
    redundant = bool(opts.get("redundant", False))

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
        for part in _iter_parts(document, section, include_footers=redundant):
            if part.part in processed_parts:
                continue
            processed_parts.add(part.part)
            # 先清本 part 所有本工具水印（replace 语义：仅保留本次所选类型）
            _remove_in_part(part)
            for (mark, png, nat_w, nat_h, base, angle, scale, offset_x, offset_y) in layers:
                # 同一份 PNG 在同一 part 内会被复用（按内容哈希去重），平铺不会撑大文件
                rId, img_part = part.part.get_or_add_image(BytesIO(png))
                if tile:
                    boxes = tile_layout(page_w, page_h, nat_w, nat_h,
                                        tile_rows, tile_cols, scale, offset_x, offset_y)
                else:
                    disp_w = page_w * base * scale
                    disp_h = disp_w * nat_h / nat_w if nat_w else disp_w
                    # 居中基准 + 百分比偏移；offset 正=右/下，负=左/上
                    boxes = [(page_w / 2.0 - disp_w / 2.0 + offset_x / 100.0 * page_w,
                              page_h / 2.0 - disp_h / 2.0 + offset_y / 100.0 * page_h,
                              disp_w, disp_h)]
                for (pos_x, pos_y, disp_w, disp_h) in boxes:
                    # 加固时用伪装显示名，私有标记放 descr（不影响本工具识别/清除）
                    disp_name = decoy_name(docpr_counter) if redundant else mark
                    drawing = _make_drawing(rId, int(disp_w), int(disp_h), angle,
                                            disp_name, mark, docpr_counter, pos_x, pos_y)
                    docpr_counter += 1
                    _add_drawing_to_part(part, drawing)
                    inserted += 1

    document.save(path)
    return {"ok": True, "engine": "docx", "inserted": inserted, "kinds": kinds,
            "tile": tile, "redundant": redundant}


def detect_watermark_types(path: str) -> set:
    """检测文档里存在哪些“本工具添加”的水印类型。

    返回集合，元素是 'text' / 'image' 的子集（可能同时含两者）。
    仅识别带本工具私有标记（MARK_TEXT / MARK_IMG）的图形——不把正文、
    用户图片、Word 原生水印误判为本工具水印。
    """
    document = Document(path)
    found = set()
    for section in document.sections:
        for part in _iter_parts(document, section, include_footers=True):
            for drawing in part._element.iter(qn("w:drawing")):
                mark = _mark_of(drawing)
                if mark == MARK_TEXT:
                    found.add("text")
                elif mark == MARK_IMG:
                    found.add("image")
                if found == {"text", "image"}:
                    return found
    return found


def clear_watermark(path: str, kinds: list = None) -> dict:
    """清除 .docx 中的水印。

    - kinds 为 None：清掉本工具添加的，以及 Word 原生的（衬于文字下方的图形）。
      既能清掉本工具留下的文本/图像水印，也能清掉原本就带水印的 Word 文档。
    - kinds 给定（包含 'text' / 'image'）：只删带对应标记的水印，保留其它类型，
      且**不**动 Word 原生的 behindDoc 图形（避免“只去文字、保留图片”时被误删）。
      例如 kinds=['text'] 仅去除文字水印、图片水印原样保留。
    """
    document = Document(path)
    kinds = set(kinds or [])
    only_kinds = bool(kinds)
    # 把用户传入的 'text'/'image' 映射到本工具的私有标记常量
    kind_marks = set()
    if "text" in kinds:
        kind_marks.add(MARK_TEXT)
    if "image" in kinds:
        kind_marks.add(MARK_IMG)
    removed = 0
    processed_parts = set()
    for section in document.sections:
        # 始终扫描页脚：加固模式下水印也会写进页脚，否则会清不干净
        for part in _iter_parts(document, section, include_footers=True):
            if part.part in processed_parts:
                continue
            processed_parts.add(part.part)
            if only_kinds:
                # 仅按类型删除：只删 mark 命中 kinds 的图形，绝不碰原生水印
                for drawing in list(part._element.iter(qn("w:drawing"))):
                    if _mark_of(drawing) in kind_marks:
                        _detach_drawing(drawing)
                        removed += 1
            else:
                removed += _remove_in_part(part, clear_any=True)  # 清掉所有水印类图形
    document.save(path)
    return {"ok": True, "engine": "docx", "removed": removed, "kinds": sorted(kinds)}


def has_watermark(path: str) -> bool:
    document = Document(path)
    return _has_watermark_in_doc(document)


def count_watermarks(path: str) -> int:
    """返回本工具添加的水印图形份数（守护按“份数”判断是否被删过）。"""
    document = Document(path)
    return _count_marks_in_doc(document)
