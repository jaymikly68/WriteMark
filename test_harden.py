"""
防去除加固测试：平铺满页、多份冗余嵌入、伪装显示名、清除与份数校验、守护按份数补回。

背景（为什么需要加固）：
- 守护只在本机、工具运行时盯着输出文件；文件一旦发出去，它管不到。
- PS 打不开 .docx，别人用 PS/AI 去水印必然走「docx → PDF/图片 → 修图」，
  此时水印已是像素，docx 层面的任何加固都无效——只能靠“让水印在像素层面难去掉”，
  即平铺满页（覆盖整页的密集纹理，AI/PS 要抹就得重画整页）。
- 另一条路径是在 Word/WPS 里点“删除水印”或用脚本批量清，这靠
  多份冗余嵌入 + 伪装显示名来挡。
"""
from __future__ import annotations

import os
import re
import shutil
import sys
import tempfile
import time
import zipfile

from docx import Document

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from watermark_tool import core, engine_docx


def _make_docx(path):
    doc = Document()
    doc.add_paragraph("第一段 hello world 测试文档")
    doc.add_paragraph("第二段内容，用于验证水印插入不影响正文。")
    doc.save(path)


def _text_opts(**over):
    d = {"text": "机密 CONFIDENTIAL", "font_size": 120, "color": (192, 0, 0),
         "angle": 30, "transparency": 0.6, "scale": 1.0,
         "offset_x": 0.0, "offset_y": 0.0,
         "cn_font_name": "仿宋", "latin_font_name": "Times New Roman"}
    d.update(over)
    return {"text": d}


def _part_xml(path):
    """返回 (页眉 XML 列表, 页脚 XML 列表)。"""
    z = zipfile.ZipFile(path)
    hs = [z.read(n).decode("utf8") for n in z.namelist()
          if re.match(r"word/header\d+\.xml$", n)]
    fs = [z.read(n).decode("utf8") for n in z.namelist()
          if re.match(r"word/footer\d+\.xml$", n)]
    return hs, fs


def _names(xml_list):
    out = []
    for x in xml_list:
        out += re.findall(r'<wp:docPr[^>]*name="([^"]+)"', x)
    return out


def _descrs(xml_list):
    out = []
    for x in xml_list:
        out += re.findall(r'<wp:docPr[^>]*descr="([^"]+)"', x)
    return out


def _positions(xml_list):
    """用 lxml 解析出各水印图形的绝对定位坐标集合 {(x, y), ...}（EMU 字符串）。"""
    from lxml import etree
    WP = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
    out = set()
    for x in xml_list:
        root = etree.fromstring(x.encode("utf8"))
        for anchor in root.iter("{%s}anchor" % WP):
            h = anchor.find(".//{%s}positionH/{%s}posOffset" % (WP, WP))
            v = anchor.find(".//{%s}positionV/{%s}posOffset" % (WP, WP))
            if h is not None and v is not None:
                out.add((h.text, v.text))
    return out


def test_standard_mode_unchanged():
    """默认（不开加固）应与历史行为一致：页眉 1 份、名字就是标记名。"""
    d = tempfile.mkdtemp(prefix="_hard_")
    p = os.path.join(d, "a.docx")
    _make_docx(p)
    res = engine_docx.insert_watermark(p, ["text"], **_text_opts())
    assert res["inserted"] == 1, res
    hs, fs = _part_xml(p)
    assert _names(hs) == [engine_docx.MARK_TEXT], _names(hs)
    assert not fs, "标准模式不应写入页脚"
    print("[标准] 单份 + 页眉 + 原名，与历史行为一致")


def test_tile_produces_grid():
    """平铺：一个 part 内写入 rows×cols 份。"""
    d = tempfile.mkdtemp(prefix="_hard_")
    p = os.path.join(d, "b.docx")
    _make_docx(p)
    res = engine_docx.insert_watermark(p, ["text"], tile=True,
                                       tile_rows=4, tile_cols=3, **_text_opts())
    assert res["inserted"] == 12, res
    assert res["tile"] is True
    hs, _ = _part_xml(p)
    assert len(_names(hs)) == 12, len(_names(hs))
    # 平铺必须真的分散到不同坐标（不是 12 份叠在正中）
    pos = _positions(hs)
    assert len(pos) == 12, f"平铺 12 份应有 12 个不同坐标，实际 {len(pos)}"
    xs = {p[0] for p in pos}
    ys = {p[1] for p in pos}
    assert len(xs) == 3 and len(ys) == 4, f"应为 3 列 × 4 行，实际 {len(xs)}×{len(ys)}"
    print(f"[平铺] 4×3 = 12 份，分布在 {len(xs)} 列 × {len(ys)} 行的 {len(pos)} 个坐标")


def test_redundant_writes_footer_only():
    """冗余嵌入：水印同时写入页脚；且显示名不含 watermark 字样，标记藏在 descr。"""
    d = tempfile.mkdtemp(prefix="_hard_")
    p = os.path.join(d, "c.docx")
    _make_docx(p)
    res = engine_docx.insert_watermark(p, ["text"], redundant=True, **_text_opts())
    assert res["inserted"] == 2, res          # 页眉 1 + 页脚 1
    hs, fs = _part_xml(p)
    assert fs, "冗余模式应写入页脚"
    assert len(_names(hs)) == 1 and len(_names(fs)) == 1
    all_names = _names(hs) + _names(fs)
    assert all("watermark" not in n.lower() for n in all_names), all_names
    # 私有标记仍可在 descr 中读到（保证本工具能识别与清除）
    all_descrs = _descrs(hs) + _descrs(fs)
    assert all(d == engine_docx.MARK_TEXT for d in all_descrs), all_descrs
    print(f"[冗余] 页眉 1 + 页脚 1；显示名 {all_names}（无 watermark 字样），descr 保留标记")


def test_names_not_watermark_like():
    """伪装名不应命中“按名字匹配水印”的常见规则。"""
    d = tempfile.mkdtemp(prefix="_hard_")
    p = os.path.join(d, "d.docx")
    _make_docx(p)
    engine_docx.insert_watermark(p, ["text"], tile=True, tile_rows=3, tile_cols=2,
                                 redundant=True, **_text_opts())
    hs, fs = _part_xml(p)
    names = _names(hs) + _names(fs)
    assert len(names) == 12, len(names)     # 6 份 × 2 个 part
    for n in names:
        low = n.lower()
        for bad in ("watermark", "水印", "wbg_watermark", "powerplus"):
            assert bad not in low, f"{n} 命中关键字 {bad}"
    print(f"[伪装] 12 份图形名均无 watermark/水印 等关键字；样例 {names[:3]}")


def test_clear_removes_tiled_and_footer():
    """一键清除必须能清掉平铺 + 页脚里的全部水印。"""
    d = tempfile.mkdtemp(prefix="_hard_")
    p = os.path.join(d, "e.docx")
    _make_docx(p)
    engine_docx.insert_watermark(p, ["text"], tile=True, tile_rows=4, tile_cols=2,
                                 redundant=True, **_text_opts())
    before = engine_docx.count_watermarks(p)
    assert before == 16, before
    res = engine_docx.clear_watermark(p)
    assert res["removed"] == 16, res
    assert engine_docx.count_watermarks(p) == 0
    assert not engine_docx.has_watermark(p)
    hs, fs = _part_xml(p)
    assert not _names(hs) and not _names(fs)
    print(f"[清除] 平铺+页脚共 {before} 份一次清空，页眉/页脚无残留")


def test_count_watermarks():
    d = tempfile.mkdtemp(prefix="_hard_")
    p = os.path.join(d, "f.docx")
    _make_docx(p)
    assert engine_docx.count_watermarks(p) == 0
    engine_docx.insert_watermark(p, ["text"], **_text_opts())
    assert engine_docx.count_watermarks(p) == 1
    engine_docx.insert_watermark(p, ["text"], tile=True, tile_rows=5, tile_cols=3,
                                 redundant=True, **_text_opts())
    assert engine_docx.count_watermarks(p) == 30
    print("[计数] 份数统计正确（0 → 1 → 30）")


def test_watchdog_refills_partial_removal():
    """守护按份数校验：水印被删掉一部分也应整体补齐。"""
    from watermark_tool import watchdog
    d = tempfile.mkdtemp(prefix="_hard_")
    src = os.path.join(d, "src.docx")
    out = os.path.join(d, "out.docx")
    _make_docx(src)

    opts = {**_text_opts(), "image": {}, "tile": True,
            "tile_rows": 4, "tile_cols": 2, "redundant": True}
    res = core.insert_watermark(src, ["text"], output_path=out, **opts)
    total = res["inserted"]
    assert total == 16, total

    logs = []
    wd = watchdog.WatermarkWatchdog(src, out, ["text"], opts, interval=0.3,
                                    log=lambda m, verbose=False: logs.append(m),
                                    expected=total)
    wd.start()
    try:
        time.sleep(0.6)
        # 模拟“别人删掉一半”：只留 4 份
        doc = Document(out)

        def _strip(part, keep):
            root = part._element
            drawings = list(root.iter(engine_docx.qn("w:drawing")))
            for dr in drawings[keep:]:
                engine_docx._detach_drawing(dr)

        sec = doc.sections[0]
        _strip(sec.header, 4)      # 页眉 8 份砍到 4 份
        _strip(sec.footer, 0)      # 页脚 8 份全删
        doc.save(out)
        partial = engine_docx.count_watermarks(out)
        assert partial == 4, partial

        def wait_for(pred, timeout=4.0):
            t0 = time.time()
            while time.time() - t0 < timeout:
                try:
                    if pred():
                        return True
                except Exception:
                    pass
                time.sleep(0.15)
            return False

        ok = wait_for(lambda: engine_docx.count_watermarks(out) >= total)
        assert ok, f"守护未补齐（当前 {engine_docx.count_watermarks(out)}/{total}）"
        assert any("部分移除" in m for m in logs), logs
    finally:
        wd.stop()
    print(f"[守护] 16 份被删到 4 份后自动补齐到 {engine_docx.count_watermarks(out)} 份")


if __name__ == "__main__":
    test_standard_mode_unchanged()
    test_tile_produces_grid()
    test_redundant_writes_footer_only()
    test_names_not_watermark_like()
    test_clear_removes_tiled_and_footer()
    test_count_watermarks()
    test_watchdog_refills_partial_removal()
    print("\n防去除加固测试全部通过 ✅")
