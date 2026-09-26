"""
二次安全审查（commit 84a773e 之后）—— 主题 1：watchdog + preserve_existing。

要回答的问题：output.docx 被用户改过正文 / 表格 / 图片 / 段落格式之后，
“水印被删 → 守护自动补回”这一轮，**用户的修改是否完整保留**。

以及更底层的一条：preserve_existing=True 是否真的**始终**在现有 output 上操作，
而不是在某些分支里悄悄回退成「从 source 重新生成」。

背景：84a773e 之前，守护补回一律走 core.insert_watermark 的默认值，即从源文件
重新 copy 一份再加水印 —— 用户在输出文件上的所有正文修改会被整份覆盖。
84a773e 引入 preserve_existing 参数后，这条路径才改成“在现有 output 上先清后插”。

本文件同时覆盖审查 2 的一条相关断言：守护补回时的“先清”步骤究竟清掉了什么。
"""
from __future__ import annotations

import io
import os
import sys
import tempfile

from docx import Document
from docx.shared import Pt
from PIL import Image

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from watermark_tool import core, engine_docx, watchdog as wd_mod


def _png_bytes(color=(10, 20, 30), size=(40, 40)):
    img = Image.new("RGB", size, color)
    bio = io.BytesIO()
    img.save(bio, "PNG")
    return bio.getvalue()


def _make_source(path):
    """造一份「用户原始素材」：正文段落 + 带格式的段落 + 表格 + 正文图片。"""
    doc = Document()
    p = doc.add_paragraph("原始正文第一段")
    p.runs[0].bold = True
    p.runs[0].font.size = Pt(18)
    doc.add_paragraph("原始正文第二段")

    t = doc.add_table(rows=2, cols=2)
    t.cell(0, 0).text = "表A1"
    t.cell(0, 1).text = "表A2"
    t.cell(1, 0).text = "表B1"
    t.cell(1, 1).text = "表B2"

    doc.add_picture(io.BytesIO(_png_bytes()), width=Pt(120))
    doc.save(path)
    return path


def _text_opts(**over):
    d = {"text": "机密 CONFIDENTIAL", "font_size": 120, "color": (192, 0, 0),
         "angle": 30, "transparency": 0.6, "scale": 1.0,
         "offset_x": 0.0, "offset_y": 0.0}
    d.update(over)
    return {"text": d}


def _snapshot(path):
    """把 output 的“用户可见内容”抓出来，用于补回后逐项比对。

    只记录**非水印**的用户资产：正文文字、字体加粗/字号、表格、正文图片数量。
    水印图形本身不在比对范围内（它本来就期望被重插）。
    """
    doc = Document(path)
    body = []
    for p in doc.paragraphs:
        runs = [(r.text, bool(r.bold), r.font.size.pt if r.font.size else None)
                for r in p.runs]
        body.append((p.text, runs))
    tables = [[[c.text for c in row.cells] for row in t.rows] for t in doc.tables]
    imgs = len(doc.inline_shapes)
    return {"body": body, "tables": tables, "inline_images": imgs}


def _hf_snapshot(doc):
    """页眉/页脚快照：行内图片数、浮动 drawing 数、文字。

    用来验证“页眉里的普通图片（Logo）没有被守护补回流程破坏”。
    """
    q = engine_docx.qn
    snap = []
    for part in (doc.sections[0].header.part, doc.sections[0].footer.part):
        el = part._element
        snap.append({
            "inline": sum(1 for _ in el.iter(q("wp:inline"))),
            "drawing": sum(1 for _ in el.iter(q("w:drawing"))),
            "text": [p.text for p in el.findall(q("w:p"))],
        })
    return snap


def _full_snapshot(path):
    doc = Document(path)
    snap = _snapshot(path)
    snap["header_footer"] = _hf_snapshot(doc)
    return snap


def test_preserve_existing_keeps_user_body_table_image_format():
    """preserve_existing=True 单独使用：用户资产必须原样留下，只补水印。"""
    tmp = tempfile.mkdtemp()
    src = os.path.join(tmp, "src.docx")
    out = os.path.join(tmp, "out.docx")
    _make_source(src)

    core.insert_watermark(src, ["text"], output_path=out, **_text_opts())
    assert core.watermark_count(out) > 0

    # ---- 用户开始编辑输出文件（守护的目标文件）----
    doc = Document(out)
    doc.paragraphs[0].runs[0].text = "我改过的第一段"
    doc.paragraphs[1].text = "新增的一整段"
    doc.add_paragraph("用户在输出文件里追加的第三段")
    t = doc.tables[0]
    t.cell(0, 0).text = "被用户改成 XYZ"
    t.add_row().cells[0].text = "用户加的行"
    doc.add_picture(io.BytesIO(_png_bytes((200, 30, 30))), width=Pt(60))
    doc.save(out)

    before = _snapshot(out)

    # ---- 用户把水印删掉（这就是守护要响应的场景）----
    engine_docx.clear_watermark(out, kinds=["text"])

    # ---- 守护补回 ----
    res = core.insert_watermark(src, ["text"], output_path=out,
                                preserve_existing=True, **_text_opts())
    assert res.get("ok")
    assert core.watermark_count(out) > 0, "补回后水印应重新出现"

    after = _snapshot(out)
    assert after == before, (
        "preserve_existing 补回后用户资产发生变化——说明发生了覆盖。\n"
        f"before={before}\n\nafter={after}"
    )
    print("[review1] preserve_existing 单独使用：正文/表格/图片/段落格式均完整保留")


def test_watchdog_reinsert_roundtrip_preserves_user_edits():
    """走真实 WatermarkWatchdog._reinsert()，验证“删水印→自动补回”不覆盖用户修改。"""
    tmp = tempfile.mkdtemp()
    src = os.path.join(tmp, "src.docx")
    out = os.path.join(tmp, "out.docx")
    _make_source(src)

    core.insert_watermark(src, ["text"], output_path=out, **_text_opts())

    # 用户在输出文件上的修改
    doc = Document(out)
    doc.paragraphs[0].runs[0].text = "用户改过的标题"
    doc.add_paragraph("用户在守护运行期间追加的段落")
    doc.tables[0].cell(1, 1).text = "用户改过的单元格"
    doc.save(out)
    before = _snapshot(out)

    wd = wd_mod.WatermarkWatchdog(src, out, ["text"], _text_opts(), interval=0.2)
    wd.expected = core.watermark_count(out)

    # 1) 水印被整体删除
    engine_docx.clear_watermark(out, kinds=["text"])
    assert core.watermark_count(out) == 0

    # 2) 守护补回
    wd._reinsert()
    assert core.watermark_count(out) > 0

    # 3) 只删掉其中一份（加固/平铺场景）
    wd._reinsert()
    wd._reinsert()

    after = _snapshot(out)
    assert after == before, (
        "守护补回后用户修改被覆盖。\n"
        f"before={before}\n\nafter={after}"
    )
    print("[review1] 守护 _reinsert 往返：用户修改 100% 保留，水印补回")


def test_watchdog_full_user_scenario_end_to_end():
    """按实际使用场景走一遍，逐条核对用户列出的验收点。

    1) 造带正文的 docx  2) 用 WriteMark 加水印
    3) 用户修改 output 正文（新增一段 + 改一段）
    4) 模拟用户删除水印  5) 触发 watchdog._reinsert()
    6) 验收：水印恢复 / 新增正文在 / 原正文在 / 表格未变 / 图片与页眉页脚未破坏
    """
    tmp = tempfile.mkdtemp()
    src = os.path.join(tmp, "src.docx")
    out = os.path.join(tmp, "out.docx")
    _make_source(src)                                     # 1)
    core.insert_watermark(src, ["text"], output_path=out, **_text_opts())   # 2)
    assert core.watermark_count(out) > 0

    # 页眉放一张普通图片（Logo）、页脚放一段说明文字
    doc = Document(out)
    doc.sections[0].header.paragraphs[0].add_run().add_picture(
        io.BytesIO(_png_bytes((7, 7, 7))))
    doc.sections[0].footer.paragraphs[0].text = "页脚：内部资料"
    doc.save(out)

    doc = Document(out)
    doc.add_paragraph("用户在 output 里新增的段落")          # 3) 新增一段
    doc.paragraphs[0].runs[0].text = "原始正文第一段（用户改过）"
    doc.tables[0].cell(1, 1).text = "用户改过的单元格"
    doc.save(out)
    before = _full_snapshot(out)

    engine_docx.clear_watermark(out, kinds=["text"])      # 4) 用户删水印
    assert core.watermark_count(out) == 0, "前置条件：水印应被删掉"

    wd = wd_mod.WatermarkWatchdog(src, out, ["text"], _text_opts(), interval=0.2)
    wd._reinsert()                                        # 5)

    # 6) 逐条验收
    assert core.watermark_count(out) > 0, "水印必须恢复"

    doc2 = Document(out)
    texts = [p.text for p in doc2.paragraphs]
    assert "用户在 output 里新增的段落" in texts, "用户新增的正文段落必须还在"
    assert "原始正文第一段（用户改过）" in texts, "原正文（被用户改过）必须还在"
    assert "原始正文第二段" in texts, "未被改动的原正文也必须还在"
    assert doc2.tables[0].cell(1, 1).text == "用户改过的单元格", "表格内容必须没变"
    assert len(doc2.tables[0].rows) == 2, "表格行数必须没变"

    after = _full_snapshot(out)
    assert after == before, (
        "守护补回后用户资产发生变化——说明发生了覆盖。\n"
        f"before={before}\n\nafter={after}")
    print("[review3] watchdog 全场景：水印恢复 + 用户新增/原正文/表格/页眉页脚全部保留")


def test_watchdog_partial_deletion_reinserts_without_touching_user_edits():
    """用户只删掉「一部分」水印（平铺模式）时，补齐不能动用户的正文修改。

    守护的 _loop 里部分删除走的是另一条分支（cnt < expected），
    这条分支同样必须保证 preserve_existing 生效。
    """
    tmp = tempfile.mkdtemp()
    src = os.path.join(tmp, "src.docx")
    out = os.path.join(tmp, "out.docx")
    _make_source(src)
    tile_opts = _text_opts()
    tile_opts.update({"tile": True, "tile_rows": 2, "tile_cols": 2})   # 平铺参数是顶层键
    core.insert_watermark(src, ["text"], output_path=out, **tile_opts)

    doc = Document(out)
    doc.paragraphs[0].runs[0].text = "用户改过的第一段"
    doc.add_paragraph("用户新增的段落")
    doc.save(out)
    before = _full_snapshot(out)
    expected = core.watermark_count(out)
    assert expected >= 2, f"平铺模式应写入多份水印，实际 {expected}"

    # 模拟用户只删掉其中一份：直接摘掉一个带标记的 drawing
    doc = Document(out)
    hdr_el = doc.sections[0].header._element
    marked = [d for d in hdr_el.iter(engine_docx.qn("w:drawing"))
              if engine_docx._mark_of(d) is not None]
    assert len(marked) == expected, "页眉里应能找到全部本工具水印"
    engine_docx._detach_drawing(marked[0])
    doc.save(out)
    assert core.watermark_count(out) == expected - 1

    wd = wd_mod.WatermarkWatchdog(src, out, ["text"], tile_opts,
                                  interval=0.2, expected=expected)
    wd._reinsert()

    assert core.watermark_count(out) == expected, "部分删除后应被补齐到期望份数"
    doc2 = Document(out)
    assert "用户改过的第一段" in [p.text for p in doc2.paragraphs]
    assert "用户新增的段落" in [p.text for p in doc2.paragraphs]
    assert _full_snapshot(out)["body"] == before["body"], \
        "部分删除补齐后，用户的正文修改与表格内容不应变化"
    print(f"[review3] watchdog 部分删除（{expected-1}/{expected}）补齐，正文零改动")


def test_watchdog_reinsert_never_pulls_source_content_into_output():
    """硬约束：watchdog 补回时绝不能把 source.docx 的内容搬进 output.docx。"""
    tmp = tempfile.mkdtemp()
    src = os.path.join(tmp, "src.docx")
    out = os.path.join(tmp, "out.docx")
    _make_source(src)
    core.insert_watermark(src, ["text"], output_path=out, **_text_opts())

    # 让 source 与 output 正文明显不同
    doc = Document(src)
    doc.add_paragraph("【只应存在于 source 的内容】")
    doc.save(src)
    doc = Document(out)
    doc.add_paragraph("【只应存在于 output 的内容】")
    doc.save(out)

    wd = wd_mod.WatermarkWatchdog(src, out, ["text"], _text_opts(), interval=0.2)
    engine_docx.clear_watermark(out, kinds=["text"])
    wd._reinsert()

    texts = [p.text for p in Document(out).paragraphs]
    assert "【只应存在于 output 的内容】" in texts, "output 原有内容必须保留"
    assert "【只应存在于 source 的内容】" not in texts, (
        "watchdog 补回把 source.docx 的正文搬进了 output.docx——"
        "preserve_existing 被破坏，用户在输出文件上的修改会被静默覆盖。"
    )
    print("[review3] watchdog 补回不从 source 恢复正文，preserve_existing 未被破坏")


def test_watchdog_output_deleted_rebuilds_from_source():
    """输出文件被整体删除时，守护应当从 source 重新生成（这条路径本来就该重建）。"""
    tmp = tempfile.mkdtemp()
    src = os.path.join(tmp, "src.docx")
    out = os.path.join(tmp, "out.docx")
    _make_source(src)

    core.insert_watermark(src, ["text"], output_path=out, **_text_opts())

    # 用户在 output 上的修改，随后 output 被整体删除
    doc = Document(out)
    doc.add_paragraph("这段修改会随 output 删除一起消失，属预期")
    doc.save(out)
    os.remove(out)

    wd = wd_mod.WatermarkWatchdog(src, out, ["text"], _text_opts(), interval=0.2)
    wd._reinsert()

    assert os.path.exists(out), "output 被删后应重新生成"
    assert core.watermark_count(out) > 0
    doc2 = Document(out)
    texts = [p.text for p in doc2.paragraphs]
    assert "这段修改会随 output 删除一起消失，属预期" not in texts, (
        "output 被整体删除后再重建，是从 source 生成的；上面对 output 的临时修改"
        "本就应该消失——这不是 preserve_existing 的覆盖范围。"
    )
    print("[review1] output 被整体删除 → 从 source 重建，行为符合预期")


def test_preserve_existing_never_reads_source_content():
    """preserve_existing=True 时，output 的内容不应被 source 的内容“反向覆盖”。"""
    tmp = tempfile.mkdtemp()
    src = os.path.join(tmp, "src.docx")
    out = os.path.join(tmp, "out.docx")
    _make_source(src)
    core.insert_watermark(src, ["text"], output_path=out, **_text_opts())

    # 改 source（模拟“源文档后来又改了”），让 output 与 source 明显不同
    doc = Document(src)
    doc.add_paragraph("源文档后来新增的内容（不应被搬进 output）")
    doc.save(src)

    doc = Document(out)
    doc.add_paragraph("output 独有的内容（必须留下）")
    doc.save(out)
    before = _snapshot(out)

    core.insert_watermark(src, ["text"], output_path=out,
                          preserve_existing=True, **_text_opts())

    after = _snapshot(out)
    assert after == before, "output 内容不应随 source 一起变"
    doc2 = Document(out)
    assert "源文档后来新增的内容（不应被搬进 output）" not in [p.text for p in doc2.paragraphs]
    print("[review1] preserve_existing 不回读 source 内容，output 保持独立")


def test_preserve_existing_when_source_is_locked():
    """source 被 Word/WPS 占用时，补回仍应作用于已有 output（此时根本不需要读 source）。"""
    tmp = tempfile.mkdtemp()
    src = os.path.join(tmp, "src.docx")
    out = os.path.join(tmp, "out.docx")
    _make_source(src)
    core.insert_watermark(src, ["text"], output_path=out, **_text_opts())

    # 用户在 output 上改点东西
    doc = Document(out)
    doc.add_paragraph("用户在 output 上的修改")
    doc.save(out)

    # 模拟“其它程序（Word/WPS）持有 source 的独占写句柄”。
    # 注：Python 自带的 open() 在 Windows 上用 _SH_DENYNO，同进程里再开一次不会失败，
    # 因此这里替换占用检测函数，模拟的是「检测结果为被占用」这一事实本身。
    orig = core.is_file_locked
    src_abs = os.path.abspath(src)

    def _fake_locked(path):
        return True if os.path.abspath(path) == src_abs else orig(path)

    core.is_file_locked = _fake_locked
    try:
        engine_docx.clear_watermark(out, kinds=["text"])
        assert core.watermark_count(out) == 0

        # 守护补回：output 存在、且我们只读写 output，应当成功
        core.insert_watermark(src, ["text"], output_path=out,
                              preserve_existing=True, **_text_opts())
        docs = Document(out)
        assert "用户在 output 上的修改" in [p.text for p in docs.paragraphs]
        assert core.watermark_count(out) > 0, (
            "source 被占用时，只要 output 存在就不该连补回一起失败——"
            "preserve_existing 的语义正是“只操作现有 output”。"
        )
    finally:
        core.is_file_locked = orig
    print("[review1] source 被占用时，preserve_existing 仍能补回 output 水印")


def test_watchdog_reinsert_preserves_user_behinddoc_header_graphic():
    """守护补回时的“先清”步骤，不应删掉用户在页眉里自己的 behindDoc 图形。"""
    tmp = tempfile.mkdtemp()
    src = os.path.join(tmp, "src.docx")
    out = os.path.join(tmp, "out.docx")
    _make_source(src)
    core.insert_watermark(src, ["text"], output_path=out, **_text_opts())

    # 用户在页眉里放一张自己的背景图（behindDoc=1、且**没有**本工具私有标记）
    doc = Document(out)
    hdr = doc.sections[0].header
    rId, _ = hdr.part.get_or_add_image(io.BytesIO(_png_bytes((1, 100, 200))))
    drawing = engine_docx._make_drawing(rId, 500000, 500000, 0, "用户页眉背景图", "",
                                        900, 0, 0)
    engine_docx._add_drawing_to_part(hdr, drawing)
    doc.save(out)

    wd = wd_mod.WatermarkWatchdog(src, out, ["text"], _text_opts(), interval=0.2)
    engine_docx.clear_watermark(out, kinds=["text"])
    wd._reinsert()

    doc2 = Document(out)
    hdr2 = doc2.sections[0].header._element
    kept = 0
    for dr in hdr2.iter(engine_docx.qn("w:drawing")):
        anchor = dr.find(".//" + engine_docx.qn("wp:anchor"))
        behind = anchor is not None and anchor.get("behindDoc") == "1"
        mark = engine_docx._mark_of(dr)
        if behind and mark is None:
            kept += 1
    assert kept == 1, (
        "用户在页眉里自己的 behindDoc 图形被守护补回流程连带删除了——"
        "这违背了 preserve_existing「保留用户修改」的承诺。"
        f"(kept={kept})"
    )
    print("[review1/2] 守护补回保留用户页眉 behindDoc 图形")
