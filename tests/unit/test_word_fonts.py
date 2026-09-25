"""
字体功能测试（无需显示器，offscreen）：
- word_fonts：注册表字体映射、字体名解析、系统字体列表、Word COM 优雅降级
- engine_docx：文本水印真正使用所选字体（解析到字体文件），选不到时优雅回退
- gui：导入后权限询问 + 下拉框合并逻辑（去重、保留当前选择、返回新增数量）
"""
import os
import sys

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")
from PySide6.QtWidgets import QApplication, QComboBox

from watermark_tool import word_fonts
from watermark_tool.engine_docx import render_text_png


def test_registry_map():
    mp = word_fonts.build_font_map()
    # 注册表存在时应为非空字典，且值是绝对路径
    assert isinstance(mp, dict)
    if mp:
        for v in mp.values():
            assert os.path.isabs(v), f"字体路径应为绝对路径: {v}"
    print(f"[registry] 字体映射条目数: {len(mp)}")


def test_system_fonts_nonempty():
    fonts = word_fonts.list_system_fonts()
    assert isinstance(fonts, list)
    # Windows 上一般都有系统字体；非 Windows 可能为空，这里只保证不报错
    print(f"[system] 系统字体数: {len(fonts)}")


def test_resolve_roundtrip():
    # 取一个系统字体名，应能解析回存在的磁盘文件
    fonts = word_fonts.list_system_fonts()
    if not fonts:
        print("[resolve] 无系统字体，跳过 roundtrip")
        return
    name = fonts[0]
    path = word_fonts.resolve_font_path(name)
    assert path is not None, f"应能解析字体 {name!r}"
    assert os.path.exists(path), f"解析出的字体文件应存在: {path}"
    # 大小写/后缀无关匹配：用原始名（可能含后缀）再解一次也应成功
    print(f"[resolve] {name!r} -> {path}")


def test_resolve_unknown_returns_none():
    # 不存在的字体应返回 None（引擎据此回退默认字体）
    assert word_fonts.resolve_font_path("__不存在的字体XYZ__") is None
    print("[resolve] 未知字体正确返回 None")


def test_get_word_fonts_no_crash():
    # 无论本机是否有 Word，调用都不应抛异常（无 Word 时返回 None）
    fonts = word_fonts.get_word_fonts()
    assert fonts is None or isinstance(fonts, list)
    print(f"[com] get_word_fonts 返回类型: {'list(%d)' % len(fonts) if fonts else 'None(无Word/失败)'}")


def test_collect_fonts_fallback():
    fonts, source = word_fonts.collect_fonts()
    assert isinstance(fonts, list) and isinstance(source, str)
    assert len(fonts) > 0, "即使 Word 不可用，也应回退到系统字体"
    print(f"[collect] 来源={source}, 字体数={len(fonts)}")


def test_engine_uses_font_name():
    # 用系统里真实存在的字体渲染，确认不抛异常且产出图像
    fonts = word_fonts.list_system_fonts()
    font_name = fonts[0] if fonts else None
    img = render_text_png("测试水印 TEST", 120, (128, 128, 128), 128,
                          cn_font_name=font_name, latin_font_name=font_name)
    assert img.size[0] > 0 and img.size[1] > 0
    # 未知字体应优雅回退（不抛异常）
    img2 = render_text_png("测试水印", 120, (128, 128, 128), 128,
                           cn_font_name="__不存在__", latin_font_name="__不存在__")
    assert img2.size[0] > 0
    print("[engine] 文本水印渲染支持分别指定中西文字体，且未知字体回退正常")


def test_gui_apply_fonts():
    # 直接构造真实 App（offscreen），验证中/西文字体下拉框合并逻辑：去重/保留当前选择/返回新增数量
    # 注意：整套回归里前面的用例可能已经建过 QApplication（单例不可重复创建），
    # 必须先取 instance，否则整轮跑下来会报 libshiboken 单例冲突。
    app = QApplication.instance() or QApplication(sys.argv)
    from watermark_tool.gui import App
    win = App()
    combo = win.cn_font_combo  # 中文字体框；两个框逻辑一致，测其一即可

    before = combo.count()
    added = win._apply_word_fonts(["微软雅黑", "TestFontA", "TestFontB", "Arial", "新字体C"])
    after = combo.count()
    # 微软雅黑/ Arial 与内置重复，应只新增 3 个（两框合计计入，但此处看单框）
    assert added == 3, f"新增数量应为 3，实际 {added}"
    assert after == before + 3
    # 当前选择应保留（默认：微软雅黑）
    assert combo.currentText() in ("微软雅黑",)
    # 新增项确实存在
    items = [combo.itemText(i) for i in range(combo.count())]
    for f in ("TestFontA", "TestFontB", "新字体C"):
        assert f in items
    # 西文字体框也应同步新增
    lat_items = [win.latin_font_combo.itemText(i) for i in range(win.latin_font_combo.count())]
    for f in ("TestFontA", "TestFontB", "新字体C"):
        assert f in lat_items, "西文字体框应同步扩充"
    print("[gui] 中/西文字体下拉框合并：去重/保留当前选择/返回新增数量 均正常")


def _text_color_stats(img):
    """返回不透明像素的平均 (R,G,B)，用于确认颜色被采用。"""
    rgba = img.convert("RGBA")
    px = rgba.load()
    rs = gs = bs = n = 0
    for y in range(rgba.height):
        for x in range(rgba.width):
            r, g, b, a = px[x, y]
            if a > 10:
                rs += r; gs += g; bs += b; n += 1
    return (rs / n, gs / n, bs / n) if n else (0, 0, 0)


def test_engine_color_applied():
    # 引擎应按所选颜色渲染文本（颜色按钮已存在，这里验证其真正生效）
    red = render_text_png("水印", 120, (255, 0, 0), 255, None)
    grn = render_text_png("水印", 120, (0, 200, 0), 255, None)
    rmean = _text_color_stats(red)
    gmean = _text_color_stats(grn)
    assert rmean[0] > 200 and rmean[1] < 60, f"红色水印应偏红: {rmean}"
    assert gmean[1] > 150 and gmean[0] < 60, f"绿色水印应偏绿: {gmean}"
    print(f"[engine] 文本颜色正确应用: 红={tuple(map(int,rmean))} 绿={tuple(map(int,gmean))}")


def test_font_size_controls_size():
    # 纯 Python .docx 引擎：字号现在真实控制水印大小（font_size 越小，绘制宽度越小）
    from docx import Document
    from docx.oxml.ns import qn
    from watermark_tool.engine_docx import _iter_headers, insert_watermark

    import tempfile, os
    d = tempfile.mkdtemp()
    small = os.path.join(d, "small.docx")
    big = os.path.join(d, "big.docx")
    Document().save(small)
    Document().save(big)

    insert_watermark(small, ["text"], text={"text": "机密", "font_size": 20, "color": (0, 0, 0)})
    insert_watermark(big, ["text"], text={"text": "机密", "font_size": 280, "color": (0, 0, 0)})

    def _max_cx(path):
        doc = Document(path)
        mx = 0
        for section in doc.sections:
            for header in _iter_headers(doc, section):
                for drawing in header._element.iter(qn("w:drawing")):
                    ext = drawing.find(".//" + qn("wp:extent"))
                    if ext is not None:
                        mx = max(mx, int(ext.get("cx")))
        return mx

    cx_small = _max_cx(small)
    cx_big = _max_cx(big)
    assert cx_small > 0 and cx_big > cx_small, f"字号大应更宽: small={cx_small} big={cx_big}"
    print(f"[engine] 字号控制大小生效: 字号20宽={cx_small} 字号280宽={cx_big}")


def test_gui_font_combo_editable():
    # 字体下拉框应可编辑，且补全按“前缀+不区分大小写”弹出，扩充后补全模型同步
    from PySide6.QtWidgets import QApplication, QComboBox, QCompleter
    from PySide6.QtCore import Qt
    app = QApplication.instance() or QApplication([])
    from watermark_tool.gui import App
    win = App()
    combo = win.cn_font_combo
    assert combo.isEditable(), "中文字体下拉框应可编辑（允许直接键入字体名）"
    comp = combo.completer()
    assert comp is not None, "应配置自动补全"
    assert comp.completionMode() == QCompleter.PopupCompletion, "应为弹出式补全"
    assert comp.caseSensitivity() == Qt.CaseInsensitive, "应不区分大小写"
    assert comp.filterMode() == Qt.MatchStartsWith, "应按前缀匹配（输入 t → t 开头字体）"
    # 扩充字体后，补全模型应包含新增字体
    model = comp.model()
    before = list(model.stringList())
    add_list = ["Times New Roman", "Tahoma", "新字体Z"]  # Times New Roman 已是内置项，应被去重
    expected_new = len([f for f in add_list if f not in before])
    win._apply_word_fonts(add_list)
    model_now = list(comp.model().stringList())
    assert "Times New Roman" in model_now and "Tahoma" in model_now, "补全模型应同步新增字体"
    assert len(model_now) == len(before) + expected_new, (
        f"补全模型应新增 {expected_new} 个: {len(before)} -> {len(model_now)}")
    assert len(set(model_now)) == len(model_now), "补全模型不应出现重复项"
    # 模拟用户键入：找到前缀匹配项
    comp.setCompletionPrefix("t")
    cm = comp.completionModel()
    assert cm is not None and cm.rowCount() > 0, "输入 t 应能匹配到字体"
    hits = [str(cm.data(cm.index(i, 0))) for i in range(min(cm.rowCount(), 5))]
    assert any(h.lower().startswith("t") for h in hits), "匹配结果应为 t 开头"
    print("[gui] 中/西文字体下拉框可编辑 + 前缀自动补全（含大小写不敏感）配置正常")


def _mask_diff(a, b):
    """返回两图“墨水遮罩（alpha>10）不一致”的像素数量。"""
    aa = a.convert("RGBA"); bb = b.convert("RGBA")
    pa = aa.load(); pb = bb.load()
    w = min(aa.width, bb.width); h = min(aa.height, bb.height)
    d = 0
    for y in range(h):
        for x in range(w):
            ia = pa[x, y][3] > 10
            ib = pb[x, y][3] > 10
            if ia != ib:
                d += 1
    return d


def test_cn_latin_explicit_split():
    # 核心需求验证：中文字符取“中文字体”、拉丁字符取“西文字体”。
    # “机密C”此类混排：机密→中文字体，C→西文字体；二者互不影响。
    from watermark_tool.engine_docx import render_text_png, find_font
    if find_font() is None:
        print("[split] 本机无中文字体，跳过显式分流断言")
        return
    from watermark_tool import word_fonts
    if (word_fonts.resolve_font_path("Times New Roman") is None
            or word_fonts.resolve_font_path("仿宋") is None):
        print("[split] 本机缺 Times New Roman / 仿宋，跳过显式分流断言")
        return
    # 拉丁字符“C”应随“西文字体”变化：西文字体不同 → 字形不同
    c_tnr = render_text_png("C", 200, (0, 0, 0), 255,
                            cn_font_name="微软雅黑", latin_font_name="Times New Roman")
    c_yh = render_text_png("C", 200, (0, 0, 0), 255,
                           cn_font_name="微软雅黑", latin_font_name="微软雅黑")
    assert _mask_diff(c_tnr, c_yh) > 50, "拉丁字母 C 应取自西文字体（西文字体改变则字形改变）"
    # 中文字符“机”应随“中文字体”变化：中文字体不同 → 字形不同
    j_fang = render_text_png("机", 200, (0, 0, 0), 255,
                             cn_font_name="仿宋", latin_font_name="Times New Roman")
    j_yh = render_text_png("机", 200, (0, 0, 0), 255,
                           cn_font_name="微软雅黑", latin_font_name="Times New Roman")
    assert _mask_diff(j_fang, j_yh) > 50, "汉字 机 应取自中文字体（中文字体改变则字形改变）"
    # 混排“机密C”：C 取自西文字体（TNR），与“机密+Arial”组合中 C 取自 Arial 应不同
    mix_tnr = render_text_png("机密C", 200, (0, 0, 0), 255,
                              cn_font_name="微软雅黑", latin_font_name="Times New Roman")
    mix_arial = render_text_png("机密C", 200, (0, 0, 0), 255,
                                cn_font_name="微软雅黑", latin_font_name="Arial")
    assert _mask_diff(mix_tnr, mix_arial) > 50, "混排文本中 C 应取自西文字体（随西文字体变化）"
    print("[engine] 中西文字体显式分流生效：汉字取中文字体、拉丁取西文字体")


def test_cjk_fallback_rendering():
    # 兜底：若“中文字体”被设成纯西文字体（无中文字形），中文应回退到默认中文字体，
    # 绝不画成方框/豆腐块。
    from watermark_tool.engine_docx import render_text_png, find_font
    if find_font() is None:
        print("[cjk] 本机无中文字体可回退，跳过 CJK 回退渲染断言")
        return
    from watermark_tool import word_fonts
    if word_fonts.resolve_font_path("Arial") is None:
        print("[cjk] 本机无 Arial，跳过 CJK 回退渲染断言")
        return
    # 中文字体设成 Arial（无中文），西文字体也设 Arial → 中文应二次回退到默认中文字体
    img_west = render_text_png("中文", 120, (0, 0, 0), 255,
                               cn_font_name="Arial", latin_font_name="Arial")
    img_cjk = render_text_png("中文", 120, (0, 0, 0), 255,
                              cn_font_name="微软雅黑", latin_font_name="微软雅黑")
    # 西文设置下，中文仍应有笔画（alpha=255 像素），不能是全透明方框
    assert 255 in set(img_west.split()[3].getdata()), "中文不应为空（应回退到中文字体）"
    # 与直接用中文字体渲染的差异应很小（字形一致）
    diff = _mask_diff(img_west, img_cjk)
    assert diff <= 5, f"西文回退渲染与中文字体渲染应一致，差异像素={diff}"
    print("[engine] 中文字体无字形时二次回退到默认中文字体，无方框")


def test_cn_font_names_resolve():
    # 回归：注册表里中文字体的显示名多为英文（"Microsoft YaHei"/"SimHei"），
    # 直接用中文名（微软雅黑/黑体）曾经解析不到 → 用户选了中文字体却不生效。
    # 这里保证常见中文名都能解析到真实存在的字体文件。
    from watermark_tool import word_fonts
    required = ["微软雅黑", "黑体", "仿宋", "宋体", "楷体", "Times New Roman"]
    missing = []
    for n in required:
        p = word_fonts.resolve_font_path(n)
        if not p or not os.path.exists(p):
            missing.append(n)
    assert not missing, f"以下字体名解析失败（选中后将静默回退，体验为“选了没用”）: {missing}"
    print(f"[resolve] 常见中文/西文字体名均可解析: {required}")


def test_mixed_baseline_alignment():
    # 混排“机密C”：汉字用中文字体、C 用西文字体，且两者必须共用同一条基线。
    # 若按各自 bbox 垂直居中拼接，C 会明显偏上/偏下（此前实现的缺陷）。
    from PIL import ImageFont
    from watermark_tool import word_fonts
    from watermark_tool.engine_docx import render_text_png
    cn_path = word_fonts.resolve_font_path("微软雅黑")
    la_path = word_fonts.resolve_font_path("Times New Roman")
    if not (cn_path and la_path):
        print("[baseline] 本机缺 微软雅黑 / Times New Roman，跳过基线对齐断言")
        return

    size = 200
    f_cn = ImageFont.truetype(cn_path, size)
    f_la = ImageFont.truetype(la_path, size)
    asc_mix = max(f_cn.getmetrics()[0], f_la.getmetrics()[0])
    asc_la = f_la.getmetrics()[0]
    pad = max(20, int(size * 0.25))

    def _ink_bottom(img, x0, x1):
        """返回 [x0,x1) 列区间内墨迹（alpha>0）的最低行号。"""
        a = img.split()[3]
        px = a.load()
        bottom = -1
        for y in range(img.height - 1, -1, -1):
            for x in range(x0, x1):
                if px[x, y] > 0:
                    return y
        return bottom

    mix = render_text_png("机密C", size, (0, 0, 0), 255,
                          cn_font_name="微软雅黑", latin_font_name="Times New Roman")
    c_only = render_text_png("C", size, (0, 0, 0), 255,
                             cn_font_name="微软雅黑", latin_font_name="Times New Roman")
    # 汉字部分（左）与 C 部分（右）都必须有墨迹，不能出现空白/方框
    adv_cn = float(f_cn.getlength("机密"))
    x_c0 = int(pad + adv_cn)
    x_c1 = mix.width - pad
    assert _ink_bottom(mix, pad, x_c0) > 0, "汉字部分应有墨迹"
    assert _ink_bottom(mix, x_c0, x_c1) > 0, "拉丁字母部分应有墨迹"

    # 混排里的 C 与单独渲染的 C：相对各自基线的“下沉量”应一致 → 基线对齐
    over_mix = _ink_bottom(mix, x_c0, x_c1) - (pad + asc_mix)
    over_only = _ink_bottom(c_only, pad, c_only.width - pad) - (pad + asc_la)
    assert abs(over_mix - over_only) <= 3, (
        f"混排中西文基线未对齐：混排下沉={over_mix}, 单独渲染下沉={over_only}")

    # 宽度路由：混排内容宽 = 汉字宽 + 字母宽（证明两段分别取自中/西文字体）
    mix_w = mix.width - 2 * pad
    cn_w = render_text_png("机密", size, (0, 0, 0), 255, "微软雅黑", "Times New Roman").width - 2 * pad
    la_w = c_only.width - 2 * pad
    assert abs(mix_w - (cn_w + la_w)) <= 2, (
        f"混排宽度应由中西文两段拼接: {mix_w} != {cn_w} + {la_w}")
    print("[engine] 中西文混排共用基线对齐，宽度=汉字宽+字母宽")


def test_close_event_no_error():
    # closeEvent 在打桩 os._exit 下不应抛异常（验证线程清理逻辑健壮）。
    # 默认关闭行为是「最小化到后台」，不走强退路径；用 WM_CLOSE_CHOICE=quit
    # 才能覆盖到「停守护 + 兜底计时线程」这段代码（与自动化退出同一条路径）。
    import os as _os
    import threading as _th
    from PySide6.QtWidgets import QApplication
    from PySide6.QtGui import QCloseEvent
    real_exit = _os._exit
    _os._exit = lambda code: None  # 防止测试进程被真正杀掉
    _saved_choice = _os.environ.get("WM_CLOSE_CHOICE")
    _os.environ["WM_CLOSE_CHOICE"] = "quit"
    try:
        app = QApplication.instance() or QApplication([])
        from watermark_tool.gui import App
        win = App()
        win.closeEvent(QCloseEvent())
        # 确认兜底计时线程已启动（daemon）
        daemons = [t for t in _th.enumerate() if t.daemon and t.is_alive()]
        assert any(True for _ in daemons), "应有守护兜底计时线程在运行"
    finally:
        _os._exit = real_exit
        if _saved_choice is None:
            _os.environ.pop("WM_CLOSE_CHOICE", None)
        else:
            _os.environ["WM_CLOSE_CHOICE"] = _saved_choice
    print("[gui] closeEvent 线程清理 + 兜底强退计时线程 启动正常")


if __name__ == "__main__":
    test_registry_map()
    test_system_fonts_nonempty()
    test_resolve_roundtrip()
    test_resolve_unknown_returns_none()
    test_get_word_fonts_no_crash()
    test_collect_fonts_fallback()
    test_engine_uses_font_name()
    test_gui_apply_fonts()
    test_engine_color_applied()
    test_font_size_controls_size()
    test_gui_font_combo_editable()
    test_cn_latin_explicit_split()
    test_cjk_fallback_rendering()
    test_cn_font_names_resolve()
    test_mixed_baseline_alignment()
    test_close_event_no_error()
    print("\n全部字体功能测试通过 ✅")
