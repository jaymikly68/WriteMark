"""
二次安全审查（commit 84a773e 之后）—— 主题 3：word_fonts 的 TTC face 索引。

84a773e 给 build_font_map() 加了副作用：把 TTC 内的 face 索引存进模块级
`_INDEX_MAP`，再由 font_index_of() 交给 Pillow 的 ImageFont.truetype(index=)。
本文件回答两个问题：

A. `_INDEX_MAP` 在多次 build_font_map() 之后会不会残留旧字体的索引？
B. `index=N` 是否**真的**加载了第 N 个 face？（只检查 font_index_of 的返回值
   是不够的——返回值对、Pillow 用错 face，等于没修。）

B 的验证手法：Pillow 对 TTC 的 `ImageFont.truetype(path, size, index=i).getname()`
会返回**该 face 自己的家族名**（如 msyh.ttc 的第 0 面是 'Microsoft YaHei'、
第 1 面是 'Microsoft YaHei UI'）。据此可以断言“选中的确实是那一面”。
"""
from __future__ import annotations

import contextlib
import os
import sys
import types

import pytest
from PIL import ImageFont

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from watermark_tool import engine_docx, word_fonts

WIN_FONTS = os.path.join(os.environ.get("SystemRoot", "C:/Windows"), "Fonts")
_MSYH = os.path.join(WIN_FONTS, "msyh.ttc")


class _FakeReg:
    """假 winreg：只提供 build_font_map() 用到的三个东西。"""

    def __init__(self, entries):
        self.entries = list(entries)

    HKEY_LOCAL_MACHINE = 0

    def OpenKey(self, root, sub):            # noqa: N802（保持原生命名以便替换）
        return "fake-key"

    def EnumValue(self, key, index):         # noqa: N802
        if index >= len(self.entries):
            raise OSError(259, "no more data")   # ERROR_NO_MORE_ITEMS
        name, value, _t = self.entries[index]
        return name, value, 1


@contextlib.contextmanager
def _install_fake_reg(monkeypatch, entries):
    monkeypatch.setitem(sys.modules, "winreg", _FakeReg(entries))
    saved_cache = word_fonts._MAP_CACHE
    word_fonts._MAP_CACHE = None
    try:
        yield
    finally:
        # 先摘掉假注册表，再重建缓存，否则会用假数据还原成脏缓存
        monkeypatch.undo()
        word_fonts._MAP_CACHE = saved_cache
        word_fonts.refresh_font_map()      # 还原到真实注册表


def test_index_map_no_stale_residue_between_builds(monkeypatch):
    """A. 连续多次 build_font_map()，_INDEX_MAP 不应残留已卸载字体的索引。"""
    entries = [
        ("SimSun & NSimSun (TrueType)", "simsun.ttc,0", 1),
        ("RemovedFont (TrueType)", "removedfont.ttf,3", 1),
    ]
    with _install_fake_reg(monkeypatch, entries):
        m1 = word_fonts.build_font_map()
        assert m1["RemovedFont (TrueType)"].endswith("removedfont.ttf")
        assert word_fonts._INDEX_MAP["RemovedFont (TrueType)"] == 3

    # 第二次构建：RemovedFont 被卸载
    with _install_fake_reg(monkeypatch, [("SimSun & NSimSun (TrueType)", "simsun.ttc,0", 1)]):
        m2 = word_fonts.build_font_map()
        assert "RemovedFont (TrueType)" not in m2
        assert "RemovedFont (TrueType)" not in word_fonts._INDEX_MAP, (
            "_INDEX_MAP 残留了已卸载字体的索引"
        )
        assert set(word_fonts._INDEX_MAP) == set(m2), (
            "_INDEX_MAP 与 mapping 的键应始终一致"
        )

    print("[review3] _INDEX_MAP 每次 build_font_map() 都整体重建，无残留")


def test_index_map_matches_mapping_keys(monkeypatch):
    """A（补充）：索引表与路径表的键必须一一对应。"""
    entries = [
        ("SimSun & NSimSun (TrueType)", "simsun.ttc,0", 1),
        ("WeirdNoIndex (TrueType)", "some.ttf", 1),
        ("BrokenIndex (TrueType)", "weird.ttc,notanumber", 1),
        ("EmptyValue (TrueType)", "", 1),
    ]
    with _install_fake_reg(monkeypatch, entries):
        m = word_fonts.build_font_map()
        assert set(word_fonts._INDEX_MAP) == set(m), "_INDEX_MAP 与 mapping 键不一致"
        # 无法解析的索引按 0 处理（不应抛异常）
        assert word_fonts._INDEX_MAP["BrokenIndex (TrueType)"] == 0
        assert "EmptyValue (TrueType)" not in m
    print("[review3] 索引表与路径表键一致；非法索引回退 0；空值条目被跳过")


def test_ttc_face_1_loads_when_index_is_supplied():
    """B（前半，已验证）：一旦索引确实是 1，Pillow 就真的会载入第 1 个 face。

    这条不依赖 font_index_of 的返回值，直接验证「index 参数 → 实际 face」，
    从而把结论精确地定位到「机制本身可用，问题出在谁把索引算成 0」。
    """
    if not os.path.exists(_MSYH):
        pytest.skip(f"本机缺少 {_MSYH}，无法验证真实 TTC face")
    f0 = ImageFont.truetype(_MSYH, 40, index=0)
    f1 = ImageFont.truetype(_MSYH, 40, index=1)
    name0, name1 = f0.getname()[0], f1.getname()[0]
    assert name0 == "Microsoft YaHei", f"index=0 应载入 {name0!r}"
    assert name1 == "Microsoft YaHei UI", f"index=1 应载入 {name1!r}"
    assert name0 != name1, "两个 face 应确实不同，否则这条验证没有意义"
    print(f"[review3] 真实 TTC 校验：index=0 -> {name0!r}，index=1 -> {name1!r}（确为不同 face）")


@pytest.mark.xfail(reason="已知技术债：font_index_of() 对命中别名表(_ALIAS_FILES)的名字"
                          "直接 return 0，绕过了 _INDEX_MAP。别名表覆盖了 微软雅黑/宋体/黑体…"
                          "等主要字体，于是 84a773e 的 TTC 索引修复对它们完全不生效；"
                          "「微软雅黑 UI / Light」这类变体因此仍会渲染成第 0 面。", strict=False)
def test_alias_table_does_not_bypass_ttc_index(monkeypatch):
    """B（后半，未通过）：注册表明明记了 "msyh.ttc,1"，font_index_of 仍应返回 1。"""
    if not os.path.exists(_MSYH):
        pytest.skip(f"本机缺少 {_MSYH}，无法验证真实 TTC face")

    with _install_fake_reg(monkeypatch,
                           [("Microsoft YaHei UI (TrueType)", "msyh.ttc,1", 1)]):
        word_fonts.build_font_map()
        # 注册表里确实记了索引 1
        assert word_fonts._INDEX_MAP["Microsoft YaHei UI (TrueType)"] == 1
        # 但 font_index_of 因为是别名表条目而返回 0
        assert word_fonts.font_index_of("Microsoft YaHei UI") == 1, (
            "别名表条目（微软雅黑/宋体/黑体…）会绕过 _INDEX_MAP 直接返回 0，"
            "导致 TTC 分面索引对常用字体不生效"
        )


def test_real_machine_face_selection_for_multi_face_ttc():
    """B（真实机器）：本机注册表若没写索引，多 face TTC 的分面选择是否仍然正确。"""
    word_fonts.refresh_font_map()
    path = word_fonts.resolve_font_path("微软雅黑 UI")
    if not path or not path.lower().endswith((".ttc", ".ttf")):
        pytest.skip("未解析到 TTC，跳过")
    if not os.path.exists(path):
        pytest.skip(f"{path} 不存在，跳过")

    idx = word_fonts.font_index_of("微软雅黑 UI")
    face = ImageFont.truetype(path, 40, index=idx).getname()[0]
    faces = [ImageFont.truetype(path, 40, index=i).getname()[0]
             for i in range(2) if _try_face(path, i)]
    print(f"[review3] 本机 {os.path.basename(path)}：索引={idx}，"
          f"选中 face={face!r}，该文件各 face={faces}")
    # 这里只打印事实，断言留给下面两条针对性用例


def _try_face(path, i):
    try:
        ImageFont.truetype(path, 40, index=i)
        return True
    except Exception:
        return False


@pytest.mark.xfail(reason="已知技术债：本机注册表把多 face TTC 记成单纯的 'msyh.ttc'（无 ,N 索引），"
                          "split_alias_entries 又把它拆出 'Microsoft YaHei UI' 这个显示名，"
                          "但索引只能取到 0，于是选 UI 变体时实际渲染的是常规 face。", strict=False)
def test_real_machine_ui_variant_uses_ui_face():
    """「微软雅黑 UI」应真正载入 UI 那个 face（本机注册表无索引时目前做不到）。"""
    word_fonts.refresh_font_map()
    path = word_fonts.resolve_font_path("微软雅黑 UI")
    if not path or not os.path.exists(path):
        pytest.skip("未解析到字体，跳过")
    idx = word_fonts.font_index_of("微软雅黑 UI")
    face = ImageFont.truetype(path, 40, index=idx).getname()[0]
    assert face == "Microsoft YaHei UI", (
        f"选了「微软雅黑 UI」，实际载入的 face 是 {face!r}（索引 {idx}）"
    )


def test_render_text_png_passes_face_index_into_pillow(monkeypatch):
    """B（端到端，本轮新增）：把「显示名 → font_index_of → ImageFont.truetype(index=)」

    整条链路钉死。前面几条只验证了「机制可用」与「索引表干净」，
    这里验证的是**交接**：渲染水印时传给 Pillow 的 index 到底是不是
    font_index_of() 算出来的那个值，文件参数是不是那个 TTC 的路径。
    """
    if not os.path.exists(_MSYH):
        pytest.skip(f"本机缺少 {_MSYH}，无法验证真实 TTC face")

    # 注册表明文写着 "msyh.ttc,1"，并额外保留常规面条目以免路径解析落空
    with _install_fake_reg(monkeypatch, [
        ("Microsoft YaHei (TrueType)", "msyh.ttc,0", 1),
        ("ConfTtcUI (TrueType)", "msyh.ttc,1", 1),
    ]):
        word_fonts.build_font_map()
        idx = word_fonts.font_index_of("ConfTtcUI (TrueType)")
        path = word_fonts.resolve_font_path("ConfTtcUI (TrueType)")
        assert idx == 1, f"注册表记了 msyh.ttc,1，font_index_of 应返回 1，实际 {idx}"
        assert path and os.path.samefile(path, _MSYH), f"解析到的字体路径异常: {path}"

        # 拦下 Pillow 的 truetype，记录每一次调用真正收到的参数
        calls = []

        def _spy_truetype(fontpath, size, *a, **k):
            calls.append((fontpath, size, k.get("index", 0)))
            return _real_truetype(fontpath, size, *a, **k)

        # 兜底字体改成非 TTC（arial），这样记录里命中 msyh.ttc 的调用只可能是
        # 「中文那条链路」本身，不会把 find_font() 的兜底加载混进来
        arial = os.path.join(WIN_FONTS, "arial.ttf")
        if not os.path.exists(arial):
            arial = None
        monkeypatch.setattr(engine_docx, "find_font", lambda *a, **k: arial)

        _real_truetype = ImageFont.truetype
        monkeypatch.setattr(ImageFont, "truetype", _spy_truetype)
        try:
            engine_docx.render_text_png("机密", 40, (0, 0, 0), 255,
                                        cn_font_name="ConfTtcUI (TrueType)")
        finally:
            monkeypatch.undo()

    assert calls, "渲染时应当有人调用 ImageFont.truetype"
    hit = [c for c in calls if os.path.samefile(c[0], _MSYH)]
    assert hit, f"中文水印未使用 msyh.ttc 渲染，实际调用: {calls}"
    for _p, _s, i in hit:
        assert i == idx, (
            f"传给 Pillow 的 index={i} 与 font_index_of() 算出的 {idx} 不一致——"
            "TTC face 索引没有被正确透传，多 face 字体会退化成第 0 面"
        )
    print(f"[review3] 端到端透传：font_index_of('ConfTtcUI')={idx} → "
          f"ImageFont.truetype(path={os.path.basename(_MSYH)}, index={idx})")


def test_index_map_and_cache_are_coupled_characterization(monkeypatch):
    """技术债固化：直接调 build_font_map() 会刷新 _INDEX_MAP，但不会刷新 _MAP_CACHE。

    结果就是「路径表是旧的、索引表是新的」这种半新半旧状态；此时 font_index_of()
    会按旧显示名去查新索引表，查不到就静默返回 0（不会崩，但会退回第 0 面）。
    这里把当前行为固定下来，防止日后重构时不小心放大这个不一致。
    """
    entries = [("TtcFont & TtcFontUI (TrueType)", "msyh.ttc,1", 1)]
    with _install_fake_reg(monkeypatch, entries):
        word_fonts.build_font_map()                 # 填好 _INDEX_MAP
        word_fonts._MAP_CACHE = {"TtcFontUI (TrueType)": "old.ttc"}  # 模拟半新半旧
        idx = word_fonts.font_index_of("TtcFontUI (TrueType)")
        assert idx == 0, (
            "当前实现在『索引表新、路径表旧』时会静默返回 0；"
            "若这条断言开始失败，说明耦合问题已被修好（值得检查是否引入回归）"
        )
    print("[review3] 已知耦合：build_font_map() 与 _MAP_CACHE 需配套刷新，否则索引静默退化为 0")
