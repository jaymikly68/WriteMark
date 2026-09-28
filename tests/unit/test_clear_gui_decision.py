"""
GUI「一键清除水印」(`App._clear()`) 的**真实行为**测试。

不看源码文本，而是真的调用 `_clear()`，捕获它交给 `core.clear_watermark` 的调用，
必要时把捕获到的任务**真的执行一遍**，再检查最终文档状态。

要钉死的两条安全契约（都是修复前的真实缺陷）：

1. 按检测到的类型精确删除
   `{text}` -> kinds=['text']；`{image}` -> kinds=['image']。
   修复前 `types` 不等于 `{"text","image"}` 时 kinds 保持 None，
   掉进 `clear_any=True`，把页眉里用户自己的 behindDoc 图也一起删了。

2. 检测失败 -> 停止清除 + 报告，**绝不**退化成扩大删除范围
   修复前 `except Exception: kinds = None`，同样掉进 `clear_any=True`。

说明：`App` 的构造（整窗 UI + 定时器）对这类决策测试是纯开销，
因此这里用 `App.__new__(App)` 造一个不带 UI 的实例、只填 `_clear()` 真正读到的
几个属性。被调用的仍是**真实的 `App._clear()` 方法本身**。
"""
from __future__ import annotations

import os
import sys
import tempfile

import pytest
from docx import Document

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from PySide6.QtWidgets import QLineEdit

from watermark_tool import core, engine_docx
from watermark_tool import gui as gui_mod

USER_BODY = "用户正文"


# ---------------------------------------------------------------------------
# 测试替身
# ---------------------------------------------------------------------------
_WARNINGS: list[tuple[str, str]] = []


class _FakeMessageBox:
    """替身：只记录调用了什么提示框，不发真弹窗。

    注意：不能用实例方法来记，否则 `instance.warning(tuple)` 会又解析到这个
    替身自身、无限递归。这里统一写进模块级列表。
    """

    @staticmethod
    def warning(parent, title, text, *a, **k):
        _WARNINGS.append((title, str(text)))

    @staticmethod
    def information(parent, title, text, *a, **k):
        _WARNINGS.append((title, str(text)))


class _Recorder:
    """替身：记录 core.clear_watermark 的真实调用参数。"""

    def __init__(self):
        self.calls = []

    def __call__(self, src_path, output_path=None, kinds=None,
                 clear_native_only=False, **kw):
        self.calls.append({"src": src_path, "out": output_path, "kinds": kinds,
                           "clear_native_only": clear_native_only})
        return {"ok": True, "engine": "docx", "removed": 0,
                "kinds": sorted(kinds or [])}


def _stub_app(monkeypatch, file_path, out_path):
    """造一个只带 `_clear()` 所需属性的 App 实例，并拦下任务执行。"""
    app = gui_mod.App.__new__(gui_mod.App)
    app.file_path = file_path
    app.out_edit = QLineEdit(out_path)

    jobs = []
    monkeypatch.setattr(gui_mod.App, "_run", lambda self, job: jobs.append(job))
    return app, jobs


def _apply(app, jobs, monkeypatch):
    """执行 `_clear()` 排队下来的清除任务，返回替身记录到的调用参数。

    跑完**立刻 undo 所有 patch**：测试后面还要用真正的 `core.clear_watermark`
    做端到端验证，否则那次调用也会打到替身上、什么文件都不会生成。
    """
    rec = _Recorder()
    monkeypatch.setattr(core, "clear_watermark", rec)
    try:
        for job in jobs:
            job()
    finally:
        monkeypatch.undo()
    return rec.calls


# ---------------------------------------------------------------------------
# 素材
# ---------------------------------------------------------------------------
def _svg_png(tmp, name):
    from PIL import Image
    import io as _io
    b = _io.BytesIO()
    Image.new("RGB", (24, 24), (9, 9, 9)).save(b, "PNG")
    p = os.path.join(tmp, name)
    with open(p, "wb") as f:
        f.write(b.getvalue())
    return p


def _build(tmp, name, kinds, with_user_behind=False):
    """返回 (file_path, output_path)：一份带指定水印的 docx + 一个输出路径。"""
    src = os.path.join(tmp, f"{name}_in.docx")
    doc = Document()
    doc.add_paragraph(USER_BODY)
    doc.save(src)

    opts = {"text": {"text": "机密"},
            "image": {"image_path": _svg_png(tmp, f"{name}.png")}}
    wm = os.path.join(tmp, f"{name}_wm.docx")
    core.insert_watermark(src, list(kinds), output_path=wm, **opts)
    if with_user_behind:
        # 用户在页眉里放一张自己的 behindDoc 图——修复前会被 GUI 的删除一起带走
        d = Document(wm)
        rid, _ = d.sections[0].header.part.get_or_add_image(
            __import__("io").BytesIO(open(_svg_png(tmp, f"{name}b.png"), "rb").read()))
        engine_docx._add_drawing_to_part(
            d.sections[0].header,
            engine_docx._make_drawing(rid, 300000, 300000, 0, "用户背景图", "", 880, 0, 0))
        d.save(wm)
    return wm, os.path.join(tmp, f"{name}_out.docx")


def _user_behind_count(path):
    n = 0
    for dr in Document(path).sections[0].header._element.iter(
            engine_docx.qn("w:drawing")):
        anchor = dr.find(".//" + engine_docx.qn("wp:anchor"))
        if anchor is not None and anchor.get("behindDoc") == "1" \
                and engine_docx._mark_of(dr) is None:
            n += 1
    return n


# ===========================================================================
# P0：只有一种水印时，GUI 必须精确按类型删除
# ===========================================================================
def test_gui_only_text_watermark_clears_text_only(monkeypatch, app):
    """只有文字水印：GUI 必须传 kinds=['text']（而不是 None 的全清）。"""
    tmp = tempfile.mkdtemp()
    fp, out = _build(tmp, "only_text", ["text"], with_user_behind=True)

    app, jobs = _stub_app(monkeypatch, fp, out)
    # 文档同时含本工具文字水印 + 用户 behindDoc 图（被 detect_native 识别为原生水印），
    # 走四选一冲突弹窗；用户选「仅清除本工具水印」-> kinds=['text']，不动用户图。
    monkeypatch.setattr(gui_mod.App, "_ask_clear_conflict", lambda self: "tool")
    gui_mod.App._clear(app)
    calls = _apply(app, jobs, monkeypatch)

    assert len(calls) == 1, f"应当恰好发起一次清除，实际 {len(calls)}"
    assert calls[0]["kinds"] == ["text"], (
        f"只有文字水印时应传 kinds=['text']，实际 {calls[0]['kinds']}——"
        "None 会走 clear_any，连带删掉用户自己的 behindDoc 图")
    assert calls[0]["clear_native_only"] is False

    # 真的执行一遍，看最终文档
    core.clear_watermark(fp, output_path=out, kinds=["text"])
    assert engine_docx.detect_watermark_types(out) == set(), "文字水印应已清除"
    assert _user_behind_count(out) == 1, "用户自己的 behindDoc 图必须保留"
    assert USER_BODY in [p.text for p in Document(out).paragraphs]
    print("[GUI] 文字水印+用户图（冲突弹窗选「仅工具」） -> kinds=['text']；"
          "工具水印清除，用户 behindDoc 图保留")


def test_gui_only_image_watermark_clears_image_only(monkeypatch, app):
    """只有图片水印：GUI 必须传 kinds=['image']。"""
    tmp = tempfile.mkdtemp()
    fp, out = _build(tmp, "only_img", ["image"], with_user_behind=True)

    app, jobs = _stub_app(monkeypatch, fp, out)
    # 同上：冲突弹窗选「仅清除本工具水印」-> kinds=['image']，不动用户图。
    monkeypatch.setattr(gui_mod.App, "_ask_clear_conflict", lambda self: "tool")
    gui_mod.App._clear(app)
    calls = _apply(app, jobs, monkeypatch)

    assert len(calls) == 1
    assert calls[0]["kinds"] == ["image"], (
        f"只有图片水印时应传 kinds=['image']，实际 {calls[0]['kinds']}")
    assert calls[0]["clear_native_only"] is False

    core.clear_watermark(fp, output_path=out, kinds=["image"])
    assert engine_docx.detect_watermark_types(out) == set(), "图片水印应已清除"
    assert _user_behind_count(out) == 1, "用户自己的 behindDoc 图必须保留"
    print("[GUI] 图片水印+用户图（冲突弹窗选「仅工具」） -> kinds=['image']；"
          "工具水印清除，用户 behindDoc 图保留")


# ===========================================================================
# P0：同时含两类水印时弹窗，三个选项分别对应三种 kinds
# ===========================================================================
@pytest.mark.parametrize("choice,expected", [
    ("text", ["text"]),
    ("image", ["image"]),
    ("both", ["text", "image"]),
])
def test_gui_dual_watermark_dialog_maps_choice_to_kinds(monkeypatch, app, choice, expected):
    """双水印文档 -> 弹窗 -> 用户选什么就删什么。"""
    tmp = tempfile.mkdtemp()
    fp, out = _build(tmp, "dual", ["text", "image"])

    app, jobs = _stub_app(monkeypatch, fp, out)
    asked = []
    monkeypatch.setattr(gui_mod.App, "_ask_clear_choice",
                        lambda self: asked.append(True) or choice)
    gui_mod.App._clear(app)
    calls = _apply(app, jobs, monkeypatch)

    assert asked == [True], "双水印文档应该弹一次选择框"
    assert len(calls) == 1 and calls[0]["kinds"] == expected, (
        f"用户选 {choice} 时应传 {expected}，实际 {calls[0] if calls else None}")

    core.clear_watermark(fp, output_path=out, kinds=expected)
    want = set() if choice == "both" else {"text", "image"} - {choice}
    assert engine_docx.detect_watermark_types(out) == want, (
        f"选 {choice} 后应剩 {want or '无'}，实际 {engine_docx.detect_watermark_types(out)}")
    print(f"[GUI] 双水印弹窗：选{choice} -> kinds={expected} -> 残留 {want or '无'}")


def test_gui_dual_watermark_user_cancel_does_nothing(monkeypatch, app):
    """用户点“取消”时，一次清除都不许发生。"""
    tmp = tempfile.mkdtemp()
    fp, out = _build(tmp, "dual_cancel", ["text", "image"])

    app, jobs = _stub_app(monkeypatch, fp, out)
    monkeypatch.setattr(gui_mod.App, "_ask_clear_choice", lambda self: None)
    gui_mod.App._clear(app)

    assert jobs == [], "用户取消时不应排队任何清除任务"
    assert not os.path.exists(out), "用户取消时不该产生输出文件"
    print("[GUI] 双水印弹窗点取消 -> 不发起任何清除")


# ===========================================================================
# P0：检测失败 -> 停止清除、报告原因，不扩大删除范围
# ===========================================================================
def test_gui_detection_failure_does_not_expand_clear_scope(monkeypatch, app):
    """检测抛异常：绝不能退化成 kinds=None（clear_any）。

    同时验证：用户看到了提示，且异常原因被带出来（不吞异常）。
    """
    tmp = tempfile.mkdtemp()
    fp, out = _build(tmp, "fail", ["text"], with_user_behind=True)
    assert _user_behind_count(fp) == 1

    boom = RuntimeError("模拟检测崩溃：文档结构异常")
    monkeypatch.setattr(engine_docx, "detect_watermark_types", lambda p: (_ for _ in ()).throw(boom))

    _WARNINGS.clear()
    monkeypatch.setattr(gui_mod, "QMessageBox", _FakeMessageBox)

    app, jobs = _stub_app(monkeypatch, fp, out)
    gui_mod.App._clear(app)
    calls = _apply(app, jobs, monkeypatch)

    assert calls == [], f"检测失败时不应发起任何清除，实际 {calls}"
    assert not os.path.exists(out), "检测失败时不应产生输出文件"
    assert len(_WARNINGS) == 1, "必须弹窗告知用户"
    title, text = _WARNINGS[0]
    assert "取消" in text and "检测" in text, f"提示语应说明检测失败并已取消，实际：{text}"
    assert "模拟检测崩溃" in text, f"真实原因应透传给用户，实际：{text}"
    print("[GUI] 检测失败 -> 不清除、不扩大范围、弹窗报告原因")


def test_gui_detection_failure_leaves_document_untouched(monkeypatch, app):
    """检测失败后原文件必须一字未改（含用户自己的 behindDoc 图）。"""
    tmp = tempfile.mkdtemp()
    fp, out = _build(tmp, "fail2", ["image"], with_user_behind=True)

    real_detect = engine_docx.detect_watermark_types   # 断言时要用真的（下面会被替换）
    monkeypatch.setattr(engine_docx, "detect_watermark_types",
                        lambda p: (_ for _ in ()).throw(OSError("读取失败")))
    _WARNINGS.clear()
    monkeypatch.setattr(gui_mod, "QMessageBox", _FakeMessageBox)

    app, jobs = _stub_app(monkeypatch, fp, out)
    gui_mod.App._clear(app)
    _apply(app, jobs, monkeypatch)

    assert _user_behind_count(fp) == 1, "原文件的用户 behindDoc 图不应被动过"
    assert real_detect(fp) == {"image"}, "原文件水印应原样"
    print("[GUI] 检测失败后原文档零改动")


# ===========================================================================
# 空集 / 原生水印：未检测到本工具水印时的安全守门（取代旧的“无脑全清”语义）
# ===========================================================================
def _build_native_only(tmp, name):
    """返回 (file_path, output_path)：一份只有「Word 原生水印」（behindDoc 图片、无本工具标记）的 docx。"""
    import io as _io
    src = os.path.join(tmp, f"{name}.docx")
    doc = Document()
    doc.add_paragraph(USER_BODY)
    doc.save(src)
    d = Document(src)
    rid, _ = d.sections[0].header.part.get_or_add_image(
        _io.BytesIO(open(_svg_png(tmp, f"{name}.png"), "rb").read()))
    engine_docx._add_drawing_to_part(
        d.sections[0].header,
        engine_docx._make_drawing(rid, 300000, 300000, 0, "用户背景图", "", 880, 0, 0))
    d.save(src)
    return src, os.path.join(tmp, f"{name}_out.docx")


def test_gui_no_watermark_no_native_informs_and_skips(monkeypatch, app):
    """检测为空集且确实没有任何水印：提示「未检测到任何水印」并跳过，不再静默全清。

    取代旧的「保持 kinds=None 全清语义」——那会连用户页眉里自己的 behindDoc 图一起删。
    """
    tmp = tempfile.mkdtemp()
    src = os.path.join(tmp, "plain.docx")
    Document().add_paragraph(USER_BODY)
    doc = Document()
    doc.add_paragraph(USER_BODY)
    doc.save(src)
    fp, out = src, os.path.join(tmp, "plain_out.docx")

    _WARNINGS.clear()
    monkeypatch.setattr(gui_mod, "QMessageBox", _FakeMessageBox)

    app, jobs = _stub_app(monkeypatch, fp, out)
    gui_mod.App._clear(app)
    calls = _apply(app, jobs, monkeypatch)

    assert calls == [], "无任何水印时不应发起任何清除"
    assert not os.path.exists(out), "不应产生输出文件"
    assert any("未检测到任何水印" in t for _, t in _WARNINGS), (
        f"应提示无内容可清，实际：{_WARNINGS}")
    print("[GUI] 无工具水印且确无原生水印 -> 提示并跳过（不再静默全清）")


def test_gui_dual_watermark_plus_user_native_spares_user_art(monkeypatch, app):
    """WriteMark 文字+图片 与 用户自己的 behindDoc 图共存：
    即便用户选「文字和图片水印都去除」，用户的图也必须留着。

    这是规格第七条的硬边界：**不能因为存在 WriteMark 标记就把用户对象一起删**。
    kinds 模式下清除只认私有标记，原生/用户图形一概不碰。
    """
    tmp = tempfile.mkdtemp()
    fp, out = _build(tmp, "dual_user", ["text", "image"], with_user_behind=True)
    assert _user_behind_count(fp) == 1, "前置：文档里应有 1 张用户自己的 behindDoc 图"

    app, jobs = _stub_app(monkeypatch, fp, out)
    # 四选一冲突弹窗里用户选「仅清除本工具水印」：应只删本工具两种水印，
    # 用户的 behindDoc 图必须留着（marker 与用户内容始终可区分）。
    monkeypatch.setattr(gui_mod.App, "_ask_clear_conflict", lambda self: "tool")
    gui_mod.App._clear(app)
    calls = _apply(app, jobs, monkeypatch)

    assert len(calls) == 1 and set(calls[0]["kinds"]) == {"text", "image"}
    assert calls[0]["clear_native_only"] is False

    core.clear_watermark(fp, output_path=out, kinds=["text", "image"])
    assert engine_docx.detect_watermark_types(out) == set(), "两种 Watermark 水印都应清除"
    assert _user_behind_count(out) == 1, \
        "用户自己的 behindDoc 图必须保留——marker 与用户内容必须始终可区分"
    assert USER_BODY in [p.text for p in Document(out).paragraphs], "正文必须保留"
    print("[GUI] 双水印+用户图（冲突弹窗选「仅工具」） -> 只删本工具水印，用户图保留")


def test_gui_no_tool_watermark_but_native_requires_confirm(monkeypatch, app):
    """检测为空集、但有 Word 原生水印：必须先问用户，确认才全清、拒绝则不动。"""
    tmp = tempfile.mkdtemp()
    fp, out = _build_native_only(tmp, "native")
    assert engine_docx.detect_watermark_types(fp) == set(), "该文档不应含本工具水印"
    assert engine_docx.detect_native_watermark(fp) is True, "应检测到原生水印"

    # 用户拒绝 -> 一次清除都不许发生
    app, jobs = _stub_app(monkeypatch, fp, out)
    monkeypatch.setattr(gui_mod.App, "_ask_clear_native", lambda self: False)
    gui_mod.App._clear(app)
    calls = _apply(app, jobs, monkeypatch)
    assert calls == [], "用户拒绝时不应清除原生水印"
    assert not os.path.exists(out), "用户拒绝时不应产生输出文件"
    print("[GUI] 无工具水印但有原生 -> 拒绝 -> 不清除")

    # 用户确认 -> 全清（kinds=None，含原生水印）
    _WARNINGS.clear()
    app, jobs = _stub_app(monkeypatch, fp, out)
    monkeypatch.setattr(gui_mod.App, "_ask_clear_native", lambda self: True)
    gui_mod.App._clear(app)
    calls = _apply(app, jobs, monkeypatch)
    assert len(calls) == 1 and calls[0]["kinds"] is None, (
        f"确认后应全清（kinds=None，含原生水印），实际 {calls}")
    print("[GUI] 无工具水印但有原生 -> 确认 -> 全清")


# ===========================================================================
# 新需求：本工具水印 与 用户 Word 原生水印 共存时的四选一冲突弹窗
# ===========================================================================
def test_gui_conflict_dialog_only_when_both_present(monkeypatch, app):
    """仅含本工具水印、没有原生水印时，不应弹四选一冲突框（走原有按类型逻辑）。"""
    tmp = tempfile.mkdtemp()
    fp, out = _build(tmp, "tool_only", ["text"], with_user_behind=False)
    assert engine_docx.detect_native_watermark(fp) is False, "前置：不应有原生水印"

    asked = []
    monkeypatch.setattr(gui_mod.App, "_ask_clear_conflict",
                        lambda self: asked.append(True) or "tool")
    app, jobs = _stub_app(monkeypatch, fp, out)
    gui_mod.App._clear(app)
    _apply(app, jobs, monkeypatch)
    assert asked == [], "无原生水印时不应弹四选一冲突框"


@pytest.mark.parametrize("choice,expected_kinds,expected_native", [
    ("tool", ["text"], False),
    ("native", None, True),
    ("both", None, False),
])
def test_gui_conflict_dialog_maps_choice(monkeypatch, app, choice, expected_kinds, expected_native):
    """四选一：用户选什么，就传对应的 (kinds, clear_native_only)。"""
    tmp = tempfile.mkdtemp()
    fp, out = _build(tmp, "conflict", ["text"], with_user_behind=True)
    assert engine_docx.detect_native_watermark(fp) is True, "前置：应有原生水印"

    app, jobs = _stub_app(monkeypatch, fp, out)
    monkeypatch.setattr(gui_mod.App, "_ask_clear_conflict", lambda self: choice)
    gui_mod.App._clear(app)
    calls = _apply(app, jobs, monkeypatch)

    assert len(calls) == 1, f"应恰好一次清除，实际 {len(calls)}"
    assert calls[0]["kinds"] == expected_kinds, (
        f"选 {choice} 应传 kinds={expected_kinds}，实际 {calls[0]['kinds']}")
    assert calls[0]["clear_native_only"] is expected_native, (
        f"选 {choice} 应传 clear_native_only={expected_native}，"
        f"实际 {calls[0]['clear_native_only']}")
    print(f"[GUI] 冲突弹窗选 {choice} -> kinds={expected_kinds}, "
          f"native_only={expected_native}")


def test_gui_conflict_dialog_cancel_does_nothing(monkeypatch, app):
    """用户点「取消」：一次清除都不许发生。"""
    tmp = tempfile.mkdtemp()
    fp, out = _build(tmp, "conflict_cancel", ["text", "image"], with_user_behind=True)
    app, jobs = _stub_app(monkeypatch, fp, out)
    monkeypatch.setattr(gui_mod.App, "_ask_clear_conflict", lambda self: None)
    gui_mod.App._clear(app)
    assert jobs == [], "用户取消时不应排队任何清除任务"
    assert not os.path.exists(out), "用户取消时不该产生输出文件"
    print("[GUI] 冲突弹窗点取消 -> 不发起任何清除")


def test_gui_conflict_native_choice_removes_user_keeps_tool(monkeypatch, app):
    """选「仅清除 Word/用户水印」：用户的 behindDoc 图被删，本工具文字水印保留。"""
    tmp = tempfile.mkdtemp()
    fp, out = _build(tmp, "conflict_native", ["text"], with_user_behind=True)
    assert _user_behind_count(fp) == 1, "前置：1 张用户 behindDoc 图"
    assert engine_docx.detect_watermark_types(fp) == {"text"}, "前置：1 个本工具文字水印"

    app, jobs = _stub_app(monkeypatch, fp, out)
    monkeypatch.setattr(gui_mod.App, "_ask_clear_conflict", lambda self: "native")
    gui_mod.App._clear(app)
    calls = _apply(app, jobs, monkeypatch)

    # 真的执行一遍：用捕获到的参数跑真实清除
    c = calls[0]
    core.clear_watermark(fp, output_path=out, kinds=c["kinds"],
                         clear_native_only=c["clear_native_only"])
    assert _user_behind_count(out) == 0, "用户的 behindDoc 图应被清除"
    assert engine_docx.detect_watermark_types(out) == {"text"}, \
        "本工具文字水印必须保留（只删了用户水印）"
    assert USER_BODY in [p.text for p in Document(out).paragraphs]
    print("[GUI] 冲突选「仅用户水印」 -> 用户图删掉、本工具文字水印保留")


def test_gui_conflict_both_choice_removes_everything(monkeypatch, app):
    """选「两者都清除」：本工具水印与用户 behindDoc 图一并删除。"""
    tmp = tempfile.mkdtemp()
    fp, out = _build(tmp, "conflict_both", ["text", "image"], with_user_behind=True)
    assert _user_behind_count(fp) == 1
    assert engine_docx.detect_watermark_types(fp) == {"text", "image"}

    app, jobs = _stub_app(monkeypatch, fp, out)
    monkeypatch.setattr(gui_mod.App, "_ask_clear_conflict", lambda self: "both")
    gui_mod.App._clear(app)
    calls = _apply(app, jobs, monkeypatch)

    c = calls[0]
    core.clear_watermark(fp, output_path=out, kinds=c["kinds"],
                         clear_native_only=c["clear_native_only"])
    assert engine_docx.detect_watermark_types(out) == set(), "本工具水印应清除"
    assert _user_behind_count(out) == 0, "用户 behindDoc 图应清除"
    assert USER_BODY in [p.text for p in Document(out).paragraphs]
    print("[GUI] 冲突选「两者都清除」 -> 本工具水印与用户图均删掉")
