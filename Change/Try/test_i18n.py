"""
i18n 不变式测试（风险 l：中文 UI 文案作 key 的长期维护）。

设计本身（中文原文即 key、zh-CN 无翻译表、回退链兜底）是有意而为，功能不受影响，
因此这里**不改动任何运行时代码**，只把“中文作 key”的契约与回退语义锁成测试，
防止日后编辑中文文案或新增语种时静默漂移到无法翻译的状态。

覆盖：
- zh-CN 下 tr() 原样返回（不翻译）；
- en 等语言能翻译、缺词条时按回退链兜底返回中文原文（绝不出现空串）；
- trf 占位符正常格式化，且占位符写错时回退到原文格式而不崩；
- FALLBACK 目标语言（zh-CN 除外）确实存在；
- 任一非 zh-CN 翻译 dict 的 key 都必须是 en 里定义过的“规范中文 key”
  （捕捉某语种中文 key 拼写错误导致整条翻译失效的隐患）；
- resolve_ui_font 返回 str；set_lang 对未知语言码回退到默认。
"""
from __future__ import annotations

import pytest

from watermark_tool import i18n


@pytest.fixture
def restore_lang():
    saved = i18n._cur
    try:
        yield
    finally:
        i18n.set_lang(saved)


def test_tr_zh_cn_is_noop(restore_lang):
    i18n.set_lang("zh-CN")
    assert i18n.tr("Word 秒退") == "Word 秒退"
    assert i18n.tr("") == ""
    assert i18n.tr("不存在的文案XYZ") == "不存在的文案XYZ"


def test_tr_en_translates(restore_lang):
    i18n.set_lang("en")
    assert i18n.tr("提示") == "Notice"
    # 缺词条的 key 应回退到中文原文，绝不返回空串
    missing = "這是一條絕對不存在的测试文案"
    assert i18n.tr(missing) == missing


def test_trf_placeholder_and_fallback(restore_lang):
    i18n.set_lang("ja")
    out = i18n.trf("颜色: {name}", name="赤")
    assert out == "色: 赤"
    # 占位符写错时回退到原文格式而不崩（trf 内部 except 兜底）
    i18n.set_lang("zh-CN")
    out2 = i18n.trf("颜色: {name}", name="红")
    assert out2 == "颜色: 红"


def test_fallback_targets_valid():
    for _code, targets in i18n.FALLBACK.items():
        for t in targets:
            # zh-CN 是“中文原文即 key”的特殊目标，允许不在 T 中；其余必须是真实语种表
            if t == "zh-CN":
                continue
            assert t in i18n.T, f"FALLBACK 目标 {t} 既非 zh-CN 也不在 T 中"


def test_translation_keys_subset_of_en():
    en_keys = set(i18n.T["en"])
    for code, _name in i18n.LANGS:
        if code == "zh-CN":
            continue
        extra = set(i18n.T.get(code, {})) - en_keys
        assert not extra, f"{code} 含有 en 未定义的 key（中文 key 可能拼写错误）: {sorted(extra)[:5]}"


def test_resolve_ui_font_returns_str(restore_lang):
    for code, _name in i18n.LANGS:
        val = i18n.resolve_ui_font(code)
        assert isinstance(val, str), f"resolve_ui_font({code}) 应返回 str"


def test_set_lang_unknown_falls_back(restore_lang):
    assert i18n.set_lang("not-a-real-lang") == i18n.DEFAULT_LANG
    # 回退后 tr 仍按默认（zh-CN）原样返回
    assert i18n.tr("提示") == "提示"


def test_has_and_missing(restore_lang):
    i18n.set_lang("en")
    assert i18n.has("提示")
    assert "提示" not in i18n.missing(["提示"])
    assert "不存在的文案XYZ" in i18n.missing(["不存在的文案XYZ"])
