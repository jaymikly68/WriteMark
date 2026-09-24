"""
校验打包后的 exe 是否真正包含了本次更新后的源码（避免“旧 exe 被运行”的问题）。

做法：用 PyInstaller 的归档读取器打开 exe，定位 PYZ.pyz，反序列化出
watermark_tool.gui / watermark_tool.engine_docx 的代码对象，检查关键符号是否存在。
"""
import sys
import os

from PyInstaller.archive.readers import CArchiveReader
from PyInstaller.archive.readers import ZlibArchiveReader


def load_pyz_modules(exe_path):
    car = CArchiveReader(exe_path)
    if "PYZ.pyz" not in car.toc:
        raise RuntimeError("PYZ.pyz 未找到")
    pyz = car.open_embedded_archive("PYZ.pyz")
    modules = {}
    for name in pyz.toc:
        try:
            modules[name] = pyz.extract(name)
        except Exception:
            pass
    return modules


def find_str_consts(co):
    """递归收集代码对象中的所有字符串常量（含元组/列表/集合等容器常量里的字符串）。

    注意：像 `{"仿宋": ("simfang.ttf",)}` 这样的字典常量，其值元组在字节码里是
    一个整体常量对象，必须深入容器才能取到 "simfang.ttf"，否则会误判“未编入”。
    """
    found = []
    stack = [co]
    while stack:
        c = stack.pop()
        for k in getattr(c, "co_consts", []):
            if isinstance(k, str):
                found.append(k)
            elif hasattr(k, "co_consts"):
                stack.append(k)
            elif isinstance(k, (tuple, list, frozenset, set)):
                for item in k:
                    if isinstance(item, str):
                        found.append(item)
                    elif hasattr(item, "co_consts"):
                        stack.append(item)
    return found


def collect_names(co):
    """递归收集代码对象及其嵌套代码对象中的所有 co_names（含类/方法名）。"""
    names = set()
    stack = [co]
    while stack:
        c = stack.pop()
        names.update(getattr(c, "co_names", ()))
        for k in getattr(c, "co_consts", []):
            if hasattr(k, "co_names"):
                stack.append(k)
    return names


def main():
    exe = sys.argv[1] if len(sys.argv) > 1 else "dist/WordWatermark.exe"
    if not os.path.exists(exe):
        print("exe 不存在:", exe)
        sys.exit(1)
    mods = load_pyz_modules(exe)

    gui = mods.get("watermark_tool.gui")
    docx = mods.get("watermark_tool.engine_docx")
    wf = mods.get("watermark_tool.word_fonts")
    ok = True

    if gui is None:
        print("[缺失] watermark_tool.gui 未在归档中找到")
        ok = False
    else:
        names = collect_names(gui)
        consts = find_str_consts(gui)
        has_new = ("_make_slider_spin" in names and "_gather_kinds" in names
                   and "_maybe_load_word_fonts" in names
                   and "_apply_word_fonts" in names and "FontWorker" in names
                   and "setEditable" in names and "setCompleter" in names
                   and "setCompletionMode" in names and "setFilterMode" in names
                   and "QStringListModel" in names)
        # 中西文字体各自一个下拉框（Word 式）：界面上有“中文字体/西文字体”两个标签
        has_split_ui = (any("中文字体:" == c for c in consts)
                        and any("西文字体:" == c for c in consts)
                        and "cn_font_combo" in names and "latin_font_combo" in names)
        print(f"watermark_tool.gui: 含字体相关 _maybe_load_word_fonts/_apply_word_fonts/FontWorker"
              f" + 可编辑下拉框(Editable/Completer/CompletionMode/FilterMode/Model) -> {has_new}; "
              f"中西文字体双下拉框 -> {has_split_ui}")
        if not has_new or not has_split_ui:
            ok = False

    if docx is None:
        print("[缺失] watermark_tool.engine_docx 未在归档中找到")
        ok = False
    else:
        consts = find_str_consts(docx)
        has_new_mark = any("WB_WATERMARK_TEXT" in c for c in consts)
        has_old_mark = any(c == "WB_WATERMARK" for c in consts)
        has_font_param = any("font_name" in c for c in consts)
        has_word_fonts_ref = "word_fonts" in collect_names(docx)
        has_text_base = "text_base" in collect_names(docx)
        # 中文回退渲染的缺字形参照字符 \uffff 作为字符串常量编入（比函数名更可靠）
        has_cjk_fallback = any("\uffff" in c for c in consts)
        # 中西文分别选字体：cn_font_name / latin_font_name 透传 + 逐字符分流函数 is_cjk
        names_docx = collect_names(docx)
        has_split = (any("cn_font_name" in c for c in consts)
                     and any("latin_font_name" in c for c in consts)
                     and "is_cjk" in names_docx)
        # 混排基线对齐：用 getlength 计算各字体字宽、anchor="ls" 共基线
        has_baseline = "getlength" in names_docx and any("ls" == c for c in consts)
        print(f"watermark_tool.engine_docx: 含新标记 WB_WATERMARK_TEXT -> {has_new_mark}; "
              f"含旧标记 WB_WATERMARK -> {has_old_mark}; 透传 font_name -> {has_font_param}; "
              f"引用 word_fonts 模块 -> {has_word_fonts_ref}; 字号控制大小 text_base -> {has_text_base}; "
              f"中文回退渲染(缺字形参照\\uffff) -> {has_cjk_fallback}; "
              f"中西文分别选字体(cn/latin/is_cjk) -> {has_split}; 混排基线对齐(getlength/anchor) -> {has_baseline}")
        if not has_new_mark or has_old_mark or not (has_font_param and has_word_fonts_ref
                                                    and has_text_base and has_cjk_fallback
                                                    and has_split and has_baseline):
            ok = False

    if wf is None:
        print("[缺失] watermark_tool.word_fonts 未在归档中找到")
        ok = False
    else:
        names = collect_names(wf)
        consts = find_str_consts(wf)
        has_wf = "get_word_fonts" in names and "resolve_font_path" in names
        # 中文别名表：中文显示名（微软雅黑/黑体/仿宋…）能解析到字体文件，否则“选了不生效”
        has_alias = ("_ALIAS_FILES" in names
                     and any("微软雅黑" == c for c in consts)
                     and any("simfang.ttf" == c for c in consts))
        print(f"watermark_tool.word_fonts: 含 get_word_fonts/resolve_font_path -> {has_wf}; "
              f"中文名字体别名表 _ALIAS_FILES(微软雅黑/simfang.ttf) -> {has_alias}")
        if not has_wf or not has_alias:
            ok = False

    wdog = mods.get("watermark_tool.watchdog")
    if wdog is None:
        print("[缺失] watermark_tool.watchdog 未在归档中找到")
        ok = False
    else:
        names = collect_names(wdog)
        consts = find_str_consts(wdog)
        # 守护修复：输出文件被整个删除时能从原文件重建（此前只抛 Package not found 永不补回）
        has_recreate = (any("输出文件已被删除" in c for c in consts)
                        and "exists" in names)
        print(f"watermark_tool.watchdog: 输出文件被删除后自动重建 -> {has_recreate}")
        if not has_recreate:
            ok = False

    print("\n校验结果:", "通过 ✅" if ok else "存在问题 ❌")
    sys.exit(0 if ok else 2)


if __name__ == "__main__":
    main()
