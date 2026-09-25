"""
校验打包后的产物是否真正包含了本次更新后的源码（避免“旧 exe 被运行”的问题）。

现在交付形态是两层：
    dist/WordWatermark.exe        启动器（极小 onefile，尾部追加了 payload zip）
    dist/WriteMarkApp/WriteMarkApp.exe   真正的运行体（onedir，代码在 PYZ 里）
因此：
    python test_exe_verify.py                      # 同时校验两者
    python test_exe_verify.py <运行体exe>          # 只校验运行体

做法：用 PyInstaller 的归档读取器打开 exe，定位 PYZ.pyz，反序列化出
watermark_tool.gui / watermark_tool.engine_docx 的代码对象，检查关键符号是否存在。
"""
import sys
import os
import zipfile

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


def verify_launcher(path):
    """校验启动器：既能被当成 PE 执行，也在尾部挂着可被 zipfile 读穿的运行体包。"""
    print("\n--- 启动器校验 ---")
    if not os.path.exists(path):
        print(f"[缺失] 启动器不存在：{path}")
        return False
    try:
        with zipfile.ZipFile(path) as z:
            names = z.namelist()
    except Exception as e:
        print(f"[错误] 启动器尾部 payload 不可读：{e}")
        return False
    has_app = any(n.endswith("WriteMarkApp.exe") for n in names)
    size_mb = os.path.getsize(path) / 1048576
    print(f"启动器 {os.path.basename(path)}：{size_mb:.1f} MB，"
          f"内嵌运行体条目 {len(names)} 个，含 WriteMarkApp.exe -> {has_app}")
    if not has_app:
        return False
    # 启动器自身只应有很小的运行时，保证退出时不会有大量临时文件要删
    mods = load_pyz_modules(path)
    gui_leaked = any(m.startswith("watermark_tool") for m in mods)
    print(f"启动器未把应用代码打进自身 PYZ（避免临时目录膨胀） -> {not gui_leaked}")
    return not gui_leaked


def main():
    app_exe = sys.argv[1] if len(sys.argv) > 1 else \
        os.path.join("dist", "WriteMarkApp", "WriteMarkApp.exe")
    if not os.path.exists(app_exe):
        print("exe 不存在:", app_exe)
        sys.exit(1)
    mods = load_pyz_modules(app_exe)

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
        # 防去除加固：平铺/冗余开关 + 系统托盘常驻
        has_harden = ("tile_chk" in names and "redundant_chk" in names
                      and "tile_rows_spin" in names and "_on_harden_changed" in names
                      and any("平铺满页" in c for c in consts))
        has_tray = ("QSystemTrayIcon" in names and "_setup_tray" in names
                    and "_quit_app" in names and "_show_from_tray" in names
                    and "_stop_watch_from_tray" in names)
        # v1.4.2 版式：两列底边平行（_EqualBottoms/_col_eq）+ 视频输入输出行上移
        # v1.4.3 版式：Word 秒退框并入文本水印列、与防去除加固同带（_band_eq）
        has_layout42 = ("_EqualBottoms" in names and "_col_eq" in names
                        and "v_src_edit" in names and "v_out_edit" in names
                        and "file_edit" in names and "out_edit" in names)
        has_layout43 = ("_band_eq" in names and "_exit_info" in names
                        and "_exit_detail" in names and "_btn_fix" in names
                        and "_btn_revert" in names)
        print(f"watermark_tool.gui: 含字体相关 _maybe_load_word_fonts/_apply_word_fonts/FontWorker"
              f" + 可编辑下拉框(Editable/Completer/CompletionMode/FilterMode/Model) -> {has_new}; "
              f"中西文字体双下拉框 -> {has_split_ui}; "
              f"防去除加固(平铺/冗余) -> {has_harden}; 系统托盘常驻 -> {has_tray}; "
              f"v1.4.2 版式(两列等高/视频行上移) -> {has_layout42}; "
              f"v1.4.3 版式(秒退框并入左列/与加固同带) -> {has_layout43}")
        if (not has_new or not has_split_ui or not has_harden or not has_tray
                or not has_layout42 or not has_layout43):
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
        # 防去除加固：平铺布局 + 伪装显示名 + 份数统计 + 页脚冗余
        has_harden_docx = ("tile_layout" in names_docx and "decoy_name" in names_docx
                           and "DECOY_NAMES" in names_docx
                           and "count_watermarks" in names_docx
                           and "_iter_footers" in names_docx
                           and any("redundant" == c for c in consts))
        print(f"watermark_tool.engine_docx: 含新标记 WB_WATERMARK_TEXT -> {has_new_mark}; "
              f"含旧标记 WB_WATERMARK -> {has_old_mark}; 透传 font_name -> {has_font_param}; "
              f"引用 word_fonts 模块 -> {has_word_fonts_ref}; 字号控制大小 text_base -> {has_text_base}; "
              f"中文回退渲染(缺字形参照\\uffff) -> {has_cjk_fallback}; "
              f"中西文分别选字体(cn/latin/is_cjk) -> {has_split}; 混排基线对齐(getlength/anchor) -> {has_baseline}; "
              f"防去除加固(tile/decoy/count/footer) -> {has_harden_docx}")
        if not has_new_mark or has_old_mark or not (has_font_param and has_word_fonts_ref
                                                    and has_text_base and has_cjk_fallback
                                                    and has_split and has_baseline
                                                    and has_harden_docx):
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
        # 按份数校验：水印被删掉一部分也要整体补齐
        has_count_check = (any("部分移除" in c for c in consts)
                           and "_reinsert" in names and "watermark_count" in names)
        print(f"watermark_tool.watchdog: 输出文件被删除后自动重建 -> {has_recreate}; "
              f"按份数校验并补齐 -> {has_count_check}")
        if not has_recreate or not has_count_check:
            ok = False

    # 关闭流程（v1.1.7）：点关闭 = 关闭页面不退出后台，不再弹询问框：
    # - _go_background（隐藏窗口保留进程）必须存在，旧的 _ask_close_choice 询问框必须移除；
    # - 一键插入/清除完成必须弹「任务已完成」；
    # - main() 关掉 quitOnLastWindowClosed + 显式 quit，避免隐藏窗口被当成退出；
    # - 托盘初始化失败不再静默吞掉。
    names_gui = collect_names(gui) if gui is not None else set()
    consts_gui = find_str_consts(gui) if gui is not None else []
    has_close_background = ("_go_background" in names_gui
                            and "_ask_close_choice" not in names_gui
                            and any("任务已完成" in c for c in consts_gui))
    has_env_override = any("WM_CLOSE_CHOICE" in c for c in consts_gui)
    # co_names 里是属性名，"(False)" 参数不会作为字符串常量出现
    has_no_autoclose = "setQuitOnLastWindowClosed" in names_gui
    has_tray_debug = ("_tray_reason" in names_gui and "_tray_tried" in names_gui)
    print(f"watermark_tool.gui: 关闭页面不退出后台(无询问框+完成弹窗) -> {has_close_background}; "
          f"WM_CLOSE_CHOICE 自动化开关 -> {has_env_override}; "
          f"main() 禁用 quitOnLastWindowClosed -> {has_no_autoclose}; "
          f"托盘失败原因可见(_tray_reason) -> {has_tray_debug}")
    if not (has_close_background and has_env_override and has_no_autoclose and has_tray_debug):
        ok = False

    # 幽灵 Word 修复（v1.1.2）：收尾模块必须一起打进去，且字体来源要过滤 @ 竖排变体
    cc = mods.get("watermark_tool.com_cleanup")
    if cc is None:
        print("[缺失] watermark_tool.com_cleanup 未在归档中找到（幽灵 Word 收尾逻辑）")
        ok = False
    else:
        cc_names = collect_names(cc)
        has_cleanup = {"winword_pids", "quit_word", "kill_pids"} <= cc_names
        print(f"watermark_tool.com_cleanup: 含 winword_pids/quit_word/kill_pids -> {has_cleanup}")
        if not has_cleanup:
            ok = False

    if wf is not None:
        wf_names = collect_names(wf)
        has_com_cleanup_ref = "com_cleanup" in wf_names
        print(f"watermark_tool.word_fonts: 引用 com_cleanup 收尾 Word -> {has_com_cleanup_ref}")
        if not has_com_cleanup_ref:
            ok = False

    # ---- 视频水印（v1.1.8）：与 Word 水印完全独立的模块 + 自带 ffmpeg 二进制 ----
    video = mods.get("watermark_tool.video")
    if video is None:
        print("[缺失] watermark_tool.video 未在归档中找到")
        ok = False
    else:
        vnames = collect_names(video)
        vconsts = find_str_consts(video)
        # 关键能力符号（模块级函数）必须编入
        has_video_api = ({"add_video_watermark", "watermark_video_formats",
                          "_mux_audio", "_scroll_pos", "_fixed_pos",
                          "_layer_specs", "_normalize_motion"} <= vnames)
        # 复用 engine_docx 的纯 PIL 文字/图片渲染（不依赖 Word）
        has_pil_reuse = ("engine_docx" in vnames
                         and ("render_text_png" in vnames or "render_text_png" in vconsts))
        # 逐帧合成/写帧的逻辑（区分于 Word 水印的实现）
        has_frame_loop = "append_data" in vnames or any("逐帧" in c for c in vconsts)
        print(f"watermark_tool.video: 视频水印 API(add_video_watermark/格式/滚动/固定/mux) -> {has_video_api}; "
              f"复用 engine_docx 纯 PIL 渲染 -> {has_pil_reuse}; 逐帧合成/写帧 -> {has_frame_loop}")
        if not has_video_api or not has_pil_reuse or not has_frame_loop:
            ok = False

    # ffmpeg 二进制：随包附带（imageio_ffmpeg 自带 ~70MB），视频解码/音频 mux 必需——
    # 打包后若不随附，视频水印功能会因找不到解码器而失效。
    # 注意：较新 PyInstaller 会把 datas 放到 onedir 的 _internal/ 子目录，
    # 因此运行体目录与 payload zip 都要兼顾 _internal/ 前缀。
    ff_found, ff_note = False, ""
    app_dir = os.path.dirname(os.path.abspath(app_exe))
    # 1) onedir 运行体目录里应存在 ffmpeg 二进制（可能在 _internal/ 下）
    for cand in ("imageio_ffmpeg/binaries",
                 "_internal/imageio_ffmpeg/binaries"):
        ff_dir = os.path.join(app_dir, cand)
        if os.path.isdir(ff_dir):
            for fn in os.listdir(ff_dir):
                if fn.startswith("ffmpeg") and fn.endswith(".exe"):
                    ff_found, ff_note = True, f"运行体目录 {ff_dir}/{fn}"
                    break
        if ff_found:
            break
    # 2) 若运行体目录里没有（例如直接对启动器校验），再看启动器 payload zip
    if not ff_found:
        try:
            with zipfile.ZipFile(app_exe) as z:
                for n in z.namelist():
                    if "imageio_ffmpeg/binaries/ffmpeg" in n:
                        ff_found, ff_note = True, f"payload zip 内 {n}"
                        break
        except Exception:
            pass
    print(f"ffmpeg 二进制随包附带(视频解码/音频保留) -> {ff_found}  ({ff_note})")
    if not ff_found:
        ok = False

    if len(sys.argv) <= 1:
        # 未指定参数时，连同启动器一起校验
        if not verify_launcher(os.path.join("dist", "WordWatermark.exe")):
            ok = False

    print("\n校验结果:", "通过 ✅" if ok else "存在问题 ❌")
    sys.exit(0 if ok else 2)


if __name__ == "__main__":
    main()
