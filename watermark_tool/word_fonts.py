"""
字体来源模块：把“用户 Word 中可用的字体”作为文本水印可选字体。

两条来源（按优先级回退）：
- get_word_fonts()：通过 Word COM（Application.FontNames）读取 Word 当前可用的
  全部字体名称。需要本机安装并可用 Microsoft Word；读取时仅取名称列表，
  不打开/不修改用户任何文档内容。
- list_system_fonts() / build_font_map()：Word 不可用（例如只用 WPS、或 COM 失败）
  时，改从 Windows 注册表 Fonts 键读取“系统已安装字体”作为回退，效果与 Word 列出
  的字体基本等价（Word 列出的是系统字体）。

另外提供 resolve_font_path(font_name)：把用户在下拉框选中的字体“显示名”解析为
磁盘上的字体文件路径，供 Pillow 在纯 Python 的 .docx 引擎里真正按该字体渲染水印
（否则所选字体只是“显示用”，实际水印仍是固定字体）。
"""
from __future__ import annotations

import os
import re

# ---------------------------------------------------------------- 注册表字体映射
# 显示名可能带后缀，如 "微软雅黑 (TrueType)" / "Arial (OpenType)"，解析与匹配时去除。
_SUFFIX_RE = re.compile(
    r"\s*\((TrueType|OpenType|TrueType & OpenType|OpenType & TrueType)\)$",
    re.IGNORECASE,
)


def _norm(name: str) -> str:
    """规范化字体名：去后缀、去首尾空白、转小写，便于精确匹配。"""
    return _SUFFIX_RE.sub("", (name or "").strip()).lower()


def _strip_suffix(name: str) -> str:
    """仅去除后缀，保留原大小写与可读显示名（用于下拉框展示）。"""
    return _SUFFIX_RE.sub("", (name or "").strip())


def _fonts_dir() -> str:
    return os.path.join(os.environ.get("SystemRoot", "C:/Windows"), "Fonts")


# 中文显示名 → 字体文件名。
# 必要性：Windows 注册表 Fonts 键里的显示名多为英文（"Microsoft YaHei"、
# "SimHei"），直接用中文名（"微软雅黑"、"黑体"）查会查不到，导致用户选了
# 中文字体却仍由兜底字体渲染——表现为“选了中文字体但没变化”。
# 这里给出常见中文字体的显式别名，优先于注册表匹配。
_ALIAS_FILES: dict[str, tuple[str, ...]] = {
    "微软雅黑": ("msyh.ttc", "msyh.ttf"),
    "微软雅黑 light": ("msyhl.ttc",),
    "microsoft yahei": ("msyh.ttc", "msyh.ttf"),
    "microsoft yahei light": ("msyhl.ttc",),
    "microsoft yahei ui": ("msyh.ttc",),
    "黑体": ("simhei.ttf",),
    "simhei": ("simhei.ttf",),
    "宋体": ("simsun.ttc", "simsun.ttf"),
    "simsun": ("simsun.ttc",),
    "新宋体": ("nsimsun.ttc",),
    "nsimsun": ("nsimsun.ttc",),
    "楷体": ("simkai.ttf",),
    "楷体_gb2312": ("simkai.ttf",),
    "kaiti": ("simkai.ttf",),
    "仿宋": ("simfang.ttf",),
    "仿宋_gb2312": ("simfang.ttf",),
    "fangsong": ("simfang.ttf",),
    "隶书": ("simli.ttf",),
    "幼圆": ("simyou.ttf",),
    "华文仿宋": ("stfangso.ttf",),
    "华文宋体": ("stsong.ttf",),
    "华文楷体": ("stkaiti.ttf",),
    "华文黑体": ("stheiti.ttf",),
    "华文中宋": ("stzhongs.ttf",),
    "华文细黑": ("stxihei.ttf",),
    "方正舒体": ("fzstk.ttf",),
    "方正姚体": ("fzytk.ttf",),
}


def build_font_map() -> dict[str, str]:
    """返回 {字体显示名: 绝对字体文件路径}。

    来源：HKLM\\SOFTWARE\\Microsoft\\Windows NT\\CurrentVersion\\Fonts
    注册表值名即“显示名”，值即文件名（可能带盘符/目录，或相对 Fonts 目录，
    也可能形如 "msyh.ttc,0" 表示集合内索引）。非 Windows 或读取失败返回空字典。
    """
    mapping: dict[str, str] = {}
    try:
        import winreg
    except Exception:
        return mapping

    fonts_dir = _fonts_dir()
    try:
        key = winreg.OpenKey(
            winreg.HKEY_LOCAL_MACHINE,
            r"SOFTWARE\Microsoft\Windows NT\CurrentVersion\Fonts",
        )
    except OSError:
        return mapping

    i = 0
    while True:
        try:
            name, value, _ = winreg.EnumValue(key, i)
        except OSError:
            break
        i += 1
        if not name or not value:
            continue
        path = str(value).split(",")[0].strip()  # 去掉 ",索引"
        if not path:
            continue
        if not os.path.isabs(path):
            path = os.path.join(fonts_dir, path)
        mapping[name] = path
    return mapping


_MAP_CACHE: dict[str, str] | None = None


def _get_map() -> dict[str, str]:
    global _MAP_CACHE
    if _MAP_CACHE is None:
        _MAP_CACHE = build_font_map()
    return _MAP_CACHE


def refresh_font_map() -> None:
    """强制重建字体映射缓存（例如字体安装后）。"""
    global _MAP_CACHE
    _MAP_CACHE = build_font_map()


def list_system_fonts() -> list[str]:
    """返回系统已安装字体的“干净显示名”列表（去后缀、排序、去重）。

    作为 Word 不可用时扩充字体下拉框的回退来源。
    """
    names = []
    seen = set()
    for disp in _get_map().keys():
        clean = _strip_suffix(disp)
        if clean and clean not in seen:
            seen.add(clean)
            names.append(clean)
    names.sort()
    return names


def resolve_font_path(font_name: str | None) -> str | None:
    """把字体“显示名”解析为磁盘字体文件路径（供 Pillow 加载）。

    解析顺序：
    1) 常见中文名别名表（微软雅黑 / 黑体 / 宋体 / 楷体 / 仿宋 …）——注册表只有
       英文显示名，中文名必须靠本表才能解析到，否则“选了中文字体却不生效”；
    2) 注册表精确匹配（去后缀+小写）；
    3) 注册表子串宽松匹配；
    均失败返回 None（调用方应回退到默认字体）。
    """
    if not font_name:
        return None
    mapping = _get_map()
    target = _norm(font_name)
    if not target:
        return None

    # 1) 中文别名表（文件确实存在才采用，避免装了精简字体包时指向不存在的路径）
    for fn in _ALIAS_FILES.get(target, ()):
        p = os.path.join(_fonts_dir(), fn)
        if os.path.exists(p):
            return p

    # 2) 精确
    for disp, path in mapping.items():
        if _norm(disp) == target:
            return path
    # 3) 宽松
    for disp, path in mapping.items():
        nd = _norm(disp)
        if target in nd or nd in target:
            return path
    return None


# ---------------------------------------------------------------- Word COM 字体
def get_word_fonts() -> list[str] | None:
    """通过 Word COM 读取 Word 当前可用的全部字体名称。

    返回字体名列表；若 Word 不可用 / COM 调用失败，返回 None（调用方应回退到
    系统字体）。本函数不打开、不读取、不修改用户的任何文档内容，仅取字体名列表。

    ⚠ 注意：这条路径需要临时启动一个 Word 实例。**默认不再走它**（见 collect_fonts），
    因为它会留下“幽灵 Word”——用户后来去关那个看不见的 Word 时就会觉得“退出 Word 卡顿”。
    这里保留它仅作为注册表读取失败时的兜底，并且会用 com_cleanup 确认 Word 真的退出。
    """
    try:
        import pythoncom
        import win32com.client
    except Exception:
        return None

    from . import com_cleanup

    before = com_cleanup.winword_pids()

    # 在工作线程里也要先初始化 COM 单元（主线程通常已由 Qt 初始化过）
    try:
        pythoncom.CoInitializeEx(pythoncom.COINIT_APARTMENTTHREADED)
    except Exception:
        pass

    word = None
    created = False
    try:
        try:
            # 优先连已运行的 Word，避免额外启动一个实例
            word = win32com.client.GetActiveObject("Word.Application")
        except Exception:
            word = win32com.client.Dispatch("Word.Application")
            word.Visible = False
            word.DisplayAlerts = False
            created = True
        names = word.FontNames
        fonts = [str(n) for n in names]
        return fonts
    except Exception:
        return None
    finally:
        if word is not None and created:
            # 只收拾我们自己起的那个 Word；退出不干净时强制结束
            com_cleanup.quit_word(word, before_pids=before)
        else:
            try:
                del word
            except Exception:
                pass
            try:
                import gc
                gc.collect()
            except Exception:
                pass
        try:
            pythoncom.CoUninitialize()
        except Exception:
            pass


def collect_fonts() -> tuple[list[str], str]:
    """获取可用于下拉框的字体列表，并返回来源说明。

    来源优先级：**Word COM（本机化名称、最全）→ 系统已安装字体（注册表）**。

    为什么仍以 Word 为主：Word 的字体名是**本机化**的（中文系统上给出"宋体 / 黑体 /
    微软雅黑 / 仿宋"，数量约 1400 条），而注册表里这些只以 "SimSun / SimHei /
    Microsoft YaHei" 之类的英文名出现（中文名只有 14 条），用户按中文名找字体会很别扭。

    代价是要临时起一个 Word——早先版本正因为没收拾干净而留下"幽灵 Word"，
    导致用户后来关它时觉得"退出 Word 卡顿"。现在收尾统一走 com_cleanup：
    先 Quit、再释放引用、仍不退就只强制结束我们自己起的那一个进程（不动用户的 Word）。

    另外过滤掉以 `@` 开头的竖排字体变体（"@宋体" 之类），它们只是同一字体的竖排形式，
    对本工具的水印渲染没有意义，留着只会让下拉框里出现一堆重复项。
    """
    fonts = get_word_fonts()
    if fonts:
        cleaned = [f for f in fonts if f and not f.startswith("@")]
        if cleaned:
            return cleaned, "Word"

    sys_fonts = list_system_fonts()
    if sys_fonts:
        return sys_fonts, "系统字体"
    return [], "无"
