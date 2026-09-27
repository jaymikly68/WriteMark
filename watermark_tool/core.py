"""
统一调度层：根据文件类型选择引擎。
- .docx：默认用纯 Python(docx) 引擎，无需安装 Word 即可处理（更便携）。
- .doc ：必须用 Word COM 引擎（目标机需安装 Word）。
"""
from __future__ import annotations

import logging
import os
import shutil

from . import engine_docx, engine_com

log = logging.getLogger(__name__)


def get_desktop() -> str:
    """返回当前用户的桌面路径；若取不到则返回源文件所在目录。

    说明：此前用 `except Exception: pass` 静默吞掉所有异常，会藏匿真实的
    IO/环境错误。这里收窄为文件操作可能抛出的 (OSError, ValueError)，
    其余异常（多为代码 bug）会正常向上传播，便于排查；找不到桌面时
    仍按设计回退到源文件目录（返回 None，由调用方处理）。
    """
    try:
        desk = os.path.join(os.path.expanduser("~"), "Desktop")
        if os.path.isdir(desk):
            return desk
    except (OSError, ValueError) as e:
        log.debug("定位桌面目录失败，将回退到源文件目录: %s", e)
    return None


def default_output_path(src_path: str, suffix="WaterMark") -> str:
    """默认输出路径：桌面/源文件同名+后缀；桌面不存在时退化为源文件同目录。"""
    src_path = os.path.abspath(src_path)
    desk = get_desktop()
    base_dir = desk if desk else os.path.dirname(src_path)
    stem, ext = os.path.splitext(os.path.basename(src_path))
    return os.path.join(base_dir, f"{stem}{suffix}{ext}")


def _protect_original(src_path: str, output_path: str) -> str:
    """若输出路径与源文件完全相同（会覆盖原文件），自动加后缀保护原文件。"""
    if os.path.abspath(output_path) == os.path.abspath(src_path):
        stem, ext = os.path.splitext(output_path)
        return f"{stem}{'WaterMark'}{ext}"
    return output_path


def com_available() -> bool:
    try:
        import win32com.client  # noqa: F401
        return True
    except Exception:
        return False


LOCK_MSG = ("文件正被 WPS/Word 占用，无法写入。\n"
            "请先完全退出 WPS/Word 再操作：注意 WPS 关闭窗口后进程可能仍驻留后台并锁定文件，"
            "可在任务管理器中结束 wps.exe 后重试。")


def is_file_locked(path: str) -> bool:
    """检测文件是否被其它程序（WPS/Word 等）以写方式占用。文件不存在时返回 False。"""
    try:
        with open(path, "r+b"):
            return False
    except PermissionError:
        return True
    except FileNotFoundError:
        return False


def _ensure_not_locked(path: str):
    if is_file_locked(path):
        raise RuntimeError(f"{LOCK_MSG}\n文件: {path}")


def _use_com(path: str) -> bool:
    ext = os.path.splitext(path)[1].lower()
    if ext == ".doc":
        return True  # .doc 只能靠 Word
    # .docx / 其它：优先纯 Python；若纯 Python 失败可在此回退 COM
    return False


def insert_watermark(src_path: str, kinds, output_path: str = None,
                     preserve_existing: bool = False, **opts) -> dict:
    """
    插入水印。kinds 为包含 'text' / 'image' 的列表，可同时含两者（同时添加两类水印）。
    默认【不破坏原文件】：复制到 output_path 后在副本上操作。
    - output_path 为 None：默认输出到桌面/<原名>WaterMark<ext>。
    - 若 output_path 与源文件相同：自动加后缀，避免覆盖原文件。
    - preserve_existing=True：当 output_path 已存在时不从源文件覆盖，而是先在现有
      副本上清除本工具写入的水印图形（不动正文），再重新写入——用于后台守护“补回”
      场景，避免覆盖用户在输出文件上的正文修改。
    - 返回 dict 内含 "output" 字段，标明实际写出位置。
    """
    src_path = os.path.abspath(src_path)
    _ensure_not_locked(src_path)
    if output_path is None:
        output_path = default_output_path(src_path)
    output_path = os.path.abspath(_protect_original(src_path, output_path))
    _ensure_not_locked(output_path)
    out_dir = os.path.dirname(output_path)
    if out_dir and not os.path.isdir(out_dir):
        os.makedirs(out_dir, exist_ok=True)  # 输出目录不存在时自动创建

    if _use_com(src_path):
        if not com_available():
            raise RuntimeError("处理 .doc 需要本机安装 Microsoft Word。")
        if preserve_existing and output_path and os.path.exists(output_path):
            # 守护补回：在现有输出上先清后插，保留正文修改（.doc 仅 Windows+Word 可用）
            engine_com.clear_watermark(output_path, output_path=output_path)
            res = engine_com.insert_watermark(output_path, kinds, output_path=output_path, **opts)
        else:
            res = engine_com.insert_watermark(src_path, kinds, output_path=output_path, **opts)
    else:
        if preserve_existing and output_path and os.path.exists(output_path):
            # 守护补回：在现有输出上先清后插，保留正文修改（只清本工具的水印图形）
            engine_docx.clear_watermark(output_path)
            res = engine_docx.insert_watermark(output_path, kinds, **opts)
        else:
            shutil.copy2(src_path, output_path)  # 保留原文件，在新副本上加水印
            res = engine_docx.insert_watermark(output_path, kinds, **opts)
    if isinstance(res, dict):
        res["output"] = output_path
    return res


def clear_watermark(src_path: str, output_path: str = None, kinds: list = None) -> dict:
    """
    清除水印。同样默认【不破坏原文件】：复制到 output_path 后在副本上清除。
    kinds 给定（'text'/'image'）时只清除指定类型（仅 .docx 生效，.doc 仍整文档清除）。
    """
    src_path = os.path.abspath(src_path)
    _ensure_not_locked(src_path)
    if output_path is None:
        output_path = default_output_path(src_path)
    output_path = os.path.abspath(_protect_original(src_path, output_path))
    _ensure_not_locked(output_path)
    out_dir = os.path.dirname(output_path)
    if out_dir and not os.path.isdir(out_dir):
        os.makedirs(out_dir, exist_ok=True)

    if _use_com(src_path):
        if not com_available():
            raise RuntimeError("处理 .doc 需要本机安装 Microsoft Word。")
        res = engine_com.clear_watermark(src_path, output_path=output_path)
    else:
        shutil.copy2(src_path, output_path)  # 保留原文件，在新副本上清除水印
        res = engine_docx.clear_watermark(output_path, kinds=kinds)
    if isinstance(res, dict):
        res["output"] = output_path
    return res


def has_watermark(path: str) -> bool:
    if _use_com(path):
        if not com_available():
            raise RuntimeError("处理 .doc 需要本机安装 Microsoft Word。")
        return engine_com.has_watermark(path)
    return engine_docx.has_watermark(path)


def watermark_count(path: str) -> int:
    """统计水印图形份数：守护用它判断水印是否被“部分删除”。"""
    if _use_com(path):
        if not com_available():
            raise RuntimeError("处理 .doc 需要本机安装 Microsoft Word。")
        return engine_com.count_watermarks(path)
    return engine_docx.count_watermarks(path)
