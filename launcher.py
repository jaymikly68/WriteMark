"""单文件启动器（这就是用户拿到的 WordWatermark.exe）。

为什么要多这一层？
------------------------------------------------------------------
PyInstaller 的 **onefile** 形态有个绕不开的毛病：每次运行都要把整个运行时
（本项目约 130MB / 220+ 个文件）解包到 %TEMP%\\_MEIxxxxxx，程序窗口关掉以后
bootloader 还得把这整个目录删干净，进程才真正退出。本机实测：窗口 0.1 秒消失，
但进程要 8.7~11.6 秒才退出（用户机器上带杀软实时扫描会到 20~30 秒）——
这就是各个版本都“没修好的退出卡顿”。根因不在 Python 代码，无解于代码层面。

所以这里换成：
    启动器 exe（极小）  +  本地常驻的运行时目录
运行时只在第一次使用时解压到 %LOCALAPPDATA%\\WriteMark\\runtime\\<版本>，
之后启动直接就地运行；**关闭时不需要删除任何临时目录 → 进程立刻退出**。
附带好处：第 2 次以后的启动速度也明显变快（不用每次现场解压 130MB）。

对用户而言仍然是双击一个 WordWatermark.exe，没有其它文件要管。
"""
from __future__ import annotations

import os
import shutil
import subprocess
import sys
import zipfile

APP_VERSION = "1.1.2"
APP_EXE_NAME = "WriteMarkApp.exe"
PAYLOAD = "app_payload.zip"


def payload_zip_path() -> str:
    """运行体压缩包的位置。

    优先：同目录 payload/app_payload.zip（便于本地调试）
    否则：本 exe 自身——打包时 zip 被追加在 exe 文件尾部，
    Python 的 zipfile 本来就支持这种“前面带数据”的自解压格式。
    这样做的关键收益：压缩包不再需要随 onefile 每次展开到 %TEMP% 再删除，
    启动器自身的临时目录只剩几 MB 的 Python 运行时，退出几乎瞬间完成。
    """
    base_dir = os.path.dirname(sys.executable if getattr(sys, "frozen", False)
                               else os.path.abspath(__file__))
    local = os.path.join(base_dir, "payload", PAYLOAD)
    if os.path.isfile(local):
        return local
    return sys.executable if getattr(sys, "frozen", False) else os.path.abspath(__file__)


def runtime_dir() -> str:
    base = os.environ.get("LOCALAPPDATA") or os.path.expanduser("~")
    return os.path.join(base, "WriteMark", "runtime", APP_VERSION)


def _tk_window(title: str, text: str):
    """返回 (update_fn, close_fn)；tkinter 不可用时退化为 no-op。"""
    try:
        import tkinter
        from tkinter import ttk
    except Exception:
        return (lambda msg=None: None), (lambda: None)

    root = tkinter.Tk()
    root.title(title)
    root.resizable(False, False)
    root.attributes("-topmost", True)
    frame = ttk.Frame(root, padding=20)
    frame.pack()
    var = tkinter.StringVar(value=text)
    label = ttk.Label(frame, textvariable=var, width=46)
    label.pack()
    pb = ttk.Progressbar(frame, mode="indeterminate", length=320)
    pb.pack(pady=(10, 0))
    pb.start(12)
    root.update_idletasks()
    try:
        root.eval("tk::PlaceWindow . center")
    except Exception:
        pass
    root.update()

    def update(msg=None):
        if msg:
            var.set(msg)
        try:
            root.update()
        except Exception:
            pass

    def close():
        try:
            root.destroy()
        except Exception:
            pass

    return update, close


def ensure_app(update) -> str:
    dst = runtime_dir()
    app_exe = os.path.join(dst, APP_EXE_NAME)
    if os.path.isfile(app_exe):
        return app_exe

    zip_path = payload_zip_path()
    if not os.path.isfile(zip_path):
        raise RuntimeError(f"找不到内嵌的运行体压缩包：{zip_path}")

    update("正在准备运行环境（首次使用，仅需一次）…")
    os.makedirs(os.path.dirname(dst), exist_ok=True)
    tmp = dst + ".tmp"
    shutil.rmtree(tmp, ignore_errors=True)
    os.makedirs(tmp, exist_ok=True)
    with zipfile.ZipFile(zip_path) as z:
        names = z.namelist()
        total = len(names)
        for i, name in enumerate(names, 1):
            z.extract(name, tmp)
            if i % 40 == 0 or i == total:
                update(f"正在准备运行环境（首次使用，仅需一次）… {i}/{total}")
    update("解压完成，正在启动…")
    # 原子替换：避免中途失败留下残缺目录，导致下次启动用到坏文件
    if os.path.isdir(dst):
        shutil.rmtree(dst, ignore_errors=True)
    os.replace(tmp, dst)
    app_exe = os.path.join(dst, APP_EXE_NAME)
    if not os.path.isfile(app_exe):
        raise RuntimeError(f"解压后找不到主程序：{app_exe}\n"
                           f"请删除 {dst} 后重试。")
    return app_exe


def cleanup_old_version():
    """删除其它版本的运行时目录，避免长期占用磁盘。"""
    base = os.path.dirname(runtime_dir())
    keep = runtime_dir()
    try:
        for name in os.listdir(base):
            d = os.path.join(base, name)
            if os.path.isdir(d) and os.path.abspath(d) != os.path.abspath(keep):
                shutil.rmtree(d, ignore_errors=True)
    except Exception:
        pass


def _fatal(msg: str):
    try:
        import ctypes
        ctypes.windll.user32.MessageBoxW(0, msg, "Word 一键水印工具", 0x10)
    except Exception:
        pass


def main() -> int:
    need_gui = not os.path.isfile(os.path.join(runtime_dir(), APP_EXE_NAME))
    update = lambda msg=None: None
    close = lambda: None
    if need_gui:
        update, close = _tk_window("Word 一键水印工具", "正在准备运行环境（首次使用，仅需一次）…")
    try:
        app_exe = ensure_app(update)
    except Exception as e:
        close()
        _fatal(f"启动失败：无法准备运行环境。\n\n{e}")
        return 1
    close()

    # 独立进程组：启动器退出后应用继续独立运行
    flags = 0
    if sys.platform == "win32":
        flags = subprocess.DETACHED_PROCESS | subprocess.CREATE_NEW_PROCESS_GROUP
    try:
        subprocess.Popen([app_exe] + sys.argv[1:], creationflags=flags,
                         close_fds=True, cwd=os.path.dirname(app_exe))
    except Exception as e:
        _fatal(f"启动失败：无法运行 {app_exe}\n\n{e}")
        return 1

    # 旧版本运行时目录清理（同步执行；只在版本升级后的第一次启动有东西可删）
    cleanup_old_version()
    return 0


if __name__ == "__main__":
    sys.exit(main())
