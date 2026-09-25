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

APP_VERSION = "1.3.3"
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


STAMP_NAME = "payload.stamp"


def _swap_into_place(tmp: str, dst: str):
    """把解压好的 tmp 目录原子改名为 dst。

    踩过的坑：用户旧实例还在运行时占用了 dst 目录（exe/dll 被锁），
    os.replace 直接 WinError 5「拒绝访问」，用户只看到一个莫名其妙的报错。
    所以这里重试一段时间，仍失败就给出「请先退出正在运行的程序」的明确指引。
    """
    import time
    last_err = None
    for _ in range(10):
        if os.path.isdir(dst):
            shutil.rmtree(dst, ignore_errors=True)
        if not os.path.isdir(dst):
            try:
                os.replace(tmp, dst)
                return
            except OSError as e:
                last_err = e
        else:
            last_err = OSError(5, "拒绝访问（旧运行环境被占用）")
        time.sleep(0.6)
    raise RuntimeError(
        "无法更新运行环境：旧版本的程序可能还在运行，占用了文件。\n"
        "请先完全退出「一键水印工具」（包括任务栏托盘图标），再重新打开。\n\n"
        f"技术细节：{last_err}"
    )


def payload_fingerprint() -> str:
    """当前 payload 的指纹（大小 + 修改时间）。

    踩过的坑：runtime 目录是按 APP_VERSION 命名的，版本号不变就永远不重新解压。
    于是「明明重新构建并发布了新的 exe，用户双击却还是跑着旧代码」——
    遇到过一次（GUI 崩在 NameError 上，用户反馈"程序直接打不开"，
    查了半天才发现他手里的新 exe 一直在复用 4 分钟前解压的旧运行时）。
    所以这里用 payload 自身的指纹做陈旧检测，跟版本号彻底解耦。
    """
    try:
        st = os.stat(payload_zip_path())
        return "%d-%d" % (st.st_size, int(st.st_mtime))
    except OSError as e:
        return "err:%s" % e


def ensure_app(update) -> str:
    dst = runtime_dir()
    app_exe = os.path.join(dst, APP_EXE_NAME)
    stamp_file = os.path.join(dst, STAMP_NAME)

    def _reinstall():
        """重新解压 payload 到 dst。调用前必须自行保证 dst 不存在。"""
        zip_path = payload_zip_path()
        if not os.path.isfile(zip_path):
            raise RuntimeError(f"找不到内嵌的运行体压缩包：{zip_path}")
        update("正在准备运行环境（首次使用，仅需一次）…")
        tmp = dst + ".tmp"
        shutil.rmtree(tmp, ignore_errors=True)
        os.makedirs(tmp, exist_ok=True)
        total = 0
        with zipfile.ZipFile(zip_path) as z:
            names = z.namelist()
            total = len(names)
            for i, name in enumerate(names, 1):
                z.extract(name, tmp)
                if i % 40 == 0 or i == total:
                    update(f"正在准备运行环境（首次使用，仅需一次）… {i}/{total}")
        with open(os.path.join(tmp, STAMP_NAME), "w", encoding="utf-8") as f:
            f.write(payload_fingerprint())
        update("解压完成，正在启动…")
        try:
            _swap_into_place(tmp, dst)
        except Exception:
            shutil.rmtree(tmp, ignore_errors=True)  # 失败时清掉残留的 .tmp
            raise
        if not os.path.isfile(os.path.join(dst, APP_EXE_NAME)):
            raise RuntimeError(f"解压后找不到主程序：{os.path.join(dst, APP_EXE_NAME)}\n"
                               f"请删除 {dst} 后重试。")
        return os.path.join(dst, APP_EXE_NAME)

    # 指纹对不上 = 这一版 payload 对应不上这份运行时，必须重来
    if os.path.isfile(stamp_file):
        try:
            with open(stamp_file, encoding="utf-8") as f:
                if f.read().strip() == payload_fingerprint() and os.path.isfile(app_exe):
                    return app_exe
        except OSError:
            pass
    if os.path.isfile(app_exe):
        # 没有 stamp 的旧运行时（更早版本的产物）一律重建，避免沿用未知年代的代码
        shutil.rmtree(dst, ignore_errors=True)

    return _reinstall()


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


def runtime_is_ready() -> bool:
    """运行时目录是否真的、且是最新一版的。用于决定要不要显示进度窗口。"""
    dst = runtime_dir()
    stamp_file = os.path.join(dst, STAMP_NAME)
    if not (os.path.isfile(os.path.join(dst, APP_EXE_NAME))
            and os.path.isfile(stamp_file)):
        return False
    try:
        with open(stamp_file, encoding="utf-8") as f:
            return f.read().strip() == payload_fingerprint()
    except OSError:
        return False


def main() -> int:
    need_gui = not runtime_is_ready()
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
