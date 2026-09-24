"""实测「关闭窗口 → 进程真正结束」的耗时，并区分是 Python 侧卡住还是 onefile 临时目录清理卡住。

用法：
    python test_exit_timing.py exe            # 测 dist/WordWatermark.exe
    python test_exit_timing.py src            # 测源码版 pythonw main.py
    python test_exit_timing.py src --console  # 源码版带控制台（可看到清理日志）
"""
import glob
import os
import subprocess
import sys
import time

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

TITLE = "Word 一键水印工具"


def _import_win32():
    try:
        import win32gui
        import win32con
        return win32gui, win32con
    except ImportError:
        return None, None


def _snapshot_mei():
    tmp = os.environ.get("TEMP", "")
    if not tmp:
        return set()
    return set(glob.glob(os.path.join(tmp, "_MEI*")))


def wait_window(proc, timeout=90, exclude=()):
    win32gui, _ = _import_win32()
    if win32gui is None:
        time.sleep(8)
        return None
    t0 = time.time()
    found = {"hwnd": None}

    def _enum(hwnd, _):
        if hwnd in exclude:
            return
        if not win32gui.IsWindowVisible(hwnd):
            return
        title = win32gui.GetWindowText(hwnd) or ""
        cls = win32gui.GetClassName(hwnd) or ""
        # 用「标题 + Qt 窗口类」双重判定，避免 pid 在包装层下取不准
        if TITLE in title and cls.startswith("Qt"):
            found["hwnd"] = hwnd

    while time.time() - t0 < timeout:
        found["hwnd"] = None
        try:
            win32gui.EnumWindows(_enum, None)
        except Exception:
            pass
        if found["hwnd"]:
            return found["hwnd"]
        if proc.poll() is not None:
            return None
        time.sleep(0.3)
    return None


def win32process_get(hwnd):
    try:
        import win32process
        return win32process.GetWindowThreadProcessId(hwnd)
    except Exception:
        return 0, -1


def measure(cmd, cwd, label):
    before = _snapshot_mei()
    t0 = time.time()
    proc = subprocess.Popen(cmd, cwd=cwd)
    hwnd = wait_window(proc)
    t_up = time.time()
    if hwnd is None:
        print(f"[{label}] 未找到窗口，进程状态={proc.poll()}")
        proc.kill()
        return
    print(f"[{label}] 窗口出现耗时 {t_up - t0:.2f}s (hwnd={hwnd})")

    time.sleep(2.0)  # 让后台线程稳定

    win32gui, win32con = _import_win32()
    t_close = time.time()
    prev = _snapshot_mei() - before
    win32gui.PostMessage(hwnd, win32con.WM_CLOSE, 0, 0)

    # 轮询：窗口消失时间、进程退出时间、_MEI 目录消失时间
    win_gone = None
    proc_gone = None
    mei_gone_at = None
    while time.time() - t_close < 120:
        now = time.time()
        if win_gone is None and not win32gui.IsWindow(hwnd):
            win_gone = now
        if proc_gone is None and proc.poll() is not None:
            proc_gone = now
        if prev:
            alive = [d for d in prev if os.path.isdir(d)]
            if not alive and mei_gone_at is None:
                mei_gone_at = now
            prev = alive
        if proc_gone is not None and (not prev) and (mei_gone_at is not None or True):
            if proc_gone is not None:
                break
        time.sleep(0.1)
    end = time.time()
    print(f"[{label}] 关闭信号发出 → 窗口消失: "
          f"{(win_gone - t_close):.2f}s" if win_gone else f"[{label}] 窗口未消失")
    if proc_gone:
        print(f"[{label}] 关闭信号发出 → 进程结束: {(proc_gone - t_close):.2f}s   <<< 用户感知的卡顿")
    else:
        print(f"[{label}] 进程仍在运行（>120s），强制结束")
        proc.kill()
    if mei_gone_at:
        print(f"[{label}] _MEI 临时目录清理完成耗时: {(mei_gone_at - t_close):.2f}s")
    print(f"[{label}] 总耗时 {(end - t_close):.2f}s")


def main():
    mode = sys.argv[1] if len(sys.argv) > 1 else "exe"
    here = os.path.dirname(os.path.abspath(__file__))
    if mode == "exe":
        exe = os.path.join(here, "dist", "WordWatermark.exe")
        measure([exe], here, "EXE(onefile)")
    else:
        py = os.path.join(here, "venv", "Scripts",
                          "python.exe" if "--console" in sys.argv else "pythonw.exe")
        measure([py, os.path.join(here, "main.py")], here, "SRC(source)")


if __name__ == "__main__":
    main()
