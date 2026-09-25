"""对照实验：Word / Excel / PowerPoint 三者的「点×到进程结束」耗时。

用途：本工具只会通过 COM 碰过 Word，从不碰 Excel / PowerPoint。
若三者一样慢 → 说明这是本机 Office 与微软云端通信的共性问题，与本工具无关；
若只有 Word 慢 → 才需要怀疑是工具污染了 Word 的相关状态。

用法：
    python test_office_exit_lag.py              # 三个都测 1 轮
    python test_office_exit_lag.py word excel   # 只测指定的
"""
import os
import subprocess
import sys
import time

import win32api
import win32con
import win32gui
import win32process

APPS = {
    "word": ("WINWORD.EXE", "OpusApp"),
    "excel": ("EXCEL.EXE", "XLMAIN"),
    "ppt": ("POWERPNT.EXE", "PPTFrameClass"),
}


def app_path(exe_name):
    for tmpl in (r"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\%s",
                 r"SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\App Paths\%s"):
        try:
            h = win32api.RegOpenKeyEx(win32con.HKEY_LOCAL_MACHINE, tmpl % exe_name,
                                      0, win32con.KEY_READ)
            v, _ = win32api.RegQueryValueEx(h, "")
            win32api.RegCloseKey(h)
            if v and os.path.isfile(v):
                return v
        except Exception:
            continue
    return ""


def pids_named(name):
    out = set()
    lname = name.lower()
    for pid in win32process.EnumProcesses():
        h = None
        try:
            h = win32api.OpenProcess(
                win32con.PROCESS_QUERY_INFORMATION | win32con.PROCESS_VM_READ, False, pid)
            path = win32process.GetModuleFileNameEx(h, win32process.EnumProcessModules(h)[0])
        except Exception:
            continue
        finally:
            if h:
                try:
                    win32api.CloseHandle(h)
                except Exception:
                    pass
        if os.path.basename(path).lower() == lname:
            out.add(pid)
    return out


def wait_window(exe_name, win_class, before, timeout=60):
    t0 = time.time()
    while time.time() - t0 < timeout:
        new = pids_named(exe_name) - before
        if new:
            found = {"h": None}

            def cb(h, _):
                try:
                    if win32gui.GetClassName(h) != win_class:
                        return
                    if not win32gui.IsWindowVisible(h):
                        return
                    _, pid = win32process.GetWindowThreadProcessId(h)
                    if pid in new:
                        found["h"] = h
                except Exception:
                    pass
            try:
                win32gui.EnumWindows(cb, None)
            except Exception:
                pass
            if found["h"]:
                return found["h"], time.time() - t0
        time.sleep(0.1)
    return None, time.time() - t0


def measure(app, settle=4.0, timeout=90):
    exe_name, win_class = APPS[app]
    exe = app_path(exe_name)
    if not exe:
        print(f"  {app}: 未安装，跳过")
        return None
    before = pids_named(exe_name)
    if before:
        print(f"  {app}: 已有运行实例 {before}，跳过")
        return None

    p = subprocess.Popen([exe], close_fds=True)
    hwnd, t_open = wait_window(exe_name, win_class, before, timeout=45)
    if hwnd is None:
        print(f"  {app}: 窗口未出现")
        p.kill()
        return None
    title = win32gui.GetWindowText(hwnd)[:36]
    time.sleep(settle)

    new_pids = pids_named(exe_name) - before
    t0 = time.time()
    win32gui.PostMessage(hwnd, win32con.WM_CLOSE, 0, 0)
    dt = None
    while time.time() - t0 < timeout:
        if not (pids_named(exe_name) - before) & new_pids:
            dt = time.time() - t0
            break
        time.sleep(0.05)
    if dt is None:
        print(f"  {app:6} 启动 {t_open:.2f}s  点×→进程退出：>{timeout}s 未退出 ⚠")
    else:
        verdict = "慢！" if dt > 3 else "正常"
        print(f"  {app:6} 启动 {t_open:.2f}s  点×→进程退出 {dt:.2f}s   {verdict}"
              f"   窗口={title!r}")
    left = pids_named(exe_name) - before
    if left:
        print(f"         ⚠ 残留 {left}")
    try:
        p.wait(2)
    except Exception:
        p.kill()
    return dt


def main():
    apps = [a for a in sys.argv[1:] if a in APPS] or list(APPS)
    print("=== Office 关闭耗时对照 ===")
    res = {}
    for a in apps:
        d = measure(a)
        if d is not None:
            res[a] = d
        time.sleep(1.0)
    print("\n=== 结果 ===")
    for a, d in res.items():
        print(f"  {a:6} {d:.2f}s")
    if len(res) >= 2:
        slow = {k: v for k, v in res.items() if v > 3}
        if len(slow) == len(res):
            print("\n判据：全部都慢 → 这是本机 Office 自身的问题（关闭时要等微软云端连接），"
                  "与本工具无关 —— 本工具从未调用 Excel/PowerPoint。")
        elif "word" in slow and len(slow) == 1:
            print("\n判据：只有 Word 慢 → 才需要怀疑工具污染了 Word 的状态。")
        else:
            print("\n判据：部分慢，需要进一步排查。")


if __name__ == "__main__":
    main()
