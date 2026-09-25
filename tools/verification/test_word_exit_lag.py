"""量化「关闭 Word 需要多久」，用于定位用户反馈的「退出 Word 卡 20 秒」。

做法和用户实际操作一致：普通启动 WINWORD.EXE（不是 COM），等窗口就绪，
发 WM_CLOSE（等价点标题栏 ×），计时到进程结束。

用法：
    python test_word_exit_lag.py              # 测 2 轮 + 打印环境线索
    python test_word_exit_lag.py --rounds 4   # 指定轮数
"""
import os
import subprocess
import sys
import time

try:
    import win32api
    import win32con
    import win32gui
    import win32print
    import win32process
except ImportError:
    sys.exit("需要 pywin32（项目 venv 内已装）")

WORD_WINDOW_CLASS = "OpusApp"


# ---------------------------------------------------------------- 环境线索
def winword_exe() -> str:
    for key, sub in (
        (win32con.HKEY_LOCAL_MACHINE,
         r"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\Winword.exe"),
        (win32con.HKEY_LOCAL_MACHINE,
         r"SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\App Paths\Winword.exe"),
    ):
        try:
            h = win32api.RegOpenKeyEx(key, sub, 0, win32con.KEY_READ)
            val, _ = win32api.RegQueryValueEx(h, "")
            win32api.RegCloseKey(h)
            if val and os.path.isfile(val):
                return val
        except Exception:
            continue
    return ""


def winword_pids() -> set:
    out = set()
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
        if os.path.basename(path).lower() == "winword.exe":
            out.add(pid)
    return out


def env_clues():
    print("\n=== 环境线索 ===")
    print("WINWORD.EXE      :", winword_exe() or "(未安装 Word)")
    try:
        print("默认打印机        :", win32print.GetDefaultPrinter())
    except Exception as e:
        print("默认打印机        : 读取失败", e)

    def reg_val(root, path, name):
        try:
            h = win32api.RegOpenKeyEx(root, path, 0, win32con.KEY_READ)
            v, _ = win32api.RegQueryValueEx(h, name)
            win32api.RegCloseKey(h)
            return v
        except Exception:
            return None

    base = r"Software\Microsoft\Office"
    for ver in ("16.0", "15.0", "17.0"):
        root = r"%s\%s\Word\Data" % (base, ver)
        for name in ("Settings", "UserData", "Toolbars", "Customization"):
            v = reg_val(win32con.HKEY_CURRENT_USER, root, name)
            if v is not None:
                size = len(v) if isinstance(v, (bytes, bytearray)) else -1
                print(f"HKCU\\...\\{ver}\\Word\\Data\\{name:<14}: {size} 字节")

    appdata = os.environ.get("APPDATA", "")
    for rel, desc in (
        (r"Microsoft\Templates\Normal.dotm", "Normal.dotm"),
        (r"Microsoft\Word\STARTUP", "Word STARTUP 目录"),
    ):
        p = os.path.join(appdata, rel)
        if os.path.isfile(p):
            print(f"{desc:<18}: {p}  ({os.path.getsize(p)/1024:.0f} KB, "
                  f"修改 {time.strftime('%m-%d %H:%M', time.localtime(os.path.getmtime(p)))})")
        elif os.path.isdir(p):
            items = os.listdir(p)
            print(f"{desc:<18}: {p}  共 {len(items)} 项 -> {items[:6]}")

    local = os.environ.get("LOCALAPPDATA", "")
    for rel in (r"Microsoft\Office\16.0\OfficeFileCache",
                r"Microsoft\Office\OTele",
                r"Microsoft\Office\16.0\WEF"):
        p = os.path.join(local, rel)
        if os.path.isdir(p):
            n = sum(len(f) for _, _, f in os.walk(p))
            print(f"缓存目录 {rel:<40}: {n} 个文件")


# ---------------------------------------------------------------- 测量
def wait_word_window(pids_before, timeout=60):
    """等待属于本次启动的 Word 窗口出现；返回 (hwnd, 耗时)。"""
    t0 = time.time()
    last = None
    while time.time() - t0 < timeout:
        now = winword_pids()
        new = now - pids_before
        if new:
            found = {"h": None}

            def cb(h, _):
                if h in (None,):
                    return
                try:
                    if win32gui.GetClassName(h) != WORD_WINDOW_CLASS:
                        return
                except Exception:
                    return
                try:
                    if not win32gui.IsWindowVisible(h):
                        return
                except Exception:
                    return
                _, pid = win32process.GetWindowThreadProcessId(h)
                if pid in new:
                    found["h"] = h
            try:
                win32gui.EnumWindows(cb, None)
            except Exception:
                pass
            if found["h"]:
                return found["h"], time.time() - t0
        last = new
        time.sleep(0.1)
    return None, time.time() - t0


def measure(round_no, settle=4.0, timeout=90):
    before = winword_pids()
    exe = winword_exe()
    if not exe:
        sys.exit("没找到 WINWORD.EXE")

    print(f"\n--- 第 {round_no} 轮 ---")
    p = subprocess.Popen([exe], close_fds=True)
    hwnd, t_open = wait_word_window(before, timeout=45)
    if hwnd is None:
        print("  Word 窗口未出现（可能已有实例接管），跳过")
        return None
    title = win32gui.GetWindowText(hwnd)
    print(f"  窗口出现        : {t_open:.2f}s   标题={title!r}")

    time.sleep(settle)  # 让 Word 完成启动后的后台工作（加载项/遥测）

    new_pids = winword_pids() - before
    t0 = time.time()
    win32gui.PostMessage(hwnd, win32con.WM_CLOSE, 0, 0)

    t_window = None
    t_proc = None
    while time.time() - t0 < timeout:
        now = winword_pids()
        if t_window is None:
            alive = now & new_pids
            still_up = False
            for _, _, _ign in ():
                pass
            still_up = bool(alive)
            if not still_up:
                t_proc = time.time() - t0
                break
        time.sleep(0.05)

    if t_proc is None:
        print(f"  ⚠ 等待 {timeout}s 进程仍未退出")
    else:
        mark = "慢！" if t_proc > 3 else "正常"
        print(f"  点×到进程结束  : {t_proc:.2f}s   {mark}")

    left = winword_pids() - before
    if left:
        print(f"  ⚠ 仍有残留 WINWORD: {left}")
    return t_proc


def main():
    rounds = 2
    if "--rounds" in sys.argv:
        rounds = int(sys.argv[sys.argv.index("--rounds") + 1])
    left = winword_pids()
    if left:
        print(f"⚠ 开始前已有 WINWORD 进程: {left}  —— 先手工关闭它们再测，结果才准")
    env_clues()
    results = []
    for i in range(1, rounds + 1):
        r = measure(i)
        if r is not None:
            results.append(r)
    if results:
        print(f"\n=== 结论：{len(results)} 轮平均 {sum(results)/len(results):.2f}s，"
              f"最长 {max(results):.2f}s ===")
        print("判据：< 2s 正常；> 5s 说明 Word 关闭路径被阻塞，需看上面环境线索")


if __name__ == "__main__":
    main()
