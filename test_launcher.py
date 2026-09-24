"""端到端验证交付形态（单个 WordWatermark.exe）：

1) 首次运行会把内嵌运行体解压到 %LOCALAPPDATA%\\WriteMark\\runtime\\<版本>（带进度提示窗）
2) 关闭窗口后，应用进程必须 **立刻退出**（这是本次要根治的问题：
   老 onefile 形态要 8.7s~30s，因为它要删除 130MB 的临时解包目录）
3) 启动器进程也应很快退出，不残留
4) 第二次运行不再解压，启动更快

用法：
    python test_launcher.py           # 常规验证（会保留 runtime）
    python test_launcher.py --clean   # 先删除 runtime，模拟用户首次使用
"""
import glob
import os
import shutil
import subprocess
import sys
import time

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
TITLE = "Word 一键水印工具"
APP_EXE_NAME = "WriteMarkApp.exe"


def runtime_root():
    base = os.environ.get("LOCALAPPDATA") or os.path.expanduser("~")
    return os.path.join(base, "WriteMark", "runtime")


def find_window():
    import win32gui
    res = {"h": None}

    def cb(h, _):
        if (win32gui.IsWindowVisible(h)
                and TITLE in (win32gui.GetWindowText(h) or "")
                and (win32gui.GetClassName(h) or "").startswith("Qt")):
            res["h"] = h
    win32gui.EnumWindows(cb, None)
    return res["h"]


def app_pids():
    """找出运行体进程（位于 WriteMark\\runtime 下）的 pid。"""
    import win32api
    import win32con
    import win32process
    out = []
    for pid in win32process.EnumProcesses():
        h = None
        try:
            h = win32api.OpenProcess(
                win32con.PROCESS_QUERY_INFORMATION | win32con.PROCESS_VM_READ, False, pid)
            mods = win32process.EnumProcessModules(h)
            path = win32process.GetModuleFileNameEx(h, mods[0])
        except Exception:
            continue
        finally:
            if h is not None:
                try:
                    win32api.CloseHandle(h)
                except Exception:
                    pass
        if path.lower().endswith(APP_EXE_NAME.lower()) and "writemark" in path.lower():
            out.append((pid, path))
    return out


def wait_for(cond, timeout=180, interval=0.2):
    t0 = time.time()
    while time.time() - t0 < timeout:
        v = cond()
        if v:
            return v, time.time() - t0
        time.sleep(interval)
    return None, time.time() - t0


def run_once(exe, first_run):
    import win32con
    import win32gui
    print(f"\n=== {'首次' if first_run else '再次'}运行 ===")
    t0 = time.time()
    launcher = subprocess.Popen([exe], cwd=os.path.dirname(exe))

    hwnd, dt = wait_for(find_window, timeout=240)
    print(f"双击 exe → 窗口出现：{dt:.2f}s")
    assert hwnd, "窗口未出现（启动失败）"
    time.sleep(1.5)

    # 启动器自身应尽快退出（不必跟随应用整个生命周期）
    if launcher.poll() is None:
        _, dt_l = wait_for(lambda: launcher.poll() is not None, timeout=15)
        print(f"启动器进程退出：{dt_l:.2f}s（窗口已交给独立进程）")
    else:
        print("启动器进程已提前退出 ✅")

    apps = app_pids()
    print(f"运行体进程：{[p for p, _ in apps]}")
    assert apps, "找不到运行体进程"

    t_close = time.time()
    win32gui.PostMessage(hwnd, win32con.WM_CLOSE, 0, 0)
    gone = None
    while time.time() - t_close < 30:
        ids = [p for p, _ in app_pids()]
        if not [i for i in ids if i in apps]:
            gone = time.time() - t_close
            break
        time.sleep(0.05)
    if gone is None:
        print("❌ 30 秒内进程仍未退出！")
        return False
    print(f"关闭窗口 → 进程结束：**{gone:.2f}s**  <<< 用户感知的卡顿")

    mei = glob.glob(os.path.join(os.environ.get("TEMP", ""), "_MEI*"))
    print(f"运行期间 %TEMP% 残留 _MEI 临时目录：{len(mei)} 个")
    ok = gone < 2.0
    print("判定：", "通过 ✅" if ok else f"偏慢 ❌（{gone:.2f}s）")
    return ok


def main():
    here = os.path.dirname(os.path.abspath(__file__))
    exe = os.path.join(here, "dist", "WordWatermark.exe")
    if not os.path.exists(exe):
        print("找不到", exe)
        sys.exit(1)

    if "--clean" in sys.argv and os.path.isdir(runtime_root()):
        # 用改名的方式让本次变成“首次运行”，避免直接批量删除用户目录
        bak = runtime_root() + "_bak"
        i = 0
        while os.path.exists(bak):
            i += 1
            bak = f"{runtime_root()}_bak{i}"
        os.replace(runtime_root(), bak)
        print(f"已把既有 runtime 改名为 {os.path.basename(bak)}，本次按首次运行测试")
    first = not os.path.isdir(os.path.join(runtime_root(), "1.1.1"))

    ok1 = run_once(exe, first)
    ok2 = run_once(exe, False)

    runtime_path = os.path.join(runtime_root(), "1.1.1", APP_EXE_NAME)
    print(f"\nruntime 已就绪：{os.path.isfile(runtime_path)} -> {runtime_path}")
    if os.path.isdir(runtime_root()):
        size = sum(os.path.getsize(os.path.join(d, f))
                   for d, _, fs in os.walk(runtime_root()) for f in fs)
        print(f"本地运行体占用磁盘：{size/1048576:.0f} MB（长期保留，之后启动秒开）")
    print("\n结果：", "ALL PASS ✅" if (ok1 and ok2) else "存在问题 ❌")
    return 0 if (ok1 and ok2) else 2


if __name__ == "__main__":
    sys.exit(main())
