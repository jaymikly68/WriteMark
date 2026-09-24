"""端到端验证交付形态（单个 WordWatermark.exe）：

1) 首次运行会把内嵌运行体解压到 %LOCALAPPDATA%\\WriteMark\\runtime\\<版本>（带进度提示窗）
2) 点关闭时弹选择；用 WM_CLOSE_CHOICE 环境变量可模拟“选退出程序/最小化”
3) **关闭 → 进程结束**必须接近瞬时（用户感知的卡顿就在这一段：
   老 onefile 形态要 8.7s~30s，因为它要删除 130MB 的临时解包目录）
4) 启动器进程自身应很快退场，且退出后不留下任何本程序进程

用法：
    python test_launcher.py                 # 常规验证
    python test_launcher.py --clean         # 把既有 runtime 改名，模拟用户首次使用
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
VERSION = "1.1.3"


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
    """找出本程序相关进程：WriteMarkApp.exe（运行体）与 WordWatermark.exe（启动器）。"""
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
        low = path.lower()
        if low.endswith(APP_EXE_NAME.lower()) and "writemark" in low:
            out.append((pid, path, "app"))
        elif low.endswith("wordwatermark.exe"):
            out.append((pid, path, "launcher"))
    return out


def wait_for(cond, timeout=240, interval=0.05):
    t0 = time.time()
    while time.time() - t0 < timeout:
        v = cond()
        if v:
            return v, time.time() - t0
        time.sleep(interval)
    return None, time.time() - t0


def run_once(exe, first_run, choice="quit"):
    import win32con
    import win32gui
    tag = "首次" if first_run else "再次"
    print(f"\n=== {tag}运行（关闭时选择：{'退出程序' if choice == 'quit' else '最小化到后台'}）===")

    env = os.environ.copy()
    env["WM_CLOSE_CHOICE"] = choice      # 免去自动化脚本去点模态框
    t0 = time.time()
    launcher = subprocess.Popen([exe], cwd=os.path.dirname(exe), env=env)

    hwnd, dt = wait_for(find_window, timeout=240)
    print(f"双击 exe → 窗口出现：{dt:.2f}s")
    assert hwnd, "窗口未出现（启动失败）"

    # 启动器应在把应用交出去之后很快退场（它自己也是 onefile，有很小的临时目录要清）
    lt = None
    while time.time() - t0 < 60:
        if launcher.poll() is not None:
            lt = time.time() - t0
            break
        time.sleep(0.02)
    print(f"启动器进程存活：{lt:.2f}s" if lt else "⚠ 启动器进程 60s 仍未退出")
    assert lt and lt < 10, f"启动器退场过慢：{lt}s"

    time.sleep(1.5)
    procs = app_pids()
    apps = [p for p in procs if p[2] == "app"]
    print(f"当前本程序进程：{[(p, k) for p, _, k in procs]}")
    assert apps, "找不到运行体进程"

    if choice == "minimize":
        win32gui.PostMessage(hwnd, win32con.WM_CLOSE, 0, 0)
        time.sleep(2.0)
        still = [p for p in app_pids() if p[2] == "app"]
        hidden = not win32gui.IsWindowVisible(hwnd)
        print(f"选“最小化到后台”：窗口已隐藏={hidden}，进程仍在={bool(still)}")
        ok = hidden and bool(still)
        # 收尾：真退出（直接结束进程，避免影响后续用例）
        for pid, _, _ in still:
            subprocess.run(["taskkill", "/PID", str(pid), "/F"],
                           capture_output=True)
        print("判定：", "通过 ✅" if ok else "失败 ❌")
        return ok

    t_close = time.time()
    win32gui.PostMessage(hwnd, win32con.WM_CLOSE, 0, 0)
    gone, dt_close = wait_for(lambda: not [p for p in app_pids() if p[2] == "app"],
                              timeout=30, interval=0.02)
    if not gone:
        print("❌ 30 秒内进程仍未退出！")
        return False
    print(f"关闭窗口 → 进程结束：**{dt_close:.2f}s**  <<< 用户感知的卡顿")

    left = app_pids()
    print(f"退出后残留本程序进程：{len(left)} 个"
          + (f" {[(p, k) for p, _, k in left]}" if left else ""))
    mei = glob.glob(os.path.join(os.environ.get("TEMP", ""), "_MEI*"))
    ok = dt_close < 2.0 and not left
    print("判定：", "通过 ✅" if ok else f"存在问题 ❌（{dt_close:.2f}s，残留 {len(left)}）")
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

    first = not os.path.isdir(os.path.join(runtime_root(), VERSION))
    results = [
        run_once(exe, first, "quit"),        # 首次（含解压）→ 选退出程序
        run_once(exe, False, "quit"),        # 二次 → 选退出程序
        run_once(exe, False, "minimize"),    # 选最小化到后台 → 窗口隐藏、进程保留
    ]

    runtime_path = os.path.join(runtime_root(), VERSION, APP_EXE_NAME)
    print(f"\nruntime 已就绪：{os.path.isfile(runtime_path)} -> {runtime_path}")
    if os.path.isdir(runtime_root()):
        size = sum(os.path.getsize(os.path.join(d, f))
                   for d, _, fs in os.walk(runtime_root()) for f in fs)
        print(f"本地运行体占用磁盘：{size/1048576:.0f} MB（长期保留，之后启动秒开）")
    print("\n结果：", "ALL PASS ✅" if all(results) else "存在问题 ❌")
    return 0 if all(results) else 2


if __name__ == "__main__":
    sys.exit(main())
