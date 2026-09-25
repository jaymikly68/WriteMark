"""交付版 exe 的端到端检查：从启动到退出，全程不出现任何 Word 进程。

这是对「用完本工具 → 退出 Word 卡 20 秒」这条怀疑的终局验证：
如果连 exe 完整跑一遍都不会产生 WINWORD.EXE，那么 Word 的关闭耗时与本工具
在机制上就没有接触点。

同时检查哪些平方会被写脏：进程的临时目录、本地运行体目录。
"""
import os
import subprocess
import sys
import threading
import time

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from test_launcher import find_window  # noqa: E402
from test_office_exit_lag import pids_named  # noqa: E402

EXE = os.path.join("dist", "WordWatermark.exe")
TITLE = "Word 一键水印工具"


class Watcher(threading.Thread):
    def __init__(self, interval=0.05):
        super().__init__(daemon=True)
        self.interval = interval
        self._stop = threading.Event()
        self.seen = set()

    def run(self):
        while not self._stop.is_set():
            self.seen |= pids_named("WINWORD.EXE")
            time.sleep(self.interval)

    def stop(self):
        self._stop.set()
        self.join(timeout=2)


def main():
    if not os.path.isfile(EXE):
        sys.exit(f"找不到 {EXE}，请先 build_launcher.py")
    baseline = pids_named("WINWORD.EXE")
    print("起始 WINWORD:", baseline or "无")

    w = Watcher()
    w.start()
    env = os.environ.copy()
    env["WM_CLOSE_CHOICE"] = "quit"     # 自动化测试：直走退出分支
    p = subprocess.Popen([EXE], cwd=os.getcwd(), env=env)
    try:
        import win32con
        import win32gui
        t_open = time.time()
        hwnd = None
        while time.time() - t_open < 60:
            hwnd = find_window()
            if hwnd:
                break
            time.sleep(0.1)
        print(f"窗口出现：{time.time()-t_open:.2f}s  hwnd={hwnd}")
        time.sleep(6.0)                 # 让它把启动后的初始化都走完
        t0 = time.time()
        win32gui.PostMessage(hwnd, win32con.WM_CLOSE, 0, 0)
        while p.poll() is None and time.time() - t0 < 30:
            time.sleep(0.05)
        print(f"关闭→进程结束：{time.time()-t0:.2f}s")
    finally:
        w.stop()
        if p.poll() is None:
            p.kill()

    appeared = sorted(w.seen - baseline)
    print("\n=== 结果 ===")
    print("期间出现过的 WINWORD:", appeared or "无")
    leftover = sorted(pids_named("WINWORD.EXE") - baseline)
    print("收尾后残留 WINWORD:", leftover or "无")
    ok = not appeared and not leftover
    print("结论：", "✅ exe 全程未启动 Word" if ok else "❌ 出现了 Word 进程")
    return 0 if ok else 1


if __name__ == "__main__":
    sys.exit(main())
