"""在 Word 关闭的这段时间里采样，判断它卡在什么资源上。

每 0.4s 记录一次 WINWORD 的 CPU 时间、磁盘 IO、TCP 连接数、可见对话框，
结束后打印时间线 —— 据此区分三种典型原因：

- CPU≈0、IO≈0、有处于 SYN_SENT/ESTABLISHED 的外网连接 → 在等网络超时（Office 遥测/授权）
- IO 持续上涨 → 在写缓存/模板
- 出现 #32770 之类对话框 → 在弹隐藏提示等你响应
"""
import os
import subprocess
import sys
import threading
import time

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from test_word_exit_lag import winword_exe, winword_pids, wait_word_window  # noqa: E402
import win32api
import win32con
import win32gui
import win32process


class Sample:
    __slots__ = ("t", "cpu", "io_rw", "tcp", "dialogs")

    def __init__(self, t, cpu, io_rw, tcp, dialogs):
        self.t = t
        self.cpu = cpu
        self.io_rw = io_rw
        self.tcp = tcp
        self.dialogs = dialogs


def my_connections(pid, cache={"t": 0, "v": {}}):
    """用 netstat 查某 pid 的 TCP 连接（较慢，结果缓存 2 秒）。"""
    now = time.time()
    if now - cache["t"] < 2.0:
        return cache["v"].get(pid, [])
    cache["t"] = now
    try:
        out = subprocess.run(["netstat", "-ano", "-p", "TCP"], capture_output=True,
                             text=True, timeout=6).stdout
    except Exception:
        return cache["v"].get(pid, [])
    table = {}
    for line in out.splitlines():
        parts = line.split()
        if len(parts) < 5 or not parts[-1].isdigit():
            continue
        proto, local, foreign, state = parts[0], parts[1], parts[2], parts[3]
        try:
            table.setdefault(int(parts[-1]), []).append(f"{state} {local} -> {foreign}")
        except Exception:
            pass
    cache["v"] = table
    return table.get(pid, [])


def find_dialogs(pid):
    out = []

    def cb(h, _):
        try:
            if not win32gui.IsWindowVisible(h):
                return
            cls = win32gui.GetClassName(h)
            _, p = win32process.GetWindowThreadProcessId(h)
            if p == pid and cls in ("#32770", "NUIDialog", "NetUIHWND", "WordSubDlg"):
                out.append((cls, win32gui.GetWindowText(h)[:40]))
        except Exception:
            pass
    try:
        win32gui.EnumWindows(cb, None)
    except Exception:
        pass
    return out


def main():
    before = winword_pids()
    p = subprocess.Popen([winword_exe()])
    hwnd, t_open = wait_word_window(before, timeout=45)
    if hwnd is None:
        sys.exit("Word 窗口未出现")
    print(f"窗口 {t_open:.2f}s 出现，等待 5s 让它就绪…")
    time.sleep(5.0)

    pids = sorted(winword_pids() - before)
    print("本次 Word pid:", pids)
    main_pid = pids[0]

    def proc_times(pid):
        h = win32api.OpenProcess(win32con.PROCESS_QUERY_INFORMATION, False, pid)
        try:
            t = win32process.GetProcessTimes(h)
            io = win32process.GetProcessIoCounters(h)
        finally:
            win32api.CloseHandle(h)
        cpu = (t["UserTime"] + t["KernelTime"]) / 1e7
        io_rw = io["ReadOperationCount"] + io["WriteOperationCount"]
        return cpu, io_rw

    samples = []
    t0 = time.time()
    print("发送 WM_CLOSE，开始采样…")
    win32gui.PostMessage(hwnd, win32con.WM_CLOSE, 0, 0)

    while time.time() - t0 < 60:
        alive = [x for x in pids if x in winword_pids()]
        if not alive:
            print(f"进程在时间 {(time.time()-t0):.2f}s 退出")
            break
        try:
            cpu, io = proc_times(alive[0])
        except Exception:
            time.sleep(0.2)
            continue
        conns = my_connections(alive[0])
        ext = [c for c in conns
               if "127.0.0.1" not in c and "[::1]" not in c]
        samples.append(Sample(time.time() - t0, cpu, io, ext[:8], find_dialogs(alive[0])))
        time.sleep(0.4)

    print("\n=== 时间线（只在连接状态变化时展开明细）===")
    prev = None
    prev_tcp = None
    for s in samples:
        cpu_d = (s.cpu - prev.cpu) if prev else 0.0
        io_d = (s.io_rw - prev.io_rw) if prev else 0
        changed = tuple(s.tcp) != (prev_tcp or ())
        print(f"  t={s.t:5.2f}s  CPU+{cpu_d*1000:6.1f}ms  IO+{io_d:5d}  "
              f"连接={len(s.tcp)}  对话框={s.dialogs if s.dialogs else ''}")
        if s.tcp and changed:
            for c in s.tcp:
                print(f"          {c}")
        prev = s
        prev_tcp = tuple(s.tcp)
    if not samples:
        print("  （进程退出太快，没采到样本）")

    left = winword_pids() - before
    if left:
        print("残留 pid:", left)
    try:
        p.wait(3)
    except Exception:
        p.kill()


if __name__ == "__main__":
    main()
