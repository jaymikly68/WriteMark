# -*- coding: utf-8 -*-
"""对比不同启动模式下「关闭 Word 的耗时」，用来指认是谁在退出时联网。

模式：
    普通     WINWORD.EXE
    安全模式 WINWORD.EXE /safe    —— 不加载任何加载项与自定义模板
    干净     WINWORD.EXE /a       —— 使用默认设置，不加载加载项/自定义模板

若安全/干净模式秒退，而普通模式卡十几秒，那元凶就是某个加载项或 Normal.dotm；
若三种模式一样慢，说明是 Office 自身的联网心跳，需要从联网侧解决。
"""
from __future__ import annotations

import os
import subprocess
import sys
import time

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__)))))
from watermark_tool import office_tweak as ot  # noqa: E402

MODES = [
    ("普通模式", []),
    ("安全模式 /safe", ["/safe"]),
    ("干净模式 /a", ["/a"]),
]


def measure(args, settle=4.0, close_timeout=60):
    exe = ot.app_path(ot.WORD_EXE)
    before = ot.pids_named(ot.WORD_EXE)
    proc = subprocess.Popen([exe] + args, close_fds=True)
    hwnd = None
    t0 = time.time()
    while time.time() - t0 < 45:
        new = ot.pids_named(ot.WORD_EXE) - before
        if new:
            hwnd = ot._find_window(new, ot.WORD_WINDOW_CLASS)
            if hwnd:
                break
        time.sleep(0.1)
    if not hwnd:
        proc.kill()
        return None, "窗口没出现（参数 %s 可能被拒绝）" % args
    time.sleep(settle)
    u = ot._user32()
    t1 = time.time()
    u.PostMessageW(hwnd, 0x0010, 0, 0)
    dt = None
    while time.time() - t1 < close_timeout:
        if not (ot.pids_named(ot.WORD_EXE) - before):
            dt = time.time() - t1
            break
        time.sleep(0.05)
    try:
        proc.wait(3)
    except Exception:
        proc.kill()
    if dt is None:
        return float("inf"), "超过 %.0fs 仍未退出" % close_timeout
    return dt, "%.2fs" % dt


def main():
    if not ot.app_path(ot.WORD_EXE):
        print("没找到 Word")
        return 1
    running = ot.pids_named(ot.WORD_EXE)
    if running:
        print("请先关闭所有 Word（当前 PID：%s）" % sorted(running))
        return 1

    print("%-18s %s" % ("模式", "点× → 进程结束"))
    print("-" * 40)
    results = {}
    for name, args in MODES:
        dt, note = measure(args)
        results[name] = dt
        print("%-18s %s" % (name, note))
        time.sleep(1.0)

    base = results.get("普通模式")
    print("")
    for name, _ in MODES[1:]:
        v = results.get(name)
        if base and v is not None and base > 3:
            if v is None or v == float("inf"):
                print("  结论 %s：仍然卡住 → 与模板/加载项无关的 Office 自带联网" % name)
            elif v < 2.0:
                print("  结论 %s：秒退 → 元凶在模板/加载项，普通模式卡是它们引起的" % name)
            else:
                print("  结论 %s：%s，改善有限" % (name, "%.2fs" % v))
    return 0


if __name__ == "__main__":
    sys.exit(main())
