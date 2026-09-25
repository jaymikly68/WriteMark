# -*- coding: utf-8 -*-
"""逐个禁用 Word 的 COM 加载项，找出「关闭时联网收尾」的那一个。

完全可逆：每一步都先备份原 LoadBehavior，实验结束后自动复原。
用法：python diag_addin_ab.py [轮次]
"""
from __future__ import annotations

import os
import sys
import time
import winreg

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__)))))
from watermark_tool import office_tweak as ot  # noqa: E402

HK = winreg.HKEY_CURRENT_USER
ADDIN_ROOT = r"Software\Microsoft\Office\Word\AddIns"

TARGETS = [
    "MSOfficePLUS",
    "YunOfficeAddin.YunWordConnect",
]


def read_lb(name):
    try:
        k = winreg.OpenKey(HK, ADDIN_ROOT + "\\" + name, 0, winreg.KEY_READ)
    except OSError:
        return None
    try:
        v, _ = winreg.QueryValueEx(k, "LoadBehavior")
        return int(v)
    except OSError:
        return None
    finally:
        winreg.CloseKey(k)


def set_lb(name, value):
    try:
        k = winreg.CreateKey(HK, ADDIN_ROOT + "\\" + name)
    except OSError:
        return False
    try:
        winreg.SetValueEx(k, "LoadBehavior", 0, winreg.REG_DWORD, value)
        return True
    except OSError:
        return False
    finally:
        winreg.CloseKey(k)


def bench_once():
    return ot.benchmark(settle=3.0, open_timeout=40, close_timeout=60)


def main():
    rounds = 1
    for a in sys.argv[1:]:
        if a.isdigit():
            rounds = max(1, int(a))

    if ot.pids_named(ot.WORD_EXE):
        print("请先关闭所有 Word")
        return 1

    # 备份
    backup = {}
    for n in TARGETS:
        lb = read_lb(n)
        if lb is not None:
            backup[n] = lb
            print("备份 %-32s LoadBehavior = %s" % (n, lb))
    if not backup:
        print("注册表里没找到任何加载项，退出")
        return 1

    print("\n=== 基线（全部加载项原样）===")
    base_vals = []
    for _ in range(rounds):
        dt, note = bench_once()
        base_vals.append(dt)
        print("   %s" % note)
        time.sleep(1)
    base = sum(base_vals) / len(base_vals)
    print("   基线平均 %.2fs\n" % base)

    results = {}
    for n in backup:
        print("=== 禁用了 %s（LoadBehavior→0）===" % n)
        set_lb(n, 0)
        vals = []
        for _ in range(rounds):
            dt, note = bench_once()
            vals.append(dt)
            print("   %s" % note)
            time.sleep(1)
        v = sum(vals) / len(vals)
        results[n] = v
        print("   平均 %.2fs   （基线 %.2fs）\n" % (v, base))

    print("=== 复原中 ===")
    for n, lb in backup.items():
        if set_lb(n, lb):
            print("   已复原 %-32s LoadBehavior = %s" % (n, lb))

    print("\n================ 结论 ================")
    if base > 5:
        for n, v in sorted(results.items(), key=lambda kv: kv[1]):
            tag = "★ 命中" if v < 2.0 else "无改善"
            print("   %-32s %6.2fs   %s" % (n, v, tag))
    else:
        print("   基线本来就快，无需处理")
    print("   实验结束后所有设置已复原。")
    return 0


if __name__ == "__main__":
    sys.exit(main())
