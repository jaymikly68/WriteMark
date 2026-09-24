# -*- coding: utf-8 -*-
"""命令行入口：诊断 / 修复「联网时关闭 Word 卡顿十几秒」。

真正的逻辑在 watermark_tool/office_tweak.py，这里只做命令行包装，
保证 GUI 里的按钮和命令行走的是完全同一套实现。

用法
----
    python office_exit_fix.py status            只读体检
    python office_exit_fix.py apply             轻量档：禁用卡顿加载项 + 关遥测
    python office_exit_fix.py apply --strong    强力档（已弃用，见下）
    python office_exit_fix.py reapply-optional  可选加码：会踢登录账号，慎用
    python office_exit_fix.py revert            一键还原
    python office_exit_fix.py bench [次数]      量化「点× → 进程结束」耗时
"""
from __future__ import annotations

import os
import sys

try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    sys.stderr.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from watermark_tool import office_tweak as ot  # noqa: E402


def cmd_status():
    d = ot.diagnose()
    print("=== Word 退出卡顿 · 体检报告 ===")
    print("Office 版本分支：%s" % d["version"])
    print("\n[1] Microsoft 账户登录状态")
    if d["logged_in"]:
        print("    已登录 OneAuth 账户：%s…" % d["account"])
        print("    → 已排除：实测改这一项无效（改完 bench 仍有 18s）。")
        print("      真凶是下面 [2] 的 COM 加载项。默认修复【不会】动你的登录状态。")
    else:
        print("    未发现登录标记")
    print("\n[2] Word 加载项（实测卡顿的真正元凶所在）")
    if not d["addins"]:
        print("    注册表里没有相关加载项")
    for name, lb, desc, auto in d["addins"]:
        if auto:
            print("    ⚠ %-34s LoadBehavior=%-2s 会随 Word 启动自动加载 —— %s" % (name, lb, desc))
        else:
            print("      %-34s LoadBehavior=%-2s 不自动加载 —— %s" % (name, lb, desc))
    if d["auto_addins"]:
        print("    → 自动加载的有：%s" % ", ".join(d["auto_addins"]))
        print("    → 这类插件关闭 Word 时要跟云端收尾，联网就等十几秒，断网立刻秒退。")

    print("\n[3] 免管理员开关（HKCU\\Software\\Microsoft\\Office）")
    for it in d["items"]:
        if it["realm"] != "user":
            continue
        cur = "未设置" if it["current"] == ot.MISSING else "已设置=%r" % (it["current"],)
        print("    %-38s %-16s 期望=%-2r  %s" % (it["name"], cur, it["want"], it["desc"]))
    print("\n[4] 组策略开关（HKCU\\Software\\Policies\\Microsoft\\Office）—— 可选加码")
    if not d["policy_writable"]:
        print("    该分支被系统 ACL 锁住，普通权限无法写入（需要管理员）。")
        print("    用户级开关已能覆盖主要问题，策略级只是锦上添花。")
    for it in d["items"]:
        if it["realm"] != "policy":
            continue
        cur = "未设置" if it["current"] == ot.MISSING else "已设置=%r" % (it["current"],)
        print("    %-38s %-16s 期望=%-2r  %s" % (it["name"], cur, it["want"], it["desc"]))
    print("\n[5] 备份文件")
    print("    %s" % ("已存在：%s" % ot.BACKUP_FILE if d["backup"] else "尚无备份"))
    print("\n建议：apply 之后 bench 对比，平均 <2s 即为秒退。")
    return 0


def cmd_apply(strong):
    r = ot.apply_fix(strong=strong)
    print(r["msg"])
    if r["need_admin"]:
        print("\n需要管理员权限：请右键「以管理员身份运行」这条命令，或用 GUI 里的按钮触发 UAC。")
    return 0 if r["ok"] else 1


def cmd_revert():
    r = ot.revert_fix()
    print(r["msg"])
    if r["need_admin"]:
        print("\n需要管理员权限才能还原。")
    return 0 if r["ok"] else 1


def cmd_bench(rounds):
    times = []
    for i in range(rounds):
        dt, note = ot.benchmark()
        if dt is None:
            print("  %s" % note)
            return 1
        times.append(dt)
        print("  第 %d 轮：%s" % (i + 1, note))
        if i + 1 < rounds:
            import time
            time.sleep(1.0)
    if times:
        avg = sum(times) / len(times)
        verdict = "✅ 已修复" if avg < 2 else ("⚠ 仍慢" if avg > 5 else "有改善")
        print("\n平均 %.2fs（%d 次）   %s" % (avg, len(times), verdict))
    return 0


def main():
    args = sys.argv[1:]
    cmd = args[0] if args else "status"
    if cmd == "status":
        return cmd_status()
    if cmd == "apply":
        if "--strong" in args:
            print("⚠ --strong 已弃用：那个档位包含「禁止 Office 登录账户」，")
            print("  实测会把已登录的 Word 账号踢下线，默认不再开启。")
            print("  需要联网限制的话用 reapply-optional，并自行确认后果。")
        return cmd_apply(strong=("--strong" in args))
    if cmd == "reapply-optional":
        r = ot.apply_optional()
        for row in r[0]:
            print("  %-38s %-4s %s" % (row[1], "OK" if row[4] else "失败", row[3]))
        if r[1]:
            print("⚠ 写入失败：%s" % ", ".join(r[1]))
        return 0 if not r[1] else 1
    if cmd == "revert":
        return cmd_revert()
    if cmd == "bench":
        n = 1
        for a in args[1:]:
            if a.isdigit():
                n = max(1, int(a))
        return cmd_bench(n)
    print(__doc__)
    return 0


if __name__ == "__main__":
    sys.exit(main())
