"""Word COM 收尾工具：确保不会留下“幽灵 Word”进程。

背景（真实踩坑，用户感知为“退出 Word 卡半天”）：
用 win32com 起一个 Word（读字体、处理 .doc）后，即使调用了 `word.Quit()`，
Word 进程有时仍留在后台——COM 引用没释放、Word 还在启动中、文档还在加载都会
导致 Quit 被忽略。用户随后去关这个看不见的 Word 实例，就会觉得“退出 Word 卡顿”。

这里做三件事：
1) 记录调用前的 WINWORD.EXE 进程集合，用来区分“我们自己新起的”；
2) 调 Quit 之后显式释放引用并 gc，再给它一小段时间自行退出；
3) 仍然不退的，只强制结束“我们自己新起的”那几个 pid（绝不碰用户自己开着的那份）。
"""
from __future__ import annotations

import gc
import subprocess
import sys
import time

_CREATE_NO_WINDOW = 0x08000000


def winword_pids() -> set:
    """当前所有 WINWORD.EXE 进程的 pid 集合。"""
    try:
        import win32api
        import win32con
        import win32process
    except Exception:
        return set()

    pids = set()
    access = win32con.PROCESS_QUERY_INFORMATION | win32con.PROCESS_VM_READ
    for pid in win32process.EnumProcesses():
        h = None
        try:
            h = win32api.OpenProcess(access, False, pid)
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
        if str(path).lower().endswith("winword.exe"):
            pids.add(pid)
    return pids


def kill_pids(pids) -> list:
    """强制结束给定 pid（仅用于我们自己启动、且 Quit 无效的 Word）。"""
    killed = []
    for pid in pids:
        try:
            args = ["taskkill", "/PID", str(pid), "/F", "/T"]
            kwargs = {}
            if sys.platform == "win32":
                kwargs["creationflags"] = _CREATE_NO_WINDOW
            r = subprocess.run(args, capture_output=True, **kwargs)
            if r.returncode == 0:
                killed.append(pid)
        except Exception:
            pass
    return killed


def quit_word(word, before_pids=frozenset(), grace=2.5, log=None) -> list:
    """退出 Word 并确认它真的退出了；返回被强制结束的 pid 列表。

    before_pids：调用 Word COM 之前 `winword_pids()` 的快照，
    用来避免误杀用户自己打开着的 Word。
    """
    _log = log or (lambda *a, **k: None)

    if word is not None:
        try:
            word.Quit()
        except Exception as e:
            _log(f"Word.Quit() 未成功：{e}")

    # 释放 COM 引用：Word 会等所有外部引用消失后才真正退出，
    # 只调用 Quit 而对象仍被引用时，进程常常会赖着不走。
    try:
        del word
    except Exception:
        pass
    try:
        gc.collect()
    except Exception:
        pass

    mine = winword_pids() - set(before_pids)
    if not mine:
        return []

    deadline = time.time() + max(0.0, float(grace))
    while time.time() < deadline:
        mine = winword_pids() - set(before_pids)
        if not mine:
            return []
        time.sleep(0.2)

    killed = kill_pids(mine)
    if killed:
        _log(f"已结束残留的 Word 进程 {sorted(killed)}（COM 未能正常退出）")
    return sorted(killed)
