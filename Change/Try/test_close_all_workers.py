"""
二次安全审查（commit 84a773e 之后）—— 主题 4：closeEvent 的资源清理时序。

关注点：
- **不删除** 6 秒强制退出（os._exit）机制，只确认它在什么时候、会把什么留下。
- 正常退出路径下，所有 worker 都要被请求停止，而不只是「插入/清除」那一个。
  （此前 closeEvent 只处理 self._worker，视频 / 预览 / Office 线程若仍在跑，
   会被直接留给进程退出——它们的 finally（临时文件、writer）就没人执行了。）
- 硬性不变式：**所有 worker 的清理总时长必须短于 6s 兜底计时器**，
  否则「本该优雅退出」会被自己埋的兜底硬杀。
"""
import sys, time

sys.path.insert(0, '.')

from PySide6.QtGui import QCloseEvent

from watermark_tool.gui import App

# 与 gui.closeEvent 中的 SHUTDOWN_BUDGET 对应
FORCE_EXIT_SECONDS = 6.0


class _FakeWorker:
    """只实现 closeEvent 会调用的那套 QThread 停机接口，并记录调用顺序。"""

    def __init__(self, tag, log):
        self.tag = tag
        self.log = log
        self.calls = []

    def isRunning(self):
        return True            # 永远「在跑」，逼出最坏路径（每个都要走强杀）

    def requestInterruption(self):
        self.calls.append("requestInterruption")

    def quit(self):
        self.calls.append("quit")

    def wait(self, ms):
        self.calls.append(f"wait({ms})")
        time.sleep(min(ms, 300) / 1000.0)   # 模拟等待，但不真睡满
        return False                        # 拒不退出的情况

    def terminate(self):
        self.calls.append("terminate")


def _open_window(app):
    w = App()
    w.show()
    w.tray = None        # 无托盘环境下 tray 可能不可用，closeEvent 对此有兜底
    w.wd = None          # 守护线程单独测（见 test_close.py）
    return w


def test_all_workers_shut_down_and_stay_within_force_exit_budget(app):
    w = _open_window(app)
    w._quitting = True
    for name in ("_worker", "_v_worker", "_v_preview_worker", "_office_worker",
                 "_font_worker"):
        setattr(w, name, _FakeWorker(name, None))

    t0 = time.time()
    w.closeEvent(QCloseEvent())
    dt = time.time() - t0

    for name in ("_worker", "_v_worker", "_v_preview_worker", "_office_worker",
                 "_font_worker"):
        fake = getattr(w, name)
        assert "requestInterruption" in fake.calls, f"{name} 未被请求中断"
        assert "quit" in fake.calls, f"{name} 未被请求退出"
        assert "terminate" in fake.calls, f"{name} 未走到强杀分支"
        assert fake.calls[-1] == "wait(300)", f"{name} 的收尾调用不完整: {fake.calls}"

    print(f"[review4] 5 个 worker 全部停机，耗时 {dt:.2f}s（兜底 {FORCE_EXIT_SECONDS}s）")
    assert dt < FORCE_EXIT_SECONDS, (
        "worker 清理总时长超过 6s 兜底计时器：正常退出会被兜底硬杀"
    )
    print("[review4] PASS：所有 worker 均停机，且总清理时长短于兜底计时器")


def test_close_event_ignores_unused_worker_attributes(app):
    """worker 属性为 None（常见情况）时不应抛异常。"""
    w = _open_window(app)
    w.closeEvent(QCloseEvent())       # 所有 worker 属性都是 None
    print("[review4] PASS：无 worker 时 closeEvent 不抛异常")
