"""验证「点关闭按钮」的行为：
1) 守护运行时点关闭 → 弹询问；选“最小化到后台”应只隐藏窗口，守护继续跑，进程不退出。
2) 选“退出程序” → 停止守护、closeEvent 不阻塞。
3) 选“取消” → 窗口保持可见、守护继续跑。
4) main() 必须关闭 quitOnLastWindowClosed，否则隐藏窗口会被当成“最后窗口关闭”连带退出进程。
"""
import inspect
import os
import sys
import time
import tempfile
import shutil

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")
from PySide6.QtWidgets import QApplication
from PySide6.QtGui import QCloseEvent

from watermark_tool.gui import App
from watermark_tool import watchdog


def make_src():
    d = tempfile.mkdtemp()
    src = os.path.join(d, "src.docx")
    from docx import Document
    Document().save(src)
    return d, src


def _start_watchdog(w):
    d, src = make_src()
    out = os.path.join(d, "out.docx")
    wd = watchdog.WatermarkWatchdog(src, out, "text", {"text": "x"},
                                    interval=1.0, log=lambda *a, **k: None)
    wd.start()
    w.wd = wd
    time.sleep(0.2)
    return d, wd


def test_minimize(app):
    w = App()
    w.show()
    d, wd = _start_watchdog(w)
    w._ask_close_choice = lambda: "minimize"
    ev = QCloseEvent()
    w.closeEvent(ev)
    print(f"[1] 选“最小化到后台”: 窗口隐藏={w.isHidden()}, event={ev.isAccepted()} (应 False=忽略关闭)")
    assert w.isHidden() is True, "应隐藏窗口而不是销毁"
    assert not ev.isAccepted(), "最小化时不应接受关闭事件"
    time.sleep(0.3)
    alive = wd._thread.is_alive()
    print(f"[1] 守护线程仍在运行={alive} (应 True)")
    assert alive, "最小化到后台后守护必须继续运行"
    w.wd = None
    wd.stop()
    shutil.rmtree(d, ignore_errors=True)
    print("[1] PASS")


def test_cancel(app):
    w = App()
    w.show()
    d, wd = _start_watchdog(w)
    w._ask_close_choice = lambda: "cancel"
    ev = QCloseEvent()
    w.closeEvent(ev)
    print(f"[2] 选“取消”: 窗口可见={w.isVisible()}, 守护在跑={wd._thread.is_alive()}")
    assert w.isVisible() and not w.isHidden(), "取消后窗口应保持原样"
    assert not ev.isAccepted()
    w.wd = None
    wd.stop()
    shutil.rmtree(d, ignore_errors=True)
    print("[2] PASS")


def test_ask_even_without_watchdog(app):
    """不开守护时点关闭也必须弹询问（用户明确要求）。"""
    w = App()
    w.show()
    w.wd = None
    called = {"v": False}

    def _fake():
        called["v"] = True
        return "cancel"
    w._ask_close_choice = _fake
    ev = QCloseEvent()
    w.closeEvent(ev)
    print(f"[3] 未开守护也弹询问={called['v']} (应 True)")
    assert called["v"], "未开启守护时也必须弹窗询问"
    assert w.isVisible(), "选了取消应保持窗口"
    print("[3] PASS")


def test_minimize_without_watchdog(app):
    """未开守护时选“最小化到后台”应隐藏窗口且不退出进程。"""
    w = App()
    w.show()
    w.wd = None
    w._ask_close_choice = lambda: "minimize"
    ev = QCloseEvent()
    w.closeEvent(ev)
    print(f"[4] 未开守护选最小化：窗口隐藏={w.isHidden()}，事件被忽略={not ev.isAccepted()}")
    assert w.isHidden() and not ev.isAccepted()
    print("[4] PASS")


def test_env_override(app):
    """WM_CLOSE_CHOICE 环境变量可替代模态框（供自动化使用）。"""
    os.environ["WM_CLOSE_CHOICE"] = "quit"
    try:
        w = App()
        w.show()
        assert w._ask_close_choice() == "quit"
        os.environ["WM_CLOSE_CHOICE"] = "minimize"
        assert w._ask_close_choice() == "minimize"
    finally:
        os.environ.pop("WM_CLOSE_CHOICE", None)
    print("[5] PASS：WM_CLOSE_CHOICE 可指定选择，自动化无需点模态框")


def test_quit(app):
    w = App()
    w.show()
    d, wd = _start_watchdog(w)
    w._ask_close_choice = lambda: "quit"
    ev = QCloseEvent()
    t0 = time.time()
    w.closeEvent(ev)
    dt = time.time() - t0
    print(f"[6] 选“退出程序”: 耗时={dt*1000:.1f} ms (应 < 200ms)")
    assert dt < 0.5, "退出路径阻塞过久"
    assert wd._stop.is_set(), "守护停止信号未置位"
    time.sleep(0.3)
    print(f"[6] 守护线程已退出={not wd._thread.is_alive()} (应 True)")
    assert not wd._thread.is_alive()
    shutil.rmtree(d, ignore_errors=True)
    print("[6] PASS")


def test_auto_script_no_modal(app):
    """自动化脚本关闭 _confirm_close 后不应弹出阻塞式对话框。"""
    w = App()
    w.show()
    d, wd = _start_watchdog(w)

    called = {"v": False}

    def _boom():
        called["v"] = True
        assert False, "closeEvent 不应弹询问（_confirm_close=False）"
    w._ask_close_choice = _boom
    w._confirm_close = False
    w._quitting = True          # 明确退出
    w.closeEvent(QCloseEvent())
    print(f"[7] _confirm_close=False 时未弹窗={not called['v']}")
    assert not called["v"]
    w.wd = None
    wd.stop()
    shutil.rmtree(d, ignore_errors=True)
    print("[7] PASS")


def test_quit_on_last_window_closed_disabled():
    src = inspect.getsource(App.__module__ and __import__("watermark_tool.gui", fromlist=["main"]))
    ok = "setQuitOnLastWindowClosed(False)" in src and "QApplication.quit()" in src
    print(f"[8] main() 已禁用 quitOnLastWindowClosed 且显式 quit: {ok}")
    assert ok, "不禁用会导致隐藏窗口被当成关闭而退出整个进程"
    print("[8] PASS")


if __name__ == "__main__":
    app = QApplication([])
    test_minimize(app)
    test_cancel(app)
    test_ask_even_without_watchdog(app)
    test_minimize_without_watchdog(app)
    test_env_override(app)
    test_quit(app)
    test_auto_script_no_modal(app)
    test_quit_on_last_window_closed_disabled()
    app.quit()
    print("\nALL PASS：关闭按钮行为（最小化 / 退出 / 取消，开或不开守护）符合预期")
