"""验证「点关闭按钮」的新行为（不再弹询问框）：
1) 点关闭 → 默认“关闭页面不退出后台”：窗口隐藏、进程保留，守护继续跑。
2) WM_CLOSE_CHOICE=quit → 走真正退出路径：停守护、closeEvent 不阻塞。
3) WM_CLOSE_CHOICE=cancel → 窗口保持可见、守护继续跑。
4) main() 必须关闭 quitOnLastWindowClosed，否则隐藏窗口会被当成“最后窗口关闭”连带退出进程。
5) 一键插入/清除完成后必须弹「任务已完成」提示框（点确定或直接关闭均可）。
"""
import inspect
import os
import sys
import time
import tempfile
import shutil

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")
from PySide6.QtWidgets import QApplication, QMessageBox
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


def test_close_goes_background(app):
    """点关闭 = 关闭页面不退出后台：窗口隐藏、守护继续跑、无任何弹窗。"""
    w = App()
    w.show()
    d, wd = _start_watchdog(w)
    ev = QCloseEvent()
    t0 = time.time()
    w.closeEvent(ev)
    dt = time.time() - t0
    print(f"[1] 关闭耗时={dt*1000:.1f} ms，窗口隐藏={w.isHidden()}，事件被忽略={not ev.isAccepted()}")
    assert dt < 0.5, "关闭路径不应阻塞"
    assert w.isHidden() and not ev.isAccepted(), "默认应隐藏窗口到后台"
    time.sleep(0.3)
    assert wd._thread.is_alive(), "最小化到后台后守护必须继续运行"
    w.wd = None
    wd.stop()
    shutil.rmtree(d, ignore_errors=True)
    print("[1] PASS：关闭页面、后台保留、守护继续")


def test_close_without_watchdog(app):
    """未开守护时点关闭同样只隐藏窗口。"""
    w = App()
    w.show()
    w.wd = None
    ev = QCloseEvent()
    w.closeEvent(ev)
    print(f"[2] 未开守护：窗口隐藏={w.isHidden()}，事件被忽略={not ev.isAccepted()}")
    assert w.isHidden() and not ev.isAccepted()
    print("[2] PASS")


def test_env_override_quit(app):
    """WM_CLOSE_CHOICE=quit 走真正退出路径（自动化用）。"""
    os.environ["WM_CLOSE_CHOICE"] = "quit"
    try:
        w = App()
        w.show()
        d, wd = _start_watchdog(w)
        ev = QCloseEvent()
        t0 = time.time()
        w.closeEvent(ev)
        dt = time.time() - t0
        print(f"[3] 退出路径耗时={dt*1000:.1f} ms (应 < 500ms)")
        assert dt < 0.5, "退出路径阻塞过久"
        assert wd._stop.is_set(), "守护停止信号未置位"
        time.sleep(0.3)
        assert not wd._thread.is_alive(), "守护线程应已退出"
        shutil.rmtree(d, ignore_errors=True)
    finally:
        os.environ.pop("WM_CLOSE_CHOICE", None)
    print("[3] PASS：WM_CLOSE_CHOICE=quit 真正退出")


def test_env_override_cancel(app):
    """WM_CLOSE_CHOICE=cancel 时窗口保持可见。"""
    os.environ["WM_CLOSE_CHOICE"] = "cancel"
    try:
        w = App()
        w.show()
        d, wd = _start_watchdog(w)
        ev = QCloseEvent()
        w.closeEvent(ev)
        print(f"[4] cancel：窗口可见={w.isVisible()}，守护在跑={wd._thread.is_alive()}")
        assert w.isVisible() and not ev.isAccepted()
        w.wd = None
        wd.stop()
        shutil.rmtree(d, ignore_errors=True)
    finally:
        os.environ.pop("WM_CLOSE_CHOICE", None)
    print("[4] PASS")


def test_done_popup(app):
    """一键插入/清除成功后必须弹「任务已完成」；失败仍弹「失败」。"""
    w = App()
    w.show()
    seen = {"info": None, "crit": None}
    orig_info, orig_crit = QMessageBox.information, QMessageBox.critical
    QMessageBox.information = staticmethod(
        lambda parent, title, text, *a, **k: seen.__setitem__("info", (title, text)))
    QMessageBox.critical = staticmethod(
        lambda parent, title, text, *a, **k: seen.__setitem__("crit", (title, text)))
    try:
        w._on_result(True, "ok")
        assert seen["info"] == ("任务已完成", "任务已完成"), f"成功弹窗不对: {seen['info']}"
        assert seen["crit"] is None
        w._on_result(False, "出错了")
        assert seen["crit"] is not None, "失败必须弹「失败」框"
    finally:
        QMessageBox.information = orig_info
        QMessageBox.critical = orig_crit
    print("[5] PASS：任务完成弹「任务已完成」，失败弹「失败」")


def test_quit_on_last_window_closed_disabled():
    src = inspect.getsource(__import__("watermark_tool.gui", fromlist=["main"]))
    ok = "setQuitOnLastWindowClosed(False)" in src and "QApplication.quit()" in src
    print(f"[6] main() 已禁用 quitOnLastWindowClosed 且显式 quit: {ok}")
    assert ok, "不禁用会导致隐藏窗口被当成关闭而退出整个进程"
    print("[6] PASS")


if __name__ == "__main__":
    app = QApplication([])
    test_close_goes_background(app)
    test_close_without_watchdog(app)
    test_env_override_quit(app)
    test_env_override_cancel(app)
    test_done_popup(app)
    test_quit_on_last_window_closed_disabled()
    app.quit()
    print("\nALL PASS：关闭页面不退出后台 + 任务完成弹窗，行为符合预期")
