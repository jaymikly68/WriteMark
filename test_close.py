"""
验证“退出不再卡顿”：
1) closeEvent 不应阻塞：此前 stop() 内 join(timeout=3) 会卡最多 3 秒。
2) 守护线程（daemon）在窗口关闭后能自行退出。
3) 插入进行中关闭窗口，工作线程应被强制终止而非永久挂起。
"""
import sys, os, time, tempfile, shutil
sys.path.insert(0, '.')

from PySide6.QtWidgets import QApplication
from PySide6.QtGui import QCloseEvent
from watermark_tool.gui import App, Worker
from watermark_tool import watchdog, core


def make_src():
    d = tempfile.mkdtemp()
    src = os.path.join(d, "src.docx")
    # 用纯 Python 引擎造一个最简单的 docx（借助 python-docx）
    from docx import Document
    Document().save(src)
    return d, src


def test_watchdog_close_no_block(app):
    w = App()
    w.show()
    d, src = make_src()
    out = os.path.join(d, "out.docx")
    wd = watchdog.WatermarkWatchdog(src, out, "text", {"text": "x"},
                                    interval=1.0, log=lambda *a, **k: None)
    wd.start()
    w.wd = wd
    time.sleep(0.2)  # 让守护跑一轮

    # 关闭按钮现在会弹“最小化到后台 / 退出程序 / 取消”询问（模态框），
    # 自动化脚本里必须显式表达“要退出”，否则默认是隐藏到后台。
    w._quitting = True
    t0 = time.time()
    w.closeEvent(QCloseEvent())
    dt = time.time() - t0
    print(f"[1] closeEvent duration = {dt*1000:.1f} ms (应 < 100ms)")
    assert dt < 0.5, "closeEvent 阻塞过久，仍是卡顿根因"
    assert wd._stop.is_set(), "stop 信号未置位"

    time.sleep(0.3)
    print(f"[1] watchdog thread alive after close = {wd._thread.is_alive()} (应 False)")
    assert not wd._thread.is_alive(), "守护线程未退出"
    shutil.rmtree(d, ignore_errors=True)
    print("[1] PASS：守护关闭不卡顿、线程自动退出")


def test_worker_terminate_on_close(app):
    w = App()
    w.show()
    # 一个故意跑 6 秒的长任务（模拟大文件/卡住）
    def long_task():
        time.sleep(6)
        return "done"
    w._worker = Worker(long_task)
    w._worker.start()
    time.sleep(0.2)

    t0 = time.time()
    w._quitting = True          # 测的是“真正退出”路径下的线程强杀（点关闭默认是隐藏到后台）
    w.closeEvent(QCloseEvent())
    dt = time.time() - t0
    print(f"[2] closeEvent during running worker = {dt*1000:.1f} ms (应约 2000~2600ms, 非 6000ms)")
    assert dt < 4.0, "工作线程未被及时终止，仍存在卡顿"
    assert not w._worker.isRunning(), "工作线程仍 running"
    print("[2] PASS：进行中插入被强制终止，进程可退出")


if __name__ == "__main__":
    app = QApplication([])
    test_watchdog_close_no_block(app)
    test_worker_terminate_on_close(app)
    app.quit()
    print("\nALL PASS：退出卡顿问题已修复")
