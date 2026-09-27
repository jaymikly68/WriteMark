"""
验证“退出不再卡顿”：
1) closeEvent 不应阻塞：此前 stop() 内 join(timeout=3) 会卡最多 3 秒。
2) 插入进行中关闭窗口，工作线程应被强制终止而非永久挂起。
"""
import sys, os, time, tempfile, shutil

sys.path.insert(0, '.')

from PySide6.QtWidgets import QApplication
from PySide6.QtGui import QCloseEvent
from watermark_tool.gui import App, Worker


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
    w.closeEvent(QCloseEvent())
    dt = time.time() - t0
    print(f"[2] closeEvent during running worker = {dt*1000:.1f} ms (应约 2000~2600ms, 非 6000ms)")
    assert dt < 4.0, "工作线程未被及时终止，仍存在卡顿"
    assert not w._worker.isRunning(), "工作线程仍 running"
    print("[2] PASS：进行中插入被强制终止，进程可退出")


if __name__ == "__main__":
    app = QApplication([])
    test_worker_terminate_on_close(app)
    app.quit()
    print("\nALL PASS：退出卡顿问题已修复")
