"""
验证「关闭窗口 = 退出程序」后的界面行为：
1) 一键插入/清除完成后必须弹「任务已完成」提示框（点确定或直接关闭均可）；
2) 失败仍弹「失败」。
"""
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")
from PySide6.QtWidgets import QMessageBox

from watermark_tool.gui import App


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


if __name__ == "__main__":
    import sys as _sys
    from PySide6.QtWidgets import QApplication
    app = QApplication([])
    test_done_popup(app)
    app.quit()
    print("\nALL PASS：任务完成弹窗行为符合预期")
