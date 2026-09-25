"""渲染当前主窗口为 PNG，便于核对各分组块的实际位置（offscreen，不受遮挡）。"""
import os
os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

from PySide6.QtWidgets import QApplication

from watermark_tool.gui import App

app = QApplication.instance() or QApplication([])
w = App()
w.resize(1500, 1050)
w.show()
for _ in range(3):
    app.processEvents()

out = os.path.join(os.path.dirname(os.path.abspath(__file__)), "shot_current.png")
w.grab().save(out, "PNG")
print("saved:", out, "size:", w.size().width(), w.size().height())
