"""诊断：两个主按钮应（1）各自只相对自己的组框居中、（2）处于同一水平线（平行）、
（3）位于组框正下方、（4）宽度=宿主框宽 80%。"""
import os
os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")
import sys
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from PySide6.QtWidgets import QApplication, QLabel

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__)))))
from watermark_tool.gui import App

app = QApplication.instance() or QApplication([])
w = App()
w.resize(1400, 900)
w.show()
app.processEvents()

def center_x(g):
    return g.mapToGlobal(g.rect().center()).x()

b1, b2 = w._btn_insert, w._btn_clear
groups = {}
for lbl in w.findChildren(QLabel):
    if lbl.text() in ("文本水印", "图像水印"):
        groups[lbl.text()] = lbl.parentWidget()
ft = groups["文本水印"]; fi = groups["图像水印"]

c1 = abs(center_x(b1) - center_x(ft)) <= 4
c2 = abs(center_x(b2) - center_x(fi)) <= 4
b1y = b1.mapToGlobal(b1.rect().topLeft()).y()
b2y = b2.mapToGlobal(b2.rect().topLeft()).y()
parallel = abs(b1y - b2y) <= 2          # 同 Y：平行
below1 = b1y >= ft.mapToGlobal(ft.rect().bottomLeft()).y() - 1
below2 = b2y >= fi.mapToGlobal(fi.rect().bottomLeft()).y() - 1
w1 = abs(b1.width() - int(ft.width() * 0.8)) <= 4
w2 = abs(b2.width() - int(fi.width() * 0.8)) <= 4

print(f"insert: btn_cx={center_x(b1)}, text_cx={center_x(ft)}, centered={c1}, y={b1y}, below={below1}, w={b1.width()} ok={w1}")
print(f"clear : btn_cx={center_x(b2)}, img_cx={center_x(fi)}, centered={c2}, y={b2y}, below={below2}, w={b2.width()} ok={w2}")
print(f"parallel(same y)={parallel}")
assert c1 and c2 and parallel and below1 and below2 and w1 and w2, "FAIL"
print("ALL OK")
