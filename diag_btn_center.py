"""诊断：按钮应只相对自己的组框居中，且位于组框正下方。"""
import os
os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")
import sys
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from PySide6.QtWidgets import QApplication
from watermark_tool.gui import App

app = QApplication.instance() or QApplication([])
w = App()
w.resize(1400, 900)
w.show()
app.processEvents()

def center_x(g):
    return g.mapToGlobal(g.rect().center()).x()

b1, b2 = w._btn_insert, w._btn_clear
# 找到两个组框：按钮的宿主通过 eventFilter 记录，直接遍历找标题
ftext = b1.parent().parent()  # btn -> hb layout -> col container
# 更稳妥：遍历顶层找 "文本水印" / "图像水印" 组
groups = {}
def walk(widget):
    for c in widget.findChildren(type(widget), None, False):
        pass
from PySide6.QtWidgets import QLabel
for lbl in w.findChildren(QLabel):
    if lbl.text() in ("文本水印", "图像水印"):
        groups[lbl.text()] = lbl.parentWidget()

ft = groups["文本水印"]; fi = groups["图像水印"]
ok1 = abs(center_x(b1) - center_x(ft)) <= 2
ok2 = abs(center_x(b2) - center_x(fi)) <= 2
# 按钮应在组框下方（底部不低于组框底部）
below1 = b1.mapToGlobal(b1.rect().topLeft()).y() >= ft.mapToGlobal(ft.rect().bottomLeft()).y() - 1
below2 = b2.mapToGlobal(b2.rect().topLeft()).y() >= fi.mapToGlobal(fi.rect().bottomLeft()).y() - 1
# 宽度 = 宿主宽 80%
import math
w1_ok = abs(b1.width() - int(ft.width() * 0.8)) <= 2
w2_ok = abs(b2.width() - int(fi.width() * 0.8)) <= 2
print(f"insert: btn_cx={center_x(b1)}, text_cx={center_x(ft)}, centered={ok1}, below={below1}, width={b1.width()} vs {int(ft.width()*0.8)} ok={w1_ok}")
print(f"clear : btn_cx={center_x(b2)}, img_cx={center_x(fi)}, centered={ok2}, below={below2}, width={b2.width()} vs {int(fi.width()*0.8)} ok={w2_ok}")
assert ok1 and ok2 and below1 and below2 and w1_ok and w2_ok, "FAIL"
print("ALL OK")
