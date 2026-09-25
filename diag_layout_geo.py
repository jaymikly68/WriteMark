"""几何校验：防去除加固 与 水印预览 是否「并排」在三列下方、且互不重叠。"""
import os
os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

from PySide6.QtWidgets import QApplication, QLabel, QPushButton

from watermark_tool.gui import App

app = QApplication.instance() or QApplication([])
w = App()
w.resize(1500, 1050)
w.show()
for _ in range(3):
    app.processEvents()

# 组框是「QLabel 标题 + 内容」的自绘容器，标题的父控件即组框
boxes = {}
for lb in w.findChildren(QLabel):
    for key in ("文本水印", "图像水印", "视频水印", "防去除加固", "水印预览",
                "后台守护", "Word 秒退", "日志"):
        if lb.text().startswith(key) and key not in boxes:
            g = lb.parentWidget()
            boxes[key] = (g, lb)
for b in (w._btn_insert, w._btn_clear):
    boxes[b.text()] = (b, b)

print("=== 各块几何（窗口 1500x1050）===")
geo = {}
for key in ("文本水印", "图像水印", "视频水印", "一键插入水印", "一键清除水印",
            "防去除加固", "水印预览", "后台守护"):
    if key not in boxes:
        print(f"  [缺失] {key}")
        continue
    wid = boxes[key][0]
    tl = wid.mapTo(w, wid.rect().topLeft())
    r = (tl.x(), tl.y(), wid.width(), wid.height())
    geo[key] = r
    print(f"  {key:12s} x={r[0]:4d} y={r[1]:4d} w={r[2]:4d} h={r[3]:4d}")

ok = True
def check(cond, msg):
    global ok
    print(("PASS  " if cond else "FAIL  ") + msg)
    if not cond:
        ok = False

hard, prev = geo.get("防去除加固"), geo.get("水印预览")
cols_bottom = max(geo[k][1] + geo[k][3] for k in ("文本水印", "图像水印", "视频水印"))
btn_bottom = max(geo[k][1] + geo[k][3] for k in ("一键插入水印", "一键清除水印"))
check(hard is not None and prev is not None, "两块都必须存在")
if hard and prev:
    check(abs(hard[1] - prev[1]) <= 12, f"两块顶边应基本齐平（{hard[1]} vs {prev[1]}）")
    check(prev[0] > hard[0] + hard[2] - 8,
          f"水印预览应在防去除加固右侧（加固右缘 {hard[0]+hard[2]}，预览左缘 {prev[0]}）")
    check(prev[2] > hard[2], f"预览应更宽（{prev[2]} > {hard[2]}）")
    y_top = min(hard[1], prev[1])
    check(y_top >= btn_bottom - 4,
          f"两块应位于按钮行下方（按钮底 {btn_bottom}，两块顶 {y_top}）")
    check(y_top >= cols_bottom - 4,
          f"两块应位于三列下方（三列底 {cols_bottom}，两块顶 {y_top}）")

print("\nRESULT:", "ALL OK" if ok else "HAS FAILURES")
raise SystemExit(0 if ok else 1)
