"""底边对齐的稳定性校验：多档窗口尺寸下，文本/图像两列底边应始终平行，
且不会出现「反复 setMinimumHeight → Resize → 再 set」的抖动（sync 计数不应无限增长）。"""
import os
os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

from PySide6.QtWidgets import QApplication, QLabel

from watermark_tool.gui import App

app = QApplication.instance() or QApplication([])
w = App()
w.show()

# 统计 sync 的实际写入次数
calls = {"set": 0}
orig = App.__init__


def patch(eq):
    real = eq._a.setMinimumHeight

    def spy(h):
        calls["set"] += 1
        real(h)
    eq._a.setMinimumHeight = spy
    real_b = eq._b.setMinimumHeight

    def spyb(h):
        calls["set"] += 1
        real_b(h)
    eq._b.setMinimumHeight = spyb


patch(w._col_eq)

ok = True
for size in [(1600, 1050), (1280, 900), (1100, 800), (1920, 1200), (1400, 1000)]:
    w.resize(*size)
    calls["set"] = 0
    for _ in range(40):          # 多跑几轮事件循环：若同步抖动，这里会持续写入
        app.processEvents()
    box = {}
    for lb in w.findChildren(QLabel):
        for key in ("文本水印", "图像水印"):
            if lb.text().startswith(key) and key not in box:
                box[key] = lb.parentWidget()
    t, i = box["文本水印"], box["图像水印"]
    bt = t.mapTo(w, t.rect().topLeft()).y() + t.height()
    bi = i.mapTo(w, i.rect().topLeft()).y() + i.height()
    good = abs(bt - bi) <= 4 and calls["set"] <= 2
    ok = ok and good
    print(f"{'PASS' if good else 'FAIL'}  窗口 {size}: 文本列底 {bt}, 图像列底 {bi} "
          f"(差 {bt - bi}), 40 轮事件循环内 sync 写入 {calls['set']} 次")

print("\nRESULT:", "ALL OK" if ok else "HAS FAILURES")
raise SystemExit(0 if ok else 1)
