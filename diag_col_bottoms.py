"""等高对齐的稳定性校验：多档窗口尺寸下

  1) 文本水印列 / 图像水印列 底边始终平行；
  2) 「Word 秒退」框与「防去除加固」框 顶边、底边始终平行；
  3) 不出现「反复 setMinimumHeight → Resize → 再 set」的抖动（sync 写入次数应很快归零）。
"""
import os
os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

from PySide6.QtWidgets import QApplication, QLabel

from watermark_tool.gui import App

app = QApplication.instance() or QApplication([])
w = App()
w.show()

calls = {"set": 0}


def patch(eq):
    for widget in (eq._a, eq._b):
        real = widget.setMinimumHeight

        def spy(h, _real=real):
            calls["set"] += 1
            _real(h)
        widget.setMinimumHeight = spy


patch(w._col_eq)
patch(w._band_eq)

ok = True
for size in [(1600, 1050), (1280, 900), (1100, 800), (1920, 1200), (1400, 1000), (1750, 1150)]:
    w.resize(*size)
    calls["set"] = 0
    for _ in range(40):          # 多跑几轮事件循环：若同步抖动，这里会持续写入
        app.processEvents()
    box = {}
    for lb in w.findChildren(QLabel):
        for key in ("文本水印", "图像水印", "Word 秒退", "防去除加固"):
            if lb.text().startswith(key) and key not in box:
                box[key] = lb.parentWidget()

    def geo(k):
        wid = box[k]
        tl = wid.mapTo(w, wid.rect().topLeft())
        return tl.y(), tl.y() + wid.height()

    tt, tb = geo("文本水印")
    it, ib = geo("图像水印")
    et, eb = geo("Word 秒退")
    ht, hb = geo("防去除加固")
    cols = abs(tb - ib) <= 4
    band_tops = abs(et - ht) <= 4
    band_bots = abs(eb - hb) <= 4
    # 两条同步（两列底边 / 秒退-加固 顶底）各写一次两个控件 → 首轮最多 4 次写入；
    # 之后必须归零，否则就是 Resize↔setMinimumHeight 的抖动。
    stable = calls["set"] <= 4
    good = cols and band_tops and band_bots and stable
    ok = ok and good
    print(f"{'PASS' if good else 'FAIL'}  窗口 {size}: 两列底 {tb}/{ib}  |  "
          f"秒退 {et}-{eb} vs 加固 {ht}-{hb}  |  40 轮内 sync 写入 {calls['set']} 次")

print("\nRESULT:", "ALL OK" if ok else "HAS FAILURES")
raise SystemExit(0 if ok else 1)
