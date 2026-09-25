import sys
"""对齐稳定性校验：多档窗口尺寸下

  1) 文本水印列 / 图像水印列 底边始终平行；
  2) 「Word 秒退」框与「防去除加固」框 顶边、底边始终平行；
  3) 视频水印框底边 = 左侧「水印预览」框底边；
  4) 「视频水印预览」的帧把多余空间吃掉（明显高于最小高 203）；
  5) 不出现「反复 set 高度 → Resize → 再 set」的抖动（sync 写入次数应很快归零）。
"""
import os
os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

from PySide6.QtWidgets import QApplication, QLabel


sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__)))))
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
# 视频底边同步走的是 setFixedHeight（给整列定高）
_real_fixed = w._v_bottom_eq._fol.setFixedHeight


def spy_fixed(h):
    calls["set"] += 1
    _real_fixed(h)


w._v_bottom_eq._fol.setFixedHeight = spy_fixed

ok = True
for size in [(1600, 1050), (1280, 900), (1100, 800), (1920, 1200), (1400, 1000), (1750, 1150)]:
    w.resize(*size)
    for _ in range(10):          # 先让新尺寸下的同步收敛
        app.processEvents()
    calls["set"] = 0
    for _ in range(40):          # 再跑多轮：若同步抖动，这里会持续写入
        app.processEvents()
    box = {}
    for lb in w.findChildren(QLabel):
        for key in ("文本水印", "图像水印", "Word 秒退", "防去除加固", "水印预览"):
            if lb.text().startswith(key) and key not in box:
                box[key] = lb.parentWidget()
    vbox = w.v_frame_label.parentWidget().parentWidget()      # 视频水印组框
    pbox = box["水印预览"]

    def geo(wid):
        tl = wid.mapTo(w, wid.rect().topLeft())
        return tl.y(), tl.y() + wid.height()

    tt, tb = geo(box["文本水印"])
    it, ib = geo(box["图像水印"])
    et, eb = geo(box["Word 秒退"])
    ht, hb = geo(box["防去除加固"])
    vt, vb = geo(vbox)
    pt, pb = geo(pbox)
    cols = abs(tb - ib) <= 4
    band_tops = abs(et - ht) <= 4
    band_bots = abs(eb - hb) <= 4
    vbottom = abs(vb - pb) <= 4
    stretched = w.v_frame_label.height() > 240          # 最小 203，能被拉伸才算生效
    # 收敛后再跑 40 轮应当几乎零写入；若 Resize↔set 来回抖，这里会持续增长。
    stable = calls["set"] <= 2
    good = cols and band_tops and band_bots and vbottom and stretched and stable
    ok = ok and good
    print(f"{'PASS' if good else 'FAIL'}  窗口 {size}: 两列底 {tb}/{ib}  |  "
          f"秒退 {et}-{eb} vs 加固 {ht}-{hb}  |  视频底 {vb} vs 预览底 {pb}  |  "
          f"帧高 {w.v_frame_label.height()}  |  收敛后 sync 写入 {calls['set']} 次")

print("\nRESULT:", "ALL OK" if ok else "HAS FAILURES")
raise SystemExit(0 if ok else 1)
