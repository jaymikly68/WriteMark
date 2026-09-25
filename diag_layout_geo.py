"""几何校验（参考图版式）：

    左侧 Word 区（竖排）            右侧
    [文本水印] [图像水印]          [视频水印]整列
               [防去除加固]  ← 贴在图像水印列正下方、与该列同宽
    [水印预览  通栏跨两列       ]
    [一键插入水印] [一键清除水印]

校验点：
  1) 防去除加固 顶边紧接 图像水印 底边，且 左缘/宽度 与图像水印一致；
  2) 水印预览 在 加固下方，左缘对齐文本水印、右缘对齐图像水印（通栏两列）；
  3) 按钮行在预览下方，且分别相对 文本/图像 列居中、同一水平线；
  4) 视频水印为独立右列（在左区右侧、横向不重叠）。
"""
import os
os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

from PySide6.QtWidgets import QApplication, QLabel

from watermark_tool.gui import App

app = QApplication.instance() or QApplication([])
w = App()
w.resize(1600, 1050)
w.show()
for _ in range(5):
    app.processEvents()

boxes = {}
for lb in w.findChildren(QLabel):
    for key in ("文本水印", "图像水印", "视频水印", "防去除加固", "水印预览"):
        if lb.text().startswith(key) and key not in boxes:
            boxes[key] = lb.parentWidget()
for b in (w._btn_insert, w._btn_clear):
    boxes[b.text()] = b

print("=== 各块几何（窗口 1600x1050）===")
geo = {}
for key in ("文本水印", "图像水印", "防去除加固", "水印预览",
            "视频水印", "一键插入水印", "一键清除水印"):
    wid = boxes.get(key)
    if wid is None:
        print(f"  [缺失] {key}")
        continue
    tl = wid.mapTo(w, wid.rect().topLeft())
    geo[key] = (tl.x(), tl.y(), wid.width(), wid.height())
    print(f"  {key:8s} x={tl.x():4d} y={tl.y():4d} w={wid.width():4d} h={wid.height():4d}")

ok = True


def check(cond, msg):
    global ok
    print(("PASS  " if cond else "FAIL  ") + msg)
    if not cond:
        ok = False


t, i, h, p = (geo.get(k) for k in ("文本水印", "图像水印", "防去除加固", "水印预览"))
bi, bc, v = (geo.get(k) for k in ("一键插入水印", "一键清除水印", "视频水印"))

check(None not in (t, i, h, p, bi, bc, v), "七块都必须存在")

if None not in (t, i, h, p):
    # 1) 加固贴在图像水印正下方、同宽
    check(abs(h[1] - (i[1] + i[3])) <= 14,
          f"加固顶边应紧接图像水印底边（{h[1]} vs {i[1] + i[3]}）")
    check(abs(h[0] - i[0]) <= 4 and abs(h[2] - i[2]) <= 4,
          f"加固应与图像水印同左缘同宽（x {h[0]}/{i[0]}，w {h[2]}/{i[2]}）")
    # 2) 预览通栏：左缘对齐文本水印，右缘对齐图像水印，且在加固下方
    check(abs(p[0] - t[0]) <= 4,
          f"预览左缘应对齐文本水印左缘（{p[0]} vs {t[0]}）")
    check(abs((p[0] + p[2]) - (i[0] + i[2])) <= 6,
          f"预览右缘应对齐图像水印右缘（{p[0] + p[2]} vs {i[0] + i[2]}）")
    check(p[1] >= h[1] + h[3] - 4,
          f"预览应在防去除加固下方（加固底 {h[1] + h[3]}，预览顶 {p[1]}）")
    check(p[2] > t[2], f"预览应比单列更宽（{p[2]} > {t[2]}）")

if None not in (bi, bc, p):
    check(bi[1] >= p[1] + p[3] - 4,
          f"按钮行应在预览下方（预览底 {p[1] + p[3]}，按钮顶 {bi[1]}）")
    check(abs(bi[1] - bc[1]) <= 2, f"两按钮应同一水平线（{bi[1]} vs {bc[1]}）")

if None not in (bi, t) and None not in (bc, i):
    ci = bi[0] + bi[2] / 2 - (t[0] + t[2] / 2)
    cc = bc[0] + bc[2] / 2 - (i[0] + i[2] / 2)
    check(abs(ci) <= 6, f"「一键插入水印」应相对文本水印列居中（偏差 {ci:.0f}px）")
    check(abs(cc) <= 6, f"「一键清除水印」应相对图像水印列居中（偏差 {cc:.0f}px）")

if v is not None and t is not None:
    left_right = max(x[0] + x[2] for x in (t, i, h, p) if x)
    check(v[0] >= left_right - 4,
          f"视频水印应为独立右列（左侧区右缘 {left_right}，视频列左缘 {v[0]}）")
    check(abs(v[1] - t[1]) <= 12, f"视频列应与 Word 两列同一顶线（{v[1]} vs {t[1]}）")

print("\nRESULT:", "ALL OK" if ok else "HAS FAILURES")
raise SystemExit(0 if ok else 1)
