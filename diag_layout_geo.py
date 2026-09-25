"""几何校验（v1.4.3 版式）：

    左侧 Word 区（竖排）                  右侧视频列
    [Word 文件 ………… 浏览]                [源视频 ………… 浏览]
    [输出   ………… 浏览]                  [输出   ………… 浏览]
    [文本水印] [图像水印]                [视频水印（独立）]
    [Word 秒退] [防去除加固]  ← 并排成一条水平带：顶边同线、底边同线
    [水印预览  通栏跨两列       ]
    [一键插入水印] [一键清除水印]

校验点：
  1) Word 输入/输出行的「浏览」按钮右缘与「图像水印」组框右缘竖向对齐（两行已缩短）；
  2) 文本水印组框【底边】与图像水印组框【底边】平行（同一水平线）；
  3) Word 秒退框在文本水印框【正下方】、与文本水印框同左缘同宽；
     且其顶边/底边与右侧「防去除加固」框的顶边/底边平行；
  4) 防去除加固 顶边紧接 图像水印 底边，且 左缘/宽度 与图像水印一致；
  5) 水印预览 在两块下方，左缘对齐文本水印、右缘对齐图像水印（通栏两列）；
  6) 按钮行在预览下方，且分别相对 文本/图像 列居中、同一水平线；
  7) 视频的源/输出行在「视频水印」组框【上方】，且右缘与组框一致（列宽不变）。
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
for _ in range(5):            # 等高同步用 QTimer.singleShot(0) 延迟生效
    app.processEvents()

boxes = {}
for lb in w.findChildren(QLabel):
    for key in ("文本水印", "图像水印", "视频水印", "防去除加固", "水印预览", "Word 秒退"):
        if lb.text().startswith(key) and key not in boxes:
            boxes[key] = lb.parentWidget()
for b in (w._btn_insert, w._btn_clear):
    boxes[b.text()] = b
# 输入/输出行容器：由编辑框往上找父 QWidget
boxes["Word输入行"] = w.file_edit.parentWidget()
boxes["Word输出行"] = w.out_edit.parentWidget()
boxes["视频输入行"] = w.v_src_edit.parentWidget()
boxes["视频输出行"] = w.v_out_edit.parentWidget()

print("=== 各块几何（窗口 1600x1050）===")
geo = {}
keys = ("Word输入行", "Word输出行", "视频输入行", "视频输出行",
        "文本水印", "图像水印", "Word 秒退", "防去除加固", "水印预览",
        "视频水印", "一键插入水印", "一键清除水印")
for key in keys:
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


fi, fo, vfi, vfo = (geo.get(k) for k in ("Word输入行", "Word输出行", "视频输入行", "视频输出行"))
t, i, ex, h, p, v = (geo.get(k) for k in ("文本水印", "图像水印", "Word 秒退",
                                          "防去除加固", "水印预览", "视频水印"))
bi, bc = (geo.get(k) for k in ("一键插入水印", "一键清除水印"))

check(None not in (fi, fo, vfi, vfo, t, i, ex, h, p, v, bi, bc), "各块都必须存在")

if None not in (fi, fo, i, t):
    # 1) 两行缩短：右缘与「图像水印」组框右缘对齐；且明显窄于整窗（不再通栏到视频列）
    check(abs((fi[0] + fi[2]) - (i[0] + i[2])) <= 4,
          f"输入行右缘应对齐图像水印右缘（{fi[0] + fi[2]} vs {i[0] + i[2]}）")
    check(abs((fo[0] + fo[2]) - (i[0] + i[2])) <= 4,
          f"输出行右缘应对齐图像水印右缘（{fo[0] + fo[2]} vs {i[0] + i[2]}）")
    check(fi[0] + fi[2] < v[0] if v is not None else True,
          f"输入行右缘不应伸到视频列（行右缘 {fi[0] + fi[2]}，视频列左缘 {v}）")

if None not in (t, i):
    # 2) 文本/图像两列底边平行
    check(abs((t[1] + t[3]) - (i[1] + i[3])) <= 4,
          f"文本水印底边应与图像水印底边平行（{t[1] + t[3]} vs {i[1] + i[3]}）")

if None not in (ex, t, h, i):
    # 3) Word 秒退框：在文本水印框正下方、与其同左缘同宽
    check(abs(ex[1] - (t[1] + t[3])) <= 14,
          f"秒退框顶边应紧接文本水印底边（{ex[1]} vs {t[1] + t[3]}）")
    check(abs(ex[0] - t[0]) <= 4 and abs(ex[2] - t[2]) <= 4,
          f"秒退框应与文本水印同左缘同宽（x {ex[0]}/{t[0]}，w {ex[2]}/{t[2]}）")
    # 3b) 秒退框 顶/底 与 加固框 顶/底 平行
    check(abs(ex[1] - h[1]) <= 4,
          f"秒退框顶边应与加固框顶边平行（{ex[1]} vs {h[1]}）")
    check(abs((ex[1] + ex[3]) - (h[1] + h[3])) <= 4,
          f"秒退框底边应与加固框底边平行（{ex[1] + ex[3]} vs {h[1] + h[3]}）")
    check(ex[0] + ex[2] < h[0] + 2,
          f"秒退框在左列、不应与加固框横向重叠（秒退右缘 {ex[0] + ex[2]}，加固左缘 {h[0]}）")

if None not in (t, i, h, p):
    # 4) 加固贴在图像水印正下方、同宽
    check(abs(h[1] - (i[1] + i[3])) <= 14,
          f"加固顶边应紧接图像水印底边（{h[1]} vs {i[1] + i[3]}）")
    check(abs(h[0] - i[0]) <= 4 and abs(h[2] - i[2]) <= 4,
          f"加固应与图像水印同左缘同宽（x {h[0]}/{i[0]}，w {h[2]}/{i[2]}）")
    # 5) 预览通栏
    check(abs(p[0] - t[0]) <= 4,
          f"预览左缘应对齐文本水印左缘（{p[0]} vs {t[0]}）")
    check(abs((p[0] + p[2]) - (i[0] + i[2])) <= 6,
          f"预览右缘应对齐图像水印右缘（{p[0] + p[2]} vs {i[0] + i[2]}）")
    check(p[1] >= max(h[1] + h[3], ex[1] + ex[3]) - 4,
          f"预览应在两条水平带下方（带底 {max(h[1] + h[3], ex[1] + ex[3])}，预览顶 {p[1]}）")
    check(p[2] > t[2], f"预览应比单列更宽（{p[2]} > {t[2]}）")

if None not in (bi, bc, p):
    # 6) 按钮行
    check(bi[1] >= p[1] + p[3] - 4,
          f"按钮行应在预览下方（预览底 {p[1] + p[3]}，按钮顶 {bi[1]}）")
    check(abs(bi[1] - bc[1]) <= 2, f"两按钮应同一水平线（{bi[1]} vs {bc[1]}）")

if None not in (bi, t) and None not in (bc, i):
    ci = bi[0] + bi[2] / 2 - (t[0] + t[2] / 2)
    cc = bc[0] + bc[2] / 2 - (i[0] + i[2] / 2)
    check(abs(ci) <= 6, f"「一键插入水印」应相对文本水印列居中（偏差 {ci:.0f}px）")
    check(abs(cc) <= 6, f"「一键清除水印」应相对图像水印列居中（偏差 {cc:.0f}px）")

if None not in (vfi, vfo, v):
    # 7) 视频输入/输出行在组框上方，且与组框同宽
    check(vfi[1] + vfi[3] <= v[1] + 4,
          f"视频输入行应在视频水印框上方（行底 {vfi[1] + vfi[3]}，框顶 {v[1]}）")
    check(vfo[1] >= vfi[1] and vfo[1] + vfo[3] <= v[1] + 4,
          f"视频输出行应在输入行下方、组框上方（{vfo[1]}）")
    check(abs((vfi[0] + vfi[2]) - (v[0] + v[2])) <= 4,
          f"视频输入行右缘应对齐视频框右缘（{vfi[0] + vfi[2]} vs {v[0] + v[2]}）")
    check(abs(vfi[0] - v[0]) <= 4 and abs(vfo[0] - v[0]) <= 4,
          f"视频两行应与视频框同左缘（{vfi[0]}/{vfo[0]}/{v[0]}）")

if v is not None and t is not None:
    left_right = max(x[0] + x[2] for x in (t, i, h, p, ex) if x)
    check(v[0] >= left_right - 4,
          f"视频水印应为独立右列（左侧区右缘 {left_right}，视频列左缘 {v[0]}）")
    d_top = abs(v[1] - t[1])
    check(d_top <= 60,
          f"视频框顶应与 Word 两列大致同一水平线（视频框顶 {v[1]}，文本列顶 {t[1]}，差 {d_top}）")

print("\nRESULT:", "ALL OK" if ok else "HAS FAILURES")
raise SystemExit(0 if ok else 1)
