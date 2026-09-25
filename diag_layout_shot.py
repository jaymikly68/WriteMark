"""渲染当前主窗口为 PNG，便于核对各分组块的实际位置。

用法：
    python diag_layout_shot.py            # 默认用真实平台（文字正常，窗口会一闪）
    WM_SHOT_OFFSCREEN=1 python diag_layout_shot.py   # 免打扰的 offscreen（中文可能渲染成方框）
"""
import os

if os.environ.get("WM_SHOT_OFFSCREEN"):
    os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

from PySide6.QtWidgets import QApplication

from watermark_tool.gui import App

app = QApplication.instance() or QApplication([])
w = App()
# 给足高度，整页内容尽量一屏拍完（内容本身可滚动，不影响布局）
w.resize(1600, 1780)
w.show()
for _ in range(6):
    app.processEvents()
# 布局稳定后重画一次预览，确保预览尺寸与最终标签尺寸一致
try:
    w._render_preview()
    app.processEvents()
    print("preview label:", w.preview_label.width(), w.preview_label.height(),
          "img_rect:", w.preview_label._img_rect)
except Exception as e:      # 预览失败不影响布局核对
    print("render preview skipped:", e)

out = os.path.join(os.path.dirname(os.path.abspath(__file__)), "shot_current.png")
w.grab().save(out, "PNG")
print("saved:", out, "size:", w.size().width(), w.size().height())
