"""诊断：Word / 视频水印的「自定义位置」是否生效（offscreen 渲染，不弹窗）。"""
import os
os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

from PySide6.QtWidgets import QApplication

from watermark_tool.gui import App

app = QApplication.instance() or QApplication([])
w = App()
w.show()
app.processEvents()

ok = True


def check(cond, msg):
    global ok
    print(("PASS  " if cond else "FAIL  ") + msg)
    if not cond:
        ok = False


# ---------------- Word：位置预设 → 偏移 ----------------
w.text_pos_combo.setCurrentText("右下")
check(abs(w.text_offset_x - 35.0) < 0.01 and abs(w.text_offset_y - 40.0) < 0.01,
      f"Word 文本位置预设「右下」→ 偏移 ({w.text_offset_x}, {w.text_offset_y}) 应为 (35, 40)")
check(w.text_pos_combo.currentText() == "右下",
      f"应用预设后下拉不应被误切到自定义（当前 {w.text_pos_combo.currentText()}）")

w.img_pos_combo.setCurrentText("左上")
check(abs(w.img_offset_x + 35.0) < 0.01 and abs(w.img_offset_y + 40.0) < 0.01,
      f"Word 图像位置预设「左上」→ 偏移 ({w.img_offset_x}, {w.img_offset_y}) 应为 (-35, -40)")

# 手改偏移 → 自动切「自定义」
w.text_offx_spin.setValue(12.0)
check(w.text_pos_combo.currentText() == "自定义",
      f"手改偏移后下拉应切到「自定义」（当前 {w.text_pos_combo.currentText()}）")

# 预览拖拽 → 同步到偏移 + 自定义
w.preview_drag_chk.setChecked(True)
w._on_preview_drag(0.80, 0.30)
check(abs(w.text_offset_x - 30.0) < 1.0 and abs(w.text_offset_y + 20.0) < 1.0,
      f"预览拖拽 (0.80,0.30) → 文本偏移应约 (30,-20)，实际 ({w.text_offset_x:.1f}, {w.text_offset_y:.1f})")
check(w.text_pos_combo.currentText() == "自定义", "拖拽后下拉应为「自定义」")

# ---------------- 视频：预设 ↔ X/Y ↔ 拖拽 ----------------
w.v_text_pos_combo.setCurrentText("右下")
check(abs(w.v_text_x_spin.value() - 82) < 1 and abs(w.v_text_y_spin.value() - 85) < 1,
      f"视频预设「右下」应回填 X/Y≈(82,85)，实际 ({w.v_text_x_spin.value():.0f}, {w.v_text_y_spin.value():.0f})")

w.v_text_pos_combo.setCurrentText("自定义")
w.v_text_x_spin.setValue(10)
w.v_text_y_spin.setValue(20)
opts = w._v_gather_opts()
tp = opts["text"]["position"]
check(abs(tp[0] - 0.10) < 1e-6 and abs(tp[1] - 0.20) < 1e-6,
      f"视频自定义 X/Y=10/20 → position 应为 (0.1, 0.2)，实际 {tp}")

w.v_img_pos_combo.setCurrentText("自定义")
w.v_img_x_spin.setValue(55)
w.v_img_y_spin.setValue(66)
opts = w._v_gather_opts()
ip = opts["image"]["position"]
check(abs(ip[0] - 0.55) < 1e-6 and abs(ip[1] - 0.66) < 1e-6,
      f"视频图片自定义 X/Y=55/66 → position 应为 (0.55, 0.66)，实际 {ip}")

# 手改 X/Y → 下拉自动切自定义
w.v_text_pos_combo.setCurrentText("居中")
w.v_text_x_spin.setValue(7)
check(w.v_text_pos_combo.currentText() == "自定义",
      f"手改 X/Y 后下拉应切「自定义」（当前 {w.v_text_pos_combo.currentText()}）")

# 预览帧拖拽 → 写回 X/Y 并切「自定义」
w.v_drag_chk.setChecked(True)
w._on_video_frame_drag(0.4, 0.6)
check(abs(w.v_text_x_spin.value() - 40) < 1 and abs(w.v_text_y_spin.value() - 60) < 1,
      f"帧拖拽 (0.4,0.6) → X/Y 应为 (40,60)，实际 ({w.v_text_x_spin.value():.0f}, {w.v_text_y_spin.value():.0f})")
check(w.v_text_pos_combo.currentText() == "自定义" and w.v_img_pos_combo.currentText() == "自定义",
      "帧拖拽后文字/图片下拉都应切「自定义」")
opts = w._v_gather_opts()
check(abs(opts["text"]["position"][0] - 0.4) < 1e-6,
      f"拖拽落点应进入导出参数，实际 {opts['text']['position']}")

# ---------------- 视频预览：无源视频也能出示意帧 ----------------
w.v_src_edit.setText("")
w._v_grab_frame()
check(w.v_frame_label.hasImage(), "未选视频时也应能预览（16:9 示意画面 + 水印）")
check(getattr(w, "_v_last_frame", None) is not None, "示意帧应已生成")

print("\nRESULT:", "ALL OK" if ok else "HAS FAILURES")
raise SystemExit(0 if ok else 1)
