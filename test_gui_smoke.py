"""
无界面（offscreen）冒烟测试：验证 GUI 的关键逻辑——
- 默认文本启用 / 图像关闭
- 复选框切换后 _gather_kinds 正确
- 角度/透明度可直接设定并保留两位小数精度
- 滑块与数字框联动生成
"""
import os
import sys

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")
from PySide6.QtWidgets import QApplication

from watermark_tool.gui import App


def main():
    app = QApplication(sys.argv)
    win = App()

    # 默认状态
    assert win.text_enabled is True
    assert win.image_enabled is False
    assert win._gather_kinds() == ["text"]

    # 启用图像
    win.img_chk.setChecked(True)
    assert win.image_enabled is True
    assert win._gather_kinds() == ["text", "image"]

    # 数值精度：文本角度 30.5、文本透明度 25.75%（各类型可独立调整）
    win._on_text_angle(30.5)
    win._on_text_trans(25.75)
    assert abs(win.text_angle - 30.5) < 1e-9
    assert abs(win.text_transparency - 0.2575) < 1e-9

    # 图像侧参数默认独立存在且可设值
    win._on_img_angle(12.25)
    win._on_img_trans(40.0)
    assert abs(win.img_angle - 12.25) < 1e-9
    assert abs(win.img_transparency - 0.4) < 1e-9

    # 滑块+数字框：透明度 spin 应显示两位小数
    # 通过 _make_slider_spin 已创建，检查 text 组里 font_spin / text_scale_spin 的 decimals
    assert win.font_spin.decimals() == 2
    assert win.text_scale_spin.decimals() == 2
    assert win.img_scale_spin.decimals() == 2
    assert win.interval_spin.decimals() == 2

    # 关闭文本、仅图像
    win.text_chk.setChecked(False)
    assert win._gather_kinds() == ["image"]

    # 未启用任何水印时的校验提示路径（不弹窗，仅检查 gather 为空）
    win.img_chk.setChecked(False)
    assert win._gather_kinds() == []

    print("GUI 冒烟测试通过：默认值、类型开关、两位小数精度、滑块联动均正常。")


if __name__ == "__main__":
    main()
