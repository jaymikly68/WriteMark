"""共享 pytest 设施：提供一个跨会话的 QApplication 单例。

多个 GUI 测试需要 QApplication，但 Qt 的 QApplication 是进程内单例，
重复 `QApplication([])` 会抛 “already exists”。这里统一用 session 级
fixture 复用同一个实例，避免测试间互相踩踏。
"""
import os
import sys

# 测试已迁到 tests/{unit,integration,gui,e2e}/ 下，但源码包 watermark_tool 仍在项目根，
# 这里显式把根目录放进 sys.path，免得 pytest 的导入模式变动导致用例集体 ImportError。
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

import pytest
from PySide6.QtWidgets import QApplication


@pytest.fixture(scope="session")
def app():
    a = QApplication.instance() or QApplication([])
    yield a
