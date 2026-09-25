"""只允许一个程序实例常驻的回归测试。

背景：关窗后程序会缩到托盘继续跑，旧版本可以直接再开一个，于是可能同时存在
两个后台守护互相抢着补水印。这里用 QLocalServer 做单机单例，第二个实例只负责
把已有实例的窗口唤出来，然后自己退出。
"""
import sys

from PySide6.QtWidgets import QApplication

from watermark_tool import gui


def test_first_instance_wins_the_lock():
    """首个实例应抢到单例服务，第二个实例必须直接退出。"""
    from PySide6.QtNetwork import QLocalServer

    app = QApplication.instance() or QApplication([])      # noqa: F841

    server, should_exit = gui._acquire_single_instance()
    try:
        assert server is not None, "首个实例应抢到单例服务"
        assert should_exit is False, "首个实例不应退出"

        # 模拟第二个实例启动
        _server2, exit2 = gui._acquire_single_instance()
        assert exit2 is True, "已有一个实例在跑时，第二个实例必须直接退出"
        assert _server2 is None, "被拦截的实例不应再抢服务"
    finally:
        QLocalServer.removeServer(gui.SINGLE_INSTANCE_NAME)


def test_missing_qtnetwork_never_blocks_startup(monkeypatch):
    """QtNetwork 不可用时宁可多开一个，也不能让程序打不开。"""
    monkeypatch.setitem(sys.modules, "PySide6.QtNetwork", None)  # 让 import 失败
    server, should_exit = gui._acquire_single_instance()
    assert server is None and should_exit is False, "拿不到单例能力时应放行启动"
