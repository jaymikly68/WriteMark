"""
后台守护（看门狗）：定时检查目标文件的水印是否还在。
被他人/手动删除后，自动重新插入——这就是“反制一键清除”的能力。

说明：原始想法是“每 1ms 检查一次”，但 1ms 级轮询会疯狂读写文件、极易造成
Word 文件锁死或损坏，因此默认改为 1 秒（可在 GUI 调整，最小 0.2 秒）。
"""
from __future__ import annotations

import os
import threading
import time

from . import core


def _default_log(msg, verbose=False):
    print(msg)


class WatermarkWatchdog:
    def __init__(self, src_path, output_path, kinds, opts, interval=1.0, log=_default_log):
        self.src_path = os.path.abspath(src_path)
        self.output_path = os.path.abspath(output_path)
        self.kinds = list(kinds)
        self.opts = opts
        self.interval = max(0.2, float(interval))
        self.log = log
        self._stop = threading.Event()
        self._thread = None
        self.running = False

    def _loop(self):
        self.running = True
        self.log(f"守护已启动：每 {self.interval:.1f}s 检查一次水印（监控 {os.path.basename(self.output_path)}）。")
        was_locked = None
        while not self._stop.is_set():
            try:
                # 文件被 WPS/Word 占用时完全跳过本轮（不打开文件），避免与编辑器争抢导致卡顿
                locked = core.is_file_locked(self.output_path)
                if locked:
                    if was_locked is False:
                        self.log("· 输出文件被 WPS/Word 占用，守护暂停读写，待文件空闲后自动恢复。")
                    was_locked = True
                else:
                    if was_locked:
                        self.log("· 文件已空闲，守护恢复检查。")
                    was_locked = False
                    if not os.path.exists(self.output_path):
                        # 输出文件被整个删除（不只是水印被删）：直接从原文件重新生成并加水印。
                        # 此前这里会走 has_watermark 抛“Package not found”，导致只报错、永不补回。
                        core.insert_watermark(self.src_path, self.kinds,
                                              output_path=self.output_path, **self.opts)
                        self.log("⚠ 输出文件已被删除，已自动从原文件重新生成并加水印。")
                    elif not core.has_watermark(self.output_path):
                        core.insert_watermark(self.src_path, self.kinds,
                                              output_path=self.output_path, **self.opts)
                        self.log("⚠ 检测到水印被移除，已自动从原文件补回。")
                    else:
                        self.log("· 水印在位。", verbose=True)
            except Exception as e:
                self.log(f"守护检查出错: {e}")
            self._stop.wait(self.interval)
        self.running = False
        self.log("守护已停止。")

    def start(self):
        if self.running:
            return
        self._stop.clear()
        self._thread = threading.Thread(target=self._loop, daemon=True)
        self._thread.start()

    def stop(self):
        self._stop.set()
        if self._thread:
            self._thread.join(timeout=3)

    def stop_nowait(self):
        """仅置停止信号、不阻塞等待。守护线程为 daemon，会自动退出，
        避免在关闭窗口时阻塞主线程（这正是此前“退出卡顿”的根因）。"""
        self._stop.set()
