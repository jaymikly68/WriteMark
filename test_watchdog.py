import os, sys, time
sys.path.insert(0, os.path.dirname(__file__))
from docx import Document
from watermark_tool import core, watchdog

OUT = os.path.join(os.path.dirname(__file__), "wd_sample.docx")
opts = {"text": "机密", "color": (200, 0, 0), "angle": 45, "transparency": 0.6, "scale": 1.0}

def make():
    d = Document(); d.add_paragraph("正文"); d.save(OUT)

if __name__ == "__main__":
    make()
    wd = watchdog.WatermarkWatchdog(OUT, "text", opts, interval=0.5, log=lambda m, verbose=False: print(m))
    wd.start()
    time.sleep(1.0)
    print("守护启动后存在:", core.has_watermark(OUT))
    # 模拟被人清除（反制目标）
    core.clear_watermark(OUT)
    print("被清除后存在:", core.has_watermark(OUT))
    time.sleep(1.3)
    print("守护补回后存在:", core.has_watermark(OUT))
    wd.stop()
    print("测试结束")
