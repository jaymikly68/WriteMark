"""
后台守护单元测试（构造签名与当前 WatermarkWatchdog 一致）。

注意构造参数顺序：(src_path, output_path, kinds, opts, interval, log)。
端到端的三场景验证（被清除 / 手动删图形 / 文件被删除）见 test_watchdog_verify.py。
"""
import os
import sys
import time

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from docx import Document

from watermark_tool import core, engine_docx, watchdog

TMP = os.path.join(os.path.dirname(os.path.abspath(__file__)), "_wd_verify_tmp")
OUT = os.path.join(TMP, "unit_输出.docx")
SRC = os.path.join(TMP, "unit_源.docx")

OPTS = {
    "text": {"text": "机密", "cn_font_name": "微软雅黑", "latin_font_name": "Times New Roman",
             "font_size": 120, "color": (200, 0, 0), "angle": 45, "transparency": 0.6,
             "scale": 1.0, "offset_x": 0, "offset_y": 0},
    "image": {},
}


def safe_has_watermark(path):
    """守护重建文件途中去读会抛 PackageNotFoundError（文件还没写完），按“无水印”处理。"""
    try:
        return core.has_watermark(path)
    except Exception:
        return False


def wait_for(predicate, timeout=3.0):
    end = time.time() + timeout
    while time.time() < end:
        if predicate():
            return True
        time.sleep(0.1)
    return False


if __name__ == "__main__":
    os.makedirs(TMP, exist_ok=True)
    d = Document()
    d.add_paragraph("正文")
    d.save(SRC)
    if os.path.exists(OUT):
        os.remove(OUT)

    core.insert_watermark(SRC, ["text"], output_path=OUT, **OPTS)
    assert core.has_watermark(OUT), "插入后应有水印"

    wd = watchdog.WatermarkWatchdog(SRC, OUT, ["text"], OPTS, interval=0.5,
                                    log=lambda m, verbose=False: None)
    wd.start()
    try:
        assert wait_for(lambda: safe_has_watermark(OUT)), "守护启动后水印应在位"
        # 模拟被清除 → 应自动补回
        engine_docx.clear_watermark(OUT)
        assert not safe_has_watermark(OUT), "清除后应无水印"
        assert wait_for(lambda: safe_has_watermark(OUT)), "水印应被守护自动补回"

        # 模拟输出文件被删除 → 应自动重建
        os.remove(OUT)
        assert wait_for(lambda: os.path.exists(OUT) and safe_has_watermark(OUT)), \
            "输出文件被删除后应被守护自动重建"
    finally:
        wd.stop()

    print("后台守护单元测试通过：水印被清除/文件被删除后均能自动补回 ✅")
