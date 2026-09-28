"""Word 一键水印工具 —— 程序入口。"""
import sys
import os

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from watermark_tool.gui import main
from watermark_tool import video as _video


if __name__ == "__main__":
    # 自愈：回收上次崩溃 / 被强关残留的临时 mp4（best-effort，不影响启动）
    try:
        _video.cleanup_video_temp_files()
    except Exception:
        pass
    main()
