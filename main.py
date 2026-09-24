"""Word 一键水印工具 —— 程序入口。"""
import sys
import os

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from watermark_tool.gui import main

if __name__ == "__main__":
    main()
