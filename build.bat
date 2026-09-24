@echo off
REM 打包为单文件 .exe（需先 pip install pyinstaller PySide6 python-docx Pillow pywin32）
pyinstaller --onefile --windowed --name WordWatermark main.py
echo 产物位于 dist/WordWatermark.exe
pause
