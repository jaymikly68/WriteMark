@echo off
chcp 65001 >nul
powershell -NoProfile -ExecutionPolicy Bypass -Command "Start-Process -FilePath powershell.exe -Verb RunAs -ArgumentList '-NoProfile -ExecutionPolicy Bypass -File \"%~dp0fix.ps1\"'"
echo.
echo 如果上面弹出了 UAC 窗口，请点"是"。
echo 看到绿色 Done 之后，彻底关闭 Word 再重新打开即可。
pause
