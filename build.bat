@echo off
REM ============================================================
REM  构建最终交付的单个 WordWatermark.exe
REM
REM  为什么要两步（而不是一条 pyinstaller --onefile）？
REM  onefile 每次运行都会把 ~130MB 运行时解包到 %TEMP%\_MEIxxxx，
REM  关闭窗口后还要把整个目录删完进程才退出 —— 实测 8.7s（用户机上 20~30s）。
REM  现在改成：
REM    1) onedir 运行体 dist/WriteMarkApp/
REM    2) 压缩成 payload/app_payload.zip 并追加到极小启动器 exe 尾部
REM  运行体只在首次使用时解压到本地目录，关闭时没有任何清理动作 → 0.02s 退出。
REM  详见 launcher.py 顶部说明。
REM
REM  注意：build_launcher.py 会先把上一次的 dist/WriteMarkApp 改名挪走再构建。
REM  这一步不能省 —— PyInstaller 的"删除旧目录"一旦被安全策略拦下，它不会报错，
REM  产物却会静默停留在旧版本（本项目真实踩过这个坑）。
REM ============================================================
python build_launcher.py
if errorlevel 1 goto :fail
echo.
echo 产物：dist\WordWatermark.exe
echo 验证：python test_launcher.py --clean   （含首次/二次启动与关闭耗时）
pause
exit /b 0

:fail
echo 构建失败
pause
exit /b 1
