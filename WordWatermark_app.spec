# -*- mode: python ; coding: utf-8 -*-
"""目录版（onedir）构建：本工具的“真正运行体”。

退出卡顿的根因是 onefile：PyInstaller 单文件 exe 每次运行都要把 ~130MB 运行时
解包到 %TEMP%\\_MEIxxxx，程序关闭时 bootloader 还要把整个目录删完进程才算退出
（本机实测 8.7~11.6 秒，杀软实时扫描下更久）。
onedir 形态没有这一步，关闭即退出（实测 0.2 秒）。

本 spec 产出的 dist_portable/WriteMarkApp/ 会被 build_launcher.py 压缩嵌入
WordWatermark.exe 启动器，用户拿到的仍然只有一个 exe。
"""
from exclude_conf import EXCLUDES


a = Analysis(
    ['main.py'],
    pathex=[],
    binaries=[],
    datas=[],
    # office_tweak 在 gui.py 里是「try 导入（模块缺失不该拖垮 UI）」，
    # PyInstaller 的静态分析扫不到，必须显式声明，否则「Word 秒退」按钮点了没反应。
    hiddenimports=['watermark_tool.office_tweak'],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=EXCLUDES,
    noarchive=False,
    optimize=0,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name='WriteMarkApp',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)
coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='WriteMarkApp',
)
