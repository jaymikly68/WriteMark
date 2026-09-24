# -*- mode: python ; coding: utf-8 -*-
"""最终的单个 WordWatermark.exe：极小启动器 + 内嵌 zip 运行体。

不用 onefile 直接打包整个应用的原因见 launcher.py 顶部说明：onefile 关闭时要
删除 130MB 的临时目录，进程退出会卡 8~30 秒，这是各版本“退出卡顿”的真正根因。
"""
# 注意：刻意 *不* 把 zip 作为 datas 编进包里——
# 那样每次运行都要随 onefile 展开到 %TEMP% 再删除，关闭时会拖慢退出。
# 改为构建完成后把 zip 追加到 exe 文件尾部（见 build_launcher.py），
# 运行时直接用 zipfile 读自己，几乎没有临时目录开销。
a = Analysis(
    ['launcher.py'],
    pathex=[],
    binaries=[],
    datas=[],
    hiddenimports=['tkinter', 'tkinter.ttk'],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[
        # 启动器只用到：os / sys / subprocess / zipfile / shutil / tkinter
        'numpy', 'scipy', 'matplotlib', 'pandas', 'PIL', 'lxml', 'docx',
        'win32com', 'win32gui', 'win32con', 'win32api', 'pythoncom', 'pywintypes',
        'PySide6', 'shiboken6', 'PySide6.QtCore', 'PySide6.QtGui', 'PySide6.QtWidgets',
        'unittest', 'doctest', 'pydoc', 'pdb', 'pytest', 'setuptools', 'pip', 'wheel',
        'sqlite3', 'IPython',
    ],
    noarchive=False,
    optimize=0,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name='WordWatermark',
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
