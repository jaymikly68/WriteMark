# -*- mode: python ; coding: utf-8 -*-
"""单文件版构建。

重要（退出卡顿的根因）：onefile exe 运行时会把整个运行时解包到 %TEMP%\\_MEIxxxxxx，
程序关闭后 bootloader 必须把这整个目录删完，进程才算退出——
这段时间表现为“关了窗口还卡十几秒”。因此这里把依赖砍到最小，
详见 exclude_conf.py 的说明。
"""
import os
import sys
# 项目根（本 spec 与 tools/ 同级），tools/exclude_conf.py 已在下面 import
sys.path.insert(0, os.path.dirname(os.path.abspath(SPEC)))

from tools.exclude_conf import EXCLUDES


a = Analysis(
    ['main.py'],
    pathex=[],
    binaries=[],
    datas=[],
    hiddenimports=[],
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
