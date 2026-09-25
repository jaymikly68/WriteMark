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
import imageio_ffmpeg as _iff
from PyInstaller.utils.hooks import copy_metadata, collect_all

# imageio_ffmpeg 自带的 ffmpeg 静态二进制（约 70MB）必须显式打进包，
# 否则打包后视频水印功能找不到解码器。放到 imageio_ffmpeg/binaries 下，
# 使其相对路径与运行时 get_ffmpeg_exe() 的预期一致。
_ff_bin = _iff.get_ffmpeg_exe()

# imageio 在 import 时就会执行 importlib.metadata.version("imageio")，
# 不把 dist-info 元数据打进包的话，冻结环境里 import imageio 直接抛
# PackageNotFoundError（表现为视频页提示“依赖未安装”）。imageio_ffmpeg 同理带上。
# copy_metadata 返回 [(src, dest), ...] 形式的列表。
_metadata = copy_metadata('imageio') + copy_metadata('imageio_ffmpeg')

# PyMuPDF（PDF 作为水印图源）是纯 Python 包 + 编译扩展的混合体，PyInstaller 的
# 静态分析只会收二进制、漏掉它的 .py；用 collect_all 强制把整个包打进运行体。
_pymupdf_datas, _pymupdf_bins, _pymupdf_hidden = collect_all('pymupdf')
_datas_extra = _pymupdf_datas
_bins_extra = _pymupdf_bins
_hidden_extra = _pymupdf_hidden


a = Analysis(
    ['main.py'],
    pathex=[],
    binaries=_bins_extra,
    datas=[(_ff_bin, 'imageio_ffmpeg/binaries')] + _metadata + _datas_extra,
    # office_tweak 在 gui.py 里是「try 导入（模块缺失不该拖垮 UI）」，
    # PyInstaller 的静态分析扫不到，必须显式声明，否则「Word 秒退」按钮点了没反应。
    # video 同理（视频水印模块，依赖 imageio / imageio_ffmpeg / numpy）。
    hiddenimports=['watermark_tool.office_tweak', 'watermark_tool.video',
                   'watermark_tool.i18n'] + _hidden_extra,
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
