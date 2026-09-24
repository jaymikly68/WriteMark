"""打包共享配置：排除用不到的模块，缩小 onefile 运行时解压体积。

背景（重要）：单文件(onefile) exe 每次运行都会把整个运行时解压到 %TEMP%\\_MEIxxxxx，
关闭程序时 bootloader 还要把整个目录删掉——这段时间进程并不退出，
用户感知为“关了窗口还要卡十几秒到半分钟”。
实测（本机）：264 个文件 / 157.5 MB → 关闭到进程结束 11.56 秒。
所以能删的依赖越少越好：砍掉没用到的 Qt 模块与 numpy 等大件。

如需重新放某个模块，把它从 EXCLUDES 里删掉即可。
"""

# ---- Qt：本工具只用 QtCore / QtGui / QtWidgets / QtWinExtras(托盘在 QtWidgets 内) ----
_QT_UNUSED = [
    "QtBluetooth", "QtCharts", "QtConcurrent", "QtDataVisualization", "QtDBus",
    "QtDesigner", "QtDesignerComponents", "QtGraphs", "QtGraphsWidgets", "QtHelp",
    "QtHttpServer", "QtLocation", "QtMultimedia", "QtMultimediaWidgets", "QtNfc",
    "QtOpenGL", "QtOpenGLWidgets", "QtPdf", "QtPdfWidgets", "QtPositioning",
    "QtQml", "QtQmlCompiler", "QtQmlCore", "QtQmlModels", "QtQmlNetwork",
    "QtQmlWorkerScript", "QtQmlXmlListModel", "QtQuick", "QtQuick3D",
    "QtQuick3DAssetImport", "QtQuick3DAssetUtils", "QtQuick3DCore",
    "QtQuick3DEffects", "QtQuick3DHelpers", "QtQuick3DIblBaker",
    "QtQuick3DParticles", "QtQuick3DRuntimeRender", "QtQuick3DUtils",
    "QtQuickControls2", "QtQuickControls2Basic", "QtQuickControls2FluentWinUI3",
    "QtQuickControls2Fusion", "QtQuickControls2Imagine", "QtQuickControls2Impl",
    "QtQuickControls2Material", "QtQuickControls2Universal", "QtQuickDialogs2",
    "QtQuickDialogs2QuickImpl", "QtQuickDialogs2Utils", "QtQuickLayouts",
    "QtQuickTemplates2", "QtQuickTest", "QtQuickWidgets", "QtRemoteObjects",
    "QtScxml", "QtSensors", "QtSerialPort", "QtShaderTools", "QtSpatialAudio",
    "QtSql", "QtSvg", "QtSvgWidgets", "QtTest", "QtTextToSpeech", "QtUiTools",
    "QtVirtualKeyboard", "QtWebChannel", "QtWebEngineCore", "QtWebEngineQuick",
    "QtWebEngineWidgets", "QtWebSockets", "Qt3DAnimation", "Qt3DCore",
    "Qt3DExtras", "Qt3DInput", "Qt3DLogic", "Qt3DQuick", "Qt3DRender", "Qml",
]

# ---- 其它用不到的大件 / 杂项 ----
_OTHER_UNUSED = [
    "numpy",                 # 本工具完全不使用 numpy，随包进来的 47MB 纯属浪费
    "numpy.libs",
    "scipy", "pandas", "matplotlib", "imageio", "imageio_ffmpeg", "PIL.ImageQt",
    "PIL.ImageTk", "PIL.ImageShow", "PIL.ImageGrab", "PIL._tkinter_finder",
    "tkinter", "turtle", "sqlite3", "unittest", "doctest", "pydoc", "pdb",
    "pytest", "IPython", "notebook", "jupyter", "setuptools", "pip", "wheel",
    "pythonwin",             # 只用 win32com.client，不需要 pythonwin IDE 相关包
    "win32comext",           # 除 client 外的 COM 扩展不用
    "pytest", "black", "mypy", "isort", "flake8",
]

EXCLUDES = [f"PySide6.{m}" for m in _QT_UNUSED] + _OTHER_UNUSED
