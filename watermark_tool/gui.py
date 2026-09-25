"""
PySide6 (Qt) 一键 GUI：选择 Word 文件 → 设置文本/图像水印 → 一键插入/清除；
可选开启“后台守护”，水印被删自动补回（反制清除）。

选 Qt 而非 tkinter 的原因：本打包环境的标准 Python 不含 tcl/tk（tkinter），
而 PySide6 可 pip 安装、可被 PyInstaller 完整打包，最终 .exe 开箱即用。

本次改进：
- 文本水印与图像水印可【同时添加】（各自“启用”复选框，二者互不覆盖）。
- 旋转角度、透明度、偏移等除滑块拖动外，也可直接输入数字；所有数值均精确到小数点后两位。
- 文本与图像两组参数【横向并排】，整体以滚动区域承载，避免在较小屏幕上被裁切。
- 一键清除可去掉【任意】水印：本工具添加的，以及 Word 原本就带的水印（衬于文字下方的图形）。
"""
from __future__ import annotations

import os
import io
import tempfile
import sys
import threading

from PySide6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QLineEdit, QPushButton, QLabel, QFileDialog, QComboBox, QCompleter,
    QDoubleSpinBox, QCheckBox, QSlider, QTextEdit, QColorDialog, QMessageBox,
    QScrollArea, QSystemTrayIcon, QMenu, QStyle, QProgressBar,
    QButtonGroup, QRadioButton, QSizePolicy,
)
from PySide6.QtCore import Qt, QObject, QThread, Signal, QTimer, QStringListModel, QEvent, QRect, QPointF
from PySide6.QtGui import QColor, QImage, QPixmap, QIntValidator, QPainter, QPen

from . import core, watchdog, preview, word_fonts, engine_docx
try:  # office_tweak 只依赖标准库，任何情况下缺了也不该拖垮整GUI
    from . import office_tweak as otw
except Exception:  # pragma: no cover
    otw = None
try:  # 视频水印：imageio + imageio-ffmpeg（自带 ffmpeg 二进制）；缺了则视频页不可用但不拖垮整GUI
    from . import video as video_mod
    import imageio_ffmpeg  # noqa: F401  确保 ffmpeg 二进制可被打包收集
except Exception as _video_err:  # pragma: no cover
    video_mod = None
    # 把真实原因记下来：冻结环境里视频功能失效时，日志一看便知缺了什么
    try:
        import tempfile as _tf, traceback as _tb
        with open(os.path.join(_tf.gettempdir(), "WriteMark_video_import_err.log"),
                  "w", encoding="utf-8") as _f:
            _f.write("".join(_tb.format_exception(_video_err)))
    except Exception:
        pass

# 主操作按钮统一高亮样式：蓝底黑字（插入/去除/视频开始加水印等）
BTN_HL = (
    "QPushButton {"
    "  background-color:#1a73e8; color:#000; font-weight:700;"
    "  border:2px solid #0b57d0; border-radius:6px;"
    "  padding:8px 16px; min-height:34px;"
    "}"
    "QPushButton:hover { background-color:#2b82f1; }"
    "QPushButton:pressed { background-color:#0b57d0; }"
)

# 危险操作按钮：深褐底 + 红字（视频水印的「取消」），与主/次操作配色明显区分
BTN_BROWN = (
    "QPushButton {"
    "  background-color:#4a2b16; color:#ff3b30; font-weight:700;"
    "  border:2px solid #2f1a0c; border-radius:6px;"
    "  padding:8px 16px; min-height:34px;"
    "}"
    "QPushButton:hover { background-color:#5c3719; }"
    "QPushButton:pressed { background-color:#3a2110; }"
    "QPushButton:disabled {"
    "  background-color:#efe4d6; color:#b08a6a; border-color:#d8c8b6;"
    "  font-weight:400;"
    "}"
)


class _ProportionalButton(QPushButton):
    """宽度按宿主容器宽度的固定比例自适应。

    嵌在分组框内的主操作按钮需要「比所在分组框窄 20%」，但分组框宽度
    在构造时还是 0（窗口尚未布局），用固定宽度会与实际宽度脱节。
    这里监听宿主 Resize 事件持续同步，保证任何窗口宽度下比例都成立。
    """

    def __init__(self, text, host, ratio=0.8, minimum=88, host2=None, parent=None):
        super().__init__(text, parent)
        self._ratio = ratio
        self._minimum = minimum
        # 监听所有相关宿主：首次布局时各宿主的宽度是先后才定下来的，只监听一个
        # 会让后建的那个沿用旧宽度（实测出现差 1px、甚至整体偏大的情况）。
        self._hosts = [host] if host2 is None else [host, host2]
        for h in self._hosts:
            h.installEventFilter(self)
        self._sync()

    def eventFilter(self, obj, event):
        if event.type() == QEvent.Resize and any(obj is h for h in self._hosts):
            self._sync()
        return super().eventFilter(obj, event)

    def _sync(self):
        w = min((h.width() for h in self._hosts), default=0)
        if w <= 0:
            return
        self.setFixedWidth(max(self._minimum, int(w * self._ratio)))


class _EqualBottoms(QObject):
    """让两个控件的【底边】保持同一水平线。

    左右两列自然行数不同（文本水印比图像水印多行），是本版式要压平的最后
    一处参差：这里把较矮的一列的最小高度抬到两者 sizeHint 的较大值。
    不能直接改 sizePolicy 去“拉伸”——历史上 Minimum/Preferred 会吸走多余空间、
    把组标题顶得悬空（v1.3.3 踩过），这里只补最小高度，最稳。
    """

    def __init__(self, a, b, parent=None):
        super().__init__(parent if parent is not None else a)
        self._a, self._b = a, b
        for w in (a, b):
            w.installEventFilter(self)
        QTimer.singleShot(0, self.sync)

    def eventFilter(self, obj, event):
        if obj in (self._a, self._b) and event.type() in (
                QEvent.Resize, QEvent.Show, QEvent.LayoutRequest):
            QTimer.singleShot(0, self.sync)
        return False

    def sync(self):
        if self._a is None or self._b is None:
            return
        try:
            ha = self._a.minimumSizeHint().height()
            hb = self._b.minimumSizeHint().height()
        except RuntimeError:      # 控件已被销毁
            return
        target = max(int(ha), int(hb), 0)
        # 只在真的变了才写：否则 setMinimumHeight 触发 Resize → 再次 sync，死循环
        for w in (self._a, self._b):
            if target > 0 and w.minimumHeight() != target:
                w.setMinimumHeight(target)


# 视频水印锚点：位置名 -> (左上角相对帧的比例, 0..1)
# 文字与图片各有独立一份，二者可错开，避免同时固定时完全重叠。
# "自定义" 由右侧 X/Y 百分比输入决定（拖拽预览帧亦会写入 X/Y）。
_V_POSITIONS = {
    "右下": (0.82, 0.85), "右上": (0.82, 0.08), "左下": (0.08, 0.85),
    "左上": (0.08, 0.08), "居中": (0.35, 0.40), "自定义": None,
}
_V_POSITIONS_KEYS = list(_V_POSITIONS.keys())

# Word 水印位置预设：名称 -> (水平偏移%, 垂直偏移%)
# 偏移以页面中心为 0，取值 -50~50（对应 X/Y 百分比 0~100，50=居中）。
# "自定义" 表示直接由水平/垂直偏移滑块（或预览拖拽）决定，不做覆盖。
_W_POS_PRESETS = {
    "居中": (0.0, 0.0),
    "左上": (-35.0, -40.0), "上中": (0.0, -40.0), "右上": (35.0, -40.0),
    "中左": (-35.0, 0.0),                         "中右": (35.0, 0.0),
    "左下": (-35.0, 40.0),  "下中": (0.0, 40.0),  "右下": (35.0, 40.0),
    "自定义": None,
}
_W_POS_KEYS = list(_W_POS_PRESETS.keys())

# 播放方式 -> 引擎 motion 值（"跟随" 表示沿用全局单选）
_V_MOTION_MAP = {
    "固定": "fixed",
    "滚动": "scroll",
    "固定+滚动": "both",
}
_V_MOTION_LABELS = ["跟随", "固定", "滚动", "固定+滚动"]
_V_MOTION_INVERSE = {v: k for k, v in _V_MOTION_MAP.items()}


def _resolve_motion(combo, global_motion: str) -> str:
    """类型级播放方式：选了就用选的，选“跟随”就用全局单选的值。"""
    label = combo.currentText()
    if label == "跟随":
        return global_motion
    return _V_MOTION_MAP.get(label, global_motion)


# ---------------------------------------------------------------------------
# 导出能力上限：一律不超过用户当前显示器的分辨率与刷新率
# ---------------------------------------------------------------------------
def _screen_geometry() -> tuple[int, int]:
    """当前主屏幕的分辨率 (w, h)；取不到时用 1920x1080 兜底。

    必须用 screenGeometry()（不含任务栏扣除），不能用 availableGeometry()：
    后者在 Windows 上会把任务栏高度扣掉（例如 1080P 屏只剩 1040），
    导致 1080P 这一档被"不超过显示器"的过滤条件误杀。
    """
    try:
        from PySide6.QtGui import QGuiApplication
        scr = QGuiApplication.primaryScreen() or QGuiApplication.screens()[0]
        rect = scr.geometry()
        return rect.width(), rect.height()
    except Exception:
        return 1920, 1080


def _v_resolution_items() -> list[tuple[str, int, int]]:
    """可选导出分辨率；默认 1080P，2K/4K 一并提供。

    这里刻意不按显示器大小裁剪：导出分辨率与桌面显示是两回事，
    4K 档只是文件更大，不影响播放兼容性，用户有权主动选。
    """
    return [(f"480P（{854}×{480}）", 854, 480),
            (f"720P（{1280}×{720}）", 1280, 720),
            (f"1080P（{1920}×{1080}）", 1920, 1080),
            (f"2K（{2560}×{1440}）", 2560, 1440),
            (f"4K（{3840}×{2160}）", 3840, 2160)]


def _screen_refresh_rate() -> int:
    """当前主屏幕刷新率（Hz）；取不到时按 60 兜底。"""
    try:
        from PySide6.QtGui import QGuiApplication
        scr = QGuiApplication.primaryScreen() or QGuiApplication.screens()[0]
        rate = int(round(scr.refreshRate()))
    except Exception:
        rate = 60
    return rate if rate > 0 else 60


def _v_fps_items() -> list[tuple[str, int, int]]:
    """帧率档位 (label, fps, 默认索引)；默认取不超过屏幕刷新率的最大档。

    必须返回 (label, value, default_idx) 三元组：QComboBox.addItems 只认
    字符串，直接塞 int 会得到一排空白下拉项（v1.3.0 的 bug）。
    档位集合由用户指定，不给额外的"屏幕限高"（300Hz 一档允许存在），
    非标准帧率请走旁边的自定义输入框。
    """
    cands = [60, 120, 144, 165, 240, 300]
    items = [(f"{f} Hz", f) for f in cands]
    rate = _screen_refresh_rate()
    under = [f for _lb, f in items if f <= rate]                # 不超过屏幕刷新率的档
    default_fps = max(under) if under else cands[-1]
    idx = next((i for i, (_lb, f) in enumerate(items) if f == default_fps), 0)
    return [(lb, f, idx) for lb, f in items]


class Worker(QThread):
    """在后台线程执行插入/清除，结果通过信号回传主线程。"""
    log_signal = Signal(str)
    result_signal = Signal(bool, str)

    def __init__(self, fn):
        super().__init__()
        self.fn = fn

    def run(self):
        try:
            res = self.fn()
            self.log_signal.emit(f"操作成功：{res}")
            self.result_signal.emit(True, str(res))
        except Exception as e:
            self.log_signal.emit(f"操作失败：{e}")
            self.result_signal.emit(False, str(e))


class OfficeWorker(QThread):
    """后台执行 Office 注册表操作 / 退出测速，避免 20 秒的 bench 卡死 UI。

    fn 的返回值可以是 dict（修复/还原结果）也可以是 (秒数, 说明)（测速）。
    """
    result_signal = Signal(bool, str, bool)   # (ok, msg, need_admin)

    def __init__(self, fn):
        super().__init__()
        self.fn = fn

    def run(self):
        try:
            r = self.fn()
            if isinstance(r, dict):
                self.result_signal.emit(bool(r["ok"]), r["msg"], bool(r.get("need_admin")))
            elif r is None:
                self.result_signal.emit(False, "这次没测出结果", False)
            else:
                dt, note = r
                self.result_signal.emit(True,
                                        "关闭耗时 %.2f 秒\n%s" % (dt, note) if dt is not None else note,
                                        False)
        except Exception as e:
            self.result_signal.emit(False, str(e), False)


class FontWorker(QThread):
    """后台线程：读取可用字体（系统已安装字体优先；注册表读不到时才兜底 Word COM），避免卡 UI。"""
    finished_signal = Signal(list, str)   # (字体名列表, 来源说明)
    error_signal = Signal(str)

    def run(self):
        try:
            fonts, source = word_fonts.collect_fonts()
            self.finished_signal.emit(fonts or [], source)
        except Exception as e:
            self.error_signal.emit(str(e))


class VideoWorker(QThread):
    """后台逐帧处理视频水印：支持进度回传与中途取消。"""
    result_signal = Signal(bool, str)
    progress_signal = Signal(int, int)   # (当前帧, 总帧数)

    def __init__(self, src, output, opts, max_seconds=None):
        super().__init__()
        self.src = src
        self.output = output
        self.opts = opts
        self.max_seconds = max_seconds
        self._stop = False

    def request_stop(self):
        self._stop = True

    def run(self):
        try:  # 注意 max_seconds=None 时退化为处理整段视频，与旧行为完全一致
            res = video_mod.add_video_watermark(
                self.src, self.output, self.opts,
                # 兜底：进度值强制为安全 int（极端元数据下可能传进非有限值）
                progress_fn=lambda c, t: self.progress_signal.emit(
                    int(c), int(t) if isinstance(t, (int, float)) and t == t and t > 0 else 0),
                stop_check=lambda: self._stop,
                max_seconds=self.max_seconds,
            )
            self.result_signal.emit(True, str(res))
        except Exception as e:
            self.result_signal.emit(False, str(e))


class _DragLabel(QLabel):
    """可拖拽定位的预览标签：在显示的图像上按下/拖动，发射归一化坐标 (fx, fy)∈[0,1]，
    并绘制一个位置标记。供 Word 预览与视频预览帧共用，实现“拖拽自定义水印位置”。

    图像按 KeepAspectRatio 居中显示，鼠标坐标会换算回图像内容占比，避免黑边干扰。
    """

    dragged = Signal(float, float)   # (fx, fy)：水印中心在图像中的归一化位置

    def __init__(self, parent=None):
        super().__init__(parent)
        self._base = QPixmap()
        self._img_rect = QRect()
        self._marker = None           # QPointF（图像内归一化坐标），None 时不画
        self.setMouseTracking(True)
        self.setAlignment(Qt.AlignCenter)

    # ---- 外部接口 ----
    def setBasePixmap(self, pix):
        self._base = pix if pix is not None else QPixmap()
        # 换图时保留标记（拖拽过程中会重新合成，标记仍代表当前位置）
        self.update()

    def clearImage(self):
        self._base = QPixmap()
        self._marker = None
        self.update()

    def setMarker(self, fx, fy):
        self._marker = QPointF(max(0.0, min(1.0, fx)), max(0.0, min(1.0, fy)))
        self.update()

    def clearMarker(self):
        self._marker = None
        self.update()

    def hasImage(self):
        return not self._base.isNull()

    # ---- 内部 ----
    def _compute_rect(self):
        if self._base.isNull():
            self._img_rect = QRect()
            return
        ts = self._base.size().scaled(self.size(), Qt.KeepAspectRatio)
        x = (self.width() - ts.width()) // 2
        y = (self.height() - ts.height()) // 2
        self._img_rect = QRect(x, y, ts.width(), ts.height())

    def paintEvent(self, ev):
        super().paintEvent(ev)
        if self._base.isNull():
            return
        self._compute_rect()
        p = QPainter(self)
        p.drawPixmap(self._img_rect, self._base)
        if self._marker is not None:
            mx = self._img_rect.x() + self._marker.x() * self._img_rect.width()
            my = self._img_rect.y() + self._marker.y() * self._img_rect.height()
            r = max(8, int(min(self._img_rect.width(), self._img_rect.height()) * 0.06))
            pen = QPen(QColor(225, 6, 0)); pen.setWidth(3)
            p.setPen(pen)
            p.drawEllipse(QPointF(mx, my), r, r)
            p.drawLine(QPointF(mx - r, my), QPointF(mx + r, my))
            p.drawLine(QPointF(mx, my - r), QPointF(mx, my + r))
        p.end()

    def _frac(self, pos):
        if self._img_rect.isNull() or self._img_rect.width() == 0:
            return None
        fx = (pos.x() - self._img_rect.x()) / self._img_rect.width()
        fy = (pos.y() - self._img_rect.y()) / self._img_rect.height()
        return max(0.0, min(1.0, fx)), max(0.0, min(1.0, fy))

    def mousePressEvent(self, ev):
        if ev.button() == Qt.LeftButton and self.hasImage():
            f = self._frac(ev.pos())
            if f:
                self.setMarker(*f)
                self.dragged.emit(f[0], f[1])
        super().mousePressEvent(ev)

    def mouseMoveEvent(self, ev):
        if (ev.buttons() & Qt.LeftButton) and self.hasImage():
            f = self._frac(ev.pos())
            if f:
                self.setMarker(*f)
                self.dragged.emit(f[0], f[1])
        super().mouseMoveEvent(ev)


class App(QMainWindow):
    # 线程安全日志信号（必须在类级别声明）
    log_signal = Signal(str)
    status_signal = Signal(str)

    def __init__(self):
        super().__init__()
        self.setWindowTitle("一键水印工具（Word / 视频）")
        # 默认尺寸按屏幕自适应：内容最小宽度约 1.51k（左侧 Word 两列 + 右侧视频整列），
        # 旧默认 1200/1400 会顶出横向滚动条，看起来像“界面少了一列/位置错乱”。
        _sw, _sh = _screen_geometry()
        self.resize(max(1200, min(1560, int(_sw * 0.94))),
                    max(720, min(940, int(_sh * 0.90))))

        self.file_path = ""
        self.text_enabled = True     # 文本水印默认启用
        self.image_enabled = False   # 图像水印默认关闭
        self.text = "机密 CONFIDENTIAL"
        # 中西文分别指定字体：默认都先填“微软雅黑”，保持历史默认外观一致；
        # 用户可分别把“西文字体”改成 Times New Roman 等、把“中文字体”改成仿宋等
        self.cn_font_name = "微软雅黑"
        self.latin_font_name = "微软雅黑"
        self.font_size = 120.0
        self.color = (128, 128, 128)
        # 文本水印专属参数（与图像各自独立、可分别调整）
        self.text_angle = 45.0
        self.text_transparency = 0.5      # 0=不透明, 1=全透明
        self.text_scale = 1.0
        self.text_offset_x = 0.0          # 水平偏移（占页宽百分比，0=居中）
        self.text_offset_y = 0.0          # 垂直偏移（占页高百分比，0=居中）
        # 位置预设联动：写入偏移时屏蔽「改偏移→自动切自定义」的回环
        self._pos_applying = False
        # 图像水印专属参数
        self.img_angle = 45.0
        self.img_transparency = 0.5
        self.img_scale = 1.0
        self.img_offset_x = 0.0
        self.img_offset_y = 0.0
        self.image_path = ""
        self.wd = None
        self._worker = None  # 持有 Worker 引用：QThread 若被 GC 回收而线程仍在运行，进程会直接 abort 崩溃
        self._preview_img = None
        self._last_action = ""
        self._fonts_asked = False      # 是否已问过读取字体权限（每会话一次）
        self._fonts_loaded = False     # 是否已成功扩充字体下拉框
        self._font_worker = None       # 持有引用，避免线程运行中对象被 GC 导致崩溃
        # 防去除加固
        self.tile = False              # 平铺满页（默认关，外观与历史一致）
        self.tile_rows = 5
        self.tile_cols = 2
        self.redundant = True          # 多份冗余嵌入（对可见外观零影响，默认开）
        self._last_inserted = None     # 上次插入实际写入的水印份数（守护按份数校验）
        self.tray = None
        self._quitting = False
        self._tray_tried = 0         # 托盘初始化尝试次数
        self._tray_reason = ""       # 托盘初始化失败原因

        self.log_signal.connect(self._append_log)
        self.status_signal.connect(self._set_status)

        self._build()
        self._setup_tray()

    # ---------------------------------------------------------- 线程安全日志
    def _append_log(self, msg):
        self.log_text.append(msg)
        self.log_text.verticalScrollBar().setValue(self.log_text.verticalScrollBar().maximum())

    def _set_status(self, msg):
        self.status_label.setText(msg)

    def thread_log(self, msg, verbose=False):
        """供后台线程（守护）调用，切回主线程更新 UI。"""
        if verbose:
            self.status_signal.emit(msg)
        else:
            self.log_signal.emit(msg)

    # ------------------------------------------------------------------ UI
    def _build(self):
        # 用滚动区域承载全部内容：水印类型（文本/图像）横向并排后整体仍可能较高，
        # 滚动区域保证预览、按钮、守护等控件在任何屏幕高度下都可达，不会被裁掉。
        content = QWidget()
        root = QVBoxLayout(content)
        root.setContentsMargins(10, 10, 10, 10)
        root.setSpacing(8)

        scroll = QScrollArea()
        scroll.setWidget(content)
        scroll.setWidgetResizable(True)
        scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarAsNeeded)
        scroll.setVerticalScrollBarPolicy(Qt.ScrollBarAsNeeded)
        self.setCentralWidget(scroll)

        # ---------------- 左区容器：Word 的输入/输出行也归到这一列 ----------------
        # 之前这两行是通栏（横跨整窗、含视频列），右缘会伸得很远；
        # 放进左区后宽度 = 左区宽 = 文本列 + 图像列宽的合计，
        # 于是「浏览...」按钮的右缘天然与「图像水印」组框右缘竖向对齐。
        left = QWidget()
        vl = QVBoxLayout(left)
        vl.setContentsMargins(0, 0, 0, 0)
        vl.setSpacing(8)

        # 文件
        f_file = QWidget()
        h = QHBoxLayout(f_file); h.setContentsMargins(0, 0, 0, 0)
        self.file_edit = QLineEdit(); self.file_edit.setPlaceholderText("选择 Word 文件 (.docx / .doc)")
        h.addWidget(self.file_edit, 3)
        btn = QPushButton("浏览..."); btn.clicked.connect(self._browse_file); h.addWidget(btn, 1)
        vl.addWidget(f_file)

        # 输出文件（不破坏原文件）
        f_out = QWidget()
        ho = QHBoxLayout(f_out); ho.setContentsMargins(0, 0, 0, 0)
        ho.addWidget(QLabel("输出:"))
        self.out_edit = QLineEdit(); self.out_edit.setPlaceholderText("默认保存到桌面/<原名>WaterMark.docx，可点击浏览修改")
        ho.addWidget(self.out_edit, 3)
        bout = QPushButton("浏览..."); bout.clicked.connect(self._browse_output); ho.addWidget(bout, 1)
        ho.addStretch(1)
        vl.addWidget(f_out)
        self.out_edit.textChanged.connect(self._on_output_changed)

        # ---------------- 文本水印（可启用/禁用） ----------------
        f_text = self._make_group("文本水印")
        self.text_chk = QCheckBox("启用文本水印")
        self.text_chk.setChecked(self.text_enabled)
        self.text_chk.stateChanged.connect(self._on_text_enabled)
        f_text.layout().addWidget(self.text_chk)

        self.text_ctrl_widgets = []
        vt = QVBoxLayout(); vt.setContentsMargins(16, 0, 0, 0); vt.setSpacing(6)
        row = QHBoxLayout(); row.addWidget(QLabel("水印文字:"))
        self.text_edit = QLineEdit(self.text); row.addWidget(self.text_edit, 1)
        vt.addLayout(row)
        # 中西文字体分别设置（Word 式）：汉字用“中文字体”，拉丁字母/数字用“西文字体”
        row = QHBoxLayout(); row.addWidget(QLabel("中文字体:"))
        self.cn_font_combo, self.cn_font_model = self._make_font_combo(self.cn_font_name)
        self.cn_font_combo.setToolTip(
            "只作用于汉字与中文标点，例如“机密C”里的“机密”。\n"
            "可直接键入筛选（输入“fs”或“仿”快速定位仿宋）。")
        row.addWidget(self.cn_font_combo, 1)
        vt.addLayout(row)
        row = QHBoxLayout(); row.addWidget(QLabel("西文字体:"))
        self.latin_font_combo, self.latin_font_model = self._make_font_combo(self.latin_font_name)
        self.latin_font_combo.setToolTip(
            "只作用于字母、数字与半角符号，例如“机密C”里的“C”。\n"
            "可直接键入筛选（输入“t”快速定位 Times New Roman）。")
        row.addWidget(self.latin_font_combo, 1)
        vt.addLayout(row)
        row = QHBoxLayout(); row.addWidget(QLabel("字号:"))
        self.font_spin = QDoubleSpinBox(); self.font_spin.setRange(8, 300); self.font_spin.setDecimals(2)
        self.font_spin.setSingleStep(1.0); self.font_spin.setValue(self.font_size)
        row.addWidget(self.font_spin, 1)
        vt.addLayout(row)
        self.color_btn = QPushButton("文本颜色: RGB(128,128,128)")
        self.color_btn.clicked.connect(self._pick_color); vt.addWidget(self.color_btn)
        # 文本水印专属：旋转角度 / 透明度 / 缩放（与图像各自独立可调）
        row, self.text_angle_slider, self.text_angle_spin = self._make_slider_spin(
            "旋转角度(°):", -180, 180, self.text_angle, 2, 1.0, "", self._on_text_angle)
        vt.addLayout(row)
        row, self.text_trans_slider, self.text_trans_spin = self._make_slider_spin(
            "透明度(%):", 0, 100, self.text_transparency * 100, 2, 1.0, "%", self._on_text_trans)
        vt.addLayout(row)
        row = QHBoxLayout(); row.addWidget(QLabel("缩放:"))
        self.text_scale_spin = QDoubleSpinBox(); self.text_scale_spin.setRange(0.2, 2.0)
        self.text_scale_spin.setDecimals(2); self.text_scale_spin.setSingleStep(0.01)
        self.text_scale_spin.setValue(self.text_scale)
        self.text_scale_spin.valueChanged.connect(self._on_text_scale)
        self.text_scale_spin.valueChanged.connect(self._schedule_preview)
        row.addWidget(self.text_scale_spin, 1)
        vt.addLayout(row)
        # 位置预设：九宫格一键落位；手动改偏移或在预览图上拖拽会自动切到「自定义」
        row = QHBoxLayout(); row.addWidget(QLabel("位置:"))
        self.text_pos_combo = QComboBox()
        self.text_pos_combo.addItems(_W_POS_KEYS)
        self.text_pos_combo.setCurrentText("居中")
        self.text_pos_combo.setToolTip(
            "九宫格预设位置（左上/上中/右上 … 右下）。\n"
            "手动调整下方偏移，或在预览图上拖拽水印，会自动切到「自定义」。")
        self.text_pos_combo.currentTextChanged.connect(self._on_text_pos_preset)
        row.addWidget(self.text_pos_combo, 1)
        vt.addLayout(row)
        # 文本水印专属：水平/垂直偏移（支持上下左右调整，占页宽/页高百分比）
        # 这同时也是「自定义位置」的精确输入：X=50+水平偏移，Y=50+垂直偏移
        row, self.text_offx_slider, self.text_offx_spin = self._make_slider_spin(
            "水平偏移(%):", -50, 50, self.text_offset_x, 2, 1.0, "%", self._on_text_offx)
        vt.addLayout(row)
        row, self.text_offy_slider, self.text_offy_spin = self._make_slider_spin(
            "垂直偏移(%):", -50, 50, self.text_offset_y, 2, 1.0, "%", self._on_text_offy)
        vt.addLayout(row)
        for w in (self.text_edit, self.cn_font_combo, self.latin_font_combo,
                  self.font_spin, self.color_btn, self.text_pos_combo,
                  self.text_angle_slider, self.text_angle_spin,
                  self.text_trans_slider, self.text_trans_spin, self.text_scale_spin,
                  self.text_offx_slider, self.text_offx_spin,
                  self.text_offy_slider, self.text_offy_spin):
            self.text_ctrl_widgets.append(w)
        f_text.layout().addLayout(vt)
        # 文本/图像两列不再纵向拉伸：Maximum = 高度以内容为准、不跟随最高的视频列
        # （Minimum 仍会吸收多余空间，导致组标题被拉伸悬空——v1.3.3 踩过同样的坑）
        f_text.setSizePolicy(QSizePolicy.Preferred, QSizePolicy.Maximum)

        # ---------------- 图像水印（可启用/禁用） ----------------
        f_img = self._make_group("图像水印")
        self.img_chk = QCheckBox("启用图像水印")
        self.img_chk.setChecked(self.image_enabled)
        self.img_chk.stateChanged.connect(self._on_img_enabled)
        f_img.layout().addWidget(self.img_chk)

        self.img_ctrl_widgets = []
        vi = QVBoxLayout(); vi.setContentsMargins(16, 0, 0, 0); vi.setSpacing(6)
        row = QHBoxLayout(); row.addWidget(QLabel("图片路径:"))
        self.img_edit = QLineEdit(); row.addWidget(self.img_edit, 1)
        b = QPushButton("选择图片..."); b.clicked.connect(self._browse_image); row.addWidget(b)
        vi.addLayout(row)
        for w in (self.img_edit, b):
            self.img_ctrl_widgets.append(w)
        # 图像水印专属：旋转角度 / 透明度 / 缩放（与文本各自独立可调）
        row, self.img_angle_slider, self.img_angle_spin = self._make_slider_spin(
            "旋转角度(°):", -180, 180, self.img_angle, 2, 1.0, "", self._on_img_angle)
        vi.addLayout(row)
        row, self.img_trans_slider, self.img_trans_spin = self._make_slider_spin(
            "透明度(%):", 0, 100, self.img_transparency * 100, 2, 1.0, "%", self._on_img_trans)
        vi.addLayout(row)
        row = QHBoxLayout(); row.addWidget(QLabel("缩放:"))
        self.img_scale_spin = QDoubleSpinBox(); self.img_scale_spin.setRange(0.2, 2.0)
        self.img_scale_spin.setDecimals(2); self.img_scale_spin.setSingleStep(0.01)
        self.img_scale_spin.setValue(self.img_scale)
        self.img_scale_spin.valueChanged.connect(self._on_img_scale)
        self.img_scale_spin.valueChanged.connect(self._schedule_preview)
        row.addWidget(self.img_scale_spin, 1)
        vi.addLayout(row)
        # 位置预设：九宫格一键落位；手动改偏移或在预览图上拖拽会自动切到「自定义」
        row = QHBoxLayout(); row.addWidget(QLabel("位置:"))
        self.img_pos_combo = QComboBox()
        self.img_pos_combo.addItems(_W_POS_KEYS)
        self.img_pos_combo.setCurrentText("居中")
        self.img_pos_combo.setToolTip(
            "九宫格预设位置（左上/上中/右上 … 右下）。\n"
            "手动调整下方偏移，或在预览图上拖拽水印，会自动切到「自定义」。")
        self.img_pos_combo.currentTextChanged.connect(self._on_img_pos_preset)
        row.addWidget(self.img_pos_combo, 1)
        vi.addLayout(row)
        # 图像水印专属：水平/垂直偏移（支持上下左右调整，占页宽/页高百分比）
        row, self.img_offx_slider, self.img_offx_spin = self._make_slider_spin(
            "水平偏移(%):", -50, 50, self.img_offset_x, 2, 1.0, "%", self._on_img_offx)
        vi.addLayout(row)
        row, self.img_offy_slider, self.img_offy_spin = self._make_slider_spin(
            "垂直偏移(%):", -50, 50, self.img_offset_y, 2, 1.0, "%", self._on_img_offy)
        vi.addLayout(row)
        for w in (self.img_pos_combo,
                  self.img_angle_slider, self.img_angle_spin,
                  self.img_trans_slider, self.img_trans_spin, self.img_scale_spin,
                  self.img_offx_slider, self.img_offx_spin,
                  self.img_offy_slider, self.img_offy_spin):
            self.img_ctrl_widgets.append(w)
        f_img.layout().addLayout(vi)
        f_img.layout().addStretch(1)   # 底部位留白：盒高被拉到与文本水印一样高时从这里顶出
        # 页面明示支持的图片格式（仿宋红字），让用户知晓
        f_img.layout().addWidget(self._fmt_hint(
            "支持格式：png、jpg、jpeg、webp、pdf（PDF 仅取首页）"))
        f_img.setSizePolicy(QSizePolicy.Preferred, QSizePolicy.Maximum)
        # 文本/图像两列【底边平行】：行数少的图像列补到文本列高度。
        self._col_eq = _EqualBottoms(f_text, f_img)

        # Word 主操作按钮：两个按钮「同一行、彼此平行」（同一 Y），
        # 但各自只相对自己列的组框宽度居中，互不相干、也无视视频列。
        # 做法：下方单独一行 h_action，前两个格子的宽度紧跟文本/图像组框，
        # 格子内左右 addStretch 把按钮居中；最后一格 addStretch 吃下视频列的剩余宽度，
        # 于是按钮与上方三列严格对齐、又处同一水平线。
        self._btn_insert = _ProportionalButton("一键插入水印", f_text, ratio=0.8)
        self._btn_insert.setStyleSheet(BTN_HL)
        self._btn_insert.clicked.connect(self._insert)
        self._btn_clear = _ProportionalButton("一键清除水印", f_img, ratio=0.8)
        self._btn_clear.setStyleSheet(BTN_HL)
        self._btn_clear.clicked.connect(self._clear)

        class _ColumnAlignedButton(QWidget):
            """宽度紧跟宿主组框、内部把按钮水平居中的小容器。"""
            def __init__(self, btn, host):
                super().__init__()
                self._host = host
                hb = QHBoxLayout(self)
                hb.setContentsMargins(0, 0, 0, 0)
                hb.addStretch(1)        # 只在「本列宽度」内撑开 → 只相对本列组框居中
                hb.addWidget(btn)
                hb.addStretch(1)
                host.installEventFilter(self)
                self._sync()
            def eventFilter(self, obj, event):
                if event.type() == QEvent.Resize and obj is self._host:
                    self._sync()
                return super().eventFilter(obj, event)
            def _sync(self):
                w = self._host.width()
                if w > 0:
                    self.setFixedWidth(w)

        # ---------------- 防去除加固（可选）----------------
        # 位置：嵌在「图像水印」列组框的**正下方**——该列内容比文本水印列矮，
        # 这块正好补上两者的高度差；宽度自动等于图像水印列宽，视觉上与该列一体。
        f_hard = self._make_group("防去除加固（让水印更难被删掉）")
        vh = f_hard.layout()
        self.tile_chk = QCheckBox("平铺满页水印（覆盖整页，难以抹除）")
        self.tile_chk.setChecked(self.tile)
        self.tile_chk.setToolTip(
            "把水印从“单个居中”改成整页多行多列平铺。\n"
            "稀疏的单个水印，用 PS 的内容识别填充或 AI 去水印很容易抹掉且不留痕迹；\n"
            "覆盖整页的密集纹理要去掉就得把整页重画，难度陡增。\n"
            "注意：这是提升“去不掉”程度的关键手段，但没有任何方案能保证绝对去不掉。")
        vh.addWidget(self.tile_chk)
        row = QHBoxLayout(); row.addWidget(QLabel("平铺行列:"))
        self.tile_rows_spin = QDoubleSpinBox()
        self.tile_rows_spin.setRange(1, 30); self.tile_rows_spin.setDecimals(0)
        self.tile_rows_spin.setSingleStep(1); self.tile_rows_spin.setValue(self.tile_rows)
        self.tile_rows_spin.setToolTip("纵向行数：行数越多越密，越难被修图抹掉")
        row.addWidget(self.tile_rows_spin)
        row.addWidget(QLabel("× 列"))
        self.tile_cols_spin = QDoubleSpinBox()
        self.tile_cols_spin.setRange(1, 30); self.tile_cols_spin.setDecimals(0)
        self.tile_cols_spin.setSingleStep(1); self.tile_cols_spin.setValue(self.tile_cols)
        self.tile_cols_spin.setToolTip("横向列数：列数越多、单个水印越小")
        row.addWidget(self.tile_cols_spin)
        row.addStretch(1)
        vh.addLayout(row)
        self.redundant_chk = QCheckBox("多份冗余嵌入（页眉+页脚都写入）")
        self.redundant_chk.setChecked(self.redundant)
        self.redundant_chk.setToolTip(
            "除页眉外，把水印也写进页脚；并把图形的显示名改成普通图片那样的名字。\n"
            "Word/WPS 的“删除水印”按钮和多数去水印脚本都是按名称/结构匹配水印图形的，\n"
            "这样它们就找不到、删不干净。对肉眼外观没有任何影响。")
        vh.addWidget(self.redundant_chk)

        # 预览
        f_prev = self._make_group("水印预览（示意，脱离 Word 直接查看）")
        vp = f_prev.layout()
        self.preview_label = _DragLabel(); self.preview_label.setAlignment(Qt.AlignCenter)
        # 预览通栏后宽度充裕（横跨文本+图像两列），高度给足 A4 比例约 1:1.414 的可读尺寸
        self.preview_label.setMinimumSize(280, 400)
        self.preview_label.setStyleSheet("border:1px solid #bbb; background:#fafafa;")
        self.preview_label.setText("预览不可用")
        self.preview_label.dragged.connect(self._on_preview_drag)
        vp.addWidget(self.preview_label)
        self.preview_drag_chk = QCheckBox("拖拽自定义位置（在预览图上拖动水印）")
        self.preview_drag_chk.setToolTip("勾选后，可直接在预览图上用鼠标把水印拖到任意位置；"
                                         "放开后水平/垂直偏移会同步更新。")
        vp.addWidget(self.preview_drag_chk)
        hprev = QHBoxLayout()
        bp = QPushButton("预览水印"); bp.clicked.connect(self._render_preview); hprev.addWidget(bp)
        bps = QPushButton("保存预览图"); bps.clicked.connect(self._save_preview); hprev.addWidget(bps)
        vp.addLayout(hprev)

        # ---------------- 防去除加固 + 水印预览：三列下方并排一行 ----------------
        # 左「防去除加固」（窄）、右「水印预览」（宽，预览图更大更好看点）。
        # 两块都取 Maximum 高度并顶对齐，避免被行高拉伸出空白
        # （历史上 QSizePolicy.Minimum 会吸收多余空间、把组标题顶到悬空）。
        f_hard.setSizePolicy(QSizePolicy.Preferred, QSizePolicy.Maximum)
        f_prev.setSizePolicy(QSizePolicy.Preferred, QSizePolicy.Maximum)

        # ---------------- 图像水印列 = 图像水印组框 + 其正下方的「防去除加固」----------------
        # 两者同宽（同一列容器 → 自动等宽），竖直堆叠，与左侧文本水印列底部基本齐平。
        col_img = QWidget()
        vci = QVBoxLayout(col_img)
        vci.setContentsMargins(0, 0, 0, 0)
        vci.setSpacing(8)
        vci.addWidget(f_img, 0, Qt.AlignTop)
        vci.addWidget(f_hard, 0, Qt.AlignTop)
        vci.addStretch(1)

        # ---------------- 左侧 Word 区（竖直堆叠）：两列水印 → 通栏预览 → 主按钮行 ----------------
        # （left / vl 已在文件输入行之前创建，此处顺序继续往下堆）

        h_types = QHBoxLayout()
        h_types.setSpacing(8)
        h_types.addWidget(f_text, 0, Qt.AlignTop)     # 顶对齐：列高收缩到内容后不悬在行中间
        h_types.addWidget(col_img, 0, Qt.AlignTop)
        h_types.addStretch(1)          # 只占两列宽，剩余宽度整块留给右侧视频列
        vl.addLayout(h_types)

        # 水印预览：在左侧 Word 区内**通栏**（横跨文本+图像两列），不再与加固并排
        vl.addWidget(f_prev)

        h_action = QHBoxLayout()
        h_action.setSpacing(8)
        h_action.addWidget(_ColumnAlignedButton(self._btn_insert, f_text))
        h_action.addWidget(_ColumnAlignedButton(self._btn_clear, f_img))
        h_action.addStretch(1)         # 与上方两列严格对齐、且两按钮同一水平线
        vl.addLayout(h_action)
        vl.addStretch(1)

        # ---------------- 视频水印：独立成**右侧一整列**，与左侧 Word 区并排 ----------------
        f_video = self._build_video_section()

        # 左区宽度「固定为两列自然宽」：Fixed 策略下不会去抢右侧视频列的横向空间，
        # 多余宽度全部给视频列（否则左区会把空白吃掉，导致预览框被拉得极宽）。
        left.setSizePolicy(QSizePolicy.Fixed, QSizePolicy.Maximum)

        h_main = QHBoxLayout()
        h_main.setSpacing(8)
        h_main.addWidget(left, 0, Qt.AlignTop)
        h_main.addWidget(f_video, 1)
        root.addLayout(h_main)

        # 守护
        f_watch = self._make_group("后台守护（水印被删自动补回）")
        vw = f_watch.layout()
        self.watch_chk = QCheckBox("启用守护（水印被删自动补回）")
        self.watch_chk.stateChanged.connect(self._toggle_watch); vw.addWidget(self.watch_chk)
        row = QHBoxLayout(); row.addWidget(QLabel("检查间隔(秒):"))
        self.interval_spin = QDoubleSpinBox(); self.interval_spin.setRange(0.2, 10)
        self.interval_spin.setDecimals(2); self.interval_spin.setSingleStep(0.2); self.interval_spin.setValue(1.0)
        row.addWidget(self.interval_spin)
        vw.addLayout(row)
        root.addWidget(f_watch)

        # ---------------- Word 秒退（关闭卡顿） ----------------
        f_exit = self._make_group("Word 秒退（修复关闭 Word 时的十几秒卡顿）")
        vex = f_exit.layout()
        self._exit_info = QLabel("未检测")
        self._exit_info.setWordWrap(True)
        self._exit_info.setStyleSheet("color:#444;")
        vex.addWidget(self._exit_info)
        hex_ = QHBoxLayout()
        b = QPushButton("诊断"); b.clicked.connect(self._office_diagnose); hex_.addWidget(b)
        self._btn_fix = QPushButton("一键修复"); self._btn_fix.clicked.connect(self._office_fix); hex_.addWidget(self._btn_fix)
        b = QPushButton("测速"); b.clicked.connect(self._office_bench); hex_.addWidget(b)
        self._btn_revert = QPushButton("还原"); self._btn_revert.clicked.connect(self._office_revert); hex_.addWidget(self._btn_revert)
        vex.addLayout(hex_)
        # 可选加码默认不写：这些项会改变 Office 联网行为，甚至把已登录的账号踢下线，
        # 必须由用户显式确认后才会动，不能混进「一键修复」
        hopt = QHBoxLayout()
        b = QPushButton("可选加码（会踢账号，慎用）")
        b.setStyleSheet("color:#b00;")
        b.clicked.connect(self._office_optional); hopt.addWidget(b)
        lbl = QLabel("「一键修复」只禁加载项 + 关遥测，不动登录状态")
        lbl.setStyleSheet("color:#666;")
        hopt.addWidget(lbl, 1)
        vex.addLayout(hopt)
        self._exit_detail = QTextEdit(); self._exit_detail.setReadOnly(True)
        self._exit_detail.setFixedHeight(66)
        self._exit_detail.setPlaceholderText("诊断 / 修复 / 测速结果会显示在这里")
        vex.addWidget(self._exit_detail)
        root.addWidget(f_exit)

        # 日志
        f_log = self._make_group("日志")
        vl = f_log.layout()
        self.log_text = QTextEdit(); self.log_text.setReadOnly(True); vl.addWidget(self.log_text)
        root.addWidget(f_log, 1)

        self.status_label = QLabel("就绪"); root.addWidget(self.status_label)

        # 启动即渲染一次默认预览
        self._render_preview()

        # 实时预览：任一参数变化后 200ms 自动刷新（防抖，避免拖动滑块时频繁渲染）
        self._preview_timer = QTimer(self)
        self._preview_timer.setSingleShot(True)
        self._preview_timer.setInterval(200)
        self._preview_timer.timeout.connect(self._render_preview)
        self.text_edit.textChanged.connect(self._schedule_preview)
        self.cn_font_combo.currentTextChanged.connect(self._schedule_preview)
        self.latin_font_combo.currentTextChanged.connect(self._schedule_preview)
        self.font_spin.valueChanged.connect(self._schedule_preview)
        self.img_edit.textChanged.connect(self._schedule_preview)
        self.text_chk.stateChanged.connect(self._schedule_preview)
        self.img_chk.stateChanged.connect(self._schedule_preview)
        # 加固参数：改动即刷新预览（平铺密度直接影响观感）
        self.tile_chk.stateChanged.connect(self._on_harden_changed)
        self.tile_rows_spin.valueChanged.connect(self._on_harden_changed)
        self.tile_cols_spin.valueChanged.connect(self._on_harden_changed)

    def _build_video_section(self):
        """构建右侧“视频水印”整列：输入/输出行（在组框上方） + 视频水印组框。

        源视频与输出行原先嵌在组框内部，要进框里才能看到；提到组框上方后，
        与左侧 Word 区的“Word 文件/输出”两行处在同一水平线，左右对称更好找。
        """
        # 整列容器：0 边距，让两行的左右缘严格等于组框边缘
        wrap = QWidget()
        vcol = QVBoxLayout(wrap)
        vcol.setContentsMargins(0, 0, 0, 0)
        vcol.setSpacing(8)

        g = self._make_group("视频水印（与 Word 水印独立）")
        v = g.layout()
        v.setSpacing(6)
        # 页面明示支持的视频格式（仿宋红字），让用户知晓
        v.addWidget(self._fmt_hint(
            "支持视频格式：mp4、avi、mov、mkv、wmv、flv、m4v、webm 等主流视频"))

        if video_mod is None:
            v.addWidget(QLabel("⚠ 视频依赖 imageio / imageio-ffmpeg 未安装，视频功能不可用。"))
            vcol.addWidget(g)
            return wrap

        # 源视频 / 输出 —— 放在组框【上方】，宽度与组框一致
        f_vsrc = QWidget()
        h1 = QHBoxLayout(f_vsrc); h1.setContentsMargins(0, 0, 0, 0)
        self.v_src_edit = QLineEdit(); self.v_src_edit.setPlaceholderText("选择视频文件（mp4 / mkv / avi / mov / wmv …）")
        h1.addWidget(self.v_src_edit, 3)
        b1 = QPushButton("浏览..."); b1.clicked.connect(self._v_browse_src);         h1.addWidget(b1, 1)
        # Maximum：高度只跟两个控件走，不去吃掉这一列多出来的竖直空间
        f_vsrc.setSizePolicy(QSizePolicy.Preferred, QSizePolicy.Maximum)
        vcol.addWidget(f_vsrc, 0, Qt.AlignTop)
        f_vout = QWidget()
        h2 = QHBoxLayout(f_vout); h2.setContentsMargins(0, 0, 0, 0)
        h2.addWidget(QLabel("输出:"))
        self.v_out_edit = QLineEdit(); self.v_out_edit.setPlaceholderText("默认 <原名>WaterMark.mp4，可修改")
        h2.addWidget(self.v_out_edit, 3)
        b2 = QPushButton("浏览..."); b2.clicked.connect(self._v_browse_out);         h2.addWidget(b2, 1)
        f_vout.setSizePolicy(QSizePolicy.Preferred, QSizePolicy.Maximum)
        vcol.addWidget(f_vout, 0, Qt.AlignTop)
        vcol.addWidget(g, 0, Qt.AlignTop)
        # 收口：把两行以下多余的竖直空间交给末尾弹性留白，
        # 否则Qt会把多余高度平摊给列内每个控件，输入/输出行会被拉成上百像素高。
        vcol.addStretch(1)

        # 水印类型 + 各自的播放方式（可自由搭配）
        ht = QHBoxLayout()
        self.v_text_chk = QCheckBox("文字水印"); self.v_text_chk.setChecked(True)
        self.v_text_motion_combo = QComboBox()
        self.v_text_motion_combo.addItems(_V_MOTION_LABELS)
        self.v_text_motion_combo.setCurrentText("跟随")
        self.v_img_chk = QCheckBox("图片水印")
        self.v_img_motion_combo = QComboBox()
        self.v_img_motion_combo.addItems(_V_MOTION_LABELS)
        self.v_img_motion_combo.setCurrentText("跟随")
        ht.addWidget(self.v_text_chk); ht.addWidget(QLabel("播放方式:"))
        ht.addWidget(self.v_text_motion_combo)
        ht.addSpacing(12)
        ht.addWidget(self.v_img_chk); ht.addWidget(QLabel("播放方式:"))
        ht.addWidget(self.v_img_motion_combo)
        ht.addStretch(1)
        v.addLayout(ht)

        # 文字水印参数
        htt = QHBoxLayout()
        htt.addWidget(QLabel("文字:"))
        self.v_text_edit = QLineEdit("机密 CONFIDENTIAL"); htt.addWidget(self.v_text_edit, 3)
        self.v_color = (255, 255, 255)
        self.v_color_btn = QPushButton("颜色: 白"); self.v_color_btn.clicked.connect(self._v_pick_color)
        htt.addWidget(self.v_color_btn, 1)
        v.addLayout(htt)
        hdr = QHBoxLayout()
        self.v_alpha_spin = QDoubleSpinBox(); self.v_alpha_spin.setRange(0, 100); self.v_alpha_spin.setValue(70)
        self.v_alpha_spin.setSuffix("%"); hdr.addWidget(QLabel("透明度:")); hdr.addWidget(self.v_alpha_spin)
        self.v_size_spin = QDoubleSpinBox(); self.v_size_spin.setRange(2, 20); self.v_size_spin.setValue(6)
        self.v_size_spin.setSuffix("%高"); hdr.addWidget(QLabel("字号:")); hdr.addWidget(self.v_size_spin)
        v.addLayout(hdr)
        self.v_speed_spin = QDoubleSpinBox(); self.v_speed_spin.setRange(1, 50); self.v_speed_spin.setValue(12)
        self.v_speed_spin.setSuffix("%/秒"); self.v_speed_spin.setEnabled(False)  # 纯固定时不相关

        # 图片水印参数
        hi = QHBoxLayout()
        self.v_img_edit = QLineEdit(); self.v_img_edit.setPlaceholderText("水印图片路径（可选）")
        hi.addWidget(self.v_img_edit, 3)
        bi = QPushButton("选择图片..."); bi.clicked.connect(self._v_browse_img); hi.addWidget(bi, 1)
        v.addLayout(hi)
        hir = QHBoxLayout()
        self.v_img_alpha_spin = QDoubleSpinBox(); self.v_img_alpha_spin.setRange(0, 100); self.v_img_alpha_spin.setValue(70)
        self.v_img_alpha_spin.setSuffix("%"); hir.addWidget(QLabel("透明度:")); hir.addWidget(self.v_img_alpha_spin)
        self.v_img_scale_spin = QDoubleSpinBox(); self.v_img_scale_spin.setRange(5, 60); self.v_img_scale_spin.setValue(15)
        self.v_img_scale_spin.setSuffix("%宽"); hir.addWidget(QLabel("大小:")); hir.addWidget(self.v_img_scale_spin)
        v.addLayout(hir)

        # 播放方式：固定 / 滚动 / 固定+滚动（单选，三者互斥）
        hm = QHBoxLayout()
        hm.addWidget(QLabel("播放方式:"))
        self.v_motion_group = QButtonGroup(self)
        self.v_motion_group.setExclusive(True)
        self.v_motion_fixed = QRadioButton("固定")
        self.v_motion_scroll = QRadioButton("滚动")
        self.v_motion_both = QRadioButton("固定+滚动")
        self.v_motion_group.addButton(self.v_motion_fixed, 0)
        self.v_motion_group.addButton(self.v_motion_scroll, 1)
        self.v_motion_group.addButton(self.v_motion_both, 2)
        self.v_motion_fixed.setChecked(True)
        for rb in (self.v_motion_fixed, self.v_motion_scroll, self.v_motion_both):
            hm.addWidget(rb)
        hm.addStretch(1)
        v.addLayout(hm)
        # 改全局播放方式时，把两个“跟随”下拉同步过来，避免界面看起来不一致
        self.v_motion_group.buttonClicked.connect(self._v_sync_motion_combos)
        # 类型勾选 / 播放方式变动时，实时刷新「滚动速度」是否可用
        self.v_text_chk.stateChanged.connect(self._v_refresh_motion_ui)
        self.v_img_chk.stateChanged.connect(self._v_refresh_motion_ui)
        self.v_motion_group.buttonClicked.connect(self._v_refresh_motion_ui)
        self.v_text_motion_combo.currentTextChanged.connect(self._v_refresh_motion_ui)
        self.v_img_motion_combo.currentTextChanged.connect(self._v_refresh_motion_ui)
        self._v_refresh_motion_ui()   # 建完控件后再刷一次，保证初始态正确

        # 位置 / 速度：文字与图片各自一个锚点，避免同位置完全重叠。
        # X/Y 为「水印左上角相对画面」的百分比（0,0=左上角；100,100=右下角外），
        # 选预设会把 X/Y 回填，手改 X/Y 或在预览帧上拖拽会自动切到「自定义」。
        self._v_pos_applying = False
        # 拆成两行（文字 / 图片各一行），单行控件更少 → 视频列最小宽度明显变窄，
        # 否则三列并排时整页最小宽度会超过窗口、被迫横向滚动（看起来像“少了一列”）。
        hp1 = QHBoxLayout()
        self.v_text_pos_combo = QComboBox()
        self.v_text_pos_combo.addItems(_V_POSITIONS_KEYS)
        self.v_text_pos_combo.setCurrentText("右下")
        self.v_text_pos_combo.setToolTip("预设锚点；选「自定义」后以右侧 X/Y 百分比为准")
        self.v_text_x_spin = self._v_make_xy_spin(82)
        self.v_text_y_spin = self._v_make_xy_spin(85)
        hp1.addWidget(QLabel("文字位置:"))
        hp1.addWidget(self.v_text_pos_combo)
        hp1.addWidget(QLabel("X%:")); hp1.addWidget(self.v_text_x_spin)
        hp1.addWidget(QLabel("Y%:")); hp1.addWidget(self.v_text_y_spin)
        hp1.addStretch(1)
        v.addLayout(hp1)

        hp2 = QHBoxLayout()
        self.v_img_pos_combo = QComboBox()
        self.v_img_pos_combo.addItems(_V_POSITIONS_KEYS)
        self.v_img_pos_combo.setCurrentText("左下")   # 与文字错位，默认不打架
        self.v_img_pos_combo.setToolTip("预设锚点；选「自定义」后以右侧 X/Y 百分比为准")
        self.v_img_x_spin = self._v_make_xy_spin(8)
        self.v_img_y_spin = self._v_make_xy_spin(85)
        hp2.addWidget(QLabel("图片位置:"))
        hp2.addWidget(self.v_img_pos_combo)
        hp2.addWidget(QLabel("X%:")); hp2.addWidget(self.v_img_x_spin)
        hp2.addWidget(QLabel("Y%:")); hp2.addWidget(self.v_img_y_spin)
        hp2.addSpacing(10)
        hp2.addWidget(QLabel("滚动速度:"))
        hp2.addWidget(self.v_speed_spin)
        hp2.addStretch(1)
        v.addLayout(hp2)
        # 预设 ↔ X/Y 双向联动
        self.v_text_pos_combo.currentTextChanged.connect(self._v_on_text_preset)
        self.v_img_pos_combo.currentTextChanged.connect(self._v_on_img_preset)
        for sp in (self.v_text_x_spin, self.v_text_y_spin):
            sp.valueChanged.connect(self._v_on_text_xy)
        for sp in (self.v_img_x_spin, self.v_img_y_spin):
            sp.valueChanged.connect(self._v_on_img_xy)

        # 导出规格：分辨率 / 帧率（+自定义）/ 画质
        hs = QHBoxLayout()
        hs.addWidget(QLabel("导出分辨率:"))
        self.v_res_combo = QComboBox()
        self.v_res_items = _v_resolution_items()
        self.v_res_combo.addItems([lb for lb, _w, _h in self.v_res_items])
        if self.v_res_items:
            self.v_res_combo.setCurrentIndex(
                min(2, len(self.v_res_items) - 1))        # 默认 1080P
        hs.addWidget(self.v_res_combo, 0)
        hs.addWidget(QLabel("帧率:"))
        self.v_fps_combo = QComboBox()
        self.v_fps_items = _v_fps_items()                 # [(label, fps, idx), ...]
        self.v_fps_combo.addItems([lb for lb, _f, _d in self.v_fps_items])
        if self.v_fps_items:
            self.v_fps_combo.setCurrentIndex(self.v_fps_items[0][2])
        self.v_fps_custom_value = None                    # 自定义帧率，优先于下拉
        self.v_fps_combo.currentIndexChanged.connect(self._v_clear_custom_fps)
        hs.addWidget(self.v_fps_combo, 0)
        self.v_fps_custom_edit = QLineEdit()
        self.v_fps_custom_edit.setPlaceholderText("自定义")
        self.v_fps_custom_edit.setToolTip("直接输入帧率（1~300）后回车或点「应用」，"
                                          "例如 25 / 50 / 100")
        self.v_fps_custom_edit.setMaximumWidth(80)
        # 校验器放宽到 9999，越界时交给 _v_apply_custom_fps 弹明确提示，
        # 否则用户敲 1000 会被静默吃掉，搞不清到底哪儿不对
        self.v_fps_custom_edit.setValidator(QIntValidator(1, 9999))
        self.v_fps_custom_edit.returnPressed.connect(self._v_apply_custom_fps)
        hs.addWidget(self.v_fps_custom_edit, 0)
        self.v_fps_apply_btn = QPushButton("应用")
        self.v_fps_apply_btn.setFixedWidth(60)
        self.v_fps_apply_btn.clicked.connect(self._v_apply_custom_fps)
        hs.addWidget(self.v_fps_apply_btn, 0)
        hs.addWidget(QLabel("画质:"))
        self.v_crf_combo = QComboBox()
        self.v_crf_combo.addItems(list(video_mod.CRF_PRESETS.keys()))
        self.v_crf_combo.setCurrentText("标准")
        hs.addWidget(self.v_crf_combo, 0)
        hs.addStretch(1)
        v.addLayout(hs)

        # ---------------- 预览（拖拽可自定义位置 / 可播放短片） ----------------
        f_vprev = self._make_group("视频水印预览（改参数即时重绘 / 拖拽自定义位置）")
        vpv = f_vprev.layout()
        vpv.setSpacing(6)
        self.v_custom_pos = None          # 拖拽自定义位置（图像归一化 fx,fy）
        self._v_last_frame = None        # 抓取到的原始帧（PIL RGB），用于实时重绘
        self._v_preview_out = ""
        self.v_frame_label = _DragLabel()
        self.v_frame_label.setMinimumSize(360, 203)
        self.v_frame_label.setStyleSheet("border:1px solid #bbb; background:#222;")
        self.v_frame_label.setText("点“预览水印效果”即可查看（未选视频时用示意画面）")
        self.v_frame_label.dragged.connect(self._on_video_frame_drag)
        vpv.addWidget(self.v_frame_label)
        # 拖拽自定义位置开关：开启后可在帧上拖动定位，落点写回 X/Y 并切到「自定义」
        self.v_drag_chk = QCheckBox("拖拽自定义位置（在预览帧上拖动水印）")
        self.v_drag_chk.setToolTip("勾选后可在预览帧上用鼠标把水印拖到任意位置；"
                                   "落点会自动写入上方 X/Y 百分比，下拉切到「自定义」。")
        self.v_drag_chk.stateChanged.connect(self._v_toggle_drag)
        vpv.addWidget(self.v_drag_chk)
        self.v_pos_lbl = QLabel("自定义位置：未设定（默认右下/左下）")
        self.v_pos_lbl.setWordWrap(True)
        vpv.addWidget(self.v_pos_lbl)
        hvprev = QHBoxLayout()
        self.v_grab_btn = QPushButton("预览水印效果")
        self.v_grab_btn.setStyleSheet(BTN_HL)
        self.v_grab_btn.setToolTip("按当前设置在画面上叠好水印并显示在上方预览区；"
                                   "未选视频时用 16:9 示意画面，同样能确认位置与效果。")
        self.v_grab_btn.clicked.connect(self._v_grab_frame)
        hvprev.addWidget(self.v_grab_btn)
        hvprev.addWidget(QLabel("时刻(秒):"))
        self.v_t_spin = QDoubleSpinBox()
        self.v_t_spin.setRange(0, 3600); self.v_t_spin.setDecimals(1)
        self.v_t_spin.setSingleStep(0.5); self.v_t_spin.setValue(1.0)
        self.v_t_spin.setMaximumWidth(80)
        self.v_t_spin.setToolTip("取视频第几秒的画面来预览；滚动水印可用它确认不同时刻的位置")
        self.v_t_spin.valueChanged.connect(self._v_grab_frame)
        hvprev.addWidget(self.v_t_spin)
        hvprev.addStretch(1)
        self.v_play_btn = QPushButton("播放短片预览")
        self.v_play_btn.setStyleSheet(BTN_HL)
        self.v_play_btn.setToolTip("导出前 3 秒带水印片段，并用系统默认播放器播放，"
                                   "用于确认滚动/位置等动态效果。")
        self.v_play_btn.clicked.connect(self._v_play_preview)
        hvprev.addWidget(self.v_play_btn)
        vpv.addLayout(hvprev)
        v.addWidget(f_vprev)

        # ---------------- 参数改动 → 预览即时重绘（防抖 250ms） ----------------
        # 只在「已经抓到画面」时重绘：不会偷偷去读视频，但改文字/透明度/位置等
        # 能立刻看到效果，不必反复点「预览水印效果」。
        self._v_frame_timer = QTimer(self)
        self._v_frame_timer.setSingleShot(True)
        self._v_frame_timer.setInterval(250)
        self._v_frame_timer.timeout.connect(self._v_show_frame)
        for w in (self.v_text_edit, self.v_img_edit,
                  self.v_alpha_spin, self.v_size_spin,
                  self.v_img_alpha_spin, self.v_img_scale_spin,
                  self.v_text_motion_combo, self.v_img_motion_combo):
            sig = getattr(w, "textChanged", None) or getattr(w, "valueChanged", None)
            if sig is None:
                sig = w.currentTextChanged
            sig.connect(self._v_schedule_frame)
        for rb in (self.v_motion_fixed, self.v_motion_scroll, self.v_motion_both):
            rb.toggled.connect(self._v_schedule_frame)
        self.v_text_chk.stateChanged.connect(self._v_schedule_frame)
        self.v_img_chk.stateChanged.connect(self._v_schedule_frame)

        # 运行按钮 + 进度 + 取消
        hr = QHBoxLayout()
        self.v_run_btn = QPushButton("开始加水印"); self.v_run_btn.clicked.connect(self._v_run)
        self.v_run_btn.setStyleSheet(BTN_HL)          # 与主操作一致：蓝底黑字高亮
        # 取消 = 灰底黑字高亮（次操作），与主操作「开始加水印」的蓝底区分
        self.v_cancel_btn = QPushButton("取消")
        self.v_cancel_btn.setStyleSheet(BTN_BROWN)
        self.v_cancel_btn.setEnabled(False)
        self.v_cancel_btn.clicked.connect(self._v_cancel)
        hr.addWidget(self.v_run_btn); hr.addWidget(self.v_cancel_btn)
        v.addLayout(hr)
        self.v_progress = QProgressBar()
        v.addWidget(self.v_progress)
        self.v_status_lbl = QLabel(
            "播放方式可选「固定 / 滚动 / 固定+滚动」；文字水印与图片水印还能各自单独指定，"
            "因此可以做出「文字滚动 + 图片固定」等任意搭配。"
            "导出分辨率默认 1080P，可选至 4K；帧率可选 60/120/144/165/240/300 Hz，"
            "默认取不超过你屏幕刷新率的那一档，非标准帧率可在右侧输入框直接填写（1~300）。"
            "导出帧率与原视频不同时会自动按时间轴补帧/丢帧，成片时长与原视频一致"
            "（帧率越高文件越大、处理越慢）。"
            "图片水印采用超采样渲染，放大导出时依然锐利。逐帧处理较长视频较慢属正常，原音轨会自动保留。")
        self.v_status_lbl.setWordWrap(True)
        v.addWidget(self.v_status_lbl)
        return wrap

    def _v_make_xy_spin(self, value):
        """X/Y 百分比输入（0~100，1% 一档），用于精确自定义水印落点。"""
        sp = QDoubleSpinBox()
        sp.setRange(0, 100); sp.setDecimals(0); sp.setSingleStep(1)
        sp.setSuffix("%"); sp.setValue(value)
        sp.setMaximumWidth(74)
        sp.setToolTip("自定义位置：水印左上角相对画面的百分比（0=最左/最上，100=最右/最下）")
        return sp

    def _v_position(self, combo=None, x_spin=None, y_spin=None):
        """取位置锚点（左上角占比 0..1）。

        combo 选中「自定义」时以 X/Y 百分比输入为准；不传控件则按预设解析
        （兼容只传 combo 的旧调用）。文字与图片各有独立锚点。
        """
        if combo is None:
            combo = self.v_text_pos_combo
        name = combo.currentText()
        p = _V_POSITIONS.get(name, (0.82, 0.85))
        if p is not None:
            return p
        # 「自定义」：读 X/Y 百分比
        if x_spin is None or y_spin is None:
            x_spin = self.v_text_x_spin if combo is self.v_text_pos_combo else self.v_img_x_spin
            y_spin = self.v_text_y_spin if combo is self.v_text_pos_combo else self.v_img_y_spin
        fx = max(0.0, min(1.0, float(x_spin.value()) / 100.0))
        fy = max(0.0, min(1.0, float(y_spin.value()) / 100.0))
        return (fx, fy)

    # --------------------------------------------- 视频位置：预设 ↔ X/Y 联动
    def _v_mark_custom(self, combo):
        """手改 X/Y 或在预览帧上拖拽后，把位置下拉切到「自定义」。"""
        if getattr(self, "_v_pos_applying", False):
            return
        if combo.currentText() == "自定义":
            return
        combo.blockSignals(True)
        combo.setCurrentText("自定义")
        combo.blockSignals(False)

    def _v_fill_xy(self, combo, x_spin, y_spin):
        """选了预设锚点 → 把该锚点回填进 X/Y 输入，用户可直接在此基础上微调。"""
        p = _V_POSITIONS.get(combo.currentText())
        if p is None:      # 「自定义」不覆盖
            return
        self._v_pos_applying = True
        try:
            x_spin.setValue(round(p[0] * 100))
            y_spin.setValue(round(p[1] * 100))
        finally:
            self._v_pos_applying = False
        self._v_schedule_frame()

    def _v_on_text_preset(self, _text=None):
        self._v_fill_xy(self.v_text_pos_combo, self.v_text_x_spin, self.v_text_y_spin)

    def _v_on_img_preset(self, _text=None):
        self._v_fill_xy(self.v_img_pos_combo, self.v_img_x_spin, self.v_img_y_spin)

    def _v_on_text_xy(self, _v=None):
        self._v_mark_custom(self.v_text_pos_combo)
        self._v_schedule_frame()

    def _v_on_img_xy(self, _v=None):
        self._v_mark_custom(self.v_img_pos_combo)
        self._v_schedule_frame()

    def _v_motion(self):
        """把单选的播放方式映射为引擎 motion 值。"""
        checked = self.v_motion_group.checkedButton()
        label = checked.text() if checked else "固定"
        return _V_MOTION_MAP.get(label, "fixed")

    def _v_sync_motion_combos(self):
        """全局播放方式变动时，让两个“跟随”下拉跟着走。"""
        for combo in (getattr(self, "v_text_motion_combo", None),
                      getattr(self, "v_img_motion_combo", None)):
            if combo is not None and combo.currentText() == "跟随":
                combo.setCurrentText(_V_MOTION_INVERSE.get(self._v_motion(), "跟随"))

    def _v_browse_src(self):
        p, _ = QFileDialog.getOpenFileName(self, "选择视频文件", "",
                                           "视频 (*.mp4 *.mkv *.avi *.mov *.wmv *.flv *.webm *.mpeg *.mpg *.ts *.m4v *.3gp);;All (*.*)")
        if p:
            self.v_src_edit.setText(p)
            if not self.v_out_edit.text().strip():
                base, _ = os.path.splitext(p)
                self.v_out_edit.setText(base + "WaterMark.mp4")
            self._v_grab_frame()     # 选完视频自动抓一帧预览，方便立即看效果

    def _v_browse_out(self):
        p, _ = QFileDialog.getSaveFileName(self, "选择输出视频", "", "MP4 (*.mp4);;All (*.*)")
        if p:
            self.v_out_edit.setText(p)

    def _v_browse_img(self):
        p, _ = QFileDialog.getOpenFileName(self, "选择水印图片", "",
                                           "图片 (*.png *.jpg *.jpeg *.webp *.pdf);;All (*.*)")
        if p:
            self.v_img_edit.setText(p)

    def _v_out_spec(self):
        """导出分辨率 (w, h)；不超过当前显示器可显示的上限。"""
        idx = self.v_res_combo.currentIndex()
        if 0 <= idx < len(self.v_res_items):
            _lb, w, h = self.v_res_items[idx]
            return w, h
        return _screen_geometry()

    def _v_clear_custom_fps(self):
        """改回预设档位时，丢弃用户此前输入的自定义帧率。"""
        self.v_fps_custom_value = None

    def _v_apply_custom_fps(self):
        """解析自定义帧率输入框；非法值弹提示且不生效。"""
        raw = self.v_fps_custom_edit.text().strip()
        if not raw:
            return
        try:
            val = int(raw)
        except ValueError:
            QMessageBox.warning(self, "帧率无效", "请输入 1~300 之间的整数帧率。")
            return
        if not 1 <= val <= 300:
            QMessageBox.warning(self, "帧率无效", f"帧率需在 1~300 之间，当前为 {val}。")
            return
        self.v_fps_custom_value = val
        self.v_fps_custom_edit.setStyleSheet("")          # 清掉可能的标红

    def _v_out_fps(self):
        if self.v_fps_custom_value:                       # 自定义值优先
            return self.v_fps_custom_value
        idx = self.v_fps_combo.currentIndex()
        if 0 <= idx < len(self.v_fps_items):
            return self.v_fps_items[idx][1]
        return 60

    def _v_out_crf(self):
        return video_mod.CRF_PRESETS.get(self.v_crf_combo.currentText(), 18)

    def _v_pick_color(self):
        c = QColorDialog.getColor(QColor(*self.v_color), self, "选择水印文字颜色")
        if c.isValid():
            self.v_color = (c.red(), c.green(), c.blue())
            self.v_color_btn.setText(f"颜色: RGB{self.v_color}")

    def _v_gather_opts(self):
        kinds = []
        if self.v_text_chk.isChecked():
            kinds.append("text")
        if self.v_img_chk.isChecked():
            kinds.append("image")
        # 位置：预设下拉 / X-Y 百分比 / 预览帧拖拽，三者最终都落在同一处——
        # 拖拽会把落点写回 X/Y 并把下拉切到「自定义」，所以这里只需解析下拉 + X/Y。
        tpos = self._v_position(self.v_text_pos_combo,
                               self.v_text_x_spin, self.v_text_y_spin)
        ipos = self._v_position(self.v_img_pos_combo,
                               self.v_img_x_spin, self.v_img_y_spin)
        text_cfg = {
            "text": self.v_text_edit.text(),
            "color": self.v_color,
            "alpha": int(self.v_alpha_spin.value() / 100.0 * 255),
            "size_frac": self.v_size_spin.value() / 100.0,
            "position": tpos,
        }
        image_cfg = {
            "image_path": self.v_img_edit.text().strip(),
            "alpha": int(self.v_img_alpha_spin.value() / 100.0 * 255),
            "img_frac": self.v_img_scale_spin.value() / 100.0,
            "position": ipos,
        }
        global_motion = self._v_motion()
        # 文字/图片各自可单独指定播放方式；选“跟随”时才用全局的
        if "image" in kinds:
            image_cfg["motion"] = _resolve_motion(self.v_img_motion_combo, global_motion)
        if "text" in kinds:
            text_cfg["motion"] = _resolve_motion(self.v_text_motion_combo, global_motion)
        opts = {
            "kinds": kinds,
            "motion": global_motion,
            "scroll_speed": self.v_speed_spin.value() / 100.0,
            "text": text_cfg,
            "image": image_cfg,
            "out_size": self._v_out_spec(),
            "fps": self._v_out_fps(),
            "crf": self._v_out_crf(),
        }
        # 只有确实存在滚动图层时，滚动速度才可用（引擎侧会兜底，这里置 0 只是不再显示进度）
        if "scroll" not in video_mod._effective_motions(opts):
            opts["scroll_speed"] = 0
        self._v_refresh_motion_ui()
        return opts

    def _v_refresh_motion_ui(self):
        """按当前勾选，刷新「滚动速度」是否可用。"""
        kinds = [k for k, chk in (("text", self.v_text_chk), ("image", self.v_img_chk))
                 if chk.isChecked()]
        opts = {"kinds": kinds, "motion": self._v_motion(),
                "text": {"motion": _resolve_motion(self.v_text_motion_combo, self._v_motion())},
                "image": {"motion": _resolve_motion(self.v_img_motion_combo, self._v_motion())}}
        self.v_speed_spin.setEnabled("scroll" in video_mod._effective_motions(opts))

    # --------------------------------------------------- 视频预览 / 拖拽定位
    def _v_toggle_drag(self, state):
        """拖拽开关：开启后在预览帧上拖动即可定位，落点写回 X/Y（下拉切「自定义」）。

        注意：stateChanged 在 PySide6 里传的是枚举，而部分环境/老代码会按 int 比较，
        用 int(state) != 0 兼容两种形态，避免判断失效。
        """
        # PySide6 的 CheckState 是独立枚举（既不等于 int，也不能直接 int()），
        # PyQt6 / 老接口又可能传 int；统一取 .value（没有就当 int）再判非零。
        _s = getattr(state, "value", state)
        on = (int(_s) != 0)
        # 不再停用预设下拉：拖拽与下拉/X-Y 是同一份状态，互为入口更直观
        if on and self.v_custom_pos is None:
            p = self._v_position(self.v_text_pos_combo,
                                 self.v_text_x_spin, self.v_text_y_spin)
            self.v_custom_pos = (p[0], p[1])
        self._v_update_pos_label()
        if on:
            self._v_show_frame()      # 立即按自定义位置重绘

    def _v_update_pos_label(self):
        """位置摘要：优先显示「自定义」落点，否则显示当前预设对应的 X/Y。"""
        if getattr(self, "v_custom_pos", None):
            self.v_pos_lbl.setText(
                f"自定义位置：X {self.v_custom_pos[0]*100:.0f}%  "
                f"Y {self.v_custom_pos[1]*100:.0f}%")
            return
        tpos = self._v_position(self.v_text_pos_combo,
                                self.v_text_x_spin, self.v_text_y_spin)
        ipos = self._v_position(self.v_img_pos_combo,
                                self.v_img_x_spin, self.v_img_y_spin)
        self.v_pos_lbl.setText(
            f"文字 X {tpos[0]*100:.0f}% Y {tpos[1]*100:.0f}% ｜ "
            f"图片 X {ipos[0]*100:.0f}% Y {ipos[1]*100:.0f}%")

    def _v_preview_time(self):
        """预览取第几秒的画面（滚动水印可借此看不同时刻的位置）。"""
        sp = getattr(self, "v_t_spin", None)
        return float(sp.value()) if sp is not None else 1.0

    def _v_placeholder_frame(self):
        """未选视频时的 16:9 示意画面：深灰底 + 浅色网格，便于确认水印位置与大小。"""
        from PIL import Image as _PILImage, ImageDraw as _PILDraw
        w, h = 1280, 720
        img = _PILImage.new("RGB", (w, h), (48, 48, 52))
        d = _PILDraw.Draw(img)
        step = 80
        for x in range(0, w, step):
            d.line([(x, 0), (x, h)], fill=(70, 70, 76), width=1)
        for y in range(0, h, step):
            d.line([(0, y), (w, y)], fill=(70, 70, 76), width=1)
        d.rectangle([w // 2 - 90, h // 2 - 26, w // 2 + 90, h // 2 + 26],
                    fill=(90, 90, 96))
        return img

    def _v_grab_frame(self):
        """取一帧画面并按当前设置叠好水印显示。

        已选视频 → 抓该视频第 N 秒的真实帧；未选视频 → 用 16:9 示意画面，
        这样在挑视频之前也能先确认水印的位置/大小/透明度。
        """
        src = self.v_src_edit.text().strip()
        t = self._v_preview_time()
        if not src or not os.path.exists(src):
            # 无源视频：用示意画面，保证「先预览后选片」也走得通
            self._v_last_frame = self._v_placeholder_frame()
            self._v_show_frame()
            self.v_status_lbl.setText(
                "未选择视频，当前为 16:9 示意画面；选好视频后会自动换成真实帧。")
            return
        opts = self._v_gather_opts()
        try:
            raw = video_mod.grab_raw_frame(src, opts, t_sec=t)
        except Exception as e:
            self.v_frame_label.clearImage()
            self.v_frame_label.setText(f"抓取失败：{e}")
            return
        self._v_last_frame = raw
        self._v_show_frame()

    def _v_show_frame(self):
        """在已抓取的画面上，按当前设置（含自定义位置）重绘带水印的预览。"""
        if getattr(self, "_v_last_frame", None) is None:
            return
        opts = self._v_gather_opts()
        try:
            img = video_mod.compose_on_frame(self._v_last_frame, opts,
                                             t_sec=self._v_preview_time())
        except Exception as e:
            self.v_frame_label.setText(f"叠水印失败：{e}")
            return
        bio = io.BytesIO(); img.save(bio, "PNG"); bio.seek(0)
        qimg = QImage.fromData(bio.getvalue())
        self.v_frame_label.setBasePixmap(QPixmap.fromImage(qimg))
        if getattr(self, "v_custom_pos", None):
            self.v_frame_label.setMarker(*self.v_custom_pos)
        self._v_update_pos_label()

    def _v_schedule_frame(self):
        """参数改动后防抖重绘预览（已抓到画面时才重绘，不会自动去读视频）。"""
        if getattr(self, "_v_last_frame", None) is None:
            return
        timer = getattr(self, "_v_frame_timer", None)
        if timer is None:
            return
        timer.start()

    def _on_video_frame_drag(self, fx, fy):
        """视频预览帧拖拽：把落点写回 X/Y 并切到「自定义」，随后实时重绘。

        文字与图片统一采用该落点（与 Word 侧拖拽行为一致），用户若想错开，
        可在拖拽后单独改图片的 X/Y。
        """
        if not self.v_drag_chk.isChecked():
            return
        self.v_custom_pos = (fx, fy)
        self._v_pos_applying = True
        try:
            for sp in (self.v_text_x_spin, self.v_img_x_spin):
                sp.setValue(round(fx * 100))
            for sp in (self.v_text_y_spin, self.v_img_y_spin):
                sp.setValue(round(fy * 100))
            for cb in (self.v_text_pos_combo, self.v_img_pos_combo):
                cb.blockSignals(True)
                cb.setCurrentText("自定义")
                cb.blockSignals(False)
        finally:
            self._v_pos_applying = False
        self._v_update_pos_label()
        self._v_show_frame()

    def _v_play_preview(self):
        """导出前 3 秒带水印片段，并用系统默认播放器播放，用于确认动态效果。"""
        src = self.v_src_edit.text().strip()
        if not src or not os.path.exists(src):
            QMessageBox.warning(self, "提示", "请先选择有效的视频文件。")
            return
        opts = self._v_gather_opts()
        kinds = list(opts["kinds"])
        if not kinds:
            QMessageBox.warning(self, "提示", "请至少启用一种水印（文字或图片）。")
            return
        if "image" in kinds and not self.v_img_edit.text().strip():
            QMessageBox.warning(self, "提示", "已启用图片水印，但还未选择水印图片。")
            return
        if "text" in kinds and not self.v_text_edit.text().strip():
            QMessageBox.warning(self, "提示", "已启用文字水印，但水印文字为空。")
            return
        stem, _ = os.path.splitext(os.path.basename(src))
        out = os.path.join(tempfile.gettempdir(), f"WriteMark_preview_{stem}.mp4")
        self._v_preview_out = out
        self.v_status_lbl.setText("正在生成预览短片（前 3 秒），请稍候…")
        self.v_play_btn.setEnabled(False)
        self.v_grab_btn.setEnabled(False)
        w = VideoWorker(src, out, opts, max_seconds=3)
        self._v_preview_worker = w
        w.result_signal.connect(self._v_preview_done)
        w.progress_signal.connect(self._v_on_progress)
        w.start()

    def _v_preview_done(self, ok, msg):
        self.v_play_btn.setEnabled(True)
        self.v_grab_btn.setEnabled(True)
        if ok:
            out = getattr(self, "_v_preview_out", "")
            if out and os.path.exists(out):
                try:
                    os.startfile(out)      # Windows：用系统默认播放器打开
                    self.v_status_lbl.setText("已生成预览短片，正在用系统播放器打开…")
                except Exception as e:
                    self.v_status_lbl.setText(f"预览短片已生成，但无法自动打开：{e}")
            else:
                self.v_status_lbl.setText("预览短片生成成功，但找不到输出文件。")
        else:
            self.v_status_lbl.setText("预览生成失败：" + msg)

    def _v_run(self):
        # 冻结版没有控制台，槽函数里任何异常都会被 Qt 静默吞掉、表现为“点了没反应”，
        # 因此这里统一兜底：出错也要把原因显示给用户，绝不留一个死按钮。
        try:
            self._v_run_inner()
        except Exception as e:
            import traceback
            self.v_run_btn.setEnabled(True)
            self.v_cancel_btn.setEnabled(False)
            self.v_status_lbl.setText("失败：" + str(e))
            QMessageBox.critical(self, "失败", f"视频水印处理失败：\n{e}")
            try:
                with open(os.path.join(tempfile.gettempdir(),
                                       "WriteMark_video_err.log"), "w",
                          encoding="utf-8") as f:
                    f.write(traceback.format_exc())
            except Exception:
                pass

    def _v_run_inner(self):
        src = self.v_src_edit.text().strip()
        if not src or not os.path.exists(src):
            QMessageBox.warning(self, "提示", "请先选择有效的视频文件。")
            return
        # 注意：方法名带下划线前缀 _v_gather_opts
        opts = self._v_gather_opts()
        kinds = list(opts["kinds"])
        if not kinds:
            QMessageBox.warning(self, "提示", "请至少启用一种水印（文字或图片）。")
            return
        if "image" in kinds and not self.v_img_edit.text().strip():
            QMessageBox.warning(self, "提示", "已启用图片水印，但还未选择水印图片。")
            return
        if "text" in kinds and not self.v_text_edit.text().strip():
            QMessageBox.warning(self, "提示", "已启用文字水印，但水印文字为空。")
            return
        out = self.v_out_edit.text().strip()
        if not out:
            base, _ = os.path.splitext(src)
            out = base + "WaterMark.mp4"
            self.v_out_edit.setText(out)

        motions = video_mod._effective_motions(opts)
        desc = "+".join({"fixed": "固定", "scroll": "滚动"}.get(m, m) for m in motions)
        kinds_desc = "文字" if kinds == ["text"] else ("图片" if kinds == ["image"] else "文字+图片")
        # 这里解包失败过一次（out_size 曾被解析成 3 元组）——即便异常也不能中断任务，
        # 规格只影响提示文案，取不到就退化成不带数字的说明。
        try:
            ow, oh = opts["out_size"]
            spec = f"{ow}×{oh} · {int(opts['fps'])}fps · crf{int(opts['crf'])}"
        except Exception:
            spec = "自定义"
        self.v_status_lbl.setText(
            f"正在逐帧处理（{kinds_desc} · {desc} · 导出 {spec}），"
            f"请稍候（长视频较慢属正常）…")

        self.v_run_btn.setEnabled(False)
        self.v_cancel_btn.setEnabled(True)
        self.v_progress.setValue(0)
        w = VideoWorker(src, out, opts)
        self._v_worker = w
        w.result_signal.connect(self._v_on_result)
        w.progress_signal.connect(self._v_on_progress)
        w.finished.connect(self._v_on_finished)
        w.start()

    def _v_on_progress(self, cur, total):
        if total:
            pct = int(cur / total * 100)
            self.v_progress.setValue(pct)
            self.v_status_lbl.setText(f"处理中：{cur}/{total} 帧（{pct}%）")

    def _v_on_result(self, ok, msg):
        if ok:
            self.v_status_lbl.setText("任务已完成：" + msg)
            QMessageBox.information(self, "任务已完成", "任务已完成")
        else:
            self.v_status_lbl.setText("失败：" + msg)
            QMessageBox.critical(self, "失败", msg)

    def _v_on_finished(self):
        self._v_worker = None
        self.v_run_btn.setEnabled(True)
        self.v_cancel_btn.setEnabled(False)

    def _v_cancel(self):
        if getattr(self, "_v_worker", None) and self._v_worker.isRunning():
            self._v_worker.request_stop()
            self.v_status_lbl.setText("已请求取消，正在收尾…")

    @staticmethod
    def _fmt_hint(text):
        """页面格式说明：仿宋 + 红色，让用户一眼看清支持的格式。"""
        lbl = QLabel(text)
        lbl.setStyleSheet("color:#e10600; font-family:'FangSong','仿宋'; font-size:11px;")
        lbl.setWordWrap(True)
        return lbl

    def _make_group(self, title):
        """生成一个带标题的边框分组容器，返回该 QWidget（其 layout 已建好、垂直）。"""
        g = QWidget()
        v = QVBoxLayout(g)
        v.setContentsMargins(8, 8, 8, 8)
        g.setStyleSheet("QWidget{border:1px solid #cccccc; border-radius:4px;}")
        title_lbl = QLabel(title)
        title_lbl.setStyleSheet("font-weight:bold; color:#333;")
        v.addWidget(title_lbl)
        return g

    def _schedule_preview(self, *args):
        self._preview_timer.start()

    # ------------------------------------------------ Word 秒退
    def _office_start(self, worker):
        self._office_worker = worker          # 持有引用，防止线程被 GC 后 abort
        worker.result_signal.connect(self._office_done)
        for b in (getattr(self, "_btn_fix", None), getattr(self, "_btn_revert", None)):
            if b:
                b.setEnabled(False)
        self.status_label.setText("处理中…")
        worker.start()

    def _office_done(self, ok, msg, need_admin=False):
        for b in (getattr(self, "_btn_fix", None), getattr(self, "_btn_revert", None)):
            if b:
                b.setEnabled(True)
        self.status_label.setText("就绪")
        self._exit_detail.append(msg)
        self._office_refresh_info()
        if need_admin and not otw.is_admin():
            ask = QMessageBox.question(
                self, "需要管理员权限",
                "有「组策略」级别的开关被系统 ACL 拒绝写入（用户级设置已生效，不影响修复效果）。\n\n"
                "要不要现在触发一次 UAC 提权，把其余项也写进去？\n"
                "（会弹出 Windows 的「是否允许此应用对你的设备进行更改」窗口，请点「是」）",
                QMessageBox.Yes | QMessageBox.No, QMessageBox.Yes)
            if ask == QMessageBox.Yes:
                self._office_elevate()

    def _office_elevate(self):
        """拉起一个提权的自己去执行修复，然后轮询它的结果文件。"""
        launched, code = otw.elevate("office-fix")
        if not launched:
            QMessageBox.warning(self, "提权失败",
                                "未能启动提权进程（UAC 被拒绝或环境不允许）。\n"
                                "可改用：右键本程序 →「以管理员身份运行」，再点一次「一键修复」。")
            return
        self.status_label.setText("等待管理员操作…")

        def poll():
            res = otw.read_result()
            if res is None:                     # 提权进程还在跑
                return True                     # 继续等
            ok, msg = res
            self._office_done(ok, "【管理员权限执行】\n" + msg, False)
            return False                        # 停掉定时器
        t = QTimer(self)
        t.setInterval(400)
        t.timeout.connect(poll)
        t.start(400)
        self._elevate_timer = t

    def _office_refresh_info(self):
        if otw is None:
            self._exit_info.setText("诊断模块不可用")
            return
        try:
            d = otw.diagnose()
        except Exception as e:
            self._exit_info.setText("诊断失败：%s" % e)
            return
        bits = ["Office %s" % d["version"],
                "已登录账户" if d["logged_in"] else "未登录账户",
                "免管理员项 %d/%d" % (d["applied_user"], d["total_user"])]
        if not d["policy_writable"]:
            bits.append("组策略被锁（需管理员）")
        bits.append("管理员权限：是" if d["admin"] else "管理员权限：否")
        self._exit_info.setText("  ·  ".join(bits))

    def _office_diagnose(self):
        if otw is None:
            return
        def job():
            d = otw.diagnose()
            L = ["Office 版本分支：%s" % d["version"],
                 "登录状态：%s" % ("已登录微软账户 %s…（已排除：实测改它无效，"
                                  "18s 依旧，真凶是加载项）" % d["account"]
                                  if d["logged_in"] else "未发现登录标记"),
                 "管理员权限：%s" % ("有" if d["admin"] else "无"),
                 "组策略分支：%s" % ("可写" if d["policy_writable"] else "被系统 ACL 锁住，写入需管理员"),
                 "",
                 "【加载项】实测卡顿的真正元凶就在这里"]
            for name, lb, desc, auto in d["addins"]:
                L.append("  %-38s LoadBehavior=%-2s %s —— %s"
                         % (name, lb, "随 Word 启动自动加载" if auto else "不自动加载", desc))
            L.append("")
            L.append("【用户级开关】免管理员，修复主力")
            for it in d["items"]:
                if it["realm"] != "user":
                    continue
                cur = "未设置" if it["current"] == otw.MISSING else repr(it["current"])
                L.append("  %-38s %-22s → %-3r  %s" % (it["name"], cur, it["want"], it["desc"]))
            L.append("")
            L.append("【可选加码】默认不动，需手动勾选才会写")
            for path, name, want, desc in otw.optional_items(d["version"]):
                cur = "未设置" if otw.read_val(path, name) == otw.MISSING else repr(otw.read_val(path, name))
                L.append("  %-38s %-22s %s" % (name, cur, desc))
            L.append("")
            L.append("【组策略开关】可选加码")
            for it in d["items"]:
                if it["realm"] != "policy":
                    continue
                cur = "未设置" if it["current"] == otw.MISSING else repr(it["current"])
                L.append("  %-38s %-22s → %-3r" % (it["name"], cur, it["want"]))
            return {"ok": True, "msg": "\n".join(L), "need_admin": not d["policy_writable"]}
        self._office_start(OfficeWorker(job))

    def _office_fix(self):
        if otw is None:
            return
        warn = QMessageBox.question(
            self, "一键修复",
            "卡顿的真正原因是 Word 的 COM 加载项（微软 OfficePLUS、百度网盘插件等）：\n"
            "它们随 Word 启动，关闭 Word 时要做云端收尾，联网就等超时、断网就秒退。\n\n"
            "将要写入（会自动备份，可随时用「还原」退回）：\n\n"
            "• 把上述加载项设为「不自动加载」\n"
            "• 停止诊断遥测\n\n"
            "不会改动你的登录状态，已登录的 Word 账号仍然保持登录。\n"
            "代价：需要时得手动启用这些插件。\n\n"
            "继续吗？",
            QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
        if warn != QMessageBox.Yes:
            return
        self._office_start(OfficeWorker(lambda: otw.apply_fix()))

    def _office_optional(self):
        """可选加码：会踢账号，必须显式二次确认。"""
        if otw is None:
            return
        warn = QMessageBox.warning(
            self, "可选加码",
            "⚠ 这一档会改变 Office 的账户与联网行为，请先看清楚：\n\n"
            "• SignInOptions：禁止 Office 登录任何账户\n"
            "  → 会把「已登录的 Word 账号踢下线」，用 OneDrive 云保存每次都要重新登录\n"
            "• DisconnectedState 等：关闭连接体验、在线内容下载\n"
            "  → 在线模板 / 智能查找 / 翻译 / 在线字体不可用\n\n"
            "本地编辑、打开、打印、另存为本地文件都不受影响。\n"
            "需要联网功能的话，之后随时可以用「还原」退回，或直接撤销本操作。\n\n"
            "确定要应用吗？",
            QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
        if warn != QMessageBox.Yes:
            return
        self._office_start(OfficeWorker(lambda: otw.apply_optional()))

    def _office_revert(self):
        if otw is None:
            return
        self._office_start(OfficeWorker(lambda: otw.revert_fix()))

    def _office_bench(self):
        if otw is None:
            return
        reply = QMessageBox.question(
            self, "测速",
            "将自动打开一次 Word，然后模拟点击右上角 ×，\n"
            "并计时「进程真正消失」的秒数。约需 30 秒。\n"
            "测量期间请勿操作，并确保当前没有其他 Word 在跑。继续吗？",
            QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
        if reply != QMessageBox.Yes:
            return
        self._office_start(OfficeWorker(lambda: otw.benchmark()))

    def _on_harden_changed(self, *args):
        """加固参数变化：同步到实例属性（供 _gather_opts 使用）并刷新预览。"""
        self.tile = self.tile_chk.isChecked()
        self.tile_rows = int(self.tile_rows_spin.value())
        self.tile_cols = int(self.tile_cols_spin.value())
        self.redundant = self.redundant_chk.isChecked()
        self._schedule_preview()

    def _make_slider_spin(self, label, lo, hi, val, decimals, step, suffix, on_change):
        """构造【滑块 + 数字输入框（精确到小数点后 decimals 位）】联动行。

        滑块拖动与数字框输入双向同步，均不会无限递归；二者精度一致。
        """
        row = QHBoxLayout()
        row.addWidget(QLabel(label))
        factor = 10 ** decimals
        slider = QSlider(Qt.Horizontal)
        slider.setRange(int(lo * factor), int(hi * factor))
        slider.setValue(int(round(val * factor)))
        spin = QDoubleSpinBox()
        spin.setRange(lo, hi)
        spin.setDecimals(decimals)
        spin.setSingleStep(step)
        spin.setValue(val)
        if suffix:
            spin.setSuffix(suffix)

        def on_slider(v):
            value = v / factor
            spin.blockSignals(True)
            spin.setValue(value)
            spin.blockSignals(False)
            on_change(value)

        def on_spin(v):
            slider.blockSignals(True)
            slider.setValue(int(round(v * factor)))
            slider.blockSignals(False)
            on_change(v)

        slider.valueChanged.connect(on_slider)
        slider.valueChanged.connect(self._schedule_preview)
        spin.valueChanged.connect(on_spin)
        spin.valueChanged.connect(self._schedule_preview)
        row.addWidget(slider, 3)
        row.addWidget(spin, 1)
        return row, slider, spin

    def _make_font_combo(self, default_text):
        """构造一个【可编辑 + 前缀自动补全】的字体下拉框（如 Word 的字体选择）。

        - 允许直接键入字体名（如输入 t 即弹出 Times New Roman、Tahoma…）；
        - 不把键入内容当作新选项插入列表；
        - 每个下拉框自带 QStringListModel + QCompleter，便于后续扩充字体时同步。
        返回 (combo, model)，model 供 _apply_word_fonts 扩充时更新。
        """
        combo = QComboBox()
        combo.setEditable(True)
        combo.setInsertPolicy(QComboBox.NoInsert)
        combo.setDuplicatesEnabled(False)
        base_fonts = ["微软雅黑", "宋体", "黑体", "楷体", "仿宋",
                      "Times New Roman", "Arial", "Calibri", "Courier New"]
        combo.addItems(base_fonts)
        combo.setCurrentText(default_text)
        model = QStringListModel(base_fonts)
        completer = QCompleter(model, self)
        completer.setCaseSensitivity(Qt.CaseInsensitive)   # 不区分大小写
        completer.setCompletionMode(QCompleter.PopupCompletion)
        completer.setFilterMode(Qt.MatchStartsWith)        # 前缀匹配
        combo.setCompleter(completer)
        return combo, model

    # ------------------------------------------------------------ callbacks
    def _on_text_angle(self, v):
        self.text_angle = float(v)
    def _on_text_trans(self, v):
        self.text_transparency = v / 100.0
    def _on_text_scale(self, v):
        self.text_scale = float(v)
    def _on_text_offx(self, v):
        self.text_offset_x = float(v)
        self._mark_pos_custom(self.text_pos_combo)
    def _on_text_offy(self, v):
        self.text_offset_y = float(v)
        self._mark_pos_custom(self.text_pos_combo)

    def _on_img_angle(self, v):
        self.img_angle = float(v)
    def _on_img_trans(self, v):
        self.img_transparency = v / 100.0
    def _on_img_scale(self, v):
        self.img_scale = float(v)
    def _on_img_offx(self, v):
        self.img_offset_x = float(v)
        self._mark_pos_custom(self.img_pos_combo)
    def _on_img_offy(self, v):
        self.img_offset_y = float(v)
        self._mark_pos_custom(self.img_pos_combo)

    # ------------------------------------------------ 位置预设 / 自定义位置
    def _mark_pos_custom(self, combo):
        """手动改了偏移（或拖了预览图）→ 位置下拉切到「自定义」，避免预设值误导。

        预设写入偏移期间由 _pos_applying 屏蔽，否则会立刻把自己打回「自定义」。
        """
        if getattr(self, "_pos_applying", False):
            return
        if combo is None or combo.currentText() == "自定义":
            return
        combo.blockSignals(True)
        combo.setCurrentText("自定义")
        combo.blockSignals(False)

    def _apply_pos_preset(self, preset, x_spin, y_spin):
        """九宫格预设 → 写入水平/垂直偏移（0=页面中心）。"""
        off = _W_POS_PRESETS.get(preset)
        if off is None:      # 「自定义」：不做任何覆盖，沿用当前偏移
            return
        self._pos_applying = True
        try:
            x_spin.setValue(off[0])
            y_spin.setValue(off[1])
        finally:
            self._pos_applying = False
        self._schedule_preview()

    def _on_text_pos_preset(self, preset):
        self._apply_pos_preset(preset, self.text_offx_spin, self.text_offy_spin)

    def _on_img_pos_preset(self, preset):
        self._apply_pos_preset(preset, self.img_offx_spin, self.img_offy_spin)

    def _on_text_enabled(self, state):
        self.text_enabled = self.text_chk.isChecked()
        for w in self.text_ctrl_widgets:
            w.setEnabled(self.text_enabled)
        self._schedule_preview()

    def _on_img_enabled(self, state):
        self.image_enabled = self.img_chk.isChecked()
        for w in self.img_ctrl_widgets:
            w.setEnabled(self.image_enabled)
        self._schedule_preview()

    def _browse_file(self):
        p, _ = QFileDialog.getOpenFileName(self, "选择 Word 文件", "",
                                           "Word 文件 (*.docx *.doc);;All (*.*)")
        if p:
            self.file_path = p; self.file_edit.setText(p)
            # 选完源文件后，自动给出默认输出路径（桌面/<原名>WaterMark），用户可改
            self.out_edit.setText(core.default_output_path(p))
            # 导入 Word 后，征询用户同意后读取其可用字体，扩充文本水印字体选择
            self._maybe_load_word_fonts()

    def _maybe_load_word_fonts(self):
        """导入 Word 后：先征询权限，同意后后台读取 Word 里的字体并扩充下拉框。

        每会话仅询问一次；用户拒绝则沿用内置字体，不再打扰。

        实测经验：Word 给出的字体名是**本机化**的（中文系统里是"宋体/黑体/仿宋/微软雅黑"
        这类中文名，约 1400 条），注册表只有英文名且中文名的仅十几条——所以这里仍走
        Word，但读取完会确保那个 Word 进程被彻底关掉（com_cleanup），不会像早先版本
        那样在后台留下一个看不见的 Word，害得用户后来关它时以为"退出 Word 卡顿"。
        """
        if self._fonts_loaded:
            return
        if self._fonts_asked:
            return
        self._fonts_asked = True
        ans = QMessageBox.question(
            self, "读取可用字体",
            "是否允许本工具读取您 Word 中可用的字体，以扩充“文本水印”的中文字体 / 西文字体选择？\n\n"
            "说明：仅读取字体名称列表，不会访问或修改您任何文档的内容；\n"
            "读取时若需要会临时启动一个不可见的 Word，读完立即关闭"
            "（若本机没有 Word，则改用系统已安装字体）。",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.No,
        )
        if ans != QMessageBox.StandardButton.Yes:
            self._append_log("已取消读取字体，文本水印沿用内置字体。")
            return
        self._append_log("正在读取可用字体（首次可能需数秒，请在日志查看结果）…")
        self._set_busy(True)
        w = FontWorker()
        self._font_worker = w  # 保住引用
        w.finished_signal.connect(self._on_fonts_loaded)
        w.error_signal.connect(self._on_fonts_error)
        w.start()

    def _on_fonts_loaded(self, fonts, source):
        self._set_busy(False)
        added = self._apply_word_fonts(fonts)
        if added:
            self._fonts_loaded = True
            self._append_log(f"已从【{source}】读取并加入 {added} 种字体，"
                             f"可在“字体”下拉框中选择更多文本水印字体。")
        else:
            self._append_log(f"从【{source}】未获取到新的字体（可能与内置重复或读取为空）。")
        self._schedule_preview()

    def _on_fonts_error(self, msg):
        self._set_busy(False)
        self._append_log(f"读取字体失败：{msg}（仍可使用内置字体）")

    def _apply_word_fonts(self, fonts):
        """把字体名合并进“中文字体 / 西文字体”两个下拉框，去重、保持各自当前选择，
        返回新增数量（两框新增集合相同，返回该集合的大小）。

        纯逻辑、无弹窗，便于直接单元测试。
        """
        if not fonts:
            return 0
        combos = [self.cn_font_combo, self.latin_font_combo]
        # 以两框已有项的并集为基准去重，得到真正需要新增的字体集合
        existing = set()
        for combo in combos:
            existing |= {combo.itemText(i) for i in range(combo.count())}
        new = []
        for f in fonts:
            if f and f not in existing and f not in new:
                new.append(f)
        if not new:
            return 0
        new.sort()
        for combo in combos:
            current = combo.currentText()
            combo.blockSignals(True)
            for f in new:
                combo.addItem(f)
            if current in (combo.itemText(i) for i in range(combo.count())):
                combo.setCurrentText(current)  # 还原当前选择（默认微软雅黑等）
            combo.blockSignals(False)
            # 同步补全模型：键入前缀过滤时也要能匹配到新加入的字体
            all_items = [combo.itemText(i) for i in range(combo.count())]
            completer = combo.completer()
            if completer is not None:
                model = completer.model()
                if isinstance(model, QStringListModel):
                    model.setStringList(all_items)
        return len(new)

    def _browse_output(self):
        d = self.out_edit.text() or (os.path.expanduser("~/Desktop") if os.path.isdir(
            os.path.expanduser("~/Desktop")) else "")
        p, _ = QFileDialog.getSaveFileName(self, "选择输出文件（保留原文件）", d,
                                           "Word 文件 (*.docx *.doc);;All (*.*)")
        if p:
            self.out_edit.setText(p)

    def _browse_image(self):
        p, _ = QFileDialog.getOpenFileName(self, "选择水印图片", "",
                                           "图片 (*.png *.jpg *.jpeg *.webp *.pdf);;All (*.*)")
        if p:
            self.image_path = p; self.img_edit.setText(p)
            self._schedule_preview()

    def _pick_color(self):
        c = QColorDialog.getColor(QColor(*self.color), self, "选择文本水印颜色")
        if c.isValid():
            self.color = (c.red(), c.green(), c.blue())
            self.color_btn.setText(f"文本颜色: RGB{self.color}")
            self._schedule_preview()

    # -------------------------------------------------------------- actions
    def _gather_opts(self):
        """收集参数。文本与图像各自的参数相互独立，分装到 "text"/"image" 子字典。"""
        return {
            "text": {
                "text": self.text_edit.text(),
                "cn_font_name": self.cn_font_combo.currentText(),
                "latin_font_name": self.latin_font_combo.currentText(),
                "font_size": self.font_spin.value(),
                "color": self.color,
                "angle": self.text_angle,
                "transparency": self.text_transparency,
                "scale": self.text_scale,
                "offset_x": self.text_offset_x,
                "offset_y": self.text_offset_y,
            },
            "image": {
                "image_path": self.img_edit.text().strip(),
                "angle": self.img_angle,
                "transparency": self.img_transparency,
                "scale": self.img_scale,
                "offset_x": self.img_offset_x,
                "offset_y": self.img_offset_y,
            },
            "text_enabled": self.text_enabled,
            "image_enabled": self.image_enabled,
            # 防去除加固（对两类水印同时生效，故放在顶层）
            "tile": self.tile_chk.isChecked() if hasattr(self, "tile_chk") else self.tile,
            "tile_rows": int(self.tile_rows_spin.value()) if hasattr(self, "tile_rows_spin") else self.tile_rows,
            "tile_cols": int(self.tile_cols_spin.value()) if hasattr(self, "tile_cols_spin") else self.tile_cols,
            "redundant": self.redundant_chk.isChecked() if hasattr(self, "redundant_chk") else self.redundant,
        }

    def _gather_kinds(self):
        kinds = []
        if self.text_enabled:
            kinds.append("text")
        if self.image_enabled:
            kinds.append("image")
        return kinds

    def _run(self, fn):
        if self._worker is not None and self._worker.isRunning():
            QMessageBox.information(self, "提示", "上一步操作仍在进行，请稍候再试。")
            return
        self._append_log("正在处理...")
        self._set_busy(True)
        w = Worker(fn)
        self._worker = w  # 关键：必须保住引用，否则线程运行中对象被 GC → 进程崩溃退出
        w.log_signal.connect(self._append_log)
        w.result_signal.connect(self._on_result)
        w.finished.connect(self._on_worker_finished)
        w.start()

    def _on_worker_finished(self):
        self._worker = None
        self._set_busy(False)

    def _set_busy(self, busy):
        for b in (getattr(self, "_btn_insert", None), getattr(self, "_btn_clear", None)):
            if b is not None:
                b.setEnabled(not busy)
        if busy:
            self.status_label.setText("处理中...")

    # --------------------------------------------------------- 系统托盘常驻
    def _setup_tray(self):
        """建立系统托盘图标——这样关掉主窗口后守护仍能在后台运行。

        托盘创建失败（部分精简版系统/资源管理器尚未就绪）时不能静默吞掉，
        否则用户点关闭会直接走“退出进程”分支——这正是此前“开了后台保护、
        点关闭却整个进程没了”的原因。这里记录原因并自动重试一次。
        """
        try:
            if not QSystemTrayIcon.isSystemTrayAvailable():
                raise RuntimeError("系统托盘不可用（QSystemTrayIcon.isSystemTrayAvailable() = False）")
            tray = QSystemTrayIcon(self)
            tray.setIcon(self.style().standardIcon(QStyle.SP_ComputerIcon))
            tray.setToolTip("Word 一键水印工具（守护运行中）")
            menu = QMenu()
            act_show = menu.addAction("显示主窗口")
            act_show.triggered.connect(self._show_from_tray)
            self._act_stop = menu.addAction("停止守护")
            self._act_stop.triggered.connect(self._stop_watch_from_tray)
            menu.addSeparator()
            act_quit = menu.addAction("退出程序")
            act_quit.triggered.connect(self._quit_app)
            tray.setContextMenu(menu)
            tray.activated.connect(self._on_tray_activated)
            tray.show()
            self.tray = tray
        except Exception as e:
            self.tray = None
            self._tray_reason = str(e)
            self._tray_tried += 1
            self._append_log(f"· 系统托盘初始化失败：{e}")
            if self._tray_tried <= 1:
                # 资源管理器（explorer）刚启动时托盘区可能尚未就绪，3 秒后重试一次
                QTimer.singleShot(3000, self._setup_tray)

    def _on_tray_activated(self, reason):
        if reason == QSystemTrayIcon.Trigger:   # 左键单击：切回窗口
            self._show_from_tray()

    def _show_from_tray(self):
        self.show()
        self.raise_()
        self.activateWindow()

    def _stop_watch_from_tray(self):
        if hasattr(self, "watch_chk"):
            self.watch_chk.setChecked(False)     # 触发 _toggle_watch 停止守护
        self._append_log("已从托盘停止守护。")

    def _quit_app(self):
        """托盘“退出程序”：真正结束进程（含守护与后台线程清理）。"""
        self._quitting = True
        if self.tray is not None:
            try:
                self.tray.hide()
            except Exception:
                pass
        if self.wd:
            self.wd.stop_nowait()
            self.wd = None
        self.close()
        QApplication.quit()

    def _go_background(self):
        """隐藏窗口、进程留在后台（守护若在运行则继续补回水印）。"""
        self.hide()
        if self.wd is not None:
            tip = ("守护已在后台继续运行，水印被删会自动补回。\n"
                   "单击/双击托盘图标可重新打开窗口，右键托盘可停止守护 / 退出程序。")
        else:
            tip = ("程序已在后台运行（未开启后台保护）。\n"
                   "单击/双击托盘图标可重新打开窗口，右键托盘可退出程序。")
        shown = False
        try:
            if self.tray is not None:
                self.tray.showMessage("Word 一键水印工具", tip,
                                      QSystemTrayIcon.Information, 4000)
                shown = True
        except Exception:
            pass
        if shown:
            self._append_log("已最小化到系统托盘，程序继续在后台运行。")
        else:
            self._append_log("窗口已隐藏，进程仍在后台运行"
                             "（本机托盘不可用，重新打开界面需再次启动程序）。")

    def closeEvent(self, event):
        """关闭窗口 = 关闭页面，不退出后台（不弹任何询问）。

        - 点标题栏关闭按钮：窗口隐藏，进程保留，托盘图标可唤回；
          守护（若开启）继续运行，水印被删会自动补回。
        - 真正退出走托盘右键「退出程序」（_quitting=True 路径）。
        - 自动化脚本可用环境变量 WM_CLOSE_CHOICE=quit|cancel 强制改变行为
          （默认 minimize-to-background，无需人工干预）。
        - 真正退出时：守护线程已是 daemon，只发停止信号、绝不阻塞等待。
        - 兜底：启动一个计时线程，超时（6s）后若进程仍未自行退出，强制 os._exit(0)。
        - 工作/字体线程若超时仍未结束则强制 terminate。
        """
        if not self._quitting:
            forced = os.environ.get("WM_CLOSE_CHOICE", "").strip().lower()
            if forced == "cancel":
                event.ignore()
                return
            if forced != "quit":
                # 默认：只关闭页面，程序留在后台
                event.ignore()
                self._go_background()
                return

        # 兜底：无论后续清理是否卡住，超时(6s)后强制结束进程。
        # 6s 是刻意留出的余量：下面等待后台线程的完整路径最长约 3s，
        # 必须让“正常清理”跑在兜底之前完成，否则明明能优雅退出却被硬杀
        # （此前的 1.2s 就会误杀正在等待线程退出的正常关闭流程）。
        # 注意：pytest 下这个兜底不能真动手——6 秒后它会把「整个测试进程」
        # 静默 os._exit(0)，表现为回归跑到一半无报错中断（v1.3.9 踩到过）。
        # 线程照常启动（行为一致、可测），只是回调在测试环境里空转。
        def _force_exit():
            if "PYTEST_CURRENT_TEST" in os.environ:
                return
            os._exit(0)

        killer = threading.Timer(6.0, _force_exit)
        killer.daemon = True
        killer.start()
        if self.tray is not None:
            try:
                self.tray.hide()
            except Exception:
                pass

        if self.wd:
            self.wd.stop_nowait()  # 仅置停止信号，不 join 阻塞
            self.wd = None
        # 字体读取线程：一次性、可安全终止
        fw = getattr(self, "_font_worker", None)
        if fw is not None and fw.isRunning():
            fw.quit()
            if not fw.wait(1000):
                fw.terminate()
                fw.wait(300)
        # 工作线程：最多等 1.2 秒，超时强杀（总清理时间必须短于上面 6s 的兜底）
        w = self._worker
        if w is not None and w.isRunning():
            w.requestInterruption()
            w.quit()
            if not w.wait(1200):
                w.terminate()
                w.wait(300)
        event.accept()
        # 已显式关闭 setQuitOnLastWindowClosed，这里必须把事件循环也退出，
        # 否则窗口关了但 app.exec() 不返回，会留下一个没有界面的僵尸进程。
        QApplication.quit()

    def _on_result(self, ok, msg):
        if not ok:
            QMessageBox.critical(self, "失败", msg)
            return
        self.status_label.setText("完成")
        if self._last_action == "insert":
            self._append_log("已生成带水印的新文件（原文件未改动）。请打开下方日志中 output 路径的文件查看；"
                             "在 Word/WPS 中需切到『页面视图/页面布局』才能看到水印。")
            self._last_action = ""
        # 一键插入 / 一键清除完成后弹提示，用户点“确定”或直接关闭弹窗均可
        QMessageBox.information(self, "任务已完成", "任务已完成")

    # ------------------------------------------------------------- 预览
    def _render_preview(self):
        opts = self._gather_opts()
        kinds = self._gather_kinds()
        try:
            img = preview.render_preview(opts, kinds)
        except Exception as e:
            # 实时预览过程中参数可能暂时不完整（如未选图片），直接在预览区提示，不弹窗打断
            self._preview_img = None
            self.preview_label.clearImage()
            self.preview_label.setText(f"预览不可用：{e}")
            return
        self._preview_img = img
        bio = __import__("io").BytesIO(); img.save(bio, "PNG"); bio.seek(0)
        qimg = QImage.fromData(bio.getvalue())
        pix = QPixmap.fromImage(qimg)
        self.preview_label.setBasePixmap(pix)

    def _on_preview_drag(self, fx, fy):
        """Word 预览拖拽：把水印中心映射到水平/垂直偏移（-50%~50%，0=居中）。"""
        if not self.preview_drag_chk.isChecked():
            return
        off_x = max(-50.0, min(50.0, (fx - 0.5) * 100.0))
        off_y = max(-50.0, min(50.0, (fy - 0.5) * 100.0))
        # 文本与图像水印同步到同一落点（预览里两者常见叠在一起）
        self.text_offx_spin.setValue(off_x)
        self.text_offy_spin.setValue(off_y)
        self.img_offx_spin.setValue(off_x)
        self.img_offy_spin.setValue(off_y)
        # 滑块的 valueChanged 已触发 _schedule_preview，这里无需再手动刷新

    def _save_preview(self):
        if self._preview_img is None:
            self._render_preview()
        if self._preview_img is None:
            return
        p, _ = QFileDialog.getSaveFileName(self, "保存预览图", "watermark_preview.png",
                                           "PNG 图片 (*.png)")
        if p:
            self._preview_img.save(p, "PNG")
            self._append_log(f"预览图已保存：{p}")

    def _insert(self):
        if not self.file_path or not os.path.exists(self.file_path):
            QMessageBox.warning(self, "提示", "请先选择有效的 Word 文件。")
            return
        kinds = self._gather_kinds()
        if not kinds:
            QMessageBox.warning(self, "提示", "请至少启用一种水印（勾选“启用文本水印”或“启用图像水印”）。")
            return
        if "image" in kinds and not self.image_path:
            QMessageBox.warning(self, "提示", "已启用图像水印，但还未选择水印图片。")
            return
        out = self.out_edit.text().strip()
        if not out:
            out = core.default_output_path(self.file_path)
            self.out_edit.setText(out)
        self._last_action = "insert"
        path, opts = self.file_path, self._gather_opts()
        self._last_inserted = None

        def job():
            res = core.insert_watermark(path, kinds, output_path=out, **opts)
            # 记住实际写入的份数：开启守护后按这个数量校验是否被删过
            if isinstance(res, dict):
                self._last_inserted = res.get("inserted")
            return res

        self._run(job)

    def _clear(self):
        if not self.file_path or not os.path.exists(self.file_path):
            QMessageBox.warning(self, "提示", "请先选择有效的 Word 文件。")
            return
        out = self.out_edit.text().strip()
        if not out:
            out = core.default_output_path(self.file_path)
            self.out_edit.setText(out)
        kinds = None
        # 仅对 .docx 检测：若文档同时含有“文字水印”和“图片水印”，弹窗让用户二选一/全选
        if not core._use_com(self.file_path):
            try:
                types = engine_docx.detect_watermark_types(self.file_path)
                if types == {"text", "image"}:
                    choice = self._ask_clear_choice()
                    if choice is None:        # 用户点了“取消”
                        return
                    kinds = {"text": ["text"], "image": ["image"],
                             "both": ["text", "image"]}[choice]
            except Exception:
                kinds = None
        self._run(lambda: core.clear_watermark(self.file_path, output_path=out, kinds=kinds))

    def _ask_clear_choice(self):
        """文档同时含文字/图片水印时弹出选择框，返回 'text' / 'image' / 'both' / None(取消)。"""
        box = QMessageBox(self)
        box.setWindowTitle("想要去除水印？")
        box.setIcon(QMessageBox.Question)
        box.setText("该文档同时含有文字水印与图片水印，请选择要去除哪一种：")
        btn_text = box.addButton("去除文字水印", QMessageBox.ActionRole)
        btn_img = box.addButton("去除图片水印", QMessageBox.ActionRole)
        btn_both = box.addButton("文字和图片水印都去除", QMessageBox.ActionRole)
        box.addButton("取消", QMessageBox.RejectRole)
        box.setDefaultButton(btn_both)
        box.exec()
        clicked = box.clickedButton()
        if clicked is btn_text:
            return "text"
        if clicked is btn_img:
            return "image"
        if clicked is btn_both:
            return "both"
        return None

    def _toggle_watch(self, state):
        if self.watch_chk.isChecked():
            if not self.file_path or not os.path.exists(self.file_path):
                QMessageBox.warning(self, "提示", "请先选择有效的 Word 文件再启用守护。")
                self.watch_chk.setChecked(False)
                return
            kinds = self._gather_kinds()
            if not kinds:
                QMessageBox.warning(self, "提示", "请至少启用一种水印再启用守护。")
                self.watch_chk.setChecked(False)
                return
            out = self.out_edit.text().strip()
            if not out:
                out = core.default_output_path(self.file_path)
                self.out_edit.setText(out)
            if os.path.abspath(out) == os.path.abspath(self.file_path):
                QMessageBox.warning(self, "提示",
                                    "输出路径与源文件相同会覆盖原文件，请先修改输出路径。")
                self.watch_chk.setChecked(False)
                return
            opts = self._gather_opts()
            expect = None
            try:
                # 份数基准：优先用本次插入的实际份数；否则以“输出文件当前份数”为基准
                if self._last_inserted:
                    expect = int(self._last_inserted)
                elif os.path.exists(out):
                    expect = core.watermark_count(out) or None
            except Exception:
                expect = self._last_inserted
            self.wd = watchdog.WatermarkWatchdog(
                self.file_path, out, kinds, opts, interval=self.interval_spin.value(),
                log=self.thread_log, expected=expect)
            self.wd.start()
            if expect:
                self._append_log(f"守护水印份数基准：{expect} 份（被删掉任何一份都会自动补齐）。")
        else:
            if self.wd:
                self.wd.stop(); self.wd = None

    def _on_output_changed(self, text):
        # 占位：输出路径变化时无需即时处理，仅保持可被外部读取
        pass


def _run_elevated(mode):
    """以管理员身份重新启动后执行修复/还原，把结果写到文件再退出。

    mode: 'office-fix' / 'office-revert'
    """
    from watermark_tool import office_tweak as otw2
    ok, msg = (False, "未知的提权模式：%s" % mode)
    if mode == "office-fix":
        ok, msg, _ = otw2.apply_fix()
    elif mode == "office-revert":
        ok, msg, _ = otw2.revert_fix()
    otw2.write_result(ok, msg)
    print(msg)
    return 0 if ok else 1


SINGLE_INSTANCE_NAME = "WordWatermarkSingleInstance"


def _bring_window_up(win):
    """把已在跑的那个实例的窗口从托盘/最小化状态唤出来。"""
    try:
        win.showNormal()
    except Exception:
        pass
    win.show()
    win.raise_()
    win.activateWindow()


def _acquire_single_instance(name=None):
    """本机只允许一个实例常驻（否则会同时存在多个后台守护，互相抢着补水印）。

    返回 (QLocalServer | None, should_exit)：
      - 已有实例在跑 → 通知它把窗口唤出，本实例直接退出；
      - 抢到服务      → 返回 server，自己就是那个唯一实例；
      - QtNetwork 缺失或双方都失败 → 返回 (None, False)，宁可多开一个也
        不能让程序打不开。

    name 默认取生产用的 SINGLE_INSTANCE_NAME；测试可传独立名字，
    这样「机器上正跑着一个真实例」也不会把单例回归测试判成失败。
    """
    if name is None:
        name = SINGLE_INSTANCE_NAME
    try:
        from PySide6.QtNetwork import QLocalServer, QLocalSocket
    except Exception:
        return None, False

    # 先探测已有实例（本机实测：无实例/有实例都只要 0.1ms，不会拖慢启动）。
    # 不能反过来用 listen 结果判断——Windows 命名管道允许多实例同时 listen
    # 同名，第二次 listen 照样成功，那样会重复起出一个主实例。
    sock = QLocalSocket()
    sock.connectToServer(name)
    woke = sock.waitForConnected(300)
    if woke:                               # 已有实例在跑，通知它把窗口唤出来
        sock.write(b"show")
        sock.waitForBytesWritten(500)
    sock.disconnectFromServer()
    if woke:
        return None, True

    # 没人应答 = 没有实例，也可能留了空壳记录；清掉后自己成为那一个
    QLocalServer.removeServer(name)
    server = QLocalServer()
    if server.listen(name):
        return server, False
    return None, False                     # 连不上又抢不到：宁可多开，也不能打不开


def main():
    # 提权后的那一趟：不做界面，只执行完把结果交给父进程读取
    if "--office-fix" in sys.argv or "--office-revert" in sys.argv:
        mode = "office-revert" if "--office-revert" in sys.argv else "office-fix"
        return _run_elevated(mode)

    app = QApplication([])
    # 关闭/隐藏主窗口不应自动结束进程：后台守护模式下窗口是隐藏的，
    # 若保持默认 True，隐藏窗口可能会被判为“最后一个窗口已关闭”而连带退出应用。
    app.setQuitOnLastWindowClosed(False)

    server, already_running = _acquire_single_instance()
    if already_running:
        # 已有实例在跑：只把它唤出来，本实例不重复起后台
        return 0

    win = App()
    if server is not None:
        server.newConnection.connect(lambda: _bring_window_up(win))
    win.show()
    app.exec()
    return 0


if __name__ == "__main__":
    main()
