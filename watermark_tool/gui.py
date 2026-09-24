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
import threading

from PySide6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QLineEdit, QPushButton, QLabel, QFileDialog, QComboBox, QCompleter,
    QDoubleSpinBox, QCheckBox, QSlider, QTextEdit, QColorDialog, QMessageBox,
    QScrollArea, QSystemTrayIcon, QMenu, QStyle,
)
from PySide6.QtCore import Qt, QThread, Signal, QTimer, QStringListModel
from PySide6.QtGui import QColor, QImage, QPixmap

from . import core, watchdog, preview, word_fonts


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


class FontWorker(QThread):
    """后台线程：读取可用字体（Word COM 优先，失败回退系统字体），避免卡 UI。"""
    finished_signal = Signal(list, str)   # (字体名列表, 来源说明)
    error_signal = Signal(str)

    def run(self):
        try:
            fonts, source = word_fonts.collect_fonts()
            self.finished_signal.emit(fonts or [], source)
        except Exception as e:
            self.error_signal.emit(str(e))


class App(QMainWindow):
    # 线程安全日志信号（必须在类级别声明）
    log_signal = Signal(str)
    status_signal = Signal(str)

    def __init__(self):
        super().__init__()
        self.setWindowTitle("Word 一键水印工具")
        self.resize(880, 640)

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
        self._confirm_close = True   # 点关闭按钮时是否弹“最小化/退出”询问（自动化脚本可关闭）
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

        # 文件
        f_file = QWidget()
        h = QHBoxLayout(f_file); h.setContentsMargins(0, 0, 0, 0)
        self.file_edit = QLineEdit(); self.file_edit.setPlaceholderText("选择 Word 文件 (.docx / .doc)")
        h.addWidget(self.file_edit, 3)
        btn = QPushButton("浏览..."); btn.clicked.connect(self._browse_file); h.addWidget(btn, 1)
        root.addWidget(f_file)

        # 输出文件（不破坏原文件）
        f_out = QWidget()
        ho = QHBoxLayout(f_out); ho.setContentsMargins(0, 0, 0, 0)
        ho.addWidget(QLabel("输出:"))
        self.out_edit = QLineEdit(); self.out_edit.setPlaceholderText("默认保存到桌面/<原名>_水印.docx，可点击浏览修改")
        ho.addWidget(self.out_edit, 3)
        bout = QPushButton("浏览..."); bout.clicked.connect(self._browse_output); ho.addWidget(bout, 1)
        root.addWidget(f_out)
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
        # 文本水印专属：水平/垂直偏移（支持上下左右调整，占页宽/页高百分比）
        row, self.text_offx_slider, self.text_offx_spin = self._make_slider_spin(
            "水平偏移(%):", -50, 50, self.text_offset_x, 2, 1.0, "%", self._on_text_offx)
        vt.addLayout(row)
        row, self.text_offy_slider, self.text_offy_spin = self._make_slider_spin(
            "垂直偏移(%):", -50, 50, self.text_offset_y, 2, 1.0, "%", self._on_text_offy)
        vt.addLayout(row)
        for w in (self.text_edit, self.cn_font_combo, self.latin_font_combo,
                  self.font_spin, self.color_btn,
                  self.text_angle_slider, self.text_angle_spin,
                  self.text_trans_slider, self.text_trans_spin, self.text_scale_spin,
                  self.text_offx_slider, self.text_offx_spin,
                  self.text_offy_slider, self.text_offy_spin):
            self.text_ctrl_widgets.append(w)
        f_text.layout().addLayout(vt)

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
        # 图像水印专属：水平/垂直偏移（支持上下左右调整，占页宽/页高百分比）
        row, self.img_offx_slider, self.img_offx_spin = self._make_slider_spin(
            "水平偏移(%):", -50, 50, self.img_offset_x, 2, 1.0, "%", self._on_img_offx)
        vi.addLayout(row)
        row, self.img_offy_slider, self.img_offy_spin = self._make_slider_spin(
            "垂直偏移(%):", -50, 50, self.img_offset_y, 2, 1.0, "%", self._on_img_offy)
        vi.addLayout(row)
        for w in (self.img_angle_slider, self.img_angle_spin,
                  self.img_trans_slider, self.img_trans_spin, self.img_scale_spin,
                  self.img_offx_slider, self.img_offx_spin,
                  self.img_offy_slider, self.img_offy_spin):
            self.img_ctrl_widgets.append(w)
        f_img.layout().addLayout(vi)

        # ---------------- 文本/图像水印横向并排，缩短整体纵向高度 ----------------
        h_types = QHBoxLayout()
        h_types.setSpacing(8)
        h_types.addWidget(f_text)
        h_types.addWidget(f_img)
        root.addLayout(h_types)

        # ---------------- 按钮
        hb = QHBoxLayout()
        self._btn_insert = QPushButton("一键插入水印"); self._btn_insert.clicked.connect(self._insert); hb.addWidget(self._btn_insert)
        self._btn_clear = QPushButton("一键清除水印"); self._btn_clear.clicked.connect(self._clear); hb.addWidget(self._btn_clear)
        root.addLayout(hb)

        # ---------------- 防去除加固（可选）
        f_hard = self._make_group("防去除加固（让水印更难被删掉）")
        vh = f_hard.layout()
        self.tile_chk = QCheckBox("平铺满页水印（覆盖整页，PS/AI 修图难以抹除）")
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
        self.redundant_chk = QCheckBox("多份冗余嵌入（页眉+页脚都写入，图形名不使用 watermark 字样）")
        self.redundant_chk.setChecked(self.redundant)
        self.redundant_chk.setToolTip(
            "除页眉外，把水印也写进页脚；并把图形的显示名改成普通图片那样的名字。\n"
            "Word/WPS 的“删除水印”按钮和多数去水印脚本都是按名称/结构匹配水印图形的，\n"
            "这样它们就找不到、删不干净。对肉眼外观没有任何影响。")
        vh.addWidget(self.redundant_chk)
        root.addWidget(f_hard)

        # 预览
        f_prev = self._make_group("水印预览（示意，脱离 Word 直接查看）")
        vp = f_prev.layout()
        self.preview_label = QLabel(); self.preview_label.setAlignment(Qt.AlignCenter)
        self.preview_label.setMinimumSize(230, 326)
        self.preview_label.setStyleSheet("border:1px solid #bbb; background:#fafafa;")
        vp.addWidget(self.preview_label)
        hprev = QHBoxLayout()
        bp = QPushButton("预览水印"); bp.clicked.connect(self._render_preview); hprev.addWidget(bp)
        bps = QPushButton("保存预览图"); bps.clicked.connect(self._save_preview); hprev.addWidget(bps)
        vp.addLayout(hprev)
        root.addWidget(f_prev)

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
    def _on_text_offy(self, v):
        self.text_offset_y = float(v)

    def _on_img_angle(self, v):
        self.img_angle = float(v)
    def _on_img_trans(self, v):
        self.img_transparency = v / 100.0
    def _on_img_scale(self, v):
        self.img_scale = float(v)
    def _on_img_offx(self, v):
        self.img_offset_x = float(v)
    def _on_img_offy(self, v):
        self.img_offset_y = float(v)

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
            # 选完源文件后，自动给出默认输出路径（桌面/同名_水印），用户可改
            self.out_edit.setText(core.default_output_path(p))
            # 导入 Word 后，征询用户同意后读取其可用字体，扩充文本水印字体选择
            self._maybe_load_word_fonts()

    def _maybe_load_word_fonts(self):
        """导入 Word 后：先征询权限，同意后后台读取 Word（或系统）字体并扩充下拉框。

        每会话仅询问一次；用户拒绝则沿用内置字体，不再打扰。
        """
        if self._fonts_loaded:
            return
        if self._fonts_asked:
            return
        self._fonts_asked = True
        ans = QMessageBox.question(
            self, "读取 Word 字体",
            "是否允许本工具读取您 Word 中可用的字体，以扩充“文本水印”的中文字体 / 西文字体选择？\n\n"
            "说明：仅读取字体名称列表，不会访问或修改您任何文档的内容。\n"
            "（若本机未安装 Word，将自动改用系统已安装字体作为来源。）",
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
                                           "图片 (*.png *.jpg *.jpeg *.bmp *.gif);;All (*.*)")
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

    def _ask_close_choice(self):
        """点击标题栏关闭按钮时弹出询问。

        返回 'minimize'（隐藏窗口、后台继续保护）/ 'quit'（退出程序）/ 'cancel'（取消）。
        """
        has_tray = self.tray is not None
        box = QMessageBox(self)
        box.setWindowTitle("关闭窗口")
        box.setIcon(QMessageBox.Question)
        if has_tray:
            box.setText("后台保护正在运行。要如何处理这个窗口？")
            box.setInformativeText(
                "【最小化到后台】窗口隐藏到系统托盘，守护继续运行，水印被删会自动补回。\n"
                "单击/双击托盘图标可重新打开窗口，右键托盘可停止守护或退出程序。\n\n"
                "【退出程序】停止后台保护并完全退出进程。")
        else:
            box.setText("后台保护正在运行，但本机系统托盘不可用。")
            box.setInformativeText(
                f"托盘初始化失败原因：{self._tray_reason or '未知'}\n\n"
                "【最小化到后台】窗口隐藏，进程留在后台守护（无托盘图标，重新打开需再次启动程序）。\n"
                "【退出程序】停止后台保护并完全退出进程。")
        btn_min = box.addButton("最小化到后台" if has_tray else "最小化到后台（无托盘图标）",
                                QMessageBox.AcceptRole)
        btn_quit = box.addButton("退出程序", QMessageBox.DestructiveRole)
        btn_cancel = box.addButton("取消", QMessageBox.RejectRole)
        box.setDefaultButton(btn_min)
        box.exec()
        clicked = box.clickedButton()
        if clicked is btn_quit:
            return "quit"
        if clicked is btn_cancel:
            return "cancel"
        return "minimize"

    def closeEvent(self, event):
        """关闭窗口时清理后台线程，避免进程无法退出（表现为关闭后卡顿/驻留）。

        要点：
        - 守护运行时：先弹询问（最小化到托盘继续保护 / 退出程序 / 取消），
          选择“最小化”则隐藏窗口，守护继续后台运行，
          这样“水印被删自动补回”不再依赖主窗口一直开着。
        - 真正退出时：守护线程已是 daemon，只发停止信号、绝不阻塞等待。
        - 兜底：启动一个计时线程，1.2s 后若进程仍未自行退出，强制 os._exit(0)。
        - 工作/字体线程若超时仍未结束则强制 terminate。
        """
        # 守护运行中：先问用户要不要留后台继续保护
        if self.wd is not None and not self._quitting:
            if getattr(self, "_confirm_close", True) and self.isVisible():
                choice = self._ask_close_choice()
            elif self._quitting:
                # 自动化脚本 / 托盘“退出程序”已明确表达要退出，无需再问
                choice = "quit"
            else:
                choice = "minimize"
            if choice == "cancel":
                event.ignore()
                return
            if choice == "minimize":
                event.ignore()
                self.hide()
                tip = ("守护已在后台继续运行，水印被删会自动补回。\n"
                       "单击/双击托盘图标可重新打开窗口，右键托盘可停止守护 / 退出程序。")
                shown = False
                try:
                    if self.tray is not None:
                        self.tray.showMessage("Word 一键水印工具", tip,
                                              QSystemTrayIcon.Information, 4000)
                        shown = True
                except Exception:
                    pass
                if shown:
                    self._append_log("已最小化到系统托盘，守护继续在后台运行。")
                else:
                    self._append_log("窗口已隐藏，进程仍在后台守护（本机托盘不可用，"
                                     "重新打开界面需再次启动程序）。")
                return

        # 兜底：无论后续清理是否卡住，超时(6s)后强制结束进程。
        # 6s 是刻意留出的余量：下面等待后台线程的完整路径最长约 3s，
        # 必须让“正常清理”跑在兜底之前完成，否则明明能优雅退出却被硬杀
        # （此前的 1.2s 就会误杀正在等待线程退出的正常关闭流程）。
        killer = threading.Timer(6.0, lambda: os._exit(0))
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
        else:
            self.status_label.setText("完成")
            if self._last_action == "insert":
                self._append_log("已生成带水印的新文件（原文件未改动）。请打开下方日志中 output 路径的文件查看；"
                                 "在 Word/WPS 中需切到『页面视图/页面布局』才能看到水印。")
                self._last_action = ""

    # ------------------------------------------------------------- 预览
    def _render_preview(self):
        opts = self._gather_opts()
        kinds = self._gather_kinds()
        try:
            img = preview.render_preview(opts, kinds)
        except Exception as e:
            # 实时预览过程中参数可能暂时不完整（如未选图片），直接在预览区提示，不弹窗打断
            self._preview_img = None
            self.preview_label.setPixmap(QPixmap())  # 清空旧图
            self.preview_label.setText(f"预览不可用：{e}")
            return
        self._preview_img = img
        bio = __import__("io").BytesIO(); img.save(bio, "PNG"); bio.seek(0)
        qimg = QImage.fromData(bio.getvalue())
        pix = QPixmap.fromImage(qimg)
        scaled = pix.scaled(self.preview_label.size(), Qt.KeepAspectRatio, Qt.SmoothTransformation)
        self.preview_label.setPixmap(scaled)

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
        self._run(lambda: core.clear_watermark(self.file_path, output_path=out))

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


def main():
    app = QApplication([])
    # 关闭/隐藏主窗口不应自动结束进程：后台守护模式下窗口是隐藏的，
    # 若保持默认 True，隐藏窗口可能会被判为“最后一个窗口已关闭”而连带退出应用。
    app.setQuitOnLastWindowClosed(False)
    win = App()
    win.show()
    app.exec()


if __name__ == "__main__":
    main()
