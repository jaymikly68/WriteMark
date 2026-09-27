## WriteMark v1.6.0

**下载 [WordWatermark.exe](https://github.com/jaymikly68/WriteMark/releases/download/v1.6.0/WordWatermark.exe)**
124.7 MB，双击即可运行，无需安装 Python。首次运行会自行解包到临时目录，请稍候数秒。

### 本次更新（v1.5.0 → v1.6.0）

**视频水印**

- **文字水印可分别指定中文字体 / 西文字体**（与 Word 文本水印一致）：汉字用中文字体、
  拉丁字母与数字用西文字体，渲染按字符类别自动分流。视频没有“导入文档”动作，
  字体来源与 Word 相同——**本机已安装字体**（启动即扫描 Windows 字体注册表扩充下拉框，
  不启动 Word、无需授权）。
- **多页 PDF 作图片水印时可「PDF取任意页」**：新增页码选择框（只有选中的是 PDF 才可用，
  上限自动取该 PDF 的实际页数），悬停「PDF取任意页」会显示蓝色小字说明。
  此前固定只取第 1 页。
- **新增「生效时段」**：填一行 `1.0-2.0`，水印**只在 `[起, 止)` 这段时间出现**，
  其余时刻画面保持原样——适合只遮住视频里某几秒才出现的敏感信息
  （例：3 秒视频只想遮第 2 秒 → 填 `1.0-2.0`）。分隔符 `-`、`~`、「到」、逗号、空格都认；
  填完下方即时回述「水印只在第 x ~ y 秒出现」，填错给红字提示。
  预览与导出共用同一套判定，所见即所得。

**一处点击 BUG 的根治**

- 「时刻（秒）」的 QDoubleSpinBox 自带上下箭头，在部分 Windows 样式/宽度下
  **点击热区与三角形错位**，表现为“点向上的三角形却落进输入框、变成手动输入”，
  而向下的那个能正常 -0.5 秒。现已改为关闭原生箭头、使用一对**独立的 ▲/▼ 按钮**：
  上 +0.5 秒、下 -0.5 秒，按住可连续步进，热区就是按钮本身。

**界面**

- 「浏览...」「选择图片...」等 6 处次级按钮统一为淡蓝底 + 深蓝字。
- 悬停说明短语不再同时弹出系统 tooltip（原先会出现一红一蓝两句重复文案）。
- i18n 清理 7 条已失效的多语言文案键。

**测试**

- 新增 `tests/gui/{test_video_fonts,test_pdf_page,test_video_time_range,test_time_stepper,test_ui_polish,test_drag_target,test_clear_confirm}.py`
  与 `tests/unit/test_detect_native_watermark.py`；步进按钮用 `QTest` 真实鼠标点击验证。
- 全量：**159 passed, 3 xfailed**。

### 校验

MD5 ：以当次构建为准（exe 每次构建都会因内嵌时间戳而不同，不记录固定值）。

### 使用注意

- 升级前请从系统托盘区退出旧实例，否则新 exe 会因单实例占用而起不来
- 源码见本仓库 `main` 分支（本版提交 `dc2392d`）；后续发版执行
  `python release_upload.py dist/WordWatermark.exe --tag vX.Y.Z`
