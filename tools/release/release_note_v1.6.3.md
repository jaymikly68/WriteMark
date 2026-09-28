# WriteMark v1.6.3 更新说明

## 主要更新

- **清除水印的「冲突处理」**：当文档同时含有「本工具（WriteMark）添加的水印」和
  「Word 自带 / 你自己添加的水印」时，一键清除会弹出**四选一对话框**，让你明确选择：
  - **仅清除本工具水印**
  - **仅清除 Word/用户水印**（本工具自己插入的水印原样保留）
  - **两者都清除**
  - **取消**（不动任何内容）

  此前这类共存文档只能整体确认后全清，现在可以精确区分、按需删除，
  避免误删你自己的水印，也避免误删本工具的水印。

## 底层改动

- `core.clear_watermark` / `engine_docx.clear_watermark` 新增 `clear_native_only=True`
  模式：只删「Word 原生 / 用户自制水印」，保留本工具添加的水印（按私有标记识别）。
- GUI `_clear` 在检测到「本工具水印 ∩ 原生水印」共存时改走四选一冲突弹窗；
  仅有本工具水印（无原生）时仍按原有「文字/图片/两者」逻辑，行为不变。
- 配套单测覆盖四选一对话框的四种选择及其端到端效果（含 v1.6.2 起写在正文层的水印）。

## 验证

- 全量单测：`132 passed / 3 xfailed / 0 failed`（含新增的冲突弹窗与 native_only 测试）。
- `tests/unit` 全绿。

## 下载

[WordWatermark.exe](https://github.com/jaymikly68/WriteMark/releases/download/v1.6.3/WordWatermark.exe)
