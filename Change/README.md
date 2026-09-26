# Change 文件夹

本文件夹是 **watermark_tool「保持行为不变」工程化重构** 的变更内容打包，便于在仓库里单独查看/评审，
而不直接改动仓库现有源码树。

## 背景
- 重构目标：在不新增功能、不改变用户可见行为与默认行为、不改变 Word 水印 XML 格式 / 清除逻辑 /
  视频输出逻辑、不删除现有测试 的前提下，逐项核查并最小修复 12 个风险点（a–l）。
- 实际落地的修改（对应本地提交 `84a773e`）：
  - 风险 f：watchdog 补回水印时不再从源文件覆盖用户对输出正文所做的修改。
  - 风险 g：engine_docx 对 python-docx 私有 API `r._r.append` 的调用包裹为可读异常。
  - 风险 i：engine_com 约 19 处 `except: pass` 改为 `log.warning/debug`，不再静默吞 COM 错误。
  - 风险 j：video 帧循环后先 `writer.close()` 再 mux，临时 mp4 在 `finally` 统一清理。
  - 风险 k：word_fonts 保留 TTC face 索引，多 face TTC（如“微软雅黑 Light”）不再错加载成第 0 面。
  - 风险 l：新增 i18n 不变式测试，锁住“中文 UI 文案作 key”契约。
- 经核查无实际风险、按“不为好看而改”原则未动的：gui.py 的 Worker 生命周期、closeEvent 兜底、
  过度膨胀/self.xxx 数量、非关键裸 except（风险 a/b/c/d/e）。
- 用户约束内未动的：`clear_any` 清除逻辑（风险 h）。

## 文件夹内容
```
Change/
├── README.md                      本说明
├── refactor.patch                 重构提交 84a773e 的完整 diff（相对 6075453）
├── watermark_tool/
│   ├── core.py                    风险 e/f：get_desktop 收窄 except + insert_watermark 新增 preserve_existing
│   ├── watchdog.py                风险 f：_reinsert 改用 preserve_existing=True
│   ├── engine_docx.py             风险 g：私有 API 调用包裹为可读 RuntimeError
│   ├── engine_com.py              风险 i：裸 except 改 logging
│   ├── video.py                   风险 j：writer.close 前置 + 临时文件清理
│   └── word_fonts.py              风险 k：保留 TTC face 索引 + font_index_of()
└── tests/unit/
    └── test_i18n.py               风险 l：新增 8 个不变式测试
```

## 如何使用
- **仅查看**：直接阅读上面的文件，与仓库对应路径的源码逐项对照即可。
- **应用变更（二选一）**：
  - 把 `Change/watermark_tool/*` 与 `Change/tests/unit/test_i18n.py` 覆盖回仓库对应位置；
  - 或执行 `git apply Change/refactor.patch`（该 patch 基于提交 `6075453`；若 main 已前进，
    可能需 `git apply --3way Change/refactor.patch` 或手动合并）。
- 应用后运行测试：`pytest -q`，预期 **53 passed**（原 45 + 新 i18n 8），无回归。

## 说明
- 此文件夹为「变更快照」，未直接修改仓库现有源码。
- 若要把重构正式合入主分支，请将本地 `main` 上的提交 `84a773e` 推送/合并，或把本文件夹内容应用后提交。
