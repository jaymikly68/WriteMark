# Change 文件夹

本文件夹是 **watermark_tool「保持行为不变」工程化重构**（提交 `84a773e`）的完整变更打包：
**改动前原始文件 + 改动后重构文件 + 完整 diff + 说明**，便于在仓库里单独查看/评审，
而不直接改动仓库现有源码树。

## 背景
- 重构目标：在不新增功能、不改变用户可见行为与默认行为、不改变 Word 水印 XML 格式 / 清除逻辑 /
  视频输出逻辑、不删除现有测试 的前提下，逐项核查并最小修复 12 个风险点（a–l）。
- 实际落地的修改（对应本地提交 `84a773e`，其父提交即改动前基线 `6075453`）：
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
├── README.md                        本说明
├── refactor.patch                   重构提交 84a773e 的完整 diff（相对 6075453）
├── original/                        【改动前】重构前的原始文件（基线 = 84a773e 的父提交 6075453）
│   └── watermark_tool/
│       ├── core.py                  150 行
│       ├── engine_com.py            404 行
│       ├── engine_docx.py           695 行
│       ├── video.py                 488 行
│       ├── watchdog.py              103 行
│       └── word_fonts.py            393 行
├── watermark_tool/                  【改动后】重构后的文件（= 84a773e 版本）
│   ├── core.py                      173 行（get_desktop 收窄 except；insert_watermark 加 preserve_existing）
│   ├── engine_com.py                410 行（裸 except → log.warning/debug）
│   ├── engine_docx.py               714 行（私有 API 包裹为可读 RuntimeError）
│   ├── video.py                     495 行（writer.close 前置 + 临时文件 finally 清理）
│   ├── watchdog.py                  110 行（_reinsert 用 preserve_existing=True）
│   └── word_fonts.py                440 行（保留 TTC face 索引 + font_index_of()）
└── tests/unit/
    └── test_i18n.py                 【新增文件】8 个 i18n 不变式测试（改动前不存在此文件）
```

> 注：`tests/unit/test_i18n.py` 是本次重构**新增**的测试文件，改动前不存在，
> 因此 `original/` 下没有它的“改动前版本”；`original/tests/` 目录因此省略。

## 对照关系
| 文件 | 改动前（original/） | 改动后（根下同名路径） | 差异 |
|---|---|---|---|
| watermark_tool/core.py | 150 行 | 173 行 | 见 refactor.patch |
| watermark_tool/engine_com.py | 404 行 | 410 行 | 见 refactor.patch |
| watermark_tool/engine_docx.py | 695 行 | 714 行 | 见 refactor.patch |
| watermark_tool/video.py | 488 行 | 495 行 | 见 refactor.patch |
| watermark_tool/watchdog.py | 103 行 | 110 行 | 见 refactor.patch |
| watermark_tool/word_fonts.py | 393 行 | 440 行 | 见 refactor.patch |
| tests/unit/test_i18n.py | （不存在） | 92 行（新增） | 见 refactor.patch |

## 如何使用
- **查看改动**：`original/` 与根下同名路径逐文件对照；或直接阅读 `refactor.patch`（完整 diff）。
- **应用变更（二选一）**：
  - 把 `Change/watermark_tool/*` 与 `Change/tests/unit/test_i18n.py` 覆盖回仓库对应位置；
  - 或执行 `git apply Change/refactor.patch`（该 patch 基于提交 `6075453`；当前 main 仅多了一个
    与这些文件无关的 `fix` 文件提交，patch 可正常应用；若仍有冲突可用
    `git apply --3way Change/refactor.patch`）。
- 应用后运行测试：`pytest -q`，预期 **53 passed**（原 45 + 新 i18n 8），无回归。

## 说明
- 此文件夹为「变更快照」，未直接修改仓库现有源码（仓库根下的 `watermark_tool/` 仍是改动前版本）。
- 若要把重构正式合入主分支，请将本地 `main` 上的提交 `84a773e` 推送/合并，或把本文件夹内容应用后提交。
