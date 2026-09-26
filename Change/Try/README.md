# Change/Try —— 修复后测试与验证结果归档

本文件夹归档 **提交 `84a773e`（保持行为不变重构）之后** 连续四轮「验证 / 审查」所产生的
**测试代码**与**测试结果**，用于在不改动产品代码的前提下留存可复现的证据链。

与同级 `Change/` 的区别：

| 目录 | 内容 |
|---|---|
| `Change/original/`、`Change/watermark_tool/`、`Change/refactor.patch` | **一次性的变更快照**：重构前 / 后的源码与 diff |
| `Change/Try/`（本目录） | **验证过程的证据**：四轮审查新增的测试代码 + 全量测试运行结果 + 用例清单 |

## 背景与约束

- 审查轮次：review 2（二次安全审查）→ review 3（clear_any 形状误删）→
  review 4（按水印类型区分的删除回归）→ review 5（video 失败路径 / 替换语义 / 字体别名）。
- 每轮约束一致：**不新增功能、不做大规模重构、不改 GUI 结构 / i18n 架构 / Word 水印 XML 格式 / 视频格式**；
  已确认无实际风险的项不动代码，只补测试把现状钉死。

## 归档内容

### 1. 测试代码（8 个文件，共 56 个用例）

| 文件 | 用例数 | 验证什么 |
|---|---:|---|
| `test_clear_any_scope.py` | 9 | `clear_any` 的误删边界：正文 / 表格 / 行内图 / 非 behindDoc 浮动图不受影响；wps 形状不被误删；Word 原生水印按设计被清除 |
| `test_clear_kinds_matrix.py` | 15 | **按水印类型分别**验证三个删除入口：只删文字 / 只删图片 / 删全部，并直接改 XML 模拟「用户在 Word 里手动删」后触发 watchdog 真实线程链路 |
| `test_watchdog_preserve.py` | 9 | watchdog `preserve_existing` 不回退到从 source 重建：正文、表格、图片、段落格式、用户页眉 behindDoc 图形在补回后逐项存活 |
| `test_ttc_index.py` | 8 | 多 face TTC 的 face index 真实传给 `ImageFont.truetype(..., index=...)`；`_INDEX_MAP` 多次 `build_font_map()` 无残留 |
| `test_video_cleanup.py` | 4 | video 正常完成 / 降级 / 异常三条路径的临时 mp4 清理与 writer 关闭时序 |
| `test_video_failure_paths.py` | 4 | **修复点实测**：mux 失败且改名失败时不得再返回 `ok: True`；`get_writer` 构造失败不留 0 字节孤儿临时文件 |
| `test_insert_replacement_semantics.py` | 5 | 钉死 `insert_watermark` 是「替换语义」而非「追加语义」：先插 text 再单独插 image，文字水印确实消失（GUI 侧始终传完整集合，无风险） |
| `test_close_all_workers.py` | 2 | `closeEvent` 覆盖全部 5 类 worker，总清理时间落在 6 秒强制退出预算内（`os._exit(0)` 机制按约定保留未删） |

> 上述文件在仓库中的原始路径分别为 `tests/unit/<同名>` 或 `tests/gui/<同名>`。

### 2. 测试结果

| 文件 | 说明 |
|---|---|
| `pytest_full_result.log` | 全量 `pytest -p no:faulthandler --tb=short -q -rX` 的完整输出（含 warnings summary 与 xfail 清单） |
| `pytest_cases_inventory.txt` | 全量用例清单（111 个 / 15 个文件），按文件列出 nodeid，新增与原有分开标注 |

当前整轮结果：**107 passed, 4 xfailed, 1 warning in ~60s**。

> 4 个 xfailed 是「按约定已确认但本轮未修」的风险的可执行记录
> （`clear_any` 删用户自己的 behindDoc 页眉图片、同 run 兄弟文字连带删除、
> 别名表绕过 TTC 索引、注册表无 `,N` 时多 face 只能取第 0 面）。
> **它们不代表风险已消失**，只代表风险被固化在测试里，将来收紧实现会变成 XPASS 提醒收口。

### 3. 审查报告（4 份）

按时间顺序，与上面四轮一一对应：

| 文件 | 轮次 | 核心结论 |
|---|---|---|
| `REVIEW2_REPORT.md` | 二次安全审查 | 已修复 `preserve_existing` 补回时连带删用户页眉图形、source 被占用导致补不回；`closeEvent` 漏掉 4 类 worker。提出 D1–D6 |
| `REVIEW3_REPORT.md` | clear_any 边界 | **修复真实数据丢失**：`clear_any` 会删用户自己画的 wps 形状（OOXML 里同属 `w:drawing`），最小收窄为「只有图片才允许被推定为水印」 |
| `REVIEW4_REPORT.md` | 按类型区分的删除回归 | 15 个新用例；**用 `84a773e~1` 旧代码实跑同一套测试得到 9 failed / 6 passed**，证明测试真能抓到旧 bug |
| `REVIEW5_REPORT.md` | video 失败路径等 | 修复 video.py 静默假成功（`os.replace` 失败被吞却返回成功）与 0 字节孤儿临时文件；清理 `_ALIAS_FILES` 重复 key；钉死替换语义 |

## 如何复现

测试需在项目根目录下运行（依赖仓库的 `conftest.py` 与 `watermark_tool` 包路径）：

```bash
cd word-watermark-tool
venv/Scripts/python.exe -m pytest -q -p no:faulthandler
# 预期：107 passed, 4 xfailed
```

本目录内的测试副本仅为**归档**，直接在本目录运行会缺少 `conftest.py` 与包上下文，
请从仓库 `tests/` 下的原始路径运行；或在此目录放置一个
`sys.path.insert(0, "<仓库根目录>")` 的 `conftest.py` 后运行。

## 「改动前 vs 改动后」对照（证据强度说明）

`REVIEW4_REPORT.md` 中已给出可复现的对照方法：把 `84a773e~1` 的代码 `git archive` 导出到独立目录，
把本系列测试原样复制进去跑，同一套用例在旧版上失败 9 条，最典型的两条：

```
1) test_scenario_delete_text_watermark_only
   AssertionError: watchdog 补回后用户新增的正文必须还在
   assert 'USER_EDIT_12345' in ['原始正文第一段', '原始正文第二段']
   → 旧 _reinsert 从 source 覆盖 output，用户正文被整份抹掉

2) test_clear_all_preserves_body_table_user_picture_shape_and_header_footer
   AssertionError: [kinds=None(clear_any)] 全清时用户自己的页眉形状被误删了
   assert 0 == 1
```

## 仍未收口的技术债务（需产品决策，未擅自修改）

- **D1** `clear_any` 会删用户自己在页眉/页脚设成「衬于文字下方」的图片 —— 属设计取舍，收口会改变公开清除行为。
- **D2** 页眉某 `w:r` 内文字与图形并存时，`_detach_drawing` 摘整条 run，兄弟文字一起消失。
- **D3** `font_index_of` 别名表短路，导致常用字体的 TTC 分面修复实际不生效。
- **D4** 注册表无 `,N` 时，多 face TTC 只能取第 0 面。
- **D5** `_INDEX_MAP` 与 `_MAP_CACHE` 半新半旧时静默退化为 0。
- **D6** python-docx 非原子落盘（`ZipFile(..., "w")` 直写目标路径），`os._exit` / 强杀可能留下写坏的半截 docx。
- **D7** COM worker 被强杀可能残留 WINWORD.EXE。
- **D8** `core.clear_watermark` 是「从 src 复制到 out」语义，GUI 一键清除会覆盖用户在 out 上的编辑。
- **D9** `watermark_count` 只数总份数、不分类别，图片水印意外残留时 watchdog 不会补回。
