# 二次安全审查报告（针对 commit `84a773e`）

范围：只做本轮指定的 4 项，**不新增功能、不做架构重构、不改 UI 与 Word 水印格式**。
结论按「已修复 / 风险降低 / 仍存在的技术债务」三类分开列，**测试通过 ≠ 风险消失**。

- 测试基线：`84a773e`（上一轮重构后：53 passed）
- 本轮新增测试：`tests/unit/test_watchdog_preserve.py`、`tests/unit/test_clear_any_scope.py`、
  `tests/unit/test_ttc_index.py`、`tests/gui/test_close_all_workers.py`
- 本轮改动源码：`watermark_tool/core.py`、`watermark_tool/gui.py`（gui.py 仅改 closeEvent 的资源收尾）
- 最终全量：`pytest -q` → **71 passed, 4 xfailed**（原 53 + 新增 22）

> 关于 xfailed：本轮有 4 个复现型用例刻意保持失败态（xfail），它们是**已确认但按约定未修**的
> 风险的证据。将来收紧实现，它们会变成 XPASS，提醒回来收口。

---

## 主题 1：watchdog + preserve_existing

### 结论：**已修复 2 个真实缺陷**

先确认你最关心的一点：**`preserve_existing=True` 在所有分支里都只操作现有 output**，
不存在「某些路径偷偷从 source 重新生成」。我逐条查过并写了测试锁定：
`core.insert_watermark` 的 `.doc`(COM) 与 `.docx` 两个分支，判定条件均为
`preserve_existing and output.exists(output_path)`；只有 output 被**整体删除**时才会
从 source 重建（这是设计意图，见 `watchdog._loop` 第 73–77 行），由
`test_watchdog_output_deleted_rebuilds_from_source` 固定。
`test_preserve_existing_never_reads_source_content` 进一步验证：改了 source 也不会被搬进 output。

**修复 1 — 补回时连带删除用户在页眉/页脚的 behindDoc 图形**（`core.py`）

- 风险：`preserve_existing` 的“先清”步骤原本是 `engine_docx.clear_watermark(output_path)`，
  `kinds=None` 会走 `clear_any=True` 分支，于是**不光清本工具水印，还清掉页眉/页脚里任何
  `behindDoc=1` 的浮动图形**。也就是说“保留用户修改”这句承诺在页眉层面并不成立。
- 修复（1 行）：改成 `clear_watermark(output_path, kinds=list(kinds))`，只清本工具自己标记的图形。
  既有公开清除 API 的行为**完全未变**，变的只是这次新增的补回路径。
- 复现→修复前：`test_watchdog_reinsert_preserves_user_behinddoc_header_graphic` 失败（kept=0）；
  修复后通过。

**修复 2 — source 被 Word/WPS 占用时，守护连补回都做不了**（`core.py`）

- 风险：`insert_watermark` 在函数最开头就 `_ensure_not_locked(src_path)`，而此时已经决定
  “根本不读 source”。结果是：用户在 Word 里打开 **source** 时，守护每次补回都抛
  “文件正被 WPS/Word 占用”，水印补不回来——恰好是最该生效的场景反而失效。
- 修复（2 行）：判出 `reuse_existing` 后跳过 source 占用检查（output 仍照常检查）。
- 复现→修复前 `test_preserve_existing_when_source_is_locked` 失败；修复后通过。

**验证覆盖（6 个用例全过）**：正文/表格/图片/段落格式全量快照比对；
守护真实 `_reinsert()` 往返（删水印、补回、再补两轮后比对快照）；output 被删时从 source 重建；
不回读 source；source 被占用时仍能补回；用户页眉 behindDoc 图形被保留。

**风险降低**：`test_preserve_existing_keeps_user_body_table_image_format` 覆盖的
“正文/表格/图片/字体加粗与字号”四类资产，此前**没有任何测试守护**，现在被快照比对锁死。

---

## 主题 2：clear_any 安全边界

### 结论：**已复现，按你的要求未改实现**（2 个 xfail 复现用例）

`clear_any` 的识别条件（`engine_docx._remove_in_part`，`clear_any=True` 时）：

1. 扫描范围 = 每个 section 的**页眉 + 页脚**（含首页/偶数页）内的 `w:drawing`；
   **正文与表格不在范围内**（安全）。
2. 命中即删，两条任一成立：
   - 带本工具私有标记（`wp:docPr/@descr` 或 `@name` ∈ {WB_WATERMARK_TEXT, WB_WATERMARK_IMG}）；
   - **或** 是 `wp:anchor[behindDoc="1"]` 的浮动图形 —— **不校验归属**。
3. 删除手法 `_detach_drawing` 会把整条 `w:r` 一起摘掉，不是只摘图形。
4. 只认 `wp:anchor`；`wp:inline`（页眉 Logo）与非 behindDoc 的浮动图形不参与（安全）。

**已复现的误删风险（未改实现）**

- **R-1**：用户在页眉里放的、设为“衬于文字下方”的背景图/Logo 会被清除。
  Word 原生水印确实长这样（所以按设计要清），但用户自己的同类图形无法区分。
  触发路径是 GUI「清除水印」——`gui._clear()` 在文档不是“文字+图片双水印”时把 `kinds` 留成
  `None`，于是直接走 `clear_any=True`。这是**常见路径**，不是边缘情况。
  复现用例：`test_clear_any_deletes_user_own_behinddoc_header_picture`（xfail）。
- **R-2（放大效应）**：若页眉某个 `w:r` 里同时有文字和 behindDoc 图形，整条 run 连文字一起消失。
  复现用例：`test_clear_any_deletes_whole_run_including_sibling_text`（xfail）。

**不是风险、已写用例固定**：正文/表格不动；行内图片不动；非 behindDoc 浮动图形不动；
Word 原生水印按设计被清除；`kinds` 限定路径只删本工具标记（GUI「只去文字水印」是安全的）。

**待你定夺的修法（本次未动）**：在 `_remove_in_part` 里给“无标记的 behindDoc 图形”加一层
白名单判定（例如名字/尺寸/位置像系统水印才清），或在对这类文档清除前弹二次确认。
这会改变公开清除行为，属于产品决策，需要你点头。

---

## 主题 3：word_fonts TTC index

### 结论：**无残留**（你问的第一点，已用测试证明）；**但发现索引机制被别名表短路**（未改）

**A. `_INDEX_MAP` 是否会残留旧字体索引 → 不会。**
`build_font_map()` 开头就 `_INDEX_MAP = {}` 整体重建，不会累积。
`test_index_map_no_stale_residue_between_builds` 用假注册表模拟“第二次构建时字体被卸载”，
断言索引表与路径表键完全一致、无残留；`test_index_map_matches_mapping_keys` 另外固定了
“非法索引回退 0 / 空值条目跳过”的既有行为。

**B. `index=1` 是否真的载入第 1 个 face → 机制本身可用，但被别名表短路。**
先用真实 `msyh.ttc` 证明：Pillow 的 `ImageFont.truetype(path, 40, index=1).getname()`
返回 `'Microsoft YaHei UI'`，`index=0` 返回 `'Microsoft YaHei'`，两者确实不同
（`test_ttc_face_1_loads_when_index_is_supplied`）。

但 `font_index_of()` 里有一句 `if target in _ALIAS_FILES: return 0`：
别名表覆盖了**微软雅黑 / 宋体 / 黑体 / 仿宋 / 楷体**等主要字体，
因此 84a773e 的 TTC 分面修复对它们**完全不生效**；
「微软雅黑 UI / Light」这类变体会渲染成第 0 面（常规字重）。
复现用例：`test_alias_table_does_not_bypass_ttc_index`（xfail）。

**另一处残留债务（次要）**：本机注册表把多 face TTC 记成单纯的 `msyh.ttc`（无 `,N`），
`split_alias_entries` 又从 `A & B` 拆出 `Microsoft YaHei UI` 这个显示名，
索引只能取 0 → 选 UI 变体仍渲染常规面（`test_real_machine_ui_variant_uses_ui_face`，xfail）。
本机实测输出：`msyh.ttc：索引=0，选中 face='Microsoft YaHei'，该文件各 face=['Microsoft YaHei', 'Microsoft YaHei UI']`。

**耦合债务（已固化，待修）**：`build_font_map()` 刷新 `_INDEX_MAP` 但不刷新 `_MAP_CACHE`，
二者可能“半新半旧”，此时 `font_index_of` 静默返回 0（`test_index_map_and_cache_are_coupled_characterization`）。
修法是让两者配套刷新，属小改动，等你决定。

---

## 主题 4：closeEvent / os._exit

### 结论：**未删除 6 秒强制退出机制**；**已修复 1 个真实缺陷**（worker 收尾不全）

**修复 — closeEvent 只收尾了 `self._worker`，其余 4 类 worker 被漏掉**（`gui.py`）
`_v_worker`（视频导出）、`_v_preview_worker`、`_office_worker`（COM/.doc）、`_font_worker`
此前没有任何停机处理。视频导出没走完就退出时，它们的 `finally`
（video.py 的临时 mp4 清理、writer 关闭）没人执行。现已统一收尾，并加了共享预算：

```python
SHUTDOWN_BUDGET = 2.5          # 所有 worker 的等待合计不得超过它
deadline = time.monotonic() + SHUTDOWN_BUDGET
```

> 为什么必须共享预算：我第一版让 5 个 worker 各等满 1.2s，串行最坏 7.5s，会**越过 6s 兜底**
> —— 等于自己制造硬杀。现在总清理 << 6s，`test_all_workers_shut_down_and_stay_within_force_exit_budget`
> 实测 3.4s < 6s，把这个不变式钉住了。

**正常退出路径（不触发 os._exit）的时序**

| 资源 | 发生了什么 | 是否安全 |
|---|---|---|
| watchdog | `stop_nowait()` 只置停止信号，daemon 线程自行退出（≤1 个间隔） | 是 |
| 5 类 worker | 现在都会 requestInterruption → quit → 超时 terminate | 基本是（见下） |
| COM / Word | 正常结束时 `engine_com` 的 `finally` 调 `_close_word()` 强杀 WINWORD | 是 |
| video 临时文件 | `finally` 里 `os.remove(tmp_vid)` | 仅是当线程能正常退出 |

**超时强杀路径（`os._exit(0)`）会在进程外留下的东西**

`os._exit` 会跳过 `finally` / `atexit` / `__del__` / 缓冲刷新 / Qt 析构，因此：

1. **video.py 的临时 mp4**（`tempfile.mkstemp`）→ 孤立在 `%TMP%`，进程内无法补救。
2. **正在写的输出文件**——特别注意：python-docx 的落盘是
   `ZipFile(pkg_file, "w")`，**直接写目标路径，不用临时文件、也不原子改名**
   （已核对 `venv/.../docx/opc/phys_pkg.py:109`）。所以硬杀发生在 `document.save()` 期间时，
   目标 `.docx/.doc` 可能留下**被写坏的半截文件**；被 `terminate()` 强杀的 docx worker 同理。
3. **COM / Word**：若 `OfficeWorker` 被 `terminate()`，COM 调用被中断，`_close_word()` 来不及跑，
   可能残留 WINWORD.EXE 进程。
4. **守护线程正在补写 output** 时硬杀 → 同 2。

**这些都无法在进程内修复**——`os._exit` 连 `atexit` 都不跑，唯一的办法是缩短窗口：
本轮已把「所有 worker 的收尾」纳入 6s 之前完成（`SHUTDOWN_BUDGET=2.5s`），
但实际上只要主线程不卡死，兜底计时器根本不会触发。

**建议（本次未改，等你确认）**：给 `engine_docx` 的保存加一个原子落盘封装
（同目录临时文件 → `os.replace`），这样第 2、4 条从“可能写坏目标文件”降级为“最多留临时文件”。
改动面：保存路径三处，不涉及任何 XML 格式。

---

## 汇总

**已修复（3 处）**

1. `core.py`：补回时的“先清”不再误删用户页眉/页脚的 behindDoc 图形。
2. `core.py`：source 被占用时，守护仍能补回 output 水印。
3. `gui.py`：closeEvent 覆盖全部 5 类 worker，且总收尾时间守住 6s 兜底。

**风险降低（4 处，均有测试锁定）**

- preserve_existing 不会被回退成从 source 重建（含 output 被删的预期分支）
- 用户正文/表格/图片/段落格式在守护往返后逐字节快照一致
- `clear_any` 的扫描边界（正文/行内/非 behindDoc 不动）被用例固定，防止日后无意扩大
- `_INDEX_MAP` 无残留、与路径表键一致

**仍存在的技术债务（未修，已复现或固化）**

| # | 债务 | 影响 | 处置建议 |
|---|---|---|---|
| D1 | `clear_any` 会删用户自己的 behindDoc 页眉图形（R-1），并可能连带删同 run 文字（R-2） | 误删用户内容 | 加白名单或二次确认；**改公开清除行为，需你定夺** |
| D2 | `font_index_of` 对别名表条目直接 `return 0`，TTC 分面修复对常用字体不生效 | UI/Light 变体渲染成常规字重 | 别名命中后回查 `_INDEX_MAP` |
| D3 | 注册表无 `,N` 时多 face TTC 只能取第 0 面 | 同上（次要） | 按 face 家族名匹配选面 |
| D4 | `build_font_map()` 与 `_MAP_CACHE` 可能半新半旧 | 索引静默退化 0 | 两者配套刷新 |
| D5 | python-docx 非原子落盘 + `os._exit`/`terminate` | 极端路径下输出文件可能写坏 | 原子落盘封装 |
| D6 | COM worker 被强杀时可能残留 WINWORD.EXE | 孤儿进程 | 需 redesign 收尾，暂维持 |

**测试**：本轮 22 个新用例（其中 4 个为 xfail 复现/固化）。最终
`pytest -q` → **71 passed, 4 xfailed, 0 failed**。
