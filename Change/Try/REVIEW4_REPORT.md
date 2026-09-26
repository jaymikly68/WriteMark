# 本轮交付：按类型区分的水印删除回归测试

范围：在 `84a773e` 之后，补充「删除文字水印 / 删除图片水印 / 删除全部水印」三类入口的真实回归测试，
并验证 watchdog 对这三种删除的响应。**未修改任何产品代码**（`tests/` 之外只有既有未提交的改动）。

结论先行：

- 三类删除入口按代码确认**各自调用不同的逻辑**，行为正确，未发现新的误删风险。
- watchdog 对三种删除的响应正确，且**不回退成从 source 恢复正文**。
- 用 `84a773e~1` 的旧代码实跑同一套测试：**9 个用例失败**，证明这些测试确实能抓住实际问题，
  其中两条失败直接对应 84a773e 修复的两个真实 bug。
- 新增 15 个用例；完整 `pytest -q` → **96 passed, 4 xfailed**（原有测试一条未减）。

---

## 一、按当前代码确认：水印如何识别、三个删除入口走哪条逻辑

以下全部来自当前代码，非按函数名猜测。

### 1.1 识别方式

| 概念 | 代码依据 | 判据 |
|---|---|---|
| 文字水印 | `engine_docx.MARK_TEXT = "WB_WATERMARK_TEXT"` | 浮动图形的 `wp:docPr/@descr`（兼容旧版 `@name`）等于该标记 —— 见 `_mark_of()` |
| 图片水印 | `engine_docx.MARK_IMG = "WB_WATERMARK_IMG"` | 同上 |
| 类型汇总 | `detect_watermark_types()` | 遍历所有页眉/页脚 part，按 `MARK_TEXT` / `MARK_IMG` 归类，返回 `{"text"}` / `{"image"}` / 两者 / 空集 |
| 份数 | `count_watermarks()` → `_count_marks_in_doc()` | 带私有标记的图形**总数**（**不区分类型**，见 1.4） |

关键：识别**只看私有标记**，不看 `behindDoc`、不看 `a:graphicData/@uri`。
两者是"与"关系上的补充：只有 `clear_any` 这条路径才额外看结构特征。

### 1.2 三个删除入口的实际调用链

```
GUI「一键清除水印」  ->  gui._clear()
    └─ core._use_com(file)? 否
       └─ engine_docx.detect_watermark_types(file)
          ├─ 结果恰为 {"text","image"}  -> 弹 _ask_clear_choice() 三选一
          │     "去除文字水印"     -> kinds = ["text"]
          │     "去除图片水印"     -> kinds = ["image"]
          │     "文字和图片都去除" -> kinds = ["text","image"]
          └─ 其它（只有一种或无）  -> kinds = None
       └─ core.clear_watermark(src, output_path=out, kinds=kinds)
             └─ shutil.copy2(src, out)          # 「不破坏原文件」语义
                └─ engine_docx.clear_watermark(out, kinds=kinds)
                       ├─ kinds 非空  -> only_kinds 分支：只删 mark ∈ kind_marks 的图形
                       └─ kinds 为空  -> _remove_in_part(clear_any=True)
                                        删「本工具标记的」+「behindDoc 的**图片**」
```

也就是说：

| 用户选择 | 实际逻辑 | 会删掉什么 |
|---|---|---|
| 删除文字水印 | `only_kinds` 分支，`kind_marks={MARK_TEXT}` | 只删文字水印 |
| 删除图片水印 | `only_kinds` 分支，`kind_marks={MARK_IMG}` | 只删图片水印 |
| 删除全部水印 | `kinds=["text","image"]` 或 `None` | 前者删两类标记；`None` 走 `clear_any`，另加删 Word 原生水印 |

### 1.3 watchdog 的清除逻辑

```
watdog._loop()  ->  core.watermark_count(out)
                     ├== 0            -> _reinsert()
                     └< expected      -> _reinsert()
                        └ core.insert_watermark(src, kinds, output_path=out,
                                                 preserve_existing=True)
                           └ reuse_existing = preserve_existing and out 存在
                              → engine_docx.clear_watermark(out, kinds=list(kinds))   ★
                                engine_docx.insert_watermark(out, kinds, **opts)
                           └ (仅当 out 已被整体删除时才 shutil.copy2(src, out))
```

★ 这一行是重点：`_reinsert` 的"先清"用的是 **kinds 限定清除**，不是 `clear_any`。
所以 watchdog 补回时，**绝不会**连带删掉 Word 原生水印或用户在页眉里的 behindDoc 图形。

### 1.4 一条必须写清楚的实现事实

`watermark_count()` 统计的是**总份数，不区分类型**。因此 watchdog 的判据是：

- 两类水印共 2 份 → 用户删掉文字水印 → 剩 1 份 → `cnt < expected` → 触发补回 ✓
- 但若只启用文字水印（expected=1），图片水印是被意外留在文档里的，watchdog 认为"在位"，不补回

本轮测试按第一种（真实且常见）建模。第二种作为技术债务 D9 记录。

---

## 二、新增的测试（15 个，文件：`tests/unit/test_clear_kinds_matrix.py`）

素材统一为「正文 + 表格 + 文字水印 + 图片水印」的 `output.docx`，并在正文中注入唯一字符串
`USER_EDIT_12345`；页眉额外放一个用户自己的 `wps` 形状（衬于文字下方），用于盯 `clear_any`。
watchdog 用**真实线程**触发（`wd.start()` → 等 1.2s → `wd.stop()`），不是测试自己调 `_reinsert()` 假装触发过。

### A 组：按代码确认识别与入口（4 个，`test_*` 名见下）

| 用例 | 验证什么 |
|---|---|
| `test_text_and_image_watermarks_are_distinguished_by_private_marks` | 两类水印靠 `MARK_TEXT` / `MARK_IMG` 区分；`detect_watermark_types` 正确返回两者；每个本工具图形都带私有标记 |
| `test_three_user_clear_choices_map_to_three_distinct_kinds` | 从 `gui.App._clear` 源码里抠出 kinds 映射表求值——三个按钮必须映射到三种不同 kinds；且只有"文档同时含两类"时才弹窗 |
| `test_three_clear_entries_remove_different_watermark_sets` | 三个入口实跑：只删文字 → 只剩 `{"image"}`；只删图片 → 只剩 `{"text"}`；删全部 → `set()` |
| `test_core_clear_entry_rebuilds_output_from_source` | 文档化 `core.clear_watermark` 的"不破坏原文件"语义会从 src 重建 output（用户编辑被覆盖）——见 D8 |

### B 组：watchdog 三场景（3 个，分别验证，不合并）

| 用例 | 验证什么 |
|---|---|
| `test_scenario_delete_text_watermark_only` | 场景 1：删文字后 `types={"image"}`、USER_EDIT 在；watchdog 触发后 `types={"text","image"}`、USER_EDIT/原正文/表格全在 |
| `test_scenario_delete_image_watermark_only` | 场景 2：删图片后 `types={"text"}`；watchdog 后两类都在、USER_EDIT 在、表格未变 |
| `test_scenario_delete_all_watermarks` | 场景 3：删全部后 `types=set()`；watchdog 后两类都补回，且**正文逐字比对**与删除前完全一致 |

### C 组：模拟"用户在 Word 里直接删除"（3 个，不调用 WriteMark 的任何删除 API）

| 用例 | 验证什么 |
|---|---|
| `test_user_manually_deletes_text_watermark_via_xml` | 直接摘掉 `MARK_TEXT` 图形（总数 2→1）→ watchdog 补回文字，图片水印与正文不受影响 |
| `test_user_manually_deletes_image_watermark_via_xml` | 直接摘掉 `MARK_IMG` → watchdog 补回图片，文字水印与正文不受影响 |
| `test_user_manually_deletes_both_watermarks_via_xml` | 两类全摘 → watchdog 全量补回，USER_EDIT 与原正文都在 |

### D 组：`clear_watermark` / `clear_any` 的误删边界（4 个）

| 用例 | 验证什么 |
|---|---|
| `test_clear_text_kind_does_not_touch_image_watermark_or_user_objects` | 只删文字：图片水印保留；正文/表格/用户页眉形状零变化 |
| `test_clear_image_kind_does_not_touch_text_watermark_or_user_objects` | 只删图片：文字水印保留；正文/表格/用户页眉形状零变化 |
| `test_clear_all_preserves_body_table_user_picture_shape_and_header_footer` | 两种"全清"入口（`kinds=["text","image"]` 与 `kinds=None`→clear_any）分别跑：正文、表格、用户正文图片、页眉普通图片、页脚文字、**用户 wps 形状**全部保留 |
| `test_watchdog_scenarios_do_not_touch_user_assets` | 把三场景各跑一遍，逐一核对用户资产（表格/形状/USER_EDIT）零变化 |

### E 组：反向验证（1 个）

| 用例 | 验证什么 |
|---|---|
| `test_legacy_reinsert_without_preserve_existing_loses_user_edit` | 用 monkeypatch 把 `core.insert_watermark` 强制退回旧行为，断言"用户编辑确实会丢失"，并断言现行行为保住它 |

---

## 三、哪些测试证明 84a773e 修了真问题：用旧代码实跑对照

这不是推断——把 `84a773e~1` 的代码导出到独立目录（`git archive`，不动工作区），
在同一套测试文件上运行：

```
git archive 84a773e~1 watermark_tool | tar -x -C /tmp/wm_legacy
cd /tmp/wm_legacy && <项目 venv>/python -m pytest -q test_clear_kinds_matrix.py
```

旧包确认生效（打印 `core.py` 路径为临时目录，`insert_watermark` 无 `preserve_existing` 参数）。

**旧版结果：9 failed, 6 passed。**

两条最有代表性的失败：

1. `test_scenario_delete_text_watermark_only`
   ```
   AssertionError: watchdog 补回后用户新增的正文必须还在
   assert 'USER_EDIT_12345' in ['原始正文第一段', '原始正文第二段']
   ```
   —— 旧版 `_reinsert` 从 source 覆盖 output，用户新增的正文被整份抹掉。这正是 84a773e
   的 `preserve_existing` 所修的问题，现在由 B1 常驻守门。

2. `test_clear_all_preserves_...`
   ```
   AssertionError: [kinds=None(clear_any)] 全清时用户自己的页眉形状被误删了
   assert 0 == 1
   ```
   —— 旧版 `clear_any` 只看 `behindDoc`，会删掉用户页眉里的 wps 形状。这正是
   review3（F1）加 `_is_picture()` 收窄所修的问题，现在由 D3 常驻守门。

其余失败：`test_scenario_delete_image_watermark_only`、`test_scenario_delete_all_watermarks`、
C1/C2/C3、D4、E1 —— 全部源于上述同一条根因（source 覆盖 / clear_any 过宽）。

**同一套测试在当前代码上：15 passed。**

---

## 四、只是"正常行为"测试的那些

以下用例在本轮**没有**发现缺陷，只是把正确行为钉住，防止将来改坏：

- A1、A2 —— 识别常量与 GUI 映射表（A2 直接对源码求值，改映射会立刻红）
- A3 —— 三个入口的删除范围
- A4 —— `core.clear_watermark` 从 src 重建 output 的既有语义（属 D8，非本轮修）
- B2、B3 —— 与 B1 同源，同为正确行为
- C3 —— 用户删两类后补回
- D1、D2 —— 按类型清除互不干扰

也就是说：**本轮没有发现新的误删风险**；`clear_watermark` / `clear_any` 的已知误删面
就是 review2/review3 已记录并（部分）修复的那几项，见下。

---

## 五、`clear_watermark` / `clear_any` 的误删风险现状

| 项 | 状态 | 依据 |
|---|---|---|
| 删用户自己的页眉 behindDoc **图片** | **仍存在**（review2 的 R-1） | `clear_any` 只认"behindDoc 的**图片**"；用户在页眉设"衬于文字下方"的背景图/Logo 会被删。修法会改变公开清除行为，未擅自改。本轮 C/D 组刻意**不放置**这类对象，避免在"设计行为"上再加断言 |
| 删同 run 内的兄弟文字（R-2） | **仍存在**（review2 的 R-2） | `_detach_drawing` 摘整条 `w:r`。未改 |
| 删用户自己的 `wps` 形状 | **已修复**（review3 的 F1） | `_is_picture()` 收窄；本轮 D3 用真实 wps 形状回归，旧版会红 |
| 删 Word 原生水印 | **按设计保留** | `clear_any` 的目的就是清掉 Word 原生水印 |

本轮额外确认（无风险）：`only_kinds` 分支（`kinds=["text"]` / `["image"]`）**完全不看**
`behindDoc`、不看 `a:graphicData/@uri`，因此它绝不会碰到 Word 原生水印，也不会碰到用户对象。
watchdog 补回用的正是这条分支，这是"补回不误伤"的根本原因。

---

## 六、测试运行结果

```
$ pytest -q
96 passed, 4 xfailed, 1 warning in 57.28s
```

- 原有测试：**一条未减**（基线 81 passed + 4 xfailed → 现在 96 passed + 4 xfailed）
- 新增：15 个（本文件的 15 个用例）
- 4 个 xfailed 仍是 review2 记录的"已知风险的可执行记录"，不是"测试通过即无风险"

---

## 七、仍存在的技术债务（本轮未修，部分为既有）

| 编号 | 问题 | 影响 | 建议 |
|---|---|---|---|
| D8 | `core.clear_watermark(src, output_path=out)` 先把 src 复制成 out 再清除；GUI「一键清除」走的就是这条 | 用户在 output 上改过正文后再点清除，**编辑会被 src 覆盖**（已由 A4 钉住） | 改语义属产品决策；至少应在 GUI 提示"清除会覆盖输出文件" |
| D9 | `watermark_count()` 只数总数、不分类别 | 若用户只启用文字水印，图片水印意外留在文档里时 watchdog 不补回（总份数没变） | 守护判据改为"各类份数"，会牵动 `_loop` |
| D1 | `clear_any` 会删用户自己的 behindDoc 页眉图片 | 误删用户内容 | 加白名单或二次确认 |
| D2 | `_detach_drawing` 可能连带删同 run 文字 | 误删文字 | 改为只摘 drawing 节点 |
| D3 | `font_index_of` 被 `_ALIAS_FILES` 短路，TTC 分面对常用字体不生效 | 渲染成错误的字面 | 别名表命中后回落到 `_INDEX_MAP` |
| D4 | 注册表无 `,N` 时多 face TTC 只能取第 0 面 | 中文字形偏差 | 解析 `,N` 备选 |
| D5 | `_INDEX_MAP` 与 `_MAP_CACHE` 半新半旧时静默退化 0 | 字体错配 | 统一失效粒度 |
| D6 | python-docx 非原子落盘（`ZipFile(...,"w")` 直写目标） | `os._exit` / `terminate` 撞上 `save()` 会留下半截 .docx | 加临时文件 + `os.replace` 封装 |
| D7 | COM worker 被 `terminate()` 强杀可能残留 WINWORD.EXE | 进程残留 | 拉起前记录 pid，退出后复核 |

D1–D7 与 review2/review3 登记一致，本轮未改动。
