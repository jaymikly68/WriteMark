# 验证轮报告：84a773e 的修改是否真的安全

这一轮的任务是**验证**，不是继续加功能。约束：不新增功能、不做架构重构、
不改 `gui.py` 结构、不改 i18n、不改 Word XML 水印格式、不改视频格式；
只分析、最小修复真实风险，不为"代码更漂亮"动代码。

- 本轮改动源码：**仅 `watermark_tool/engine_docx.py` 一处**（14 行，clear_any 收窄）。
- 本轮新增测试：`tests/unit/test_clear_any_scope.py`（+2）、`tests/unit/test_watchdog_preserve.py`（+3）、
  `tests/unit/test_ttc_index.py`（+1）、`tests/unit/test_video_cleanup.py`（+4，新文件），共 +10。
- 最终全量：`pytest -q` → **81 passed, 4 xfailed**（基线 53 → 现 85 条，原有测试一条未减）。

---

## 一、实际发现并修复的问题

### F1（已修复）`clear_watermark()` 会删掉用户自己画的形状

**问题**：`clear_any=True` 分支只凭 `wp:anchor[behindDoc="1"]` 这一个结构特征判定"疑似水印"，
完全不看图形类型。而 Word 保存 2010+ 形状时用的正是
`mc:AlternateContent → mc:Choice → w:drawing → wp:anchor → a:graphicData(uri=…/wordprocessingShape) → wps:wsp`
——也就是说**用户插入的形状在 OOXML 里同样是一个 `w:drawing`**，只要"置于底层"
（形状默认的层次是浮于文字上方，设成衬于文字下方就会写 `behindDoc="1"`），
点一次 GUI「清除水印」就会被删掉。

这是真实的数据丢失，不是边缘情况。复现测试在修复前报 `left=0`（用户形状被清零）。

**最小修复（不改清除主逻辑，只加一个判定条件）**：

```python
def _is_picture(drawing) -> bool:
    """图形是否为「图片」类（a:graphicData/@uri == drawingml/2006/picture）。…"""
    gd = drawing.find(".//" + qn("a:graphicData"))
    return gd is not None and gd.get("uri") == PIC
```
并把判定收窄成：
```python
elif mark is not None or (clear_any and behind and _is_picture(drawing)):
```

只加了 `_is_picture` 这一个与运算。副作用范围清晰：
- `mark is not None` 分支（本工具自己写的水印，以及 `kinds=…` 限定路径）**完全未动**；
- Word 原生水印本身就是图片（`uri=…/drawingml/2006/picture`），**照旧按设计被清除**
  （`test_clear_any_still_removes_word_native_watermark_by_design` 通过）；
- 只有"非图片类图形凭 behindDoc=1 被当作水印"这一条被掐断，正是用户形状的处境。

**验证**：修复后 `test_clear_any_keeps_user_drawing_shape` 通过（形状保留、原生水印仍被清），
且 `test_clear_any_keeps_front_floating_shape` 作为安全对照仍成立。

---

## 二、测试证明没有问题的地方

### 2.1 watchdog + preserve_existing：按你列的场景逐步回归

新增 `test_watchdog_full_user_scenario_end_to_end`，严格按你说的六步：

| 步骤 | 实现 | 断言 |
|---|---|---|
| 1 造带正文的 docx | 正文 + 带格式段落 + 2×2 表格 + 正文图片 | — |
| 2 WriteMark 加水印 | `core.insert_watermark(src, ["text"], output_path=out)` | `watermark_count > 0` |
| 3 修改 output 正文 | 新增一段 + 改一段 + 改单元格 | 快照 `before` |
| 4 模拟用户删水印 | `clear_watermark(out, kinds=["text"])` | `watermark_count == 0` |
| 5 触发 `_reinsert()` | `WatermarkWatchdog(src, out, …)._reinsert()` | — |
| 6 验收 | 见下 | 见下 |

第 6 步逐条断言：水印恢复；**用户新增段落**在；**原正文（被改过 + 未被改动的）**都在；
表格单元格内容与行数未变；页眉行内图片数、页眉/页脚 drawing 数、页脚文字全部未变。

**反向验证（关键）——测试确实有咬合力**：我把 `preserve_existing` 强制改回 `False`
（即旧行为"从 source 重建"）后重跑同一场景，结果
`watermark_count=1` 但"用户在 output 里新增的段落"变成 `False`。
也就是说这条回归测试能在 preserve_existing 被破坏时准确失败，不是摆设。

另外两条：

- `test_watchdog_partial_deletion_reinserts_without_touching_user_edits`：
  平铺模式写入 4 份水印，模拟用户只删掉其中 1 份，走 `_loop` 里 `cnt < expected`
  那条分支补齐，验证补齐后正文/表格零改动、份数回到 4。
- `test_watchdog_reinsert_never_pulls_source_content_into_output`：
  source 与 output 各放一条独有正文，补回后 output 只保留自己的那条，
  source 的内容**不得**进来——这是"watchdog 不能重新从 source.docx 恢复正文"的硬约束。

上轮已修的两条也在本轮复测通过：source 被 Word/WPS 占用时仍能补回；
output 被整体删除时从 source 重建（属设计内行为，由
`test_watchdog_output_deleted_rebuilds_from_source` 固定）。

### 2.2 clear_watermark 的识别边界（只说明，不改代码）

**WriteMark 水印怎么被识别的**（`engine_docx._remove_in_part`）：

1. **扫描范围**：每个 section 的页眉 + 页脚（含首页/偶数页，按 part 去重）里的 `w:drawing`。
   **正文 document body、表格不在范围内**——结构上就触达不到。
2. **命中即删**，两条任一成立：
   - `wp:docPr/@descr` 或 `@name` ∈ {`WB_WATERMARK_TEXT`, `WB_WATERMARK_IMG`}（本工具私有标记）；
   - **或**（仅 `clear_any=True`）图形是 `wp:anchor[behindDoc="1"]` 的**图片**图形（本轮 F1 收窄后）。
3. **删除手法** `_detach_drawing`：整条 `w:r` 一起摘掉，不只摘图形。

**普通 Word 对象为什么不会被识别**（逐条已被测试覆盖）：

| 普通对象 | 为什么安全 | 测试 |
|---|---|---|
| 页眉里的行内图片（最常见 Logo） | 是 `wp:inline`，代码只认 `wp:anchor`，结构性不匹配 | `test_clear_any_leaves_inline_and_front_floating_graphics` |
| 浮动但浮于文字上层的图形 | 没有 `behindDoc=1` | 同上 |
| VML 图形（`w:pict` / `v:shape`，老文档常见） | 根本不在 `w:drawing` 命名空间下 | 同上 |
| **用户自己画的形状（wps / wpg）** | 本轮 F1 后，`a:graphicData/@uri` 非 picture 不再参与 | `test_clear_any_keeps_user_drawing_shape` |
| 正文 / 表格 / 正文图片 | 扫描范围不含 body | `test_clear_any_keeps_body_and_table_intact` |
| GUI「只去文字水印」（`kinds=['text']`） | 走 `only_kinds` 分支，只看标记 | `test_clear_with_kinds_limited_is_safe` |

**仍会误判的两种情况**（= 下文的 D1 / D2，故意保留未改）：
用户自己在页眉放的背景图/Logo（图片 uri + behindDoc=1）；以及同 run 内
"文字 + 图形"并存时文字被连带删掉。

### 2.3 TTC face index：端到端透传已证明

新增 `test_render_text_png_passes_face_index_into_pillow`，把链路钉死：
注册表明文写 `"msyh.ttc,1"` → `build_font_map()` → `font_index_of()` 得 1
→ 拦截 Pillow 的 `ImageFont.truetype` 记录真实入参
→ 断言传给它的 `index=` 就是那个 1、`path` 就是 `msyh.ttc`，文件大小参数正确。

配合已有的 `test_ttc_face_1_loads_when_index_is_supplied`（真实 `msyh.ttc`：
`index=0` → `'Microsoft YaHei'`，`index=1` → `'Microsoft YaHei UI'`，两者确为不同 face），
"字体文件 + face index 确实被正确传给 `ImageFont.truetype(index=…)`"这一条已有直接证据，
不再依赖 `font_index_of` 的返回值本身。

### 2.4 video.py：收尾语义按条验证，结论是所有约定都成立

| 你的问题 | 结论 | 依据 |
|---|---|---|
| 正常完成时 writer 是否正确关闭 | 是。正文循环结束后立刻 `writer.close()` 刷盘 | `test_video_temp_removed_after_success`（真实 writer + 真实 ffmpeg） |
| `_mux_audio()` 是否一定在 writer flush 后 | 是。事件序列实测 `frame … → close → mux → close`，`close` 严格早于 `mux` | `test_writer_closed_before_mux_and_twice_is_safe` |
| 正常成功后临时文件是否删除 | 是。成功分支 `os.remove(tmp_vid)`；mux 失败分支 `os.replace(tmp_vid, output)` | `test_video_temp_removed_after_success`、`test_video_temp_removed_when_mux_falls_back` |
| 中途异常后临时文件是否删除 | 是。`finally` 无条件 `os.remove`，异常照常上抛不被吞 | `test_video_temp_removed_on_midway_exception` |
| `finally` 里重复 `writer.close()` 是否安全 | 安全。writer 是 imageio 的 generator，`close()` 第二次调用是 no-op；且已被 try/except 包住 | 同上 + 事件序列以 `close` 收尾 |

没有为此改动 video.py 的任何一行——现有写法是对的，重复 close 是必要的兜底。

### 2.5 closeEvent：只验证，未动结构

`os._exit(0)` 与 6 秒兜底计时器**原样保留**。本轮确认：`tests/gui/test_close.py`、
`test_close_all_workers.py`、`test_close_confirm.py` 全部通过，即
5 类 worker 都能停机、watchdog 可用 `stop_nowait()` 非阻塞停掉、总收尾时间守住 6s 预算；
没有发现能证实"退出死锁"的新事实，因此按你的要求保持现有结构。

---

## 三、仍存在的技术债务（未修）

| # | 债务 | 影响 | 建议 |
|---|---|---|---|
| D1 | `clear_any` 仍会删用户自己在页眉放的背景图/Logo（图片 uri + behindDoc=1） | 误删用户内容（GUI「清除水印」是常见路径） | 加归属白名单或清除前二次确认；会改变公开清除行为，**需你定夺** |
| D2 | `_detach_drawing` 摘整条 `w:r`，页眉"文字+图形"同 run 时文字一起消失 | 误删文字 | 只摘 `w:drawing`、保留同 run 文字；同样改公开行为，需定夺 |
| D3 | `font_index_of()` 命中别名表（`_ALIAS_FILES`）直接 `return 0`，绕过 TTC 索引 | 微软雅黑/宋体/黑体等常用字体的分面修复不生效 | 别名命中后回查 `_INDEX_MAP` |
| D4 | 注册表无 `,N` 时多 face TTC 只能取第 0 面 | 同上（次要） | 按 face 家族名匹配选面 |
| D5 | `build_font_map()` 刷新 `_INDEX_MAP` 但不刷新 `_MAP_CACHE`，可能半新半旧 | 索引静默退化 0 | 两者配套刷新 |
| D6 | python-docx 非原子落盘（`ZipFile` 直写目标路径）+ `os._exit`/`terminate` | 极端路径下可能留下写坏的半截输出文件 | 加原子落盘封装（临时文件 + `os.replace`） |
| D7 | COM worker 被强杀时可能残留 WINWORD.EXE | 孤儿进程 | 需 redesign 收尾，暂维持 |

> 说明：D1 的范围已因本轮 F1 而收窄——用户**自己画的形状**不再属于 D1，
> D1 现在只覆盖"用户自己的 behindDoc **图片**"。这两个复现用例
> （`test_clear_any_deletes_user_own_behinddoc_header_picture`、
> `test_clear_any_deletes_whole_run_including_sibling_text`）仍保持 xfail 状态，
> 作为"已确认、按约定未修"的证据。

---

## 四、最终测试

```
$ pytest -q
81 passed, 4 xfailed, 1 warning in 39.24s
```

原有测试一条未减（基线 53 → 现 85 条共 85 项，其中 4 项为刻意的 xfail 复现）。
本轮新增的 watchdog / clear_watermark / TTC / video 测试全部通过。

**提醒**：4 个 xfailed 是"已知风险的可执行记录"，不是"测试通过了所以没问题"。
只有把它们变成 XPASS（或删掉），才算真正收口。
