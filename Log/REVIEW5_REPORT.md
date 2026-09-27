# 第 5 轮审查报告：video / insert 语义 / 字体别名表

基于 commit `84a773e`，本轮范围严格限定为你列出的四项，只做**最小修复 + 补测试**，
不新增功能、不改 GUI 结构、不动 i18n、不改 Word 水印 XML 格式。

**结论速览**

| 项 | 结论 |
|---|---|
| video.py 静默假成功 | ✅ 真实缺陷，已修复 + 4 条新测试（含旧行为反向验证） |
| video.py 0 字节孤儿临时文件 | ✅ 真实缺陷，已修复 + 4 条新测试（含旧行为反向验证） |
| insert_watermark 替换语义 | ⚪ 设计如此，GUI 无风险；新增 5 条测试**钉死行为**，未改代码 |
| `_ALIAS_FILES` 重复 key | ✅ 已清理（值一致，功能无影响）+ 2 条防回归测试 |

全量 `pytest -q` → **107 passed, 4 xfailed** （上一轮 96 + 本轮新增 11，原有测试一条未减）——
其中 1 条 warning 来自既有用例 `test_cjk_fallback_rendering` 的 Pillow 弃用提示，本轮前就有。

---

## 一、video.py：mux 失败 + 改名也失败 → 照样返回成功

### 现状（修复前）

```python
else:
    try:
        if os.path.abspath(tmp_vid) != os.path.abspath(output):
            os.replace(tmp_vid, output)
    except Exception:
        pass          # ← 吞掉
...
return {"ok": True, ...}          # ← 不管改名成功没有
```

`os.replace` 是 mux 降级路径里**唯一**让 `output` 出现的动作。它失败（输出路径被占用 /
只读 / 磁盘满 / 路径过长）时异常被 `except: pass` 吞掉，函数照常返回 `{"ok": True}`。
GUI 侧 `VideoWorker.run()` 拿到的是 `self.result_signal.emit(True, str(res))`，
用户看到「成功」，磁盘上却没有成品。属于**静默假成功**，比直接报错糟糕得多。

### 修复（最小，落在 `video.py` 的 mux 降级分支）

按你给的两种契约选了「抛异常」这一条，与函数里既有风格一致
（`无法读取该视频…`、`未能从视频中读出任何帧…` 都是 `RuntimeError`），
且 GUI 的 `VideoWorker.run()` 本来就有 `except Exception as e: emit(False, str(e))`，
不需要改调用方：

```python
if os.path.abspath(tmp_vid) == os.path.abspath(output):
    raise RuntimeError("音频封装失败，且临时文件与输出路径相同（无法改名兜底）…")
try:
    os.replace(tmp_vid, output)
except Exception as e:
    raise RuntimeError(
        f"音频封装失败，也无法把无声视频改名为输出文件（{e}）："
        f"输出文件未生成，请先确认输出路径可写、磁盘未满。"
    ) from e
```

顺带处理了一个相邻的诊断分支：如果临时文件恰好就是输出路径（旧代码直接跳过改名），
随后 `finally` 会把 `tmp_vid` 删掉，等于「报成功又把成品删了」。现在明确报错。
该路径实际不可达（`mkstemp` 返回随机名），但旧行为是错误的，顺手改掉。

### 测试（`tests/unit/test_video_failure_paths.py`）

| 用例 | 验证什么 |
|---|---|
| `test_mux_fail_and_replace_fail_must_not_return_ok` | 按你要求：**mock `_mux_audio`→False 且 mock `os.replace` 抛异常**，断言不得返回 `ok: True`。契约允许「抛 RuntimeError」或「返回 ok:False」，测试锁的是结果不是实现 |
| `test_mux_fail_replace_fail_error_names_output` | 错误信息必须点明「输出文件未生成」并带上底层原因（`No space left on device`） |
| 上述两条均额外断言 | `output` 不存在、`tmp` 临时文件已被清理 |

### 反向验证（证明测试真能抓 bug）

把 `video.py` 临时改回旧写法后跑同一批测试，全部失败：

```
E   AssertionError: os.replace 改名失败，output 并未生成，却仍返回成功：
    {'ok': True, 'engine': 'video', 'frames': 6, 'output': '…\\out.mp4', ...}
```

---

## 二、video.py：`mkstemp` / `get_writer` 在 try 之外 → 0 字节孤儿

### 现状（修复前）

```python
tmp_fd, tmp_vid = tempfile.mkstemp(suffix=".mp4")   # 建出 0 字节文件
os.close(tmp_fd)
writer = iio.get_writer(...)                        # ← 在 try/finally 之外
try:                                                # ← 真正的 try 从这里才开始
    ...
finally:
    writer.close()                                  # writer 未绑定时会 NameError
    ...
    os.remove(tmp_vid)
```

`iio.get_writer` 自己抛异常（自带的 ffmpeg 二进制缺失/损坏、参数不被接受）时，
那个刚建好的 0 字节 `.mp4` 落在 try 之外，没人清理，永久留在 `%TMP%`。
（异常原因本身**不会**被掩盖——`finally` 里的 `writer.close()` 已被
`except Exception: pass` 包住，我实测确认过，所以注释里没有夸大这一点。）

### 修复（最小）

```python
try:
    writer = iio.get_writer(
        tmp_vid, "ffmpeg", fps=out_fps, macro_block_size=1, quality=None,
        output_params=["-crf", str(crf), "-preset", "veryfast", "-pix_fmt", "yuv420p"],
    )
except Exception:
    try:
        os.remove(tmp_vid)
    except OSError:
        pass
    raise
```

### 测试

| 用例 | 验证什么 |
|---|---|
| `test_get_writer_failure_removes_empty_tmp_and_keeps_reason` | 前置断言「mkstemp 建的确实是 0 字节 `.mp4`」（正是会留下的孤儿）；随后断言异常原因原样上抛、**临时文件已被删除** |
| `test_get_writer_failure_leaves_no_orphan_across_repeats` | 连续两次失败（用户在 ffmpeg 坏掉后反复点导出）后，`%TMP%` 下相对基线**零新增** `.mp4` |

### 反向验证

改回旧写法后：

```
E   AssertionError: get_writer 构造失败后残留 0 字节孤儿临时文件（会堆在 %TMP%）:
    C:\Users\666\AppData\Local\Temp\wm_review_fail_36408.mp4
E   AssertionError: 失败后 %TMP% 下残留了 .mp4: {'wm_review_rep_23336.mp4'}
```

---

## 三、engine_docx.insert_watermark 是「替换语义」

### 按代码确认的实际行为

`engine_docx.insert_watermark(path, kinds, **opts)` 里，对每个页眉/页脚 part 先：

```python
_remove_in_part(part)            # 清掉本工具**所有**已标记的水印图形
```

再按本次 `kinds` 重新插入。**所以：先插 `["text"]`、再单独插 `["image"]`，文字水印会消失。**

### GUI 侧是否有风险：**没有**

`gui.App._gather_kinds()`（gui.py:2475）由 `self.text_enabled` 与 `self.image_enabled`
两个勾选框**整体**决定返回集合，不做增量传参；插入入口（gui.py:2740）也是
`kinds = self._gather_kinds()`。因此 GUI 不存在「只传部分 kinds 就插入」的调用点，
替换语义在 GUI 上表现为「按当前勾选重新生成一份水印」。

> 顺带记一个测试踩过的坑，值得留在代码注释里：`core.insert_watermark` 在
> 「输出路径 == 源文件」时会经 `_protect_original()` 把输出改写为 `xxxWaterMark.docx`，
> 所以验证「在原文件上再插一次」必须直接调 `engine_docx.insert_watermark`。

### 测试（`tests/unit/test_insert_replacement_semantics.py`，5 条，代码未改）

把「陷阱行为」钉死，防止将来有人当成 bug 改掉、或反过来在 GUI 里踩坑：

| 用例 | 钉住什么 |
|---|---|
| `test_insert_text_then_image_only_removes_text_watermark` | 先 `["text"]` 再单独 `["image"]` → 只剩图片水印 |
| `test_insert_both_then_text_only_removes_image_watermark` | 先 `["text","image"]` 再单独 `["text"]` → 只剩文字水印 |
| `test_replacement_is_idempotent_not_accumulating` | 反复重插的份数与单次插入完全一致（清理彻底、不累加） |
| `test_replacement_does_not_touch_body_or_table` | 三轮不同 kinds 重插后，正文（含 `USER_EDIT`）、表格逐格一致 |
| `test_gui_gather_kinds_derives_full_set_from_both_toggles` | GUI 的 kinds 由两个勾选框整体给出，不存在「只传部分」的调用点 |

---

## 四、word_fonts：`_ALIAS_FILES` 重复 key

`新宋体`（原 58/80 行）与 `华文中宋`（原 72/86 行）各定义两次，值完全相同。
字典字面量里后写的静默覆盖先写的，因此**功能无差别**，但风险在于：哪天有人只改了
其中一份，另一份会被无声吃掉，且不会有任何报错。

处理：删掉末尾两条重复项，保留靠前的定义，并在表后加注释说明为什么不允许重复；
同时由 `test_alias_table_has_no_duplicate_keys` 用 **AST 解析源码**把守
（直接读 dict 对象拿不到重复信息——构造时就被丢掉了），并断言去重没有丢失
`新宋体` / `华文中宋` 等字体名；`test_alias_values_unchanged_after_dedup`
断言各别名指向的字体文件未被误改。

---

## 五、本轮改动清单

源码（2 个文件，共约 +30 行，均为最小修改）：

- `watermark_tool/video.py`：`get_writer` 纳入 try 并清理空临时文件；mux 降级分支的
  `os.replace` 失败不再吞掉，改为抛 `RuntimeError`；顺带覆盖「临时文件即输出路径」分支。
- `watermark_tool/word_fonts.py`：删除 `_ALIAS_FILES` 两条重复项 + 注释。

测试（新增 11 条，2 个新文件 + 1 个既有文件扩 2 条）：

- `tests/unit/test_video_failure_paths.py`（新增，4 条）
- `tests/unit/test_insert_replacement_semantics.py`（新增，5 条）
- `tests/unit/test_word_fonts.py`（+2 条重复 key 防回归）

未改动：`engine_docx.py` 的水印 XML、`gui.py`、i18n、清除逻辑的用户行为。

## 六、仍存在的技术债务（本轮未动，沿用前几轮）

| 编号 | 债务 | 影响 | 建议 |
|---|---|---|---|
| D1 | `clear_any` 仍会删用户自己在页眉设成「衬于文字下方」的图片（形状部分已在 review3 F1 收窄） | 误删用户内容 | 加白名单/二次确认，属产品决策 |
| D2 | `_detach_drawing` 若某 `w:r` 内文字与图形并存，文字会一并消失 | 误删页眉文字 | 改为只摘 `w:drawing` 节点 |
| D3 | `font_index_of` 命中 `_ALIAS_FILES` 直接 `return 0`，绕过 TTC 索引 | 微软雅黑 UI/Light 等仍渲染第 0 面 | 别名表项也参与索引判定 |
| D4 | 注册表值无 `,N` 时多 face 字体只能取第 0 面 | 渲染偏细/偏粗 | 按 face 数枚举 |
| D5 | `_INDEX_MAP` 与 `_MAP_CACHE` 可能半新半旧，`font_index_of` 静默退化 0 |  Wrong index | 两者同生共死 |
| D6 | python-docx 非原子落盘（`ZipFile(pkg_file,"w")` 直写目标路径） | `os._exit` / `terminate` 撞上 `save()` 会留下写坏的 .docx | 加临时文件 + `os.replace` 封装 |
| D7 | COM worker 被强杀可能残留 WINWORD.EXE | 用户「退出 Word 卡顿」 | 记录 PID 后清理 |
| D8 | `core.clear_watermark` 是「从 src 复制到 output」语义 | GUI「一键清除」会覆盖用户编辑 | 已由 A4 用例钉住现状 |
| D9 | `watermark_count` 只数总数、不分类别 | 图片水印意外残留时 watchdog 不补回 | 按类型计数 |

**4 条 xfailed 是「已确认但按约定未修」的可执行记录，不代表风险已消失。**
