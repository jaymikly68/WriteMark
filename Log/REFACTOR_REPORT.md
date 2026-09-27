# watermark_tool 行为不变工程化重构 —— 逐模块报告（风险 a–l）

> 原则：先发现问题 → 给风险 → 最小修改 → 跑现有测试 → 下一项。未一次性重写、未删测试、
> 未改 Word 水印 XML 格式 / 清除逻辑 / 视频输出逻辑。
> 全部变更已本地提交 `84a773e`（未 push）。全量测试：`pytest -q` → **53 passed**（原 45 + 新 i18n 8），无回归。

---

## #48 core.py
1. **修改**：`get_desktop()` 的 `except Exception: pass` 收窄为 `except (OSError, ValueError) as e: log.debug(...)`；`insert_watermark()` 新增参数 `preserve_existing: bool = False`，为 True 且输出已存在时改为"在现有输出上 clear 再 insert"，默认 False。
2. **为什么**：(e) 裸 except 会吞掉真实路径错误；为修复 #49 的 watchdog 覆盖问题预留开关。
3. **行为改变**：否。默认 `preserve_existing=False` 与旧行为逐字一致，GUI 未传该参数。
4. **测试**：`pytest -q` 全量 + 相关单测。
5. **结果**：53 passed。
6. **已知风险**：`preserve_existing` 路径复用下游 clear/insert，依赖 engine 的 clear 逻辑（见 #50 h）；`get_desktop` 仅 `log.debug`，排查需开启 logging。

## #49 watchdog.py
1. **修改**：`_reinsert()` 调用 `core.insert_watermark(..., preserve_existing=True, **self.opts)`。
2. **为什么**：(f) 修复 watchdog 在用户删除/部分删除水印后，从源文件重生成输出、覆盖用户对输出正文改动的问题；现仅补回水印、保留正文。
3. **行为改变**：是（仅 watchdog 重插路径，且更正确）。正常插入流程与用户编辑后补回水印的语义均更合理。
4. **测试**：`pytest -q` 全量。
5. **结果**：53 passed。
6. **已知风险**：依赖 engine 的 clear 不过度删除（h 未动）；watchdog 对 docx/com 路径生效。

## #50 engine_docx.py
1. **修改**：`_add_drawing_to_part()` 中 `r._r.append(drawing)` 包进 `try/except AttributeError`，抛 `RuntimeError("python-docx 内部结构变动…")`；新增 docstring 说明私有 API 依赖。
2. **为什么**：(g) `r._r` 是 python-docx 私有 API，版本升级若变动，旧代码抛裸 AttributeError 难定位，改为可读异常。
3. **行为改变**：否。成功路径与生成 XML 格式完全不变，仅异常信息更清晰。
4. **测试**：`pytest -q` 全量。
5. **结果**：53 passed。
6. **已知风险**：(g) 根本风险仍在——依赖私有 API；升级 python-docx 仍需回归。(h) `clear_any` 清除逻辑按用户约束未动，仅注释说明。

## #51 engine_com.py
1. **修改**：加 `import logging` + `log`；约 19 处 `except Exception: pass` 改为 `log.warning`（_tag 的 Name/AlternativeText、图片透明度、居中、定位、删除水印）或 `log.debug`（_header_indices、_is_ours、字体回退、名称读取、_iter_doc_parts）。`finally` 中 `doc.Close` 清理块保留裸 except（由 `_close_word` 兜底）。
2. **为什么**：(i) COM 错误被静默吞掉，导致"没加上/没删掉"却无提示。
3. **行为改变**：否。功能路径不变，仅把被吞错误显式记录便于排查。
4. **测试**：`pytest -q` 全量。
5. **结果**：53 passed。
6. **已知风险**：`finally` 中 `doc.Close` 仍裸 except（故意，避免清理异常干扰退出）；warning 需配置 logging 输出才可见。

## #52 video.py
1. **修改**：帧循环后，在已有 `try` 内、其 `finally` 之前先 `writer.close()` 再 `_mux_audio`/`os.replace`/`return`；`finally` 统一清理临时 mp4（未被改名成品、也未被 mux 删除时删除）。
2. **为什么**：(j) 修复回归——此前把 mux 移到 finally 前导致读到未 flush 的临时文件（5 测试失败），经"先 close 再 mux"修复；并彻底解决异常路径下 `mkstemp` 临时 mp4 孤立残留。
3. **行为改变**：否（功能性）。输出视频内容/时长/格式逻辑不变，仅异常路径不留垃圾、mux 读已刷盘完整文件。
4. **测试**：`pytest tests/unit/test_video.py`（12 例）+ 全量。
5. **结果**：video 12 passed，全量 53 passed（此前 5 failed 已修复）。
6. **已知风险**：清理依赖 finally 执行；若在 `writer.close()` 前被 `os._exit` 强杀仍可能残留（极端）。

## #53 word_fonts.py
1. **修改**：`build_font_map()` 不再 `split(",")[0]` 丢弃 TTC 索引，写入新模块级 `_INDEX_MAP[name]=idx`（默认 0）；新增 `font_index_of(font_name)` 按与 `resolve_font_path` 相同顺序解析返回索引。
2. **为什么**：(k) 多 face TTC（如"微软雅黑 Light"=msyh.ttc 第 1 面）此前一律加载第 0 面，导致 Light 变 Regular / 字符错位。
3. **行为改变**：是（仅多 face TTC 的加载正确性）。非 TTC 字体 index=0 与旧一致；`engine_docx.render_text_png` 的 `_load` 改用 `ImageFont.truetype(p, size, index=...)`，index 默认 0，既有行为不变。
4. **测试**：`pytest tests/unit/test_word_fonts.py`（含 CJK 回退渲染）+ 全量。
5. **结果**：53 passed。
6. **已知风险**：`_INDEX_MAP` 仅在 `build_font_map` 调用后填充；未 build 时 `font_index_of` 返回 0（安全默认）。

## #54 tests/unit/test_i18n.py（新增）
1. **修改**：新增 8 个不变式测试（tr_zh_cn_is_noop、tr_en_translates、trf_placeholder_and_fallback、fallback_targets_valid、translation_keys_subset_of_en、resolve_ui_font_returns_str、set_lang_unknown_falls_back、has_and_missing）。
2. **为什么**：(l) 将"中文 UI 文案作 key"契约与回退语义锁成测试，捕捉某语种中文 key 拼写错误导致整条翻译失效（静默漂移）的隐患。零运行时代码改动。
3. **行为改变**：否。纯新增测试。
4. **测试**：本文件 8 例 + 全量。
5. **结果**：8 passed，全量 53 passed。
6. **已知风险**：无。误改 i18n 结构/拼错 key 时由本测试拦截，正是其目的。

## #55 gui.py（核查，未修改）
1. **修改**：无代码修改。
2. **为什么**：
   - (b) 五类 worker **均持有持久 `self._xxx` 引用**：`self._worker`(2489)、`self._font_worker`(2362)、`self._v_preview_worker`(1860)、`self._v_worker`(1943)、`self._office_worker`(1999)。先前"VideoWorker/OfficeWorker 未持有"系 grep 误读（赋值在下一行 / 经 `_office_start` 持有）。**无 GC 崩溃风险。**
   - (c) `closeEvent` 默认最小化到后台（守护继续）；真正退出时非阻塞停守护、线程 graceful wait 后超时强杀；6s `os._exit(0)` 兜底且用 `PYTEST_CURRENT_TEST` 守卫以免杀测试进程。有意为之的安全网，非 bug。
   - (a)/(d) 过度膨胀与 self.xxx 数量：属代码风格/结构，非实际风险。
   - (e) 剩余裸 `except Exception: pass` 全部位于非关键 best-effort 路径（QSettings/QLocale 取语言、i18n 控件扫描、托盘 hide/showMessage、窗口唤出、PySide6.QtNetwork 可选导入 fail-safe、日志写入/规格串/ensureWidgetVisible 兜底），不隐藏用户需知晓的失败。
   按用户"如果只是代码风格而非实际风险，不要为了代码看起来漂亮而修改"的原则，不做改动。
3. **行为改变**：否（无修改）。
4. **测试**：`pytest -q` 全量。
5. **结果**：53 passed，无回归（证明未引入破坏）。
6. **已知风险**：(a)/(d) 若要拆分 gui.py 可单列任务；其余核对为无风险。

---

## 收尾结论
- 已处理：核心实际风险 f/g/i/j/k/l（watchdog 覆盖、私有 API、COM 吞错、视频临时文件、TTC 索引、i18n 契约）。
- 经核查未改（无实际风险/纯风格）：a/b/c/d/e（gui.py 膨胀、worker 生命周期、closeEvent 兜底、self.xxx 数量、非关键裸 except）。
- 用户约束内未动：`clear_any` 清除逻辑（h）——已注释说明，不属本次范围。
- 测试：全量 **53 passed**，无回归；已本地提交 `84a773e`（未 push）。
