# Word 水印删除 —— 安全边界修复报告（第 6 轮）

> 归档位置：`Change/Try/CLEAR_SAFETY_REPORT.md`
> 对应仓库文件：`watermark_tool/engine_docx.py`、`watermark_tool/gui.py`、
> `tests/unit/test_clear_safety_scope.py`、`tests/unit/test_clear_gui_decision.py`、
> `tests/unit/test_clear_repeat_and_multisec.py`、`tests/unit/test_clear_any_scope.py`、
> `tests/unit/test_clear_kinds_matrix.py`

## 0. 本轮目标与边界

「一键清除水印」必须满足：

1. 支持三种情形：**只有文字水印 / 只有图片水印 / 两者都有**；
2. **只删本工具自己识别出来的水印**——绝不动用户的正文、普通图片、Logo、形状、页眉页脚内容。

本轮只做两件事：**修掉被实测证明会误删用户的两处代码**，并把安全契约用测试钉死。
其余按约定一律不动（见第 8 节）。

## 1. 修改了哪些生产代码（共 2 处）

### 1.1 `watermark_tool/engine_docx.py` → `_detach_drawing()`

改动前：

```python
if run is not None and run.tag == qn("w:r") and run.getparent() is not None:
    run.getparent().remove(run)      # 整条 w:r 一起摘掉
```

改动后：

```python
run = drawing.getparent()  # w:r
if run is not None and run.tag == qn("w:r"):
    run.remove(drawing)
    # 还有内容（用户文字 / 另一张图 / w:br 等）就必须保留这条 run
    if len(run) == 0 and run.getparent() is not None:
        run.getparent().remove(run)
```

即：**先把目标 `w:drawing` 摘下来，只有该 run 剥光后才连 run 一起删**。
正常水印 run 正是「空 run」，所以常规路径结果与改动前完全一致。

### 1.2 `watermark_tool/gui.py` → `_clear()` 的 `types`/`kinds` 决策块

同时补了 `import logging` / `log = logging.getLogger(__name__)`。

改动前的核心缺陷是「单类型不落 kinds」与「异常兜底成 `kinds = None`」：

```python
if types == {"text", "image"}:
    ...
except Exception:
    kinds = None          # ← 扩大删除范围
```

改动后：

```python
except Exception as e:
    # 检测失败绝不退化成"扩大删除范围"：直接取消本次清除并向用户报告
    log.exception("水印检测失败，已取消本次清除: %s", e)
    QMessageBox.warning(self, "提示",
        "水印检测失败，为避免误删内容，已取消本次清除。\n\n"
        f"错误：{type(e).__name__}: {e}")
    return
if types == {"text", "image"}:
    choice = self._ask_clear_choice()          # 双类型仍弹窗让用户选
    if choice is None:
        return
    kinds = {"text": ["text"], "image": ["image"],
             "both": ["text", "image"]}[choice]
elif types == {"text"}:
    kinds = ["text"]                            # 只有文字水印：只删文字
elif types == {"image"}:
    kinds = ["image"]                           # 只有图片水印：只删图片
# 空集：文档里没有本工具识别到的水印，保持 kinds=None 的"全清"语义
```

双类型分支的行为与改动前一致；变化只发生在**单类型**与**检测异常**两条路径。

## 2. 为什么修改

规格里点名的三个「嫌疑点」**没有靠读代码猜**，而是先写探针脚本实测复现，确认是真实缺陷才动手：

| 嫌疑点 | 实测结果（改动前） |
|---|---|
| GUI `_clear()` 中 `types == {"text"}` 不落 `kinds` → `kinds=None` → `clear_any=True` | 用户自己放在页眉的「衬于文字下方」图片：删除前 = 1 → **删除后 = 0，被删掉** |
| `detect_watermark_types()` 的 `except Exception: kinds = None` | 同上，同样掉进 `clear_any=True`，误删同一张用户图片，且**完全静默、无日志** |
| `_detach_drawing()` 整条 `w:r` 一起摘 | 页眉文字 `['页眉必须保留的说明文字']` → `['']`，**用户文字被连带删除** |

三条都成立，所以是本轮必改项；改动后同样用探针复核，三项的「用户内容」全部存活。

## 3. 修改前存在什么风险

- **会丢用户的文字**：只要水印图形和用户文字被 Word 合并进同一条 `w:r`
  （重新排版、旧版本工具残留都会造成），点一次「清除水印」就永久丢字。
- **会丢用户的图片**：文档只有文字水印、或水印检测抛异常时，
  `clear_any=True` 会把页眉里用户自己衬于文字下方的图片一并清掉。
- **且不可逆、无痕迹**：`_protect_original` 只防止输出覆盖原文件，不提供回滚；
  异常路径又不写日志，用户事后无从判断发生了什么。

## 4. 新增 / 修改了哪些测试

### 4.1 新增 3 个文件（27 个用例）

| 文件 | 用例数 | 覆盖什么 |
|---|---:|---|
| `tests/unit/test_clear_safety_scope.py` | 10 | P0-1~P0-8 的安全契约：三种水印组合的清除矩阵；**P0-6** 用户文字与文字水印同 run 存活；**P0-6b** 与图片水印同 run 存活；**P0-7** 用户 Logo + 衬底图在「只删图片」时存活；**P0-8** 用户衬底图在三种删除选择下全部存活；单类型文档不得吃掉用户衬底图 |
| `tests/unit/test_clear_gui_decision.py` | 9 | 直接驱动 `App._clear()`：单类型只清对应类型、双类型弹窗映射（`text`/`image`/`both`）、用户取消不动作、**检测失败不扩大删除范围**、检测失败时文档未被改动 |
| `tests/unit/test_clear_repeat_and_multisec.py` | 8 | P1：无目标时重复删除、删除→重新插入往返、两类水印交错插入互不干扰、多 Section 下按类型清除的回归 |

断言一律落在**最终文档状态**上（水印标记是否还在、类型是否正确、用户文字/图片是否还在、
文档能否被 python-docx 正常打开、必要的 XML 结构），没有一条是「只 assert 不抛异常」的空测试，
也没有为制造可测性而写假实现。

### 4.2 修改 2 个文件

- `tests/unit/test_clear_any_scope.py`：把原先标记 xfail 的
  「整条 run 包含兄弟文字一起被删」升级为**正式回归断言**
  （`test_clear_any_deletes_whole_run_including_sibling_text`）。
- `tests/unit/test_clear_kinds_matrix.py`：删除用 `inspect.getsource()` + `eval`
  抠 `_clear()` 源码文本的**过度绑定用例**，并移除随之无用的 `import re`；
  真实行为改由 `test_clear_gui_decision.py` 端到端覆盖。

### 4.3 反向验证（证明测试真的有牙）

用 `git stash` 把生产修复临时撤掉，用同一套新测试实跑：

| 测试文件 | 撤掉修复（旧代码） | 恢复修复 |
|---|---|---|
| `test_clear_safety_scope.py` | **2 failed / 8 passed** | 10 passed |
| `test_clear_gui_decision.py` | **3 failed / 6 passed** | 9 passed |

失败的三条正是 `_detach_drawing` 与 `_clear` 的两个缺陷对应的用例，
说明新测试不是「为了绿而写」。

## 5. 测试总数量

全量 `pytest`：

```
$ venv/Scripts/python.exe -m pytest -q
137 collected
```

其中本轮归档范围内的用例：**82 个**（11 个文件）。

## 6. passed / failed / xfailed

```
137 → 134 passed, 3 xfailed, 0 failed, 1 warning in ~63s
```

3 个 xfailed 与上一轮相比减少 1 个：「同 run 兄弟文字连带删除」已修，随之升级为正式断言。

## 7. 是否还有已知风险

1. **检测返回空集时仍走 `kinds=None` 全清**：文档里没有本工具识别到的水印时，
   依然会删除任何 `behindDoc=1` 的图片。这是「清掉别人加的水印」这一既有需求的代价，本轮未再收窄。
2. **`clear_any` 只按 `a:graphicData/@uri == .../picture` 收窄**：用户用 Word 画的形状
   （`wordprocessingShape`）能保住，但用户**插入的图片**若设为「衬于文字下方」仍会被删——当前设计的既定取舍。
3. **API 层语义未变**：`core.clear_watermark(path, kinds=None)` 依旧是「全清」，
   收窄只发生在 GUI 调用侧；直接调用 API 的脚本性用法仍承担上述风险。
4. 全量跑时 `tests/gui/test_close.py` 会打印一行
   `Windows fatal exception: code 0x800706ba`，属既有的 GUI 线程退出噪音
   （该机制按约定保留未改），不影响通过。

## 8. 哪些问题明确没有修改，以及为什么

| 未改动的项 | 原因 |
|---|---|
| watchdog 全部（`watchdog.py`、`preserve_existing`、相关测试、「防 AI/PS 去除」相关讨论） | 本轮范围明确排除 |
| 视频 / TTC 字体 / i18n / COM 日志 / GUI 线程退出机制 | 本轮范围明确排除 |
| `core.clear_watermark` 的 `kinds=None` = 「全清」语义 | 改它会破坏「去掉原本就带水印的 Word」这一既有需求；GUI 侧已收窄，按最小改动原则不动引擎契约 |
| 插入链路（`clear_any=False` 的替换路径） | 行为与改动前一致，未受影响 |
| 第 1 节之外的任何代码 | 「先测后改、最小改动」约束；三个嫌疑点里两个确认成立并修复，另一个经复核只在特定 XML 结构下成立，已用测试钉死 |
| 产物 README、根 README 等文档 | 只归档到本目录，未改动仓库正文文档 |
