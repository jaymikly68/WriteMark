# tests/e2e —— 交付物（exe）端到端用例

这一层测的是**打包出来的 `WordWatermark.exe`**，而不是源码模块：启动器解包、
首次/二次启动、关闭退出耗时、exe 内符号是否包含本版本的修复、单实例互斥等。

目前这些检查以**可重复执行的脚本**形式存在（不是 pytest 用例，因为它们需要真实
exe、会占用托盘、耗时较长）：

| 脚本 | 检查内容 |
| --- | --- |
| `tools/verification/test_launcher.py` | 单 exe 交付形态：首次/二次启动、退出耗时 |
| `tools/verification/test_exe_verify.py` | 反解 exe 的 PyZ，校验本版本新增的符号是否真的编进去 |
| `tools/verification/test_exe_no_word.py` | 从启动到退出全程不出现 Word 进程 |

运行方式（先执行 `build.bat` 产出 `dist/WordWatermark.exe`）：

```bash
python tools/verification/test_launcher.py --clean
python tools/verification/test_exe_verify.py
```

等到这些脚本稳定为用例后，再迁到这里改成 pytest 形式（迁移时保持 `--clean` 等参数兼容）。
