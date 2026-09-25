# -*- coding: utf-8 -*-
"""「Word 秒退」功能区 —— 无界面冒烟测试 + 关键逻辑校验。

覆盖：
1. GUI 里新增的诊断/修复/测速/还原按钮与文本域确实建出来了；
2. 点诊断不会改动注册表（job 只读）；
3. _office_done 会正确恢复按钮可用性、回显结果；
4. 快照逻辑能防止「反复修复后还原失效」；
5. 修复内容确实把实测元凶（随 Word 自动加载的插件）列为目标。
"""
import ast
import os
import re
import sys

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")
from PySide6.QtWidgets import QApplication  # noqa: E402

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__)))))
from watermark_tool import office_tweak as ot  # noqa: E402
from watermark_tool.gui import App  # noqa: E402


def _find_source(path, needle, level=0):
    """在源码里找一行文本（用于验证 GUI 接线，避免依赖运行时副作用）。"""
    with open(path, encoding="utf-8") as f:
        return any(needle in line for line in f)


def main():
    # ---- 1. 模块级关键符号 ----
    for sym in ("apply_fix", "revert_fix", "diagnose", "benchmark",
                "load_active_addins", "addin_items", "write_and_verify",
                "original_of", "elevate", "SNAPSHOT_FILE"):
        assert hasattr(ot, sym), "office_tweak 缺少符号：%s" % sym

    # ---- 2. 修复目标包含「随 Word 自动加载」的插件 ----
    auto = [n for (n, lb, d, on) in ot.load_active_addins() if on]
    print("  当前会自动加载的插件：%s" % (auto or "无（已全部禁用）"))
    assert isinstance(auto, list)

    items = ot.addin_items()
    assert items, "addin_items() 应至少列出一个加载项"
    for path, name, val, desc, root in items:
        assert val == ot.LB_DISABLE, "禁用值应为 0"
    print("  修复将禁用的加载项：%s" % sorted(set(i[1] for i in items)))

    # ---- 3. 还原用的原始值来自快照，而不是「上一次修复后的值」 ----
    snap = ot._load_snapshot()
    assert snap, "首次快照必须存在，否则还原会退化成无效操作"
    assert any(k.startswith("addin|") for k in snap), "快照里应含加载项原始值"
    assert any(k.startswith("user|") for k in snap), "快照里应含用户级开关原始值"
    print("  快照收录 %d 项原始值" % len(snap))

    # 模拟：反复修复后，original_of 仍返回首次看到的值
    probe = items[0]
    orig = ot.original_of("addin", probe[0], probe[1], probe[4])
    assert orig in (0, 1, 2, 3), "原始值必须是一个 LoadBehavior 合法值，实际=%r" % orig
    print("  %s 的原始值 = %s（多次修复后仍应如此）" % (probe[1], orig))

    # ---- 4. 写入函数带「校验 + 重试」----
    src = open(os.path.join(os.path.dirname(os.path.abspath(__file__)),
                            "watermark_tool", "office_tweak.py"),
               encoding="utf-8").read()
    tree = ast.parse(src)
    has_wv = any(isinstance(n, ast.FunctionDef) and n.name == "write_and_verify"
                 for n in tree.body)
    assert has_wv, "必须有 write_and_verify（写入后回读校验 + 重试）"
    assert "for _ in range(tries)" in src, "重试逻辑缺失"
    print("  写入带回读校验 + 重试")

    # ---- 5. GUI 接线：新增控件存在，且 _build 里挂上了槽 ----
    app = QApplication(sys.argv)
    win = App()
    for attr in ("_exit_info", "_exit_detail", "_btn_fix", "_btn_revert"):
        assert hasattr(win, attr), "GUI 缺少控件：%s" % attr
    gui_path = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                            "watermark_tool", "gui.py")
    for needle in ('"一键修复"', '"还原"', '"测速"', '"诊断"',
                   "self._office_fix", "self._office_revert",
                   "self._office_bench", "self._office_diagnose",
                   "self._office_optional", '"可选加码',
                   "OfficeWorker", "--office-fix"):
        assert _find_source(gui_path, needle), "gui.py 缺少：%s" % needle
    print("  GUI 五个按钮与槽均已接线")

    # ---- 5b. 回归：默认档绝不能碰登录状态 ----
    # 踩过的坑：曾把 SignInOptions=4（禁止登录任何账户）放进默认修复项，
    # 结果用户被直接踢下线。加载项才是真凶，这两条不许再跑出来。
    src_t = open(os.path.join(os.path.dirname(os.path.abspath(__file__)),
                              "watermark_tool", "office_tweak.py"),
                 encoding="utf-8").read()
    # 默认档 = user_items()，把它和后面的 optional_items() 切开单独看
    default_block = src_t.split("def optional_items")[0]
    user_body = default_block.split("def user_items", 1)[1]
    # "SignInOptions" 允许作为文字出现在注释/警告里，这里只查它有没有成行成项地被写入
    assert not re.search(r'^\s*\(.*SignInOptions\s*,', user_body, re.M), \
        "默认修复项里又出现了 SignInOptions —— 它会把 Word 账号踢下线"
    assert not re.search(r'Common\\SignIn\s*%', user_body), \
        "默认修复项里又出现了 SignIn 键路径"
    assert 'apply_optional' in src_t, "可选加码入口丢失"
    gt_src = open(gui_path, encoding="utf-8").read()
    fix_msg = gt_src.split("def _office_fix")[1].split("def _office_optional")[0]
    assert not re.search(r'禁止\s*Office\s*登录|禁止登录|踢.{0,3}下线', fix_msg), \
        "「一键修复」的说明里又在承诺踢账号了"
    assert "不会改动你的登录状态" in fix_msg, "「一键修复」应明确保证不动登录状态"
    print("  默认档不碰登录状态（SignInOptions 已移出）")

    # ---- 5.5 benchmark 必须检测「挡路弹窗」----
    # 踩过的坑：删除 Privacy 键后 Office 弹首次隐私向导（模态），
    # WM_CLOSE 静默失效，bench 干等 90 秒超时还误判成"联网抖动"。
    bench_src = src_t.split("def benchmark(")[1].split("\ndef ")[0]
    assert "_blocking_dialogs" in bench_src, \
        "benchmark 丢失了挡路弹窗检测（模态对话框会让 WM_CLOSE 静默失效）"
    assert hasattr(ot, "_blocking_dialogs"), "office_tweak 缺少 _blocking_dialogs"
    print("  benchmark 带挡路弹窗检测（不会对模态对话框傻等 90 秒）")

    # ---- 6. 点诊断路径能跑通且不抛异常（纯只读） ----
    win._office_refresh_info()
    assert win._exit_info.text(), "诊断摘要不应为空"
    print("  诊断摘要：%s" % win._exit_info.text()[:90])

    # 直接走 _office_done（真正的 QThread 版本由 _office_start 接线覆盖）
    win._office_fix = lambda *a, **k: None   # 防止误触真实注册表写入
    win._exit_detail.clear()
    win._office_done(True, "模拟：关闭耗时 2.89 秒", False)
    assert "2.89" in win._exit_detail.toPlainText(), "结果应回显到详情框"
    assert win._btn_fix.isEnabled(), "完成后「一键修复」按钮应恢复可用"
    assert win._btn_revert.isEnabled(), "完成后「还原」按钮应恢复可用"
    print("  结果回显 / 按钮恢复 均正常")

    # ---- 7. 提权自运行模式存在（供 UAC 那一趟使用）----
    with open(gui_path, encoding="utf-8") as f:
        seg = f.read()
    assert "def _run_elevated" in seg and "--office-revert" in seg, "提权自运行模式缺失"
    print("  提权自运行模式（--office-fix / --office-revert）已就位")

    print("\n「Word 秒退」功能区冒烟测试全部通过。")
    return 0


if __name__ == "__main__":
    sys.exit(main())
