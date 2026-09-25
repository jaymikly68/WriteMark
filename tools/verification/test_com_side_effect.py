import os
"""副作用对照实验：工具通过 COM 启动一次 Word（并干净收尾）之后，
会不会让用户随后关闭 Word 变得更慢？

这一步是判断「用完本工具 → 退出 Word 卡 20 秒」是否成立的关键：
如果工具即使真的用过 Word，之后 Word 的关闭耗时也和平时一样，
那就说明关闭卡顿与工具无关。

顺序：
  A) 先正常测一次 Word 关闭耗时（基线）
  B) 跑一次 COM 读字体（= 旧版工具一定会做的动作），确认收尾干净
  C) 立刻再测 Word 关闭耗时 ← 如果与 A 相当，说明 COM 调用没有副作用
"""
import sys
import time


sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__)))))
from watermark_tool import com_cleanup, word_fonts  # noqa: E402
from test_office_exit_lag import measure  # noqa: E402


def main():
    print("=== A：基线（未做任何 COM 调用）===")
    base = measure("word")

    print("\n=== B：调用 COM 读取 Word 字体（旧版工具的行为）===")
    before = com_cleanup.winword_pids()
    t0 = time.time()
    fonts = word_fonts.get_word_fonts()
    dt = time.time() - t0
    after = com_cleanup.winword_pids()
    print(f"读到 {len(fonts) if fonts else 0} 种字体，耗时 {dt:.2f}s")
    print(f"调用前 WINWORD={before or '无'}   调用后={after or '无'}")
    leftover = after - before
    if leftover:
        print(f"⚠ 收尾后仍有残留 {leftover}")
    else:
        print("✅ 收尾干净：没有留下任何 Word 进程")

    print("\n=== C：COM 调用之后立即复测 Word 关闭耗时 ===")
    time.sleep(2.0)
    after_com = measure("word")

    print("\n=== 结论 ===")
    if base is None or after_com is None:
        print("测量未完成（Word 启动异常），请重试")
        return 1
    print(f"  COM 调用前：{base:.2f}s")
    print(f"  COM 调用后：{after_com:.2f}s")
    diff = after_com - base
    if abs(diff) < 2.0:
        print(f"  差值 {diff:+.2f}s —— 在噪声范围内，"
              f"**COM 调用不影响之后关闭 Word 的耗时**")
        print("  说明：两者都慢的话，慢的是本机 Word 自身，与本工具无关。")
    else:
        print(f"  差值 {diff:+.2f}s —— 存在明显差异，需要进一步排查")
    return 0


if __name__ == "__main__":
    sys.exit(main())
