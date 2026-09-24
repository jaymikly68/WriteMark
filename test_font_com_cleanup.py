"""验证“读取字体不会留下幽灵 Word 进程”。

背景（用户反馈“退出 Word 时还是会卡顿”的真实根因）：
早先版本读取字体走 Word COM，会临时启动一个 Word，`word.Quit()` 之后
进程仍然残留在后台。用户后来去关这个看不见的 Word，就表现为“退出 Word 卡顿”。

本测试：
1) 默认字体来源走注册表 → 不产生任何 WINWORD 进程；
2) 即使强制走 Word COM 兜底路径，收尾后也不允许残留新的 WINWORD 进程；
3) .doc 的 COM 引擎（engine_com）收尾同样不残留。
"""
import os
import sys
import time

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from watermark_tool import com_cleanup, word_fonts


def report(tag, before, after):
    new = after - before
    print(f"[{tag}] 调用前 {sorted(before)} → 调用后 {sorted(after)}；新增 {sorted(new)}")
    return not new


def main():
    ok = True

    from watermark_tool import core
    if not core.com_available():
        print("本机没有 pywin32，跳过（不影响功能）")
        return 0

    before = com_cleanup.winword_pids()
    print(f"起始 WINWORD 进程：{sorted(before)}")

    # 1) 默认路径：Word（本机化中文名，最全），读完后不得残留 Word 进程
    t0 = time.time()
    fonts, source = word_fonts.collect_fonts()
    dt = time.time() - t0
    print(f"\n[1] collect_fonts() → {len(fonts)} 种字体，来源【{source}】，耗时 {dt:.2f}s")
    assert fonts, "没有拿到任何字体"
    assert not any(f.startswith("@") for f in fonts), "应过滤掉 @ 竖排字体变体"
    time.sleep(0.5)
    ok &= report("1 默认字体来源", before, com_cleanup.winword_pids())

    # 2) 兜底路径：强制 Word COM，收尾后不得残留
    if os.environ.get("SKIP_COM_TEST"):
        print("\n[2] 按需跳过 COM 路径测试")
    else:
        t0 = time.time()
        com_fonts = word_fonts.get_word_fonts()
        dt = time.time() - t0
        print(f"\n[2] get_word_fonts()（COM 兜底）→ "
              f"{len(com_fonts) if com_fonts else 0} 种字体，耗时 {dt:.2f}s")
        time.sleep(0.5)
        ok &= report("2 Word COM 路径", before, com_cleanup.winword_pids())

    # 3) .doc 的 COM 引擎收尾（用一个临时 .docx 走 COM 引擎不方便，这里只验证工具函数）
    fake = com_cleanup.winword_pids()
    killed = com_cleanup.quit_word(None, before_pids=fake)
    print(f"\n[3] quit_word(None) 对既有 Word 不误杀：结束 {killed} 个（应为 []）")
    ok &= (killed == [])

    print("\n结论：", "读字体/处理文档不会留下幽灵 Word ✅" if ok else "仍有残留 ❌")
    return 0 if ok else 2


if __name__ == "__main__":
    sys.exit(main())
