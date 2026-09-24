"""证明「本工具处理 .docx 的整个过程中不会产生任何 Word 进程」。

这是判断「用完工具 → 退出 Word 变慢」是否真有因果关系的关键：
如果工具从来不曾启动过 WINWORD.EXE，那 Word 的关闭耗时与本工具在机制上就没有接触点。

⚠ 关键点：**必须持续采样，不能只在前后各查一次。**
Word 即使被拉起后几秒内又退出，事后检查也看不出差别 —— 本工具的早期版本正是
"事后看起来很干净"地掩盖了它每次都会启动一个 Word 的事实。
因此这里开一条监控线程，以 50ms 间隔记录整个测试期间出现过的所有 WINWORD 进程。

覆盖路径：
  1. 读取字体列表（GUI 询问后的那条路径）
  2. 生成文字水印图片 + 插入水印（文字 / 图片 / 平铺 + 冗余）
  3. 统计水印份数 + 清除水印
  4. 守护式重复写入
"""
import os
import shutil
import subprocess
import sys
import tempfile
import threading
import time

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from watermark_tool import core, engine_docx, word_fonts  # noqa: E402
from test_office_exit_lag import app_path, pids_named  # noqa: E402


class WordWatcher(threading.Thread):
    """按固定间隔轮询，记录期间出现过的所有 WINWORD 进程快照。"""

    def __init__(self, interval=0.05):
        super().__init__(daemon=True)
        self.interval = interval
        self._stop = threading.Event()
        self.sightings = []          # [(相对时间, pid集合)]
        self.t0 = None

    def run(self):
        self.t0 = time.time()
        prev = frozenset()
        while not self._stop.is_set():
            now = frozenset(pids_named("WINWORD.EXE"))
            if now != prev:
                self.sightings.append((round(time.time() - self.t0, 3), set(now)))
                prev = now
            time.sleep(self.interval)

    def stop(self):
        self._stop.set()
        self.join(timeout=2)


def check(label, watcher, baseline):
    appeared = set()
    for _, pids in watcher.sightings:
        appeared |= (pids - baseline)
    if appeared:
        print(f"  ❌ {label}: 期间出现 WINWORD 进程 {appeared}")
        return False
    print(f"  ✅ {label}: 全程未见任何 WINWORD 进程")
    return True


def make_sample_doc(path):
    import docx
    d = docx.Document()
    d.add_paragraph("测试文档：用于验证水印工具是否启动 Word 进程。")
    d.save(path)
    return path


def main():
    baseline = pids_named("WINWORD.EXE")
    print("=== 前置：当前 WINWORD 进程 ===")
    print(" ", baseline or "无")

    watcher = WordWatcher()
    watcher.start()
    ok = True
    tmp = tempfile.mkdtemp(prefix="_nwp_")
    try:
        doc = make_sample_doc(os.path.join(tmp, "sample.docx"))

        # 1) 字体来源 —— 历史上这里会 COM 启动 Word
        print("\n[1] 获取字体列表")
        t0 = time.time()
        fonts, source = word_fonts.collect_fonts()
        print(f"  来源={source}，{len(fonts)} 种，耗时 {time.time()-t0:.3f}s")
        print(f"  中文字体名可用：{[x for x in fonts if '一' <= x[0] <= '鿿']}")
        ok = check("collect_fonts()", watcher, baseline) and ok
        for cn in ("宋体", "微软雅黑", "仿宋", "等线"):
            p = word_fonts.resolve_font_path(cn)
            if not p or not os.path.exists(p):
                print(f"  ❌ 中文字体 {cn} 解析到无效路径：{p}")
                ok = False
        print("  ✅ 中文字体名均可解析到实际字体文件")

        # 2) 插入水印
        print("\n[2] 插入水印（默认加固参数：平铺 + 冗余）")
        img_pil = engine_docx.render_text_png(
            "机密C", 48, "#FF0000", 40, cn_font_name="仿宋",
            latin_font_name="Times New Roman")
        img = os.path.join(tmp, "wm.png")
        img_pil.save(img)
        out1 = os.path.join(tmp, "out1.docx")
        # 结构与 GUI 的 _gather_opts() 完全一致
        opts = {
            "text": {"text": "机密C", "cn_font_name": "仿宋",
                     "latin_font_name": "Times New Roman", "font_size": 48,
                     "color": "#FF0000", "angle": 45, "transparency": 60,
                     "scale": 100, "offset_x": 0, "offset_y": 0},
            "image": {"image_path": img, "angle": 0, "transparency": 60,
                      "scale": 60, "offset_x": 0, "offset_y": 0},
            "text_enabled": True, "image_enabled": True,
            "tile": True, "tile_rows": 5, "tile_cols": 2, "redundant": True,
        }
        kinds = ["text", "image"]
        # 注意：core.insert_watermark 在「输出路径 == 源路径」时会自动改名以保护原文件，
        # 真实输出位置必须取返回值的 output，不能想当然当成 out1。
        res = core.insert_watermark(doc, kinds, out1, **opts)
        out1 = res.get("output", out1)
        print(f"  输出文件：{os.path.basename(out1)}")
        ok = check("insert_watermark(平铺+冗余)", watcher, baseline) and ok

        # 3) 统计 + 清除
        print("\n[3] 统计与清除水印")
        cnt = core.watermark_count(out1)
        print(f"  文档中水印份数={cnt}")
        if cnt <= 0:
            print("  ❌ 插入后份数为 0，水印可能并未写入")
            ok = False
        # 必须显式给出输出端：不给的话会按「另存保护原文件」规则写到桌面上
        out2 = os.path.join(tmp, "out2_cleared.docx")
        res = core.clear_watermark(out1, output_path=out2)
        out2 = res.get("output", out2) if isinstance(res, dict) else out2
        cnt2 = core.watermark_count(out2)
        print(f"  清除后份数={cnt2}  输出={os.path.basename(out2)}")
        if cnt2 != 0:
            print("  ❌ 清除后仍有残留水印")
            ok = False
        ok = check("watermark_count + clear_watermark", watcher, baseline) and ok

        # 4) 守护式连续写入
        print("\n[4] 模拟后台守护连续 3 次检查/补回")
        for _ in range(3):
            core.insert_watermark(out1, kinds, out1, **opts)
            core.watermark_count(out1)
        ok = check("守护循环 x3", watcher, baseline) and ok

        # 5) 监控有效性自检
        # 必须证明监控确实"抓得住" Word，否则前面那条"未见任何 WINWORD"可能只是
        # 采样没生效的空结论（早期版本就因为是事后单点检查而漏掉了瞬时启动的 Word）。
        canary_mark = len(watcher.sightings)      # 自检之前的采样到此为止
        print("\n[5] 自检：人为起一个 Word，验证监控器抓得住")
        exe = app_path("WINWORD.EXE")
        if not exe:
            print("  本机未安装 Word，跳过自检")
        else:
            p = subprocess.Popen([exe], close_fds=True)
            t0 = time.time()
            canary = None
            while time.time() - t0 < 30:
                new_pid = pids_named("WINWORD.EXE") - baseline
                if new_pid:
                    canary = new_pid
                    break
                time.sleep(0.1)
            time.sleep(0.4)
            caught = any(pids - baseline for _, pids in watcher.sightings)
            left = []
            if canary:
                try:
                    p.kill()
                except Exception:
                    pass
                time.sleep(2.0)
                left = sorted(pids_named("WINWORD.EXE") - baseline)
            print(f"  人为启动 pid={sorted(canary) if canary else None}，"
                  f"监控是否记录={caught}，结束后残留={left}")
            if not caught:
                print("  ❌ 监控器没能抓到自己启动的 Word —— 前面的结论不成立！")
                ok = False
            else:
                print("  ✅ 监控器工作正常，前面的“全程未见 Word 进程”是可信结论")
    finally:
        watcher.stop()
        shutil.rmtree(tmp, ignore_errors=True)

    # 判定只看「自检之前」的采样，避免把自检自己造的那个 Word 算进工具账上
    appeared = set()
    for _, pids in watcher.sightings[:canary_mark]:
        appeared |= (pids - baseline)
    print(f"\n=== 结果：{'通过 ✅ 全流程从未出现 Word 进程' if ok and not appeared else '失败 ❌'} ===")
    print(f"监控采样 {len(watcher.sightings[:canary_mark])} 次，"
          f"业务过程中出现过的 WINWORD: {appeared or '无'}")
    left = pids_named("WINWORD.EXE") - baseline
    if left:
        print(f"⚠ 收尾后仍有 WINWORD 残留: {left}")
    return 0 if ok and not appeared else 1


if __name__ == "__main__":
    sys.exit(main())
