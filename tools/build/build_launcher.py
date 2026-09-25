"""构建最终交付的单文件 WordWatermark.exe（启动器版）。

流程：
    0) 把上一次的 onedir 目录改名挪走（关键！见 prepare_app_dir 的说明）
    1) PyInstaller 用 WordWatermark_app.spec 生成全新的 dist/WriteMarkApp
    2) 校验运行体确实比源码新（防止“静默沿用旧产物”）
    3) 压缩为 payload/app_payload.zip 并追加到启动器 exe 尾部 → dist/WordWatermark.exe

产出：用户双击一个 exe 即可；首次运行把运行体解压到 %LOCALAPPDATA%，
之后直接运行；关闭时无需清理临时目录，进程立即退出。
"""
import os
import shutil
import subprocess
import sys
import time
import zipfile
from pathlib import Path

# 本脚本位于 tools/build/，下面所有路径（dist/ payload/ build/ watermark_tool/ main.py）
# 都是相对**项目根**的，所以这里取根目录而不是脚本自身所在目录。
HERE = Path(__file__).resolve().parent.parent.parent
APP_DIR = HERE / "dist" / "WriteMarkApp"
PAYLOAD_DIR = HERE / "payload"
PAYLOAD_ZIP = PAYLOAD_DIR / "app_payload.zip"
PAYLOAD_MARKER = b"\n#WRITEMARK-PAYLOAD#\n"


def _read_tail_marker(path):
    with open(path, "rb") as f:
        f.seek(-len(PAYLOAD_MARKER), os.SEEK_END)
        return f.read() == PAYLOAD_MARKER


def prepare_app_dir():
    """把旧的 onedir 目录改名挪走，让 PyInstaller 每轮都从零创建。

    为什么必须这么做：PyInstaller 重建时会先 `Removing dir dist/WriteMarkApp`；
    如果这次删除被安全策略拦下（沙箱/杀软/权限都可能），它不会报错，而是继续
    “构建完成”，产物却**静默停留在旧版本**——本项目真实踩过一次，导致测试全绿
    但跑的还是旧程序。改名不涉及批量删除，必然成功。
    """
    if not APP_DIR.exists():
        return
    # 全程只用 os.replace（重命名），**不做任何删除**：
    # 批量删除上百个文件会被环境的安全策略拦下，那会让构建半途失败。
    prev = APP_DIR.with_name(f"{APP_DIR.name}_prev_{time.strftime('%m%d_%H%M%S')}")
    os.replace(APP_DIR, prev)
    print(f"已把上一次的运行体改名为 {prev.name}（确保本轮全新构建）")
    print("  注：这些 _prev_* 目录可随时手工删除，不影响任何功能。")


def newest_source_mtime():
    files = list((HERE / "watermark_tool").rglob("*.py")) + [HERE / "main.py"]
    return max(f.stat().st_mtime for f in files if f.exists())


def check_freshness():
    exe = APP_DIR / "WriteMarkApp.exe"
    if not exe.exists():
        sys.exit(f"构建失败：没有生成 {exe}")
    # 允许 2 秒的时钟/写盘误差
    if exe.stat().st_mtime < newest_source_mtime() - 2:
        sys.exit("构建异常：运行体比源码还旧，说明本轮没有真正重写。\n"
                 f"  exe : {time.ctime(exe.stat().st_mtime)}\n"
                 f"  源码: {time.ctime(newest_source_mtime())}\n"
                 "请检查是否有删除被安全策略拦截，或先删除 dist/WriteMarkApp 再重试。")
    print(f"运行体已更新：{time.ctime(exe.stat().st_mtime)}")


def build_payload():
    if not APP_DIR.is_dir():
        sys.exit(f"找不到 onedir 运行体：{APP_DIR}")
    check_freshness()
    PAYLOAD_DIR.mkdir(exist_ok=True)
    # 注意：不要在这里 unlink 旧的 app_payload.zip。下文以 "w" 打开即等于覆盖写，
    # 而"删除一个 50MB+ 文件"会被环境的批量删除保护拦下，导致构建半途失败。
    # zip 内部直接以 app 目录内容为根（不要多包一层 WriteMarkApp/），
    # 这样 runtime_dir() 本身就是应用目录，APP_EXE_NAME 就在它下面
    files = sorted(p for p in APP_DIR.rglob("*") if p.is_file())
    print(f"压缩 {len(files)} 个文件 → {PAYLOAD_ZIP.name} …")
    with zipfile.ZipFile(PAYLOAD_ZIP, "w", zipfile.ZIP_DEFLATED, compresslevel=6) as z:
        for i, p in enumerate(files, 1):
            z.write(p, p.relative_to(APP_DIR).as_posix())
            if i % 200 == 0:
                print(f"  {i}/{len(files)}")
    size_mb = PAYLOAD_ZIP.stat().st_size / 1048576
    print(f"payload 完成：{size_mb:.1f} MB")


def build_launcher():
    """先把 launcher.py 打成极小的 onefile exe，再把 zip 追加到它尾部。

    追加后 exe 依然是可执行的 Windows PE，同时自身又是一个合法 zip
    （Python zipfile 支持前置数据）。这样启动器运行时不需要把几十 MB 的
    压缩包展开到临时目录，退出时也就没有任何大目录要删除。
    """
    t0 = time.time()
    launcher_exe = HERE / "dist" / "WordWatermark.exe"
    if launcher_exe.exists() and _read_tail_marker(launcher_exe):
        # 上一轮产物尾部已挂 payload：改个名挪走（PyInstaller 之后会重建同名 exe）。
        # 不能在这里 os.remove —— 大文件的删除会被批量删除保护拦下。
        old = launcher_exe.with_name(f"WordWatermark_prev_{time.strftime('%m%d_%H%M%S')}.exe")
        os.replace(launcher_exe, old)
        print(f"已移走上一版 exe：{old.name}")
    # 同样用全新的工作目录，避免 PyInstaller 去批量删除 build/launcher 里的旧缓存
    work = HERE / "build" / ("launcher_" + time.strftime("%Y%m%d_%H%M%S"))
    r = subprocess.run([sys.executable, "-m", "PyInstaller", "launcher.spec",
                        "--noconfirm", "--workpath", str(work)], cwd=str(HERE))
    if r.returncode != 0:
        sys.exit("PyInstaller 构建失败")

    exe = launcher_exe
    if not exe.exists():
        sys.exit(f"未生成 {exe}")
    if _read_tail_marker(exe):
        sys.exit("exe 尾部已存在 payload，请重新构建一个干净的启动器")

    with open(exe, "ab") as fexe, open(PAYLOAD_ZIP, "rb") as fzip:
        exe_size = exe.stat().st_size
        shutil.copyfileobj(fzip, fexe, length=1024 * 1024)
        fexe.write(PAYLOAD_MARKER)
    print(f"已附加 payload：{exe_size/1048576:.1f} MB exe + "
          f"{PAYLOAD_ZIP.stat().st_size/1048576:.1f} MB zip → "
          f"{exe.stat().st_size/1048576:.1f} MB")
    print(f"完成：{exe}  （用时 {time.time()-t0:.0f}s）")

    # 自检：确认成品既是可执行的 exe，又是能被 zipfile 读穿的压缩包
    try:
        with zipfile.ZipFile(str(exe)) as z:
            names = z.namelist()
        need = any(n.endswith("WriteMarkApp.exe") for n in names)
        print(f"自检：zipfile 可直接读取该 exe，条目 {len(names)} 个，含主程序 -> {need}")
        if not need:
            sys.exit("payload 结构异常：找不到 WriteMarkApp.exe")
    except Exception as e:
        sys.exit(f"payload 自检失败：{e}")


def main():
    if "--skip-app" not in sys.argv:
        prepare_app_dir()
        # 用**带时间戳的新工作目录**，而不是清理旧缓存（更不用 PyInstaller 的 --clean）：
        # 两者都会批量删除 build/ 下数百个文件，会被环境安全策略拦下导致构建失败。
        # 这样每次都是从零分析，效果等价，且完全不触发删除。
        work = HERE / "build" / ("app_" + time.strftime("%Y%m%d_%H%M%S"))
        print(f"PyInstaller 工作目录：build/{work.name}")
        r = subprocess.run([sys.executable, "-m", "PyInstaller",
                            "WordWatermark_app.spec", "--noconfirm",
                            "--workpath", str(work)], cwd=str(HERE))
        if r.returncode != 0:
            sys.exit("运行体构建失败")
    build_payload()
    build_launcher()


if __name__ == "__main__":
    main()
