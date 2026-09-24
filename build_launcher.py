"""构建最终交付的单文件 WordWatermark.exe（启动器版）。

流程：
    1) 把 dist/WriteMarkApp（onedir 运行体）压缩为 payload/app_payload.zip
    2) 用 PyInstaller 把 launcher.py + 该 zip 打成单个 exe：dist/WordWatermark.exe

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

HERE = Path(__file__).resolve().parent
APP_DIR = HERE / "dist" / "WriteMarkApp"
PAYLOAD_DIR = HERE / "payload"
PAYLOAD_ZIP = PAYLOAD_DIR / "app_payload.zip"


def build_payload():
    if not APP_DIR.is_dir():
        sys.exit(f"找不到 onedir 运行体：{APP_DIR}\n请先执行："
                 f"python -m PyInstaller WordWatermark_app.spec --noconfirm")
    PAYLOAD_DIR.mkdir(exist_ok=True)
    if PAYLOAD_ZIP.exists():
        PAYLOAD_ZIP.unlink()
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


PAYLOAD_MARKER = b"\n#WRITEMARK-PAYLOAD#\n"


def _read_tail_marker(path):
    with open(path, "rb") as f:
        f.seek(-len(PAYLOAD_MARKER), os.SEEK_END)
        return f.read() == PAYLOAD_MARKER


def build_launcher():
    """先把 launcher.py 打成极小的 onefile exe，再把 zip 追加到它尾部。

    追加后 exe 依然是可执行的 Windows PE，同时自身又是一个合法 zip
    （Python zipfile 支持前置数据）。这样启动器运行时不需要把几十 MB 的
    压缩包展开到临时目录，退出时也就没有任何大目录要删除。
    """
    t0 = time.time()
    r = subprocess.run([sys.executable, "-m", "PyInstaller", "launcher.spec",
                        "--noconfirm"], cwd=str(HERE))
    if r.returncode != 0:
        sys.exit("PyInstaller 构建失败")

    exe = HERE / "dist" / "WordWatermark.exe"
    if not exe.exists():
        sys.exit(f"未生成 {exe}")
    if _read_tail_marker(exe):
        sys.exit("exe 尾部已存在 payload，请重新构建一个干净的启动器 "
                 "（PyInstaller launcher.spec --noconfirm --clean）")

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
        need = "WriteMarkApp.exe" in " | ".join(names[:20]) or any(
            n.endswith("WriteMarkApp.exe") for n in names)
        print(f"自检：zipfile 可直接读取该 exe，条目 {len(names)} 个，含主程序 -> {need}")
        if not need:
            sys.exit("payload 结构异常：找不到 WriteMarkApp.exe")
    except Exception as e:
        sys.exit(f"payload 自检失败：{e}")


if __name__ == "__main__":
    build_payload()
    build_launcher()
