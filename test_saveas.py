"""验证 save-as 行为：原文件不被破坏，输出副本含水印；默认输出到桌面；自动防覆盖。"""
import sys, os, shutil
sys.path.insert(0, os.path.dirname(__file__))
from docx import Document
from watermark_tool import core, engine_docx

PRJ = os.path.dirname(__file__)
SRC = os.path.join(PRJ, "saveas_src.docx")
OUT = os.path.join(PRJ, "saveas_out.docx")

# 1) 构造一个无水印的源文件
d = Document()
d.add_paragraph("这是原始正文，不应被任何操作改动。")
d.save(SRC)
src_size = os.path.getsize(SRC)
print("原文件大小:", src_size, "原文件含水印?", engine_docx.has_watermark(SRC))

# 2) 默认输出路径应落在桌面（本机桌面存在时）
default = core.default_output_path(SRC)
print("默认输出路径:", default)
print("默认落在桌面?", os.path.dirname(default).lower().endswith("desktop"))

# 3) 插入水印（save-as 到 OUT）
res = core.insert_watermark(SRC, "text", output_path=OUT,
                            text="我是水印", font_size=120,
                            color=(128, 128, 128), angle=45,
                            transparency=0.5, scale=1.0)
print("插入结果:", res)
print("原文件大小(应不变):", os.path.getsize(SRC), "==", src_size, "->", os.path.getsize(SRC) == src_size)
print("原文件仍无水印?", not engine_docx.has_watermark(SRC))
print("输出文件含水印?", engine_docx.has_watermark(OUT))

# 4) 覆盖保护：output 与 src 相同则应自动加后缀
protected = core._protect_original(SRC, SRC)
print("防覆盖自动改名:", protected, "!= 源?", os.path.abspath(protected) != os.path.abspath(SRC))

# 5) 清除水印（save-as 到另一副本）
OUT2 = os.path.join(PRJ, "saveas_cleared.docx")
res2 = core.clear_watermark(OUT, output_path=OUT2)
print("清除结果:", res2)
print("清除副本无水印?", not engine_docx.has_watermark(OUT2))
print("原输出仍含水印?", engine_docx.has_watermark(OUT))

# 清理
for p in (SRC, OUT, OUT2):
    if os.path.exists(p):
        os.remove(p)
print("\nALL CHECKS DONE")
