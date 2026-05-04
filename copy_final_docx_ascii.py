from pathlib import Path
import shutil
base=Path(r"C:\Users\31606\Desktop")
matches=list(base.rglob("*最终修订版*.docx"))
if not matches:
    raise SystemExit("NO_FINAL_DOCX_FOUND")
src=matches[0]
dst=base / "FINAL_THESIS_REVISION.docx"
shutil.copy2(src, dst)
print("SRC=", src)
print("DST=", dst)
print("DST_EXISTS=", dst.exists())
print("DST_SIZE=", dst.stat().st_size if dst.exists() else 0)
