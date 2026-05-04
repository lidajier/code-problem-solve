from pathlib import Path
base=Path(r"C:\Users\31606\Desktop")
for p in base.rglob("*最终修订版*.docx"):
    print(str(p))
    print("size", p.stat().st_size, "mtime", p.stat().st_mtime)
print("--- system dir docx ---")
for p in (base/"系统界面").glob("*.docx"):
    print(p.name, p.stat().st_size)
