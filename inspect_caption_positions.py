from pathlib import Path
from docx import Document
p=Path(r"C:\Users\31606\Desktop\系统界面\校园垃圾自动识别与分类系统设计副本_ResidualEMA修订版.docx")
d=Document(str(p))
for i,para in enumerate(d.paragraphs):
    t=para.text.strip()
    if t.startswith('第') or t.startswith('图4.') or t.startswith('图5.') or t.startswith('表4.') or t.startswith('表5.') or '小结' in t:
        print(i, repr(t[:80]))
