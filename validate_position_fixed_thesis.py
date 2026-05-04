from pathlib import Path
from docx import Document
from zipfile import ZipFile
p=Path(r"C:\Users\31606\Desktop\系统界面\校园垃圾自动识别与分类系统设计副本_ResidualEMA位置修正版.docx")
d=Document(str(p))
print('exists', p.exists(), 'size', p.stat().st_size)
with ZipFile(p) as z:
    print('zip_bad', z.testzip(), 'media', len([n for n in z.namelist() if n.startswith('word/media/')]))
for i,para in enumerate(d.paragraphs):
    t=para.text.strip()
    if t.startswith('第一章') or t.startswith('第四章') or t.startswith('第五章') or t.startswith('4.') or t.startswith('5.') or t.startswith('图4.') or t.startswith('图5.') or t.startswith('表4.') or t.startswith('表5.'):
        print(i, t)
