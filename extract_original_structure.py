from pathlib import Path
from docx import Document
p=Path(r"C:\Users\31606\Desktop\系统界面\校园垃圾自动识别与分类系统设计副本.docx")
d=Document(str(p))
out=Path(r"C:\Users\31606\Desktop\original_thesis_structure.txt")
with out.open('w', encoding='utf-8') as f:
    for i, para in enumerate(d.paragraphs):
        t=para.text.strip()
        if t and (t.startswith('第') or t.startswith('1.') or t.startswith('2.') or t.startswith('3.') or t.startswith('4.') or t.startswith('5.') or t.startswith('6.') or t.startswith('图') or t.startswith('表') or '实验结果' in t or '系统测试' in t or '小结' in t):
            f.write(f"{i}\t{t}\n")
print(out)
